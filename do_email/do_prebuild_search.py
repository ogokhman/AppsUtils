#!/usr/bin/env python3
"""
Build search queries per user from MySQL data and combine with folder information.

Steps:
1) Read marketer_team (comma-separated users) and domain from investropipeline.
2) Map each user to the list of investor domains where that user appears.
3) Load folder information from do_user_folders.config.
4) Output a combined config file (do_final_run.config) with user, domains, search query, and folders.

MySQL credentials are loaded from .env:
- MYSQL_HOST
- MYSQL_PORT (optional, defaults to 3306)
- MYSQL_USER
- MYSQL_PASSWORD
- MYSQL_DATABASE (optional, defaults to investorpipelinedb)
- MYSQL_TABLE (required, e.g. investropipeline)
"""

import os
import re
import configparser
from collections import defaultdict

import mysql.connector
from dotenv import load_dotenv

DOMAIN_SUFFIX = "@christoffersonrobb.com"


def normalize_domain(value: str) -> str:
    """Normalize domain strings to a host-like form."""
    domain = value.strip().lower()
    domain = re.sub(r"^https?://", "", domain)
    domain = re.sub(r"^www\.", "", domain)
    domain = domain.strip().strip("/")
    return domain


def split_users(marketer_team: str) -> list[str]:
    """Split marketer_team string into individual users."""
    if not marketer_team:
        return []
    return [u.strip() for u in marketer_team.split(",") if u.strip()]


def split_domains(domain_value: str) -> list[str]:
    """Split domain field into one or more domains (comma/semicolon/space separated)."""
    if not domain_value:
        return []
    parts = re.split(r"[;,\s]+", domain_value.strip())
    domains = [normalize_domain(p) for p in parts if p.strip()]
    return [d for d in domains if d]


def build_search_query(domains: list[str]) -> str:
    """Build KQL search query for to/from using domains."""
    if not domains:
        return ""
    # Use simplified KQL syntax: combine all to/from without nested parentheses
    search_parts = []
    for domain in domains:
        search_parts.append(f"to:{domain}")
        search_parts.append(f"from:{domain}")
    return " OR ".join(search_parts)


def main() -> None:
    load_dotenv()

    host = os.getenv("MYSQL_HOST")
    port = int(os.getenv("MYSQL_PORT", "3306"))
    user = os.getenv("MYSQL_USER")
    password = os.getenv("MYSQL_PASSWORD")
    database = os.getenv("MYSQL_DATABASE", "investorpipelinedb")
    table = os.getenv("MYSQL_TABLE")

    if not all([host, user, password, database, table]):
        print(
            "ERROR: Missing MySQL credentials in .env (MYSQL_HOST, MYSQL_USER, MYSQL_PASSWORD, MYSQL_DATABASE, MYSQL_TABLE)"
        )
        return

    conn = mysql.connector.connect(
        host=host,
        port=port,
        user=user,
        password=password,
        database=database,
    )

    query = (
        "SELECT marketer_team, domain "
        f"FROM {table} "
        "WHERE marketer_team IS NOT NULL AND marketer_team <> '' "
        "AND domain IS NOT NULL AND domain <> ''"
    )

    user_domains: dict[str, set[str]] = defaultdict(set)

    try:
        cursor = conn.cursor()
        cursor.execute(query)

        for marketer_team, domain_value in cursor.fetchall():
            users = split_users(marketer_team)
            domains = split_domains(domain_value)
            if not users or not domains:
                continue
            for u in users:
                user_domains[u].update(domains)
    finally:
        conn.close()

    if not user_domains:
        print("No user/domain mappings found.")
        return

    # Load folder information from do_user_folders.config
    user_folders = load_user_folders()

    # Create output config with parameters sections
    output_config = configparser.ConfigParser()
    
    # Add configuration sections (read from do.config)
    base_config = configparser.ConfigParser()
    base_config_path = os.path.join(os.path.dirname(__file__), "do.config")
    if os.path.exists(base_config_path):
        base_config.read(base_config_path)
    
    # Copy relevant sections to output
    if base_config.has_section("dates"):
        output_config["dates"] = dict(base_config["dates"])
    else:
        output_config["dates"] = {
            "start_date": "2025-12-01",
            "end_date": "2026-01-30"
        }
    
    if base_config.has_section("messages"):
        output_config["messages"] = dict(base_config["messages"])
    else:
        output_config["messages"] = {"top": "500"}
    
    if base_config.has_section("folders"):
        output_config["folders"] = dict(base_config["folders"])
    else:
        output_config["folders"] = {"folders": "Inbox,SentItems"}
    
    # Add API method
    output_config["api"] = {"method": "search"}

    # Add user sections
    for user_name in sorted(user_domains.keys()):
        user_email = user_name if "@" in user_name else f"{user_name}{DOMAIN_SUFFIX}"
        domains = sorted(user_domains[user_name])
        search_query = build_search_query(domains)
        folders = user_folders.get(user_email, [])

        section_name = f"user_{user_name}"
        output_config[section_name] = {
            "user": user_email,
            "domains": " OR ".join(domains),
            "search_query": search_query,
            "folders": ", ".join(folders) if folders else "",
        }

        print("=" * 60)
        print(f"User: {user_email}")
        print(f"Domains ({len(domains)}): {' OR '.join(domains)}")
        print(f"Folders ({len(folders)}): {', '.join(folders) if folders else 'None'}")
        if search_query:
            print(f"Search Query: {search_query[:100]}...")
        else:
            print("Search Query: (no domains)")

    # Save to do_final_run.config
    output_path = os.path.join(os.path.dirname(__file__), "do_final_run.config")
    with open(output_path, "w") as f:
        output_config.write(f)

    print(f"\n✓ Results saved to {output_path}")
    print(f"✓ Processed {len(user_domains)} users")


def load_user_folders() -> dict[str, list[str]]:
    """Load folder information from do_user_folders.config (which has non-unique sections)"""
    user_folders: dict[str, list[str]] = {}
    folders_config_path = os.path.join(
        os.path.dirname(__file__), "do_user_folders.config"
    )

    if not os.path.exists(folders_config_path):
        return user_folders

    # Parse manually since ConfigParser doesn't handle duplicate section names
    current_user = None
    with open(folders_config_path, "r") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue

            if line == "[user]":
                # Skip the section header
                continue
            elif line == "[folders]":
                # Skip the folders header
                continue
            elif line.startswith("user = "):
                current_user = line[7:].strip()
            elif line.startswith("folders = ") and current_user:
                folders_str = line[10:].strip()
                folders = [f.strip() for f in folders_str.split(",") if f.strip()]
                user_folders[current_user] = folders

    return user_folders


if __name__ == "__main__":
    main()
