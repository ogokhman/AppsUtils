from flask import Flask, render_template, request, jsonify, session, redirect, url_for
import subprocess
import sys
import os
from datetime import datetime
import pytz


# using flash 
app = Flask(__name__)
app.secret_key = 'your-secret-key-change-this-in-production'

# EST timezone for display
est = pytz.timezone('US/Eastern')

@app.route('/')
def index():
    """Render the main page with the form."""
    return render_template('index.html')

@app.route('/login', methods=['POST'])
def login():
    """Validate user credentials."""
    auth_username = request.form.get('username', '').strip()
    auth_password = request.form.get('password', '').strip()
    
    valid_users = {
        "oleg": "Hello2026",
        "apapritz@christoffersonrobb.com": "Happy2026"
    }
    
    if auth_username in valid_users and valid_users[auth_username] == auth_password:
        return jsonify({'success': True})
    else:
        return jsonify({'success': False, 'error': 'Invalid login'})

@app.route('/get_emails', methods=['POST'])
def get_emails():
    """Execute the get_user_email.py script and return results."""
    user_email = request.form.get('user_email', '').strip()
    num_messages_raw = request.form.get('num_messages', '').strip()
    from_email_address = request.form.get('fromemailaddress', '').strip()
    earliest_date = request.form.get('earliestdate', '').strip()
    max_date = request.form.get('maxdate', '').strip()
    folder = request.form.get('folder', '').strip()
    sort_order = request.form.get('sort', 'latest').strip()  # Default to latest
    
    # Get credentials from form
    auth_username = request.form.get('auth_username', '').strip()
    auth_password = request.form.get('auth_password', '').strip()
    
    if not auth_username or not auth_password:
        return jsonify({
            'success': False,
            'error': 'Please enter your credentials'
        })
    
    # Validate credentials before executing script
    valid_users = {
        "oleg": "Hello2026",
        "apapritz@christoffersonrobb.com": "Happy2026"
    }
    
    if auth_username not in valid_users or valid_users[auth_username] != auth_password:
        return jsonify({
            'success': False,
            'error': 'Invalid login'
        })
    
    if not user_email:
        return jsonify({
            'success': False,
            'error': 'Please enter a user email address'
        })
  

    # Validate and convert number of messages (if specified)
    num_messages = None
    if num_messages_raw:
        try:
            num_messages = int(num_messages_raw)
            if num_messages < 0:
                num_messages = 3
            elif num_messages > 100:
                num_messages = 100
        except ValueError:
            num_messages = 3
    
    try:
        # Execute the get_user_email_search.py script
        script_path = os.path.join(os.path.dirname(__file__), 'get_user_email_search.py')
        # Build arguments using named parameters
        args = [sys.executable, script_path, f"user={user_email}", f"username={auth_username}", f"password={auth_password}"]
        
        if num_messages is not None:
            args.append(f"count={num_messages}")
        if from_email_address:
            args.append(f"from={from_email_address}")
        if earliest_date:
            args.append(f"mindate={earliest_date}")
        if max_date:
            args.append(f"maxdate={max_date}")
        if folder:
            args.append(f"folder={folder}")
        if sort_order:
            args.append(f"sort={sort_order}")

     
        print("scritp path:", script_path)
        print("args:", args)
        result = subprocess.run(
            args,
            capture_output=True,
            text=True,
            timeout=30
        )
        
        # Parse the output
        output_lines = result.stdout.split('\n')
        error_lines = result.stderr.split('\n')
        
        # Extract messages from the output
        messages = []
        table_started = False
        total_api_messages = None
        sender_counts = {}  # Dictionary to count messages by sender
        
        for line in output_lines:
            # Extract total messages from API
            if 'Total messages retrieved from API:' in line:
                try:
                    # Extract the number after the colon
                    total_api_messages = int(line.split(':')[1].strip())
                except:
                    pass
            elif 'Found' in line and 'recent messages' in line:
                table_started = True
                continue
            if not table_started:
                continue
            if not line.strip() or line.startswith('-') or 'Date/Time' in line:
                continue
            # Fixed-width parsing according to console table layout (6 columns after update):
            # "{idx:<4} {dt:<30} {from:<50} {to:<30} {folder:<15} {subject:<40}"
            try:
                idx_str = line[0:4].strip()
                date_time = line[5:35].strip()
                from_addr = line[36:86].strip()  # From address is now 50 chars
                to_addr = line[87:117].strip()   # To address is 30 chars
                folder = line[118:133].strip()   # Folder is 15 chars
                subject = line[134:].strip()     # Subject takes the rest
                
                # Count messages by sender
                if from_addr and from_addr != 'N/A':
                    sender_counts[from_addr] = sender_counts.get(from_addr, 0) + 1
                
                messages.append({
                    'index': idx_str,
                    'date_time': date_time,
                    'from_address': from_addr,
                    'to_address': to_addr,
                    'folder': folder,
                    'subject': subject
                })
            except Exception:
                continue
        
        return jsonify({
            'success': True,
            'messages': messages,
            'output': result.stdout,
            'error_output': result.stderr,
            'timestamp': datetime.now(est).strftime('%Y-%m-%d %H:%M:%S EST'),
            'num_messages': num_messages,
            'total_api_messages': total_api_messages,
            'sender_counts': sender_counts
        })
        
    except subprocess.TimeoutExpired:
        return jsonify({
            'success': False,
            'error': 'Request timed out. Please try again.'
        })
    except Exception as e:
        return jsonify({
            'success': False,
            'error': f'An error occurred: {str(e)}'
        })

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
