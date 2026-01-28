export k=`ps  -C web_gui.py  -o pid --no-headers`
source venv/bin/activate
python3 ./web_gui.py

