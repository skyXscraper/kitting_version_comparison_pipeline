"""Kitting version comparison pipeline - web page.

python kvc_web.py            start the page at http://127.0.0.1:8000 and open it in the browser
python kvc_web.py --port N   use another port
"""
from kvc.webapp import main

if __name__ == "__main__":
    main()
