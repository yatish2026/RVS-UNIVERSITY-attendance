import sys
import os

# Insert root folder into sys.path so server and modules can be imported smoothly
root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if root_dir not in sys.path:
    sys.path.insert(0, root_dir)

from server import app

# Vercel Serverless WSGI entrypoint
