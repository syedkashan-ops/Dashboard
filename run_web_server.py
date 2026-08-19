import os

from waitress import serve
from web_app import app

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    serve(app, host="0.0.0.0", port=port, threads=4)
