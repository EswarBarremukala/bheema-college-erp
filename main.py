import threading
from app import app  # Import your Flask app
from flask_webview import WebView # Or use Kivy's WebView

def start_flask():
    app.run(host='127.0.0.1', port=5000)

if __name__ == "__main__":
    # Start Flask in a background thread
    threading.Thread(target=start_flask, daemon=True).start()
    
    # Start the Android WebView (requires Kivy/Python-for-Android)
    # This part varies based on the bridge you use