"""
PDF to Excel Converter - Desktop Launcher
Starts the Flask server and opens the browser automatically.
"""
import os
import sys
import traceback

# Add the application directory to path
if getattr(sys, 'frozen', False):
    # Running as compiled
    APP_DIR = os.path.dirname(sys.executable)
else:
    # Running as script
    APP_DIR = os.path.dirname(os.path.abspath(__file__))

os.chdir(APP_DIR)
sys.path.insert(0, APP_DIR)

# Set environment variables
os.environ['FLASK_DEBUG'] = 'false'

# Error logging for debugging
def log_error(e):
    error_file = os.path.join(APP_DIR, 'error.log')
    with open(error_file, 'w') as f:
        f.write(f"Error: {str(e)}\n\n")
        f.write(traceback.format_exc())
    print(f"Error logged to: {error_file}")

try:
    import webbrowser
    import threading
    import time
    import socket
except Exception as e:
    log_error(e)
    input("Press Enter to exit...")
    sys.exit(1)

def find_free_port(start_port=5000):
    """Find a free port starting from start_port."""
    port = start_port
    while port < start_port + 100:
        try:
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.bind(('127.0.0.1', port))
            sock.close()
            return port
        except OSError:
            port += 1
    return start_port


def open_browser(port):
    """Open browser after a short delay."""
    time.sleep(1.5)
    webbrowser.open(f'http://127.0.0.1:{port}')


def main():
    # Find available port
    port = find_free_port(5000)

    print("=" * 50)
    print("  PDF to Excel Converter")
    print("=" * 50)
    print(f"\n  Server starting on http://127.0.0.1:{port}")
    print("  Press Ctrl+C to stop the server\n")
    print("=" * 50)

    # Open browser in background thread
    browser_thread = threading.Thread(target=open_browser, args=(port,))
    browser_thread.daemon = True
    browser_thread.start()

    # Import and run Flask app
    from pdftoexcel import app, start_heartbeat_monitor

    # Create required directories
    os.makedirs('uploads', exist_ok=True)
    os.makedirs('logs', exist_ok=True)

    # Start heartbeat monitor for auto-shutdown when browser closes
    start_heartbeat_monitor()
    print("  Auto-shutdown enabled (3 min after browser closes)\n")

    # Run the server
    app.run(host='127.0.0.1', port=port, debug=False, threaded=True)


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        log_error(e)
        print(f"\nError: {e}")
        print("\nCheck error.log for details.")
        input("\nPress Enter to exit...")
        sys.exit(1)
