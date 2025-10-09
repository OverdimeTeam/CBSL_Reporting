from waitress import serve
from app import app, ensure_directories_exist, kill_excel_instances

if __name__ == "__main__":
    kill_excel_instances()
    ensure_directories_exist()
    print("Starting Waitress server on http://0.0.0.0:5000")
    serve(app, host="0.0.0.0", port=5000)