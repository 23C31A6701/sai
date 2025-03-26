from flask import Flask, jsonify, render_template
import time
import pyttsx3
import win32gui
import win32con
import win32api

app = Flask(__name__)
engine = pyttsx3.init()

def get_latest_notification():
    window = win32gui.GetForegroundWindow()
    length = win32gui.GetWindowTextLength(window)
    title = win32gui.GetWindowText(window)
    if title and "Notification" in title:
        return title
    return "No new notifications."

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/get-notifications', methods=['GET'])
def get_notifications():
    notification = get_latest_notification()
    engine.say(notification)
    engine.runAndWait()
    return jsonify({"message": notification})

if __name__ == '__main__':
    app.run(debug=True, port=5000)
