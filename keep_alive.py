from flask import Flask
from threading import Thread
import requests
import time
import os

app = Flask(__name__)

@app.route('/')
def home():
    return "?? Бот работает 24/7!"

def ping():
    url = os.environ.get('RENDER_EXTERNAL_URL', 'https://ваш-сайт.onrender.com')
    while True:
        time.sleep(600)  # каждые 10 минут
        try:
            requests.get(url)
            print("? Пинг отправлен")
        except Exception as e:
            print(f"? Ошибка пинга: {e}")

def start_keep_alive():
    Thread(target=ping, daemon=True).start()
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port)