# coding: utf-8
from flask import Flask
from threading import Thread
import requests
import time
import os

app = Flask(__name__)

@app.route('/')
def home():
    return "🤖 Бот работает 24/7!"

def ping():
    url = os.environ.get('RENDER_EXTERNAL_URL', 'https://store-bot-b2wd.onrender.com')
    while True:
        time.sleep(600)  # êàæäûå 10 ìèíóò
        try:
            requests.get(url)
            print("? Ïèíã îòïðàâëåí")
        except Exception as e:
            print(f"? Îøèáêà ïèíãà: {e}")

def start_keep_alive():
    Thread(target=ping, daemon=True).start()
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port)
