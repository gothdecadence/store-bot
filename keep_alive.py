from flask import Flask
from threading import Thread
import os
import asyncio
import bot  # импортируем бота

app = Flask(__name__)

@app.route('/')
def home():
    return "🤖 Бот работает!"

def run_bot():
    """Запускает бота в отдельном потоке"""
    try:
        asyncio.run(bot.main())  # bot.main() — главная функция в bot.py
    except Exception as e:
        print(f"Ошибка бота: {e}")

def run_flask():
    """Запускает Flask в отдельном потоке"""
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port)

# Запускаем бота в отдельном потоке
Thread(target=run_bot, daemon=True).start()

# Запускаем Flask в основном потоке
if __name__ == "__main__":
    run_flask()
