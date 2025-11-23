import asyncio
import pandas as pd
from aiogram import Bot, Dispatcher, types
from aiogram.filters import Command

from config import TOKEN

# =======================================
# НОРМАЛИЗАЦИЯ
# =======================================
def normalize(s: str) -> str:
    if not isinstance(s, str):
        s = str(s)

    s = s.strip().upper()

    rus_to_eng = {
        "А": "A", "В": "B", "С": "C", "Е": "E", "Н": "H", "К": "K",
        "М": "M", "О": "O", "Р": "P", "Т": "T", "Х": "X", "У": "Y"
    }

    for r, e in rus_to_eng.items():
        s = s.replace(r, e)

    s = (s.replace("\u200b", "")
           .replace("\xa0", "")
           .replace(" ", "")
           .strip())

    return s


# =======================================
# ЗАГРУЗКА EXCEL
# =======================================
def load_excel_numbers(path):
    try:
        df = pd.read_excel(path, header=None)
    except Exception as e:
        print(f"❌ Ошибка чтения {path}: {e}")
        return set()

    numbers = set()

    for item in df[0]:
        if isinstance(item, str) and "(" in item:
            item = item.split("(")[0].strip()

        n = normalize(item)
        if n:
            numbers.add(n)

    print(f"✔ {path}: {len(numbers)} записей")
    return numbers


print("\n==== ЗАГРУЗКА БАЗЫ ====")

files = ["data/list.xlsx", "data/list2.xlsx", "data/list3.xlsx"]
print(f"Файлы: {files}")

blocked_data = set()
blocked_data |= load_excel_numbers("data/list.xlsx")
blocked_data |= load_excel_numbers("data/list2.xlsx")

active_data = load_excel_numbers("data/list3.xlsx")

print(f"Всего заблокировано: {len(blocked_data)}")
print(f"Всего не заблокировано: {len(active_data)}")
print("========================\n")

# =======================================
# AIOGRAM
# =======================================
bot = Bot(token=TOKEN)
dp = Dispatcher()


# =======================================
# /start
# =======================================
@dp.message(Command("start"))
async def cmd_start(message: types.Message):
    text = (
        "Приветствую!\n\n"
        "Укажите серийный номер KIT, я проверю его статус.\n\n"
        "Как написать номер:\n"
        "• латиница\n"
        "• без пробелов\n"
        "• без лишних символов\n"
        "Пример: KIT400122233 или 4PBA00745400\n\n"
        "🔴 Заблокирована — номер найден в базе блокировок.\n"
        "🟢 Не заблокирована — номер найден в списке активных.\n"
        "⚪ Отсутствует в базе — номер не найден или написан с ошибкой.\n"
    )
    await message.answer(text)


# =======================================
# ЛОГИКА ПРОВЕРКИ
# =======================================
@dp.message()
async def check(message: types.Message):
    raw = message.text.strip()
    number = normalize(raw)

    if not number:
        await message.answer("⚪ Статус: Неверный формат номера")
        return

    if number in blocked_data:
        await message.answer("🔴 Статус: Тарелка ЗАБЛОКИРОВАНА")
    elif number in active_data:
        await message.answer("🟢 Статус: Не заблокирована")
    else:
        await message.answer("⚪ Статус: Отсутствует в базе")


# =======================================
# RUN
# =======================================
async def main():
    print("Бот запущен...\n")
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())