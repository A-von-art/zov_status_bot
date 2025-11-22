import os
import asyncio
import pandas as pd
from datetime import datetime

from aiogram import Bot, Dispatcher, types
from aiogram.filters import Command, CommandStart

from config import TOKEN, ADMIN_ID


# =========================================================
# ФУНКЦИЯ НОРМАЛИЗАЦИИ СЕРИЙНИКА
# =========================================================

def normalize(text: str) -> str:
    text = text.strip().upper()
    text = text.replace(" ", "")

    # Кириллица → латиница (визуально похожие буквы)
    mapping = {
        "А": "A", "В": "B", "Е": "E", "К": "K", "М": "M",
        "Н": "H", "О": "O", "Р": "P", "С": "S", "Т": "T",
        "Х": "X", "У": "Y"
    }

    normalized = ""
    for ch in text:
        normalized += mapping.get(ch, ch)

    return normalized


# =========================================================
# ЗАГРУЗКА ВСЕХ LIST-ФАЙЛОВ
# =========================================================

def load_all_lists():
    data_dir = "data"

    blocked = set()        # list.xlsx + list2.xlsx
    not_blocked = set()    # list3.xlsx

    files = [f for f in os.listdir(data_dir) if f.endswith(".xlsx")]

    print("\n==== ЗАГРУЗКА БАЗЫ ====")
    print("Файлы:", files)

    for filename in files:
        path = os.path.join(data_dir, filename)

        try:
            df = pd.read_excel(path, header=None)

            # читаем 2 столбца
            col1 = df[0].dropna().astype(str).str.strip().str.upper().tolist()

            col2 = []
            if 1 in df.columns:
                col2 = df[1].dropna().astype(str).str.strip().str.upper().tolist()

            all_serials = col1 + col2
            name = filename.lower()

            if name in ["list.xlsx", "list2.xlsx"]:
                blocked.update(all_serials)

            elif name == "list3.xlsx":
                not_blocked.update(all_serials)

            print(f"✔ {filename}: {len(all_serials)} записей")

        except Exception as e:
            print(f"❌ Ошибка чтения {filename}: {e}")

    print("Всего заблокировано:", len(blocked))
    print("Всего не заблокировано:", len(not_blocked))
    print("========================\n")

    return blocked, not_blocked


# Загружаем при запуске
BLOCKED, NOT_BLOCKED = load_all_lists()


# =========================================================
# ИНИЦИАЛИЗАЦИЯ БОТА
# =========================================================

bot = Bot(token=TOKEN)
dp = Dispatcher()

waiting_for_file = False


# =========================================================
# /start
# =========================================================

@dp.message(CommandStart())
async def start(message: types.Message):
    text = (
        "Приветствую!\n\n"
        "Укажите серийный номер KIT, я проверю его статус.\n\n"
        "Как написать номер:\n"
        "• латиница\n"
        "• без пробелов\n"
        "• без лишних символов\n"
        "Примеры:\n"
        "KIT400122233\n"
        "4PBA00745400\n\n"
        "🔴 Заблокирована — номер найден в базе блокировок.\n"
        "🟢 Не заблокирована — номер найден в списке активных.\n"
        "⚪ Отсутствует в базе — номер не найден или написан с ошибкой."
    )
    await message.answer(text)


# =========================================================
# /update — загрузка нового Excel (только админ)
# =========================================================

@dp.message(Command("update"))
async def update_cmd(message: types.Message):
    global waiting_for_file

    if message.from_user.id != ADMIN_ID:
        await message.answer("Недостаточно прав.")
        return

    waiting_for_file = True
    await message.answer("Пришлите Excel-файл (.xlsx). Он будет добавлен в базу.")


# =========================================================
# ПРИЁМ Excel
# =========================================================

@dp.message(lambda m: m.document)
async def add_file(message: types.Message):
    global waiting_for_file, BLOCKED, NOT_BLOCKED

    if not waiting_for_file:
        return

    if message.from_user.id != ADMIN_ID:
        await message.answer("Нет прав.")
        waiting_for_file = False
        return

    filename = message.document.file_name

    if not filename.endswith(".xlsx"):
        await message.answer("Ошибка: нужен Excel-файл (.xlsx)")
        return

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    new_name = f"upload_{timestamp}.xlsx"
    save_path = f"data/{new_name}"

    file = await bot.get_file(message.document.file_id)
    await bot.download_file(file.file_path, save_path)

    await message.answer(f"Файл принят ({filename}). Обновляю базу...")

    BLOCKED, NOT_BLOCKED = load_all_lists()
    waiting_for_file = False

    await message.answer("База успешно обновлена!")


# =========================================================
# ПРОВЕРКА СЕРИЙНОГО НОМЕРА
# =========================================================

@dp.message()
async def check_serial(message: types.Message):
    serial = normalize(message.text)

    # защита от мусора
    if len(serial) < 3:
        await message.answer("Статус: Отсутствует в базе")
        return

    if serial in BLOCKED:
        await message.answer("Статус: Тарелка ЗАБЛОКИРОВАНА")
        return

    if serial in NOT_BLOCKED:
        await message.answer("Статус: Не заблокирована")
        return

    await message.answer("Статус: Отсутствует в базе")


# =========================================================
# СТАРТ БОТА
# =========================================================

async def main():
    print("Бот запущен...")
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())