
import os
import re
import asyncio
from pathlib import Path

import openpyxl
from dotenv import load_dotenv
from aiogram import Bot, Dispatcher
from aiogram.filters import CommandStart, Command
from aiogram.types import Message, FSInputFile, ReplyKeyboardMarkup, KeyboardButton
from openai import AsyncOpenAI

from smeta import generate_smeta


# ==========================================
# 1. SOZLAMALAR
# ==========================================

load_dotenv()

BOT_TOKEN = os.getenv("BOT_TOKEN")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
OWNER_ID = os.getenv("OWNER_ID", "").strip()

CATALOG_PATH = Path("data/catalog.xlsx")

if not BOT_TOKEN or not OPENAI_API_KEY:
    raise ValueError("API kalitlari topilmadi!")

if not OWNER_ID.isdigit():
    raise ValueError("OWNER_ID noto'g'ri!")

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()
ai = AsyncOpenAI(api_key=OPENAI_API_KEY)

history = {}
catalog = []

# Foydalanuvchi smeta ma'lumotlarini yuborishini kutayotgan holat
smeta_waiting = set()

# Telegram pastki menyusidagi doimiy tugma
MAIN_KEYBOARD = ReplyKeyboardMarkup(
    keyboard=[[KeyboardButton(text="📋 Smeta tayyorlash")]],
    resize_keyboard=True,
    one_time_keyboard=False,
    is_persistent=True,
    input_field_placeholder="Buyruq yoki savolingizni yozing..."
)


# ==========================================
# 2. AI SOZLAMALARI
# ==========================================

SYSTEM_PROMPT = """
Sen Hikvision Uzbekistan MCHJ kompaniyasining
professional AI yordamchisisan.

Vazifalaring:
1. Hikvision mahsulotlari haqida maslahat.
2. Kamera va NVR modellarini taqqoslash.
3. Texnik xususiyatlarni tushuntirish.
4. Excel katalogdan narx topish.
5. Videokuzatuv tizimlarini loyihalash.
6. O'zbek va rus tillarida tijorat matnlari yozish.

Narxlarni o'zingdan to'qib chiqarma.
Faqat katalogdagi narxlarga tayan.
Narx topilmasa, buni ochiq ayt.
Foydalanuvchi qaysi tilda yozsa, shu tilda javob ber.
"""


# ==========================================
# 3. KATALOGNI YUKLASH
# ==========================================

def load_catalog():

    global catalog

    if not CATALOG_PATH.exists():
        raise FileNotFoundError(
            f"Katalog topilmadi: {CATALOG_PATH}"
        )

    wb = openpyxl.load_workbook(
        CATALOG_PATH,
        read_only=True,
        data_only=True
    )

    ws = wb["Products"]

    headers = next(ws.iter_rows(values_only=True))

    catalog = []

    for row in ws.iter_rows(min_row=2, values_only=True):

        item = dict(zip(headers, row))

        if item.get("Model / name"):
            catalog.append(item)

    wb.close()

    return len(catalog)


# ==========================================
# 4. MAHSULOT QIDIRISH
# ==========================================

def words(text):

    return re.findall(
        r"[a-zA-Z0-9а-яА-ЯёЁ]+",
        str(text).lower()
    )


def normalize(text):

    return re.sub(
        r"[^a-z0-9]",
        "",
        str(text).lower()
    )


def search_products(query, limit=10):

    query_words = set(words(query))

    stop_words = {
        "qancha", "narxi", "narx", "kerak",
        "top", "ber", "chiqar", "korsat",
        "menga", "uchun", "ta", "dona",
        "model", "haqida", "malumot"
    }

    query_words -= stop_words

    qnorm = normalize(query)

    if not query_words:
        return []

    results = []

    for item in catalog:

        name = str(item.get("Model / name") or "")
        desc = str(item.get("Description") or "")
        category = str(item.get("Category") or "")

        nname = normalize(name)

        score = (
            len(query_words & set(words(name))) * 5
            + len(query_words & set(words(desc)))
            + len(query_words & set(words(category)))
        )

        if len(qnorm) >= 6 and qnorm in nname:
            score += 30

        if score >= 2:
            results.append((score, item))

    results.sort(key=lambda x: x[0], reverse=True)

    return [
        item for score, item in results[:limit]
    ]


def format_products(items):

    output = []

    for item in items:

        text = (
            f"Model: {item.get('Model / name')}\n"
            f"Kategoriya: {item.get('Category')}"
        )

        if item.get("Price USD") is not None:
            text += f"\nNarxi: ${item['Price USD']} USD"

        if item.get("Description"):
            text += f"\nTavsif: {item['Description']}"

        output.append(text)

    return "\n\n".join(output)


# ==========================================
# 5. ADMINISTRATOR
# ==========================================

def is_owner(message):

    return (
        message.from_user is not None
        and str(message.from_user.id) == OWNER_ID
    )


# ==========================================
# 6. START
# ==========================================

@dp.message(Command("menu"))
async def show_menu(message: Message):
    if not is_owner(message):
        return
    await message.answer(
        "Asosiy menyu. Smeta tayyorlash uchun pastdagi tugmani bosing.",
        reply_markup=MAIN_KEYBOARD
    )


@dp.message(CommandStart())
async def start(message: Message):

    if not is_owner(message):

        await message.answer(
            "Hikvision Uzbekistan AI yordamchisiga xush kelibsiz!"
        )
        return

    await message.answer(
        "Assalomu alaykum, Davron!\n\n"
        "Hikvision AI yordamchingiz tayyor.\n\n"
        "📦 Mahsulot qidirish\n"
        "💰 Narxlarni aniqlash\n"
        "📋 Smeta tayyorlash - maxsus tugma\n"
        "📊 /catalog - Katalog ma'lumoti\n"
        "🔄 /reload - Katalogni yangilash\n\n"
        "Savolingizni yuboring!",
        reply_markup=MAIN_KEYBOARD
    )


# ==========================================
# 7. KATALOG BUYRUQLARI
# ==========================================

@dp.message(Command("catalog"))
async def catalog_info(message: Message):

    if not is_owner(message):
        return

    await message.answer(
        f"📦 Katalog faol.\n"
        f"Jami mahsulotlar: {len(catalog)} ta",
        reply_markup=MAIN_KEYBOARD
    )


@dp.message(Command("reload"))
async def reload_catalog(message: Message):

    if not is_owner(message):
        return

    try:

        count = load_catalog()

        await message.answer(
            f"✅ Katalog yangilandi!\n"
            f"Mahsulotlar: {count} ta",
            reply_markup=MAIN_KEYBOARD
        )

    except Exception as e:

        print("KATALOG XATOLIGI:", repr(e))

        await message.answer(
            "Katalogni yangilashda xatolik."
        )


# ==========================================
# 8. AVTOMATIK SMETA

def parse_smeta_request(raw_request):
    """Mahsulot qatorlari va mijoz ma'lumotlarini ajratadi."""
    client_info = {}
    product_lines = []

    field_patterns = {
        "client": r"^(?:mijoz|klient|заказчик|клиент)\s*:\s*(.*)$",
        "phone": r"^(?:telefon|телефон|phone)\s*:\s*(.*)$",
        "object": r"^(?:obyekt|объект|manzil|адрес)\s*:\s*(.*)$",
    }

    for line in raw_request.splitlines():
        line = line.strip()
        if not line:
            continue

        found_field = False
        for field, pattern in field_patterns.items():
            match = re.match(pattern, line, flags=re.IGNORECASE)
            if match:
                client_info[field] = match.group(1).strip()
                found_field = True
                break

        if not found_field:
            product_lines.append(line)

    return "\n".join(product_lines).strip(), client_info


async def generate_and_send_smeta(message: Message, raw_request: str):
    """Smeta yaratadi va PDF hamda Excel fayllarini yuboradi."""
    request, client_info = parse_smeta_request(raw_request)

    if not request:
        await message.answer(
            "SMETA TAYYORLASH\n\n"
            "Mahsulotlarni quyidagi tartibda yuboring:\n\n"
            "NVR-108H-D 1ta\n"
            "DS-2CD1043G2-LIU 2ta\n"
            "IPC-T221HE-UC 2ta\n"
            "Mijoz: FerSystem MCHJ\n"
            "Telefon: +998911154515\n"
            "Obyekt: Farg'ona shahar",
            reply_markup=MAIN_KEYBOARD
        )
        return

    await message.answer("Mahsulotlar katalogdan qidirilmoqda...")

    try:
        pdf_file, excel_file, items, not_found = generate_smeta(
            request, catalog, client_info=client_info
        )

        if not items:
            await message.answer(
                "Mahsulotlar katalogdan topilmadi.\nModel nomlarini tekshiring.",
                reply_markup=MAIN_KEYBOARD
            )
            return

        await message.answer_document(
            FSInputFile(str(pdf_file)),
            caption="PDF tijorat smetasi"
        )
        await message.answer_document(
            FSInputFile(str(excel_file)),
            caption="Excel smeta",
            reply_markup=MAIN_KEYBOARD
        )

        if not_found:
            await message.answer(
                "Quyidagi mahsulotlar katalogdan topilmadi:\n\n"
                + "\n".join("- " + x for x in not_found)
                + "\n\nUlarni alohida tekshirish kerak.",
                reply_markup=MAIN_KEYBOARD
            )

    except Exception as e:
        print("SMETA XATOLIGI:", repr(e))
        await message.answer(
            "Smeta tayyorlashda xatolik yuz berdi.\nCMD oynasidagi xabarni tekshiring.",
            reply_markup=MAIN_KEYBOARD
        )


@dp.message(Command("smeta"))
async def smeta_command(message: Message):
    if not is_owner(message):
        return

    parts = (message.text or "").split(maxsplit=1)
    raw_request = parts[1].strip() if len(parts) > 1 else ""

    if not raw_request:
        smeta_waiting.add(message.from_user.id)
        await message.answer(
            "📋 SMETA TAYYORLASH\n\n"
            "Mahsulotlar ro'yxatini oddiy matn ko'rinishida yuboring. "
            "/smeta yozish shart emas.\n\n"
            "Masalan:\n"
            "NVR-108H-D 1ta\n"
            "DS-2CD1043G2-LIU 2ta\n"
            "IPC-T221HE-UC 2ta\n"
            "HDD 1TB WD Purple 1ta\n"
            "POE 8P+2UP Wi-Tek 1ta\n"
            "Mijoz: FerSystem MCHJ\n"
            "Telefon: +998911154515\n"
            "Obyekt: Farg'ona shahar\n\n"
            "Bekor qilish uchun /cancel yuboring.",
            reply_markup=MAIN_KEYBOARD
        )
        return

    await generate_and_send_smeta(message, raw_request)


@dp.message(lambda message: message.text == "📋 Smeta tayyorlash")
async def smeta_button(message: Message):
    if not is_owner(message):
        return

    smeta_waiting.add(message.from_user.id)
    await message.answer(
        "📋 SMETA TAYYORLASH\n\n"
        "Endi mahsulotlarni oddiy matn ko'rinishida yuboring. "
        "/smeta yozishingiz shart emas.\n\n"
        "Masalan:\n"
        "NVR-108H-D 1ta\n"
        "DS-2CD1043G2-LIU 2ta\n"
        "IPC-T221HE-UC 2ta\n"
        "HDD 1TB WD Purple 1ta\n"
        "POE 8P+2UP Wi-Tek 1ta\n"
        "Mijoz: FerSystem MCHJ\n"
        "Telefon: +998911154515\n"
        "Obyekt: Farg'ona shahar\n\n"
        "Bekor qilish uchun /cancel yuboring.",
        reply_markup=MAIN_KEYBOARD
    )


@dp.message(Command("cancel"))
async def cancel_smeta(message: Message):
    if not is_owner(message):
        return

    smeta_waiting.discard(message.from_user.id)
    await message.answer(
        "Smeta tayyorlash bekor qilindi.",
        reply_markup=MAIN_KEYBOARD
    )


@dp.message(
    lambda message: (
        message.from_user is not None
        and message.from_user.id in smeta_waiting
        and message.text is not None
        and not message.text.startswith("/")
    )
)
async def smeta_text_input(message: Message):
    if not is_owner(message):
        return

    smeta_waiting.discard(message.from_user.id)
    await generate_and_send_smeta(message, message.text)


# 9. AI SUHBAT
# ==========================================

@dp.message()
async def ai_chat(message: Message):

    if not is_owner(message):
        return

    if not message.text:
        await message.answer("Matnli xabar yuboring.", reply_markup=MAIN_KEYBOARD)
        return

    user_id = message.from_user.id

    user_history = history.get(user_id, [])

    await bot.send_chat_action(
        message.chat.id,
        action="typing"
    )

    try:

        found = search_products(message.text)

        context = ""

        if found:
            context = (
                "\n\nKATALOGDAGI MAHSULOTLAR:\n"
                + format_products(found)
            )

        response = await ai.responses.create(
            model="gpt-5.5",
            instructions=SYSTEM_PROMPT + context,
            input=user_history + [
                {
                    "role": "user",
                    "content": message.text
                }
            ],
            max_output_tokens=1800
        )

        answer = response.output_text or "Javob yaratilmadi."

        answer = re.sub(
            r"\*\*|__|`",
            "",
            answer
        )

        history[user_id] = (
            user_history
            + [
                {"role": "user", "content": message.text},
                {"role": "assistant", "content": answer}
            ]
        )[-12:]

        for i in range(0, len(answer), 4000):

            await message.answer(
                answer[i:i + 4000],
                reply_markup=MAIN_KEYBOARD
            )

    except Exception as e:

        print("AI XATOLIGI:", repr(e))

        await message.answer(
            "AI xizmatida xatolik yuz berdi.",
            reply_markup=MAIN_KEYBOARD
        )


# ==========================================
# 10. BOTNI ISHGA TUSHIRISH
# ==========================================

async def main():

    count = load_catalog()

    print("--------------------------------")
    print("HIKVISION AI ASSISTANT ISHLAYAPTI")
    print(f"Katalog: {count} ta mahsulot")
    print("Smeta generatori: FAOL")
    print("Smeta tugmasi: FAOL")
    print("--------------------------------")

    try:
        await bot.send_message(
            int(OWNER_ID),
            "✅ Hikvision AI ishga tushdi. Smeta tugmasi faollashtirildi.",
            reply_markup=MAIN_KEYBOARD
        )
    except Exception as e:
        print("MENYU TUGMASINI YUBORISHDA OGOHLANTIRISH:", repr(e))

    await dp.start_polling(bot)


if __name__ == "__main__":

    asyncio.run(main())
