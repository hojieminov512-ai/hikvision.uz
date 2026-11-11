
import re
from pathlib import Path
from datetime import datetime

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


# ==========================================
# SOZLAMALAR
# ==========================================

COMPANY = "HIKVISION UZBEKISTAN MCHJ"

OUTPUT_DIR = Path("output")
OUTPUT_DIR.mkdir(exist_ok=True)


# ==========================================
# MAHSULOTLARNI ANIQLASH
# ==========================================

def normalize(text):
    """Model va nomlarni tinish belgilarisiz bir xil ko'rinishga keltiradi."""
    return re.sub(r"[^a-z0-9]", "", str(text).lower())


def words(text):
    return re.findall(r"[a-z0-9]+", str(text).lower())


BRAND_ALIASES = {
    "hikvision": ("hikvision", "hik", "хиквижн"),
    "hilook": ("hilook", "hi-look", "хайлук", "hiluk"),
    "witek": ("wi-tek", "witek", "wi tek"),
    "wd": ("wd", "western digital", "western"),
    "seagate": ("seagate", "skyhawk"),
    "tplink": ("tp-link", "tplink", "tp link"),
    "ubiquiti": ("ubiquiti", "unifi"),
    "mikrotik": ("mikrotik", "mikro tik"),
}


def detect_brand(text):
    """Foydalanuvchi yozgan brendni yoki model prefiksidan brendni aniqlaydi."""
    raw = str(text or "").lower()
    compact = normalize(raw)

    # Avval ochiq yozilgan brendni tekshiramiz.
    for brand, aliases in BRAND_ALIASES.items():
        for alias in aliases:
            if normalize(alias) and normalize(alias) in compact:
                return brand

    # Model prefiksi bo'yicha kelishilgan qoidalar.
    model_tokens = re.findall(r"[a-z0-9]+(?:-[a-z0-9]+)*", raw)
    for token in model_tokens:
        prefix = token.upper().split("-")[0]
        if prefix == "DS":
            return "hikvision"
        if prefix in {"NS", "IPC", "THC", "NVR", "HL"}:
            return "hilook"

    return None


def catalog_brand(item):
    """Katalogdagi brendni avval model nomidan, keyin kategoriya/source dan aniqlaydi."""
    name = str(item.get("Model / name") or "")
    category = str(item.get("Category") or "")
    source = str(item.get("Source") or "")
    name_low = name.lower()
    name_compact = normalize(name)
    fallback = f"{category} {source}".lower()
    fallback_compact = normalize(fallback)

    # Model nomida ochiq ko'rsatilgan brend kategoriya qiymatidan ustun.
    if "witek" in name_compact:
        return "witek"
    if "hilook" in name_compact:
        return "hilook"
    if "hikvision" in name_compact:
        return "hikvision"
    if "seagate" in name_compact or "skyhawk" in name_compact:
        return "seagate"
    if re.search(r"\bwd\b", name_low) or "wd" in name_compact:
        return "wd"
    if "tplink" in name_compact or "tp-link" in name_low:
        return "tplink"
    if "ubiquiti" in name_compact or "unifi" in name_compact:
        return "ubiquiti"
    if "mikrotik" in name_compact:
        return "mikrotik"

    # Kelishilgan model prefikslari.
    first = re.match(r"\s*([a-z]+)", name, flags=re.I)
    if first:
        prefix = first.group(1).upper()
        if prefix == "DS":
            return "hikvision"
        if prefix in {"NS", "IPC", "THC", "NVR", "HL"}:
            return "hilook"

    # Katalog kategoriyasi/source bo'yicha zaxira aniqlash.
    if "witek" in fallback_compact:
        return "witek"
    if "hilook" in fallback_compact:
        return "hilook"
    if "hikvision" in fallback_compact:
        return "hikvision"
    if "seagate" in fallback_compact or "skyhawk" in fallback_compact:
        return "seagate"
    if "western" in fallback_compact or re.search(r"\bwd\b", fallback):
        return "wd"
    if "tplink" in fallback_compact or "tp-link" in fallback:
        return "tplink"
    if "ubiquiti" in fallback_compact or "unifi" in fallback_compact:
        return "ubiquiti"
    if "mikrotik" in fallback_compact:
        return "mikrotik"

    return None


def parse_product_line(segment):
    """
    Miqdorni satr boshidan yoki oxiridan oladi:
    2x MODEL, MODEL 2x, MODEL 2 ta, MODEL x2, MODEL 2 dona.
    """
    text = str(segment).strip()
    qty = 1

    # Satr boshidagi: 2x MODEL / 2 ta MODEL / 2 dona MODEL
    m = re.match(r"^\s*(\d+)\s*(?:x|х|ta|dona|шт)?\s+(.+?)\s*$", text, re.I)
    if m:
        qty = int(m.group(1))
        text = m.group(2).strip()
    else:
        # Satr oxiridagi: MODEL 2x / MODEL x2 / MODEL 2 ta / MODEL 2 dona
        m = re.search(r"\s+(\d+)\s*(?:x|х|ta|dona|шт)?\s*$", text, re.I)
        if m:
            qty = int(m.group(1))
            text = text[:m.start()].strip()
        else:
            m = re.search(r"\s*[xх]\s*(\d+)\s*$", text, re.I)
            if m:
                qty = int(m.group(1))
                text = text[:m.start()].strip()

    # O'lchov birliklari bilan miqdor: Kabel 100 metr / Kabel 100m.
    if qty == 1:
        m = re.search(r"\s+(\d+)\s*(?:metr|meter|метр|м|m)\s*$", text, re.I)
        if m:
            qty = int(m.group(1))
            text = text[:m.start()].strip()

    return max(qty, 1), text


def _item_search_text(item):
    return " ".join([
        str(item.get("Model / name") or ""),
        str(item.get("Description") or ""),
        str(item.get("Category") or ""),
        str(item.get("Source") or ""),
    ]).lower()


def _score_candidate(query, item, requested_brand):
    name = str(item.get("Model / name") or "")
    desc = str(item.get("Description") or "")
    category = str(item.get("Category") or "")
    full_text = _item_search_text(item)

    qnorm = normalize(query)
    nnorm = normalize(name)
    qwords = set(words(query))
    nwords = set(words(name))
    dwords = set(words(desc))
    cwords = set(words(category))

    score = len(qwords & nwords) * 10
    score += len(qwords & dwords) * 3
    score += len(qwords & cwords) * 2

    if qnorm and len(qnorm) >= 5 and qnorm == nnorm:
        score += 220
    elif qnorm and len(qnorm) >= 5 and qnorm in nnorm:
        score += 100
    if nnorm and len(nnorm) >= 6 and nnorm in qnorm:
        score += 80

    # 8P+2UP kabi port yozuvlarini katalogdagi 8+2 va uplink tavsiflari bilan solishtirish.
    qcompact = normalize(query)
    if re.search(r"8p(?:oe)?2up|8p2up|8\+2", qcompact):
        if re.search(r"8\D{0,8}(?:poe|port)", full_text) and re.search(r"2\D{0,12}(?:uplink|up)", full_text):
            score += 35
        if "8+2" in nnorm:
            score += 20

    # 4+2 / 8+2 / 16+2 kabi port sonlari.
    port_match = re.search(r"\b(4|8|16|24)\s*p(?:oe)?\s*\+?\s*(1|2|4)\s*(?:up|uplink|sfp)?\b", query, re.I)
    if port_match:
        poe_count, uplink_count = port_match.groups()
        if re.search(rf"\b{poe_count}\b", full_text) and re.search(rf"\b{uplink_count}\b", full_text):
            score += 12

    # Aniq Wi-Tek POE 8P+2UP qisqartmasi uchun kelishilgan moslik.
    if requested_brand == "witek" and re.search(r"8p(?:oe)?2up|8p2up|8\+2", qcompact):
        if "wkps210" in nnorm:
            score += 100

    return score


def extract_explicit_price(segment):
    """
    Ajratadi:
    - Xizmat 4x15$ -> miqdor 4, dona narxi 15$
    - Kabel 115m x0.6$ -> miqdor 115, dona/metr narxi 0.6$
    - Mahsulot 15$ -> miqdor keyin parse_product_line orqali olinadi
    """
    text = str(segment or "").strip()

    # Miqdor x dona narxi: 4x15$, 1X60$, 115m x0.6$
    combined_patterns = [
        r"\s+(\d+)\s*(?:m|metr|meter|метр)?\s*[xх]\s*\$?\s*(\d+(?:[.,]\d+)?)\s*(?:\$|usd)\s*$",
    ]
    for pattern in combined_patterns:
        match = re.search(pattern, text, re.I)
        if match:
            qty = int(match.group(1))
            price = float(match.group(2).replace(",", "."))
            clean = text[:match.start()].strip()
            if clean and qty > 0 and price >= 0:
                return clean, price, qty

    # Oddiy narx: 15$, $15, 15 USD.
    patterns = [
        r"\s+(\d+(?:[.,]\d+)?)\s*(?:\$|usd)\s*(?:dan)?\s*$",
        r"\s+\$\s*(\d+(?:[.,]\d+)?)\s*$",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.I)
        if match:
            price = float(match.group(1).replace(",", "."))
            clean = text[:match.start()].strip()
            if clean and price >= 0:
                return clean, price, None

    return text, None, None


def manual_position_description(name):
    """Katalogda topilmagan pozitsiya uchun smetaga tushunarli tavsif beradi."""
    q = str(name or "").strip()
    low = q.lower()

    descriptions = [
        (r"ustanov|монтаж|установка|o['’]?rnat", "Услуга по монтажу и установке оборудования видеонаблюдения, включая крепление и первичное подключение."),
        (r"kabel.*(tort|yotqiz)|прокладк.*кабел|кабель.*(монтаж|проклад)", "Работы по прокладке и укладке кабельной линии для подключения оборудования системы видеонаблюдения."),
        (r"sozlash|настройк|конфигурац", "Работы по настройке оборудования, сетевых параметров и основных функций системы видеонаблюдения."),
        (r"hdd|жестк.*диск|hard.?disk", "Жесткий диск для хранения видеозаписей системы видеонаблюдения. Емкость и совместимость уточняются по выбранной модели."),
        (r"\bpoe\b|коммутатор|switch", "Сетевой коммутатор для подключения IP-камер и передачи данных; характеристики и количество PoE-портов уточняются по модели."),
        (r"\bhdmi\b", "Кабель HDMI для передачи цифрового видео- и аудиосигнала между видеорегистратором и монитором."),
        (r"шкаф|rack|стойк", "Телекоммуникационный шкаф/стойка для размещения и защиты сетевого и видеонаблюдательного оборудования."),
        (r"кронштейн|kronshteyn|креплен", "Крепежный кронштейн для надежной установки камеры видеонаблюдения."),
        (r"блок питания|адаптер питания|power supply", "Источник питания для обеспечения электропитания оборудования видеонаблюдения."),
        (r"разъем|коннектор|connector|rj.?45", "Соединительный разъем для оконцевания и подключения кабельных линий оборудования."),
        (r"доставка|доставк", "Услуга доставки оборудования и материалов на объект заказчика."),
        (r"пусконалад|ishga tushirish|ishga tushur", "Пусконаладочные работы, проверка соединений и контроль работоспособности установленной системы."),
    ]
    for pattern, description in descriptions:
        if re.search(pattern, low, re.I):
            return description
    return (
        f"Дополнительная позиция: {q}. "
        "Наименование и технические характеристики указаны согласно заявке заказчика; "
        "точная модель и комплектация уточняются отдельно."
    )


def find_smeta_items(request, catalog):
    """
    Mahsulot nomi, brend va model prefiksi bo'yicha katalogdan qidiradi.
    Bir nechta brend/variant mos tushsa, taxmin qilmaydi.
    """
    segments = re.split(r"[,;\n]+", request)
    result = []
    not_found = []

    for original_segment in segments:
        segment = original_segment.strip()
        if not segment:
            continue

        if re.match(r"^(?:mijoz|klient|заказчик|клиент|telefon|телефон|phone|obyekt|объект|manzil|адрес)\s*:", segment, re.I):
            continue
        if any(word in segment.lower() for word in ["smeta", "tayyorla", "rus tilida", "pdf", "excel", "hisobla"]):
            continue

        # Foydalanuvchi kiritgan narx katalogdan mustaqil qabul qilinadi.
        clean_segment, explicit_price, explicit_qty = extract_explicit_price(segment)
        if explicit_qty is not None:
            qty, query = explicit_qty, clean_segment
        else:
            qty, query = parse_product_line(clean_segment)
        if not query:
            continue

        requested_brand = detect_brand(query)
        query_norm = normalize(query)
        query_words = set(words(query))

        # Ustanovka va boshqa xizmatlar katalogdagi o'xshash tovar bilan adashtirilmasin.
        service_line = bool(re.search(
            r"ustanov|установк|монтаж|o['’]?rnat|kabel tort|прокладк|sozlash|настройк|доставка|пусконалад",
            query, re.I
        ))
        if service_line and explicit_price is not None:
            result.append({
                "Model / name": query,
                "Description": manual_position_description(query),
                "Category": "Услуга / дополнительная позиция",
                "Price USD": explicit_price,
                "Quantity": qty,
                "Price source": "Пользователь",
            })
            continue

        # HDD/SSD sig'imi kiritilgan bo'lsa, boshqa sig'imdagi disklarni chiqarib tashlash.
        capacity = re.search(r"(\d+)\s*(tb|тб|gb|гб)", query, re.I)
        capacity_value = (capacity.group(1), capacity.group(2).lower()) if capacity else None

        # Katalogni brend va sig'im bo'yicha cheklash.
        pool = []
        for item in catalog:
            item_brand = catalog_brand(item)
            if requested_brand and item_brand != requested_brand:
                continue
            if capacity_value:
                cap_num, cap_unit = capacity_value
                item_text = _item_search_text(item)
                unit_variants = {"tb": r"(?:tb|тб)", "тб": r"(?:tb|тб)",
                                 "gb": r"(?:gb|гб)", "гб": r"(?:gb|гб)"}
                if not re.search(
                    rf"(?<!\d){re.escape(cap_num)}\s*{unit_variants[cap_unit]}(?![a-z])",
                    item_text, re.I
                ):
                    continue
            pool.append(item)

        # Brendsiz umumiy PoE 8+2 so'rovida ishlab chiqaruvchini taxmin qilmaymiz.
        if not requested_brand and re.search(r"poe", query, re.I) and re.search(r"8\s*p(?:oe)?\s*\+?\s*2\s*(?:up|uplink)?", query, re.I):
            switch_candidates = []
            for item in catalog:
                text = _item_search_text(item)
                has_8_poe = bool(
                    re.search(r"8.{0,35}(?:poe|port)", text)
                    or "8+2" in normalize(item.get("Model / name", ""))
                    or "8mb+2" in normalize(item.get("Model / name", ""))
                )
                has_2_uplink = bool(
                    re.search(r"2.{0,35}(?:uplink|up-link)", text)
                    or "8+2" in normalize(item.get("Model / name", ""))
                    or "8mb+2" in normalize(item.get("Model / name", ""))
                )
                if ("poe" in text or "poe" in normalize(item.get("Model / name", ""))) and has_8_poe and has_2_uplink:
                    switch_candidates.append(item)
            if switch_candidates:
                options = []
                seen = set()
                for item in switch_candidates:
                    model = str(item.get("Model / name") or "")
                    key = normalize(model)
                    if key in seen:
                        continue
                    seen.add(key)
                    price = item.get("Price USD")
                    price_text = f"${price}" if isinstance(price, (int, float)) else "narx yo'q"
                    options.append(f"{model} ({price_text})")
                    if len(options) >= 8:
                        break
                not_found.append(
                    f"{segment} — brend/modelni aniqlashtiring: " + "; ".join(options)
                )
                continue

        scored = []
        for item in pool:
            score = _score_candidate(query, item, requested_brand)
            if score >= 8:
                scored.append((score, item))

        scored.sort(key=lambda pair: pair[0], reverse=True)

        if not scored:
            if explicit_price is not None:
                # Katalogda yo'q xizmat yoki material: foydalanuvchi bergan narx bilan qo'shamiz.
                result.append({
                    "Model / name": query,
                    "Description": "Позиция добавлена вручную. Цена за единицу указана пользователем.",
                    "Category": "Пользовательская позиция",
                    "Price USD": explicit_price,
                    "Quantity": qty,
                    "Price source": "Пользователь",
                })
            else:
                # Mahsulot katalogda topilmadi va narx berilmagan:
                # nomi va tavsifi bilan smetaga kiritamiz, narx kataklari bo'sh qoladi.
                result.append({
                    "Model / name": query,
                    "Description": manual_position_description(query),
                    "Category": "Дополнительная позиция",
                    "Price USD": None,
                    "Quantity": qty,
                    "Price source": "Не найдено в каталоге",
                })
            continue

        best_score = scored[0][0]
        close = [item for score, item in scored if score >= best_score - 5]

        # Agar brend ko'rsatilmagan va bir nechta ishlab chiqaruvchi bir xil darajada mos bo'lsa,
        # tasodifiy mahsulot tanlanmaydi.
        close_brands = {catalog_brand(item) for item in close if catalog_brand(item)}
        if not requested_brand and len(close_brands) > 1:
            options = []
            for item in close[:5]:
                model = str(item.get("Model / name") or "")
                price = item.get("Price USD")
                price_text = f"${price}" if isinstance(price, (int, float)) else "narx yo'q"
                options.append(f"{model} ({price_text})")
            if explicit_price is not None:
                result.append({
                    "Model / name": query,
                    "Description": "Позиция добавлена вручную. Цена за единицу указана пользователем.",
                    "Category": "Пользовательская позиция",
                    "Price USD": explicit_price,
                    "Quantity": qty,
                    "Price source": "Пользователь",
                })
            else:
                result.append({
                    "Model / name": query,
                    "Description": manual_position_description(query),
                    "Category": "Дополнительная позиция — модель уточняется",
                    "Price USD": None,
                    "Quantity": qty,
                    "Price source": "Требуется уточнение модели",
                })
                not_found.append(
                    f"{segment} — smetaga narxsiz qo'shildi; katalogdagi ehtimoliy variantlar: "
                    + "; ".join(options)
                )
            continue

        # Bir xil brend ichida ham bir nechta yaqin variant bo'lsa, modelni taxmin qilmaymiz.
        if len(close) > 1 and (close[0].get("Model / name") != close[1].get("Model / name")):
            # To'liq model kodi kiritilgan bo'lsa, eng yuqori ballni olish mumkin.
            exact_code = any(
                len(normalize(w)) >= 7 and normalize(w) in normalize(close[0].get("Model / name", ""))
                for w in words(query)
            )
            if not exact_code and close[1] and scored[0][0] - scored[1][0] < 12:
                options = []
                for item in close[:5]:
                    model = str(item.get("Model / name") or "")
                    price = item.get("Price USD")
                    price_text = f"${price}" if isinstance(price, (int, float)) else "narx yo'q"
                    options.append(f"{model} ({price_text})")
                if explicit_price is not None:
                    result.append({
                        "Model / name": query,
                        "Description": "Позиция добавлена вручную. Цена за единицу указана пользователем.",
                        "Category": "Пользовательская позиция",
                        "Price USD": explicit_price,
                        "Quantity": qty,
                        "Price source": "Пользователь",
                    })
                else:
                    result.append({
                        "Model / name": query,
                        "Description": manual_position_description(query),
                        "Category": "Дополнительная позиция — модель уточняется",
                        "Price USD": None,
                        "Quantity": qty,
                        "Price source": "Требуется уточнение модели",
                    })
                    not_found.append(
                        f"{segment} — smetaga narxsiz qo'shildi; aniq variantni tekshiring: "
                        + "; ".join(options)
                    )
                continue

        selected = scored[0][1].copy()
        selected["Quantity"] = qty
        if explicit_price is not None:
            selected["Price USD"] = explicit_price
            selected["Price source"] = "Пользователь"
        result.append(selected)

    return result, not_found


# ==========================================
# HISOB-KITOB: FAQAT USD
# ==========================================

def calculate_totals(items):

    total_usd = 0
    missing_price = []

    for item in items:

        qty = item["Quantity"]
        price = item.get("Price USD")

        if isinstance(price, (int, float)) and not isinstance(price, bool):

            item["Total USD"] = price * qty
            total_usd += item["Total USD"]

        else:

            item["Total USD"] = None
            missing_price.append(item.get("Model / name", "Unknown"))

    return total_usd, missing_price


# ==========================================
# EXCEL SMETA
# ==========================================


def create_excel(items, total_usd, missing, client_info=None):

    from openpyxl.styles import Border, Side
    from openpyxl.worksheet.page import PageMargins

    filename = OUTPUT_DIR / (
        "Smeta_" + datetime.now().strftime("%Y%m%d_%H%M%S")
        + ".xlsx"
    )

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Коммерческое предложение"

    # Ranglar
    navy = "17365D"
    blue = "4472C4"
    light_blue = "D9EAF7"
    pale_blue = "EDF3F8"
    white = "FFFFFF"
    grey = "D6DEE8"
    dark = "243447"
    green = "E2F0D9"
    red = "FCE4D6"

    thin_grey = Side(style="thin", color=grey)
    medium_blue = Side(style="medium", color=navy)

    # Umumiy sozlamalar
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90

    # 1. Kompaniya sarlavhasi
    ws.merge_cells("A1:F1")
    ws["A1"] = COMPANY
    ws["A1"].font = Font(
        name="Arial", size=18,
        bold=True, color=white
    )
    ws["A1"].fill = PatternFill(
        "solid", fgColor=navy
    )
    ws["A1"].alignment = Alignment(
        horizontal="center", vertical="center"
    )
    ws.row_dimensions[1].height = 38

    # 2. Hujjat nomi
    ws.merge_cells("A2:F2")
    ws["A2"] = (
        "КОММЕРЧЕСКОЕ ПРЕДЛОЖЕНИЕ "
        "НА СИСТЕМУ ВИДЕОНАБЛЮДЕНИЯ"
    )
    ws["A2"].font = Font(
        name="Arial", size=12,
        bold=True, color=navy
    )
    ws["A2"].alignment = Alignment(
        horizontal="center", vertical="center",
        wrap_text=True
    )
    ws.row_dimensions[2].height = 30

    # 3. Sana va valyuta
    ws.merge_cells("A3:C3")
    ws["A3"] = (
        "Дата: " + datetime.now().strftime("%d.%m.%Y")
        + "   |   Телефон компании: +998 91 212 80 80"
    )
    ws["A3"].font = Font(
        name="Arial", size=10, color=dark
    )

    ws.merge_cells("D3:F3")
    ws["D3"] = "ВАЛЮТА: USD ($)"
    ws["D3"].font = Font(
        name="Arial", size=10,
        bold=True, color=navy
    )
    ws["D3"].alignment = Alignment(
        horizontal="right"
    )
    ws.row_dimensions[3].height = 25

    # 4. Buyurtmachi ma'lumotlari
    client_info = client_info or {}
    ws.merge_cells("A4:F4")
    client_line = "Заказчик: " + str(client_info.get("client") or "—")
    client_line += "   |   Телефон: " + str(client_info.get("phone") or "—")
    client_line += "   |   Объект: " + str(client_info.get("object") or "—")
    ws["A4"] = client_line
    ws["A4"].font = Font(name="Arial", size=9, color=dark)
    ws["A4"].alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[4].height = 25

    # 5. Jadval sarlavhasi
    headers = [
        "№",
        "Наименование товара",
        "Описание",
        "Кол-во",
        "Цена, USD",
        "Сумма, USD"
    ]

    for col, value in enumerate(headers, 1):
        cell = ws.cell(5, col, value)
        cell.font = Font(
            name="Arial", size=10,
            bold=True, color=white
        )
        cell.fill = PatternFill(
            "solid", fgColor=blue
        )
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )
        cell.border = Border(
            top=medium_blue,
            bottom=medium_blue,
            left=thin_grey,
            right=thin_grey
        )

    ws.row_dimensions[5].height = 32

    # 5. Mahsulotlar
    start_row = 6

    for number, item in enumerate(items, 1):

        row = start_row + number - 1

        values = [
            number,
            item.get("Model / name", ""),
            item.get("Description", ""),
            item["Quantity"],
            item.get("Price USD")
                if item.get("Price USD") is not None
                else None,
            item.get("Total USD")
                if item.get("Total USD") is not None
                else None
        ]

        for col, value in enumerate(values, 1):

            cell = ws.cell(row, col, value)

            cell.font = Font(
                name="Arial", size=10,
                color=dark
            )

            cell.fill = PatternFill(
                "solid",
                fgColor=white if number % 2 else pale_blue
            )

            cell.alignment = Alignment(
                horizontal=(
                    "center" if col in (1, 4)
                    else "right" if col in (5, 6)
                    else "left"
                ),
                vertical="center",
                wrap_text=True
            )

            cell.border = Border(
                left=thin_grey,
                right=thin_grey,
                top=thin_grey,
                bottom=thin_grey
            )

        # USD format
        for col in (5, 6):
            cell = ws.cell(row, col)
            if isinstance(cell.value, (int, float)):
                cell.number_format = '$#,##0.00'

        # Tavsif uzunligiga mos qator balandligi
        description = str(item.get("Description") or "")
        lines = max(2, (len(description) // 85) + 1)

        ws.row_dimensions[row].height = min(
            max(45, lines * 18), 120
        )

    # 6. Umumiy summa
    last_product_row = start_row + len(items) - 1
    total_row = last_product_row + 1

    ws.merge_cells(
        start_row=total_row,
        start_column=1,
        end_row=total_row,
        end_column=5
    )

    total_label = ws.cell(total_row, 1)
    total_label.value = "ИТОГО К ОПЛАТЕ:"
    total_label.font = Font(
        name="Arial", size=13,
        bold=True, color=navy
    )
    total_label.fill = PatternFill(
        "solid", fgColor=light_blue
    )
    total_label.alignment = Alignment(
        horizontal="right",
        vertical="center"
    )

    total_cell = ws.cell(total_row, 6)
    total_cell.value = total_usd
    total_cell.number_format = '$#,##0.00'
    total_cell.font = Font(
        name="Arial", size=14,
        bold=True, color=navy
    )
    total_cell.fill = PatternFill(
        "solid", fgColor=light_blue
    )
    total_cell.alignment = Alignment(
        horizontal="right",
        vertical="center"
    )

    for col in range(1, 7):
        ws.cell(total_row, col).border = Border(
            top=medium_blue,
            bottom=medium_blue
        )

    ws.row_dimensions[total_row].height = 38

    # 7. Izohlar
    note_start = total_row + 3

    ws.merge_cells(
        start_row=note_start,
        start_column=1,
        end_row=note_start,
        end_column=6
    )

    ws.cell(note_start, 1, "ПРИМЕЧАНИЯ")
    ws.cell(note_start, 1).font = Font(
        name="Arial", size=11,
        bold=True, color=navy
    )
    ws.cell(note_start, 1).fill = PatternFill(
        "solid", fgColor=pale_blue
    )
    ws.cell(note_start, 1).alignment = Alignment(
        vertical="center"
    )
    ws.row_dimensions[note_start].height = 25

    notes = [
        "Телефон компании: +998 91 212 80 80.",
        "Все цены указаны исключительно в долларах США (USD).",
        "Цены на оборудование указаны согласно каталогу; цены пользовательских услуг и материалов введены заказчиком.",
        "Стоимость монтажа, кабеля и дополнительных материалов не включена, если они не указаны отдельной строкой."
    ]

    if missing:
        notes.append(
            "Цена не указана для: " + ", ".join(missing)
        )

    for index, note in enumerate(notes, 1):

        row = note_start + index

        ws.merge_cells(
            start_row=row,
            start_column=1,
            end_row=row,
            end_column=6
        )

        cell = ws.cell(row, 1, "• " + note)
        cell.font = Font(
            name="Arial", size=9, color=dark
        )
        cell.alignment = Alignment(
            vertical="center",
            wrap_text=True
        )
        ws.row_dimensions[row].height = 23

    # 8. Ustun kengliklari
    widths = {
        "A": 7,
        "B": 34,
        "C": 72,
        "D": 12,
        "E": 17,
        "F": 19
    }

    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    # 9. Jadval filtri va muzlatish
    ws.freeze_panes = "A6"
    ws.auto_filter.ref = (
        f"A5:F{last_product_row}"
    )

    # 10. Chop etish sozlamalari
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    ws.page_margins = PageMargins(
        left=0.25,
        right=0.25,
        top=0.45,
        bottom=0.45,
        header=0.2,
        footer=0.2
    )

    ws.print_title_rows = "1:5"
    ws.print_area = f"A1:F{note_start + len(notes)}"

    ws.oddFooter.center.text = (
        "HIKVISION UZBEKISTAN MCHJ"
    )
    ws.oddFooter.right.text = "Страница &P из &N"

    # 11. Faylni saqlash
    wb.save(filename)

    return filename



# ==========================================
# PDF SMETA
# ==========================================

def create_pdf(items, total_usd, missing, client_info=None):

    filename = OUTPUT_DIR / (
        "Smeta_" + datetime.now().strftime("%Y%m%d_%H%M%S")
        + ".pdf"
    )

    font_regular = "Helvetica"
    font_bold = "Helvetica-Bold"

    arial = Path("C:/Windows/Fonts/arial.ttf")
    arial_bold = Path("C:/Windows/Fonts/arialbd.ttf")

    if arial.exists():
        pdfmetrics.registerFont(
            TTFont("ArialCustom", str(arial))
        )
        font_regular = "ArialCustom"

    if arial_bold.exists():
        pdfmetrics.registerFont(
            TTFont("ArialBoldCustom", str(arial_bold))
        )
        font_bold = "ArialBoldCustom"

    doc = SimpleDocTemplate(
        str(filename),
        pagesize=landscape(A4),
        rightMargin=25,
        leftMargin=25,
        topMargin=30,
        bottomMargin=30
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "TitleCustom",
        parent=styles["Title"],
        fontName=font_bold,
        fontSize=15,
        leading=19,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#17365D")
    )

    normal = ParagraphStyle(
        "NormalCustom",
        parent=styles["Normal"],
        fontName=font_regular,
        fontSize=8,
        leading=10
    )

    small = ParagraphStyle(
        "SmallCustom",
        parent=normal,
        fontSize=7,
        leading=9
    )

    story = []

    story.append(Paragraph(COMPANY, title_style))
    story.append(Spacer(1, 8))

    story.append(
        Paragraph(
            "КОММЕРЧЕСКОЕ ПРЕДЛОЖЕНИЕ НА СИСТЕМУ ВИДЕОНАБЛЮДЕНИЯ",
            title_style
        )
    )

    story.append(Spacer(1, 12))

    story.append(
        Paragraph(
            "Дата: " + datetime.now().strftime("%d.%m.%Y")
            + " | Телефон компании: +998 91 212 80 80",
            normal
        )
    )

    client_info = client_info or {}
    client_line = (
        "Заказчик: " + str(client_info.get("client") or "—")
        + " | Телефон: " + str(client_info.get("phone") or "—")
        + " | Объект: " + str(client_info.get("object") or "—")
    )
    story.append(Paragraph(client_line, normal))
    story.append(Spacer(1, 15))

    data = [[
        "№",
        "Наименование товара",
        "Описание",
        "Кол-во",
        "Цена USD",
        "Сумма USD"
    ]]

    for number, item in enumerate(items, 1):

        price = item.get("Price USD")
        total = item.get("Total USD")

        data.append([
            str(number),
            Paragraph(
                str(item.get("Model / name") or ""),
                small
            ),
            Paragraph(
                str(item.get("Description") or "")[:500],
                small
            ),
            str(item["Quantity"]),
            f"{price:,.2f}" if isinstance(price, (int, float)) else "",
            f"{total:,.2f}" if isinstance(total, (int, float)) else ""
        ])

    table = Table(
        data,
        colWidths=[25, 145, 360, 45, 85, 95],
        repeatRows=1
    )

    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0),
         colors.HexColor("#17365D")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, -1), font_regular),
        ("FONTNAME", (0, 0), (-1, 0), font_bold),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ALIGN", (0, 0), (0, -1), "CENTER"),
        ("ALIGN", (3, 0), (-1, -1), "RIGHT"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1),
         [colors.white, colors.HexColor("#EDF3F8")]),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7)
    ]))

    story.append(table)
    story.append(Spacer(1, 18))

    summary = [[
        "ИТОГОВАЯ СУММА:",
        f"{total_usd:,.2f} USD"
    ]]

    summary_table = Table(
        summary,
        colWidths=[230, 170],
        hAlign="RIGHT"
    )

    summary_table.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), font_bold),
        ("BACKGROUND", (0, 0), (-1, -1),
         colors.HexColor("#D9EAF7")),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ALIGN", (1, 0), (1, -1), "RIGHT"),
        ("PADDING", (0, 0), (-1, -1), 9)
    ]))

    story.append(summary_table)
    story.append(Spacer(1, 15))

    notes = [
        "Примечания:",
        "• Телефон компании: +998 91 212 80 80.",
        "• Все цены указаны исключительно в долларах США (USD).",
        "• Цены на оборудование указаны согласно каталогу; цены пользовательских услуг и материалов введены заказчиком.",
        "• Стоимость монтажа, кабеля и дополнительных материалов не включена, если они не указаны отдельной строкой."
    ]

    if missing:
        notes.append(
            "• Цена не указана для следующих позиций: "
            + ", ".join(missing)
        )

    for note in notes:
        story.append(Paragraph(note, normal))
        story.append(Spacer(1, 4))

    doc.build(story)

    return filename


# ==========================================
# ASOSIY SMETA FUNKSIYASI
# ==========================================

def generate_smeta(request, catalog, client_info=None):

    items, not_found = find_smeta_items(
        request,
        catalog
    )

    if not items:
        return None, None, [], not_found

    total_usd, missing = calculate_totals(items)

    pdf_file = create_pdf(
        items,
        total_usd,
        missing,
        client_info=client_info or {}
    )

    excel_file = create_excel(
        items,
        total_usd,
        missing,
        client_info=client_info or {}
    )

    return pdf_file, excel_file, items, not_found
