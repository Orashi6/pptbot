import asyncio
import io
import json
import logging
import os
import subprocess
import uuid
from datetime import datetime, timedelta, timezone
from typing import List, Optional

import aiohttp
import asyncpg
import docx
from aiohttp import web
from aiogram import Bot, Dispatcher, F, Router
from aiogram.enums import ParseMode
from aiogram.filters import Command, CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import (
    CallbackQuery,
    FSInputFile,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    KeyboardButton,
    Message,
    ReplyKeyboardMarkup,
    ReplyKeyboardRemove,
)
from google import genai
from google.genai import errors, types
from pptx import Presentation
from pptx.util import Inches, Pt
from pydantic import BaseModel, Field
from pypdf import PdfReader
from google.genai.errors import APIError

# --- SOZLAMALAR (RENDER ENV VARIABLES) ---
BOT_TOKEN = os.getenv("BOT_TOKEN", "YOUR_PPT_BOT_TOKEN")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "YOUR_GEMINI_API_KEY")
PIXABAY_API_KEY = os.getenv("PIXABAY_API_KEY", "").strip()
DATABASE_URL = os.getenv("DATABASE_URL", "postgresql://user:pass@host/db")
ADMIN_ID = int(os.getenv("ADMIN_ID", "705345765"))
GIFT_BOT_USERNAME = os.getenv("GIFT_BOT_USERNAME", "giftolishbot").lstrip("@").strip()

API_KEYS = os.getenv("API_KEYS", "YORU API KEYS")
MODELS = ["gemini-3.8-flash", "gemini-3.8-flash-lite", "gemini-3-flash-preview"]

async def generate_with_fallback(prompt: str, config=None):
    last_error = None
    for model in MODELS:
        for idx, api_key in enumerate(API_KEYS, start=1):
            client = genai.Client(api_key=api_key)
            try:
                response = await client.aio.models.generate_content(
                    model=model,
                    contents=prompt,
                    config=config,
                )
                return response
            except APIError as e:
                last_error = e
                if e.code == 404:
                    break  # Model deprecated -> skip remaining keys for this model
                if e.code == 429:
                    await asyncio.sleep(1)
                    continue
                raise
    raise RuntimeError(f"All Gemini models/keys exhausted: {last_error}")

WEEKLY_FREE_LIMIT = 3
PORT = int(os.getenv("PORT", 8080))

TEMPLATES = {
    "dark": ("🌙 Dark Modern", "templates/template_dark.pptx"),
    "business": ("💼 Business", "templates/template_business.pptx"),
    "minimal": ("⚪ Minimalist", "templates/template_minimal.pptx"),
}

LANGUAGES = {
    "uz": ("🇺🇿 O'zbek tili", "O'zbek tilida"),
    "ru": ("🇷🇺 Русский язык", "Rus tilida (на русском языке)"),
    "en": ("🇬🇧 English", "Ingliz tilida (in English)"),
}

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("ppt_bot")

gemini_client = genai.Client(api_key=GEMINI_API_KEY)
router = Router()
db_pool: Optional[asyncpg.Pool] = None


# --- PYDANTIC SCHEMA ---
class SlideContent(BaseModel):
    title: str = Field(description="Slayd sarlavhasi (qisqa va aniq)")
    points: List[str] = Field(description="Slaydning asosiy fikrlari (3-4 ta punkt)")
    image_query: str = Field(
        description="Pixabay yoki Wikimedia orqali rasm qidirish uchun 2-3 so'zdan iborat INGLIZCHA kalit so'z"
    )


class PresentationData(BaseModel):
    main_title: str = Field(description="Prezentatsiyaning asosiy sarlavhasi")
    subtitle: str = Field(description="Muqova slaydi uchun qisqa tavsif")
    slides: List[SlideContent] = Field(description="Asosiy mazmunli slaydlar ro'yxati")


# --- FSM HOLATLARI ---
class PresStates(StatesGroup):
    waiting_for_source = State()
    waiting_for_count = State()
    waiting_for_lang = State()
    waiting_for_template = State()
    waiting_for_broadcast = State()


# --- NEON POSTGRESQL BAZA ---
async def init_db():
    global db_pool
    dsn = DATABASE_URL.replace("postgres://", "postgresql://", 1)
    db_pool = await asyncpg.create_pool(dsn, min_size=1, max_size=5, statement_cache_size=0)

    async with db_pool.acquire() as conn:
        await conn.execute(
            """
            CREATE TABLE IF NOT EXISTS ppt_users (
                user_id BIGINT PRIMARY KEY,
                full_name TEXT,
                weekly_used INT DEFAULT 0,
                week_reset_at TIMESTAMPTZ DEFAULT NOW() + INTERVAL '7 days',
                is_pro BOOLEAN DEFAULT FALSE,
                total_generated INT DEFAULT 0,
                created_at TIMESTAMPTZ DEFAULT NOW()
            );
            CREATE TABLE IF NOT EXISTS bot_settings (
                key TEXT PRIMARY KEY,
                value INT NOT NULL
            );
            INSERT INTO bot_settings (key, value) VALUES ('ppt_pro', 50) ON CONFLICT (key) DO NOTHING;
            """
        )


async def get_ppt_pro_price() -> int:
    val = await db_pool.fetchval("SELECT value FROM bot_settings WHERE key='ppt_pro'")
    return int(val) if val and int(val) > 0 else 50


async def check_and_get_user(user_id: int, full_name: str) -> asyncpg.Record:
    async with db_pool.acquire() as conn:
        user = await conn.fetchrow("SELECT * FROM ppt_users WHERE user_id = $1", user_id)
        now = datetime.now(timezone.utc)

        if not user:
            user = await conn.fetchrow(
                """
                INSERT INTO ppt_users (user_id, full_name, weekly_used, week_reset_at)
                VALUES ($1, $2, 0, $3)
                RETURNING *
                """,
                user_id,
                full_name,
                now + timedelta(days=7),
            )
        elif now >= user["week_reset_at"]:
            user = await conn.fetchrow(
                """
                UPDATE ppt_users
                SET weekly_used = 0, week_reset_at = $1
                WHERE user_id = $2
                RETURNING *
                """,
                now + timedelta(days=7),
                user_id,
            )
        return user


async def increment_usage(user_id: int):
    async with db_pool.acquire() as conn:
        await conn.execute(
            """
            UPDATE ppt_users
            SET weekly_used = weekly_used + 1,
                total_generated = total_generated + 1
            WHERE user_id = $1
            """,
            user_id,
        )


# --- KLAVIATURALAR ---
def get_count_keyboard() -> ReplyKeyboardMarkup:
    buttons = [
        [KeyboardButton(text="5"), KeyboardButton(text="7"), KeyboardButton(text="10")],
        [KeyboardButton(text="12"), KeyboardButton(text="15")],
    ]
    return ReplyKeyboardMarkup(keyboard=buttons, resize_keyboard=True, one_time_keyboard=True)


def get_lang_keyboard() -> InlineKeyboardMarkup:
    buttons = [
        [InlineKeyboardButton(text=label, callback_data=f"lang:{key}")]
        for key, (label, _) in LANGUAGES.items()
    ]
    return InlineKeyboardMarkup(inline_keyboard=buttons)


def get_template_keyboard() -> InlineKeyboardMarkup:
    buttons = [
        [InlineKeyboardButton(text=label, callback_data=f"tpl:{key}")]
        for key, (label, _) in TEMPLATES.items()
    ]
    return InlineKeyboardMarkup(inline_keyboard=buttons)


async def get_pro_keyboard(user_id: int) -> InlineKeyboardMarkup:
    stars = await get_ppt_pro_price()
    pay_url = f"https://t.me/{GIFT_BOT_USERNAME}?start=pay_ppt_pro_{user_id}"
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text=f"⭐ PRO obuna olish — {stars} Stars", url=pay_url)]
        ]
    )


# --- FAYLDAN MATN AJRATIB OLISH (.PDF / .DOCX) ---
def extract_text_from_bytes(file_bytes: bytes, filename: str) -> str:
    ext = filename.lower().split(".")[-1]
    text_parts = []

    if ext == "pdf":
        reader = PdfReader(io.BytesIO(file_bytes))
        for page in reader.pages[:15]:
            extracted = page.extract_text()
            if extracted:
                text_parts.append(extracted)
    elif ext == "docx":
        document = docx.Document(io.BytesIO(file_bytes))
        for para in document.paragraphs:
            if para.text.strip():
                text_parts.append(para.text.strip())

    full_text = "\n".join(text_parts)
    return full_text[:15000]


# --- GEMINI ORQALI MATN GENERATSIYA (404 / 503 FALLBACK BILAN) ---
async def generate_with_fallback(contents, config):
    models_to_try = [
        "gemini-3.8-flash",
        "gemini-2.5-flash",
        "gemini-2.5-flash-lite",
    ]

    last_err = None
    for model_name in models_to_try:
        for attempt in range(2):
            try:
                return await gemini_client.aio.models.generate_content(
                    model=model_name,
                    contents=contents,
                    config=config,
                )
            except errors.ServerError as e:
                last_err = e
                if e.code == 503:
                    await asyncio.sleep(2)
                    continue
                raise
            except errors.ClientError as e:
                last_err = e
                if e.code == 404:
                    break
                raise

    if isinstance(last_err, errors.ServerError) and last_err.code == 503:
        raise last_err
    raise RuntimeError(f"Barcha Gemini modellarida xatolik yuz berdi: {last_err}")


async def generate_presentation_content(
    source_text: str, slide_count: int, lang_instruction: str, is_document: bool = False
) -> PresentationData:
    if is_document:
        task_desc = (
            f"Quyidagi hujjat/matn mazmunini chuqur tahlil qilib, uning asosiy g'oyalaridan "
            f"professional prezentatsiya tuzib ber:\n\n--- MATN BOSHI ---\n{source_text}\n--- MATN OXIRI ---"
        )
    else:
        task_desc = f"Mavzu: '{source_text}'. Ushbu mavzu bo'yicha professional prezentatsiya matnini tayyorla."

    prompt = (
        f"{task_desc}\n\n"
        f"Talablar:\n"
        f"1. Til: {lang_instruction} yozilsin.\n"
        f"2. Muqova (1-slayd) ni hisobga olgan holda, 'slides' ro'yxatida aynan {slide_count - 1} ta "
        f"mazmunli slayd bo'lsin (jami {slide_count} ta slayd).\n"
        f"3. Har bir slaydda 3-4 ta aniq, qisqa punkt bo'lsin (har biri 12-16 so'zdan oshmasin).\n"
        f"4. 'image_query' maydoniga rasm qidirish uchun 2-3 so'zdan iborat INGLIZCHA kalit so'z yoz."
    )

    config = types.GenerateContentConfig(
        response_mime_type="application/json",
        response_schema=PresentationData,
        temperature=0.6,
    )

    response = await generate_with_fallback(prompt, config)
    return PresentationData(**json.loads(response.text))


# --- PIXABAY + WIKIMEDIA COMMONS ORQALI RASM YUKLASH ---
async def fetch_wikimedia_image(session: aiohttp.ClientSession, query: str) -> Optional[io.BytesIO]:
    url = "https://commons.wikimedia.org/w/api.php"
    headers = {"User-Agent": "TelegramPresentationBot/1.0 (https://t.me/giftolishbot)"}
    params = {
        "action": "query",
        "format": "json",
        "generator": "search",
        "gsrsearch": f"filetype:bitmap {query}",
        "gsrnamespace": 6,
        "gsrlimit": 1,
        "prop": "imageinfo",
        "iiprop": "url",
        "iiurlwidth": 1024,
    }
    try:
        async with session.get(url, headers=headers, params=params, timeout=10) as resp:
            if resp.status != 200:
                return None
            data = await resp.json()
            pages = data.get("query", {}).get("pages", {})
            if not pages:
                return None
            first_page = next(iter(pages.values()))
            img_url = first_page.get("imageinfo", [{}])[0].get("thumburl")
            if not img_url:
                return None
            async with session.get(img_url, headers=headers, timeout=15) as img_resp:
                if img_resp.status == 200:
                    return io.BytesIO(await img_resp.read())
    except Exception:
        return None
    return None


async def fetch_slide_image(session: aiohttp.ClientSession, query: str) -> Optional[io.BytesIO]:
    if PIXABAY_API_KEY:
        url = "https://pixabay.com/api/"
        params = {
            "key": PIXABAY_API_KEY,
            "q": query,
            "image_type": "photo",
            "orientation": "horizontal",
            "per_page": 3,
            "safesearch": "true",
        }
        try:
            async with session.get(url, params=params, timeout=10) as resp:
                if resp.status == 200:
                    data = await resp.json()
                    hits = data.get("hits", [])
                    if hits:
                        img_url = hits[0]["largeImageURL"]
                        async with session.get(img_url, timeout=15) as img_resp:
                            if img_resp.status == 200:
                                return io.BytesIO(await img_resp.read())
        except Exception:
            pass

    return await fetch_wikimedia_image(session, query)


async def fetch_all_slide_images(slides: List[SlideContent]) -> List[Optional[io.BytesIO]]:
    async with aiohttp.ClientSession() as session:
        tasks = [fetch_slide_image(session, slide.image_query) for slide in slides]
        return await asyncio.gather(*tasks)


# --- AUTO-FIT SHRIFT HISOBLASH ---
def calculate_optimal_font_size(points: List[str]) -> Pt:
    total_chars = sum(len(p) for p in points)
    if total_chars > 420:
        return Pt(14)
    elif total_chars > 300:
        return Pt(16)
    elif total_chars > 200:
        return Pt(18)
    return Pt(20)


# --- .PPTX YIG'ISH ---
def build_pptx_with_template(
    data: PresentationData,
    images: List[Optional[io.BytesIO]],
    template_path: str,
    output_path: str,
) -> None:
    if os.path.exists(template_path):
        prs = Presentation(template_path)
        for i in range(len(prs.slides) - 1, -1, -1):
            rId = prs.slides._sldIdLst[i].rId
            prs.part.drop_rel(rId)
            del prs.slides._sldIdLst[i]
    else:
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)

    cover = prs.slides.add_slide(prs.slide_layouts[0])
    if cover.shapes.title:
        cover.shapes.title.text = data.main_title

    for shape in cover.placeholders:
        if shape.placeholder_format.idx == 1:
            shape.text = data.subtitle
            break

    content_layout = prs.slide_layouts[1]

    for slide_info, img_stream in zip(data.slides, images):
        slide = prs.slides.add_slide(content_layout)

        if slide.shapes.title:
            slide.shapes.title.text = slide_info.title

        body_shape = None
        for shape in slide.placeholders:
            if shape.placeholder_format.idx == 1:
                body_shape = shape
                break

        if body_shape:
            if img_stream:
                body_shape.left = Inches(0.7)
                body_shape.width = Inches(6.5)

            tf = body_shape.text_frame
            tf.clear()
            tf.word_wrap = True
            font_size = calculate_optimal_font_size(slide_info.points)

            for i, point in enumerate(slide_info.points):
                p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                p.text = point
                p.level = 0
                p.font.size = font_size

        if img_stream:
            img_stream.seek(0)
            slide.shapes.add_picture(
                img_stream,
                left=Inches(7.5),
                top=Inches(1.8),
                width=Inches(5.1),
            )

    prs.save(output_path)


# --- .PPTX NI .PDF GA O'GIRISH ---
def convert_pptx_to_pdf(pptx_path: str, output_dir: str = ".") -> Optional[str]:
    try:
        subprocess.run(
            ["soffice", "--headless", "--convert-to", "pdf", pptx_path, "--outdir", output_dir],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            timeout=30,
        )
        pdf_path = pptx_path.rsplit(".", 1)[0] + ".pdf"
        return pdf_path if os.path.exists(pdf_path) else None
    except Exception:
        return None


# --- FOYDALANUVCHI HANDLERLARI ---
@router.message(CommandStart())
async def cmd_start(message: Message, state: FSMContext):
    await state.clear()
    user = await check_and_get_user(message.from_user.id, message.from_user.full_name)

    if user["is_pro"]:
        status_text = "💎 Tarif: <b>PRO (Cheksiz)</b>"
    else:
        left = max(0, WEEKLY_FREE_LIMIT - user["weekly_used"])
        status_text = f"🆓 Haftalik bepul limit: <b>{left}/{WEEKLY_FREE_LIMIT} ta qoldi</b>"

    await message.answer(
        f"👋 Salom, <b>{message.from_user.first_name}</b>!\n"
        f"{status_text}\n\n"
        "📌 <b>Prezentatsiya mavzusini yozing</b>, uzun matn yuboring yoki tayyor "
        "<b>.pdf / .docx</b> fayl tashlang:\n\n"
        "<i>💎 Cheksiz PRO obuna uchun: /pro</i>",
        parse_mode=ParseMode.HTML,
        reply_markup=ReplyKeyboardRemove(),
    )
    await state.set_state(PresStates.waiting_for_source)


@router.message(Command("pro"))
async def cmd_pro(message: Message):
    kb = await get_pro_keyboard(message.from_user.id)
    await message.answer(
        "⭐ <b>PRO obuna imkoniyatlari:</b>\n\n"
        "• Haftalik 3 talik cheklov olib tashlanadi (Cheksiz yaratish)\n"
        "• 20 tagacha slayd yaratish imkoniyati\n"
        "• Navbatsiz tezkor generatsiya\n\n"
        f"💳 To'lov rasmiy @{GIFT_BOT_USERNAME} orqali Telegram Stars bilan amalga oshiriladi va avtomatik faollashadi:",
        parse_mode=ParseMode.HTML,
        reply_markup=kb,
    )


@router.message(PresStates.waiting_for_source, F.document)
async def process_document_source(message: Message, state: FSMContext, bot: Bot):
    doc = message.document
    if not doc.file_name.lower().endswith((".pdf", ".docx")):
        await message.answer("⚠️ Faqat <b>.pdf</b> yoki <b>.docx</b> formatdagi fayllarni yuboring.")
        return

    if doc.file_size > 10 * 1024 * 1024:
        await message.answer("⚠️ Fayl hajmi 10 MB dan kichik bo'lishi kerak.")
        return

    wait_msg = await message.answer("📄 Fayl o'qilmoqda...")
    file_io = io.BytesIO()
    await bot.download(doc, destination=file_io)

    extracted_text = await asyncio.to_thread(
        extract_text_from_bytes, file_io.getvalue(), doc.file_name
    )
    await wait_msg.delete()

    if len(extracted_text) < 100:
        await message.answer("⚠️ Fayl ichidan yetarli matn topilmadi.")
        return

    await state.update_data(
        source_text=extracted_text,
        display_title=doc.file_name,
        is_document=True,
    )
    await message.answer(
        f"✅ Fayl qabul qilindi: <b>{doc.file_name}</b>\n\n"
        "📊 <b>Nechta slayd kerak?</b> (3 dan 20 gacha):",
        parse_mode=ParseMode.HTML,
        reply_markup=get_count_keyboard(),
    )
    await state.set_state(PresStates.waiting_for_count)


@router.message(PresStates.waiting_for_source, F.text)
async def process_text_source(message: Message, state: FSMContext):
    text = message.text.strip()
    if len(text) < 3:
        await message.answer("⚠️ Mavzu juda qisqa. Batafsilroq yozing:")
        return

    is_long_text = len(text) > 150
    display_title = text[:40] + "..." if is_long_text else text

    await state.update_data(
        source_text=text,
        display_title=display_title,
        is_document=is_long_text,
    )
    await message.answer(
        f"✅ Qabul qilindi: <b>{display_title}</b>\n\n"
        "📊 <b>Nechta slayd kerak?</b> (3 dan 20 gacha):",
        parse_mode=ParseMode.HTML,
        reply_markup=get_count_keyboard(),
    )
    await state.set_state(PresStates.waiting_for_count)


@router.message(PresStates.waiting_for_count, F.text)
async def process_count(message: Message, state: FSMContext):
    text = message.text.strip()
    if not text.isdigit() or not (3 <= int(text) <= 20):
        await message.answer("⚠️ Iltimos, 3 dan 20 gacha raqam kiriting:")
        return

    await state.update_data(slide_count=int(text))
    await message.answer(
        f"📑 Slaydlar soni: <b>{text} ta</b>",
        parse_mode=ParseMode.HTML,
        reply_markup=ReplyKeyboardRemove(),
    )
    await message.answer(
        "🌐 <b>Prezentatsiya tilini tanlang:</b>",
        parse_mode=ParseMode.HTML,
        reply_markup=get_lang_keyboard(),
    )
    await state.set_state(PresStates.waiting_for_lang)


@router.callback_query(PresStates.waiting_for_lang, F.data.startswith("lang:"))
async def process_lang(callback: CallbackQuery, state: FSMContext):
    lang_key = callback.data.split(":", 1)[1]
    if lang_key not in LANGUAGES:
        await callback.answer("Xato tanlov", show_alert=True)
        return

    await state.update_data(lang_key=lang_key)
    await callback.answer()
    await callback.message.edit_text(
        f"🌐 Til: <b>{LANGUAGES[lang_key][0]}</b>\n\n"
        "🎨 <b>Dizayn shablonini tanlang:</b>",
        parse_mode=ParseMode.HTML,
        reply_markup=get_template_keyboard(),
    )
    await state.set_state(PresStates.waiting_for_template)


@router.callback_query(PresStates.waiting_for_template, F.data.startswith("tpl:"))
async def process_template(callback: CallbackQuery, state: FSMContext):
    tpl_key = callback.data.split(":", 1)[1]
    if tpl_key not in TEMPLATES:
        await callback.answer("Xato shablon", show_alert=True)
        return

    user_id = callback.from_user.id
    user = await check_and_get_user(user_id, callback.from_user.full_name)

    if not user["is_pro"] and user_id != ADMIN_ID and user["weekly_used"] >= WEEKLY_FREE_LIMIT:
        await callback.answer()
        await state.clear()
        kb = await get_pro_keyboard(user_id)
        await callback.message.edit_text(
            f"🚫 <b>Haftalik bepul limitingiz tugadi ({WEEKLY_FREE_LIMIT}/{WEEKLY_FREE_LIMIT})!</b>\n\n"
            "Yangi limit ochilishini kutishingiz yoki cheksiz foydalanish uchun "
            "<b>PRO obuna</b> olishingiz mumkin:",
            parse_mode=ParseMode.HTML,
            reply_markup=kb,
        )
        return

    tpl_label, template_path = TEMPLATES[tpl_key]
    data = await state.get_data()
    await state.clear()
    await callback.answer()

    lang_label, lang_instruction = LANGUAGES[data["lang_key"]]
    status_msg = await callback.message.edit_text(
        f"🎨 Dizayn: <b>{tpl_label}</b> | 🌐 {lang_label}\n\n"
        "⏳ <b>1/3 — Gemini AI matn tuzmoqda...</b>",
        parse_mode=ParseMode.HTML,
    )

    uid = uuid.uuid4().hex[:8]
    pptx_file = f"pres_{uid}.pptx"
    pdf_file = None

    try:
        pres_data = await generate_presentation_content(
            source_text=data["source_text"],
            slide_count=data["slide_count"],
            lang_instruction=lang_instruction,
            is_document=data["is_document"],
        )

        await status_msg.edit_text(
            f"🎨 Dizayn: <b>{tpl_label}</b> | 🌐 {lang_label}\n\n"
            "🖼 <b>2/3 — Slaydlar uchun rasmlar yuklanmoqda...</b>",
            parse_mode=ParseMode.HTML,
        )

        images = await fetch_all_slide_images(pres_data.slides)

        await status_msg.edit_text(
            f"🎨 Dizayn: <b>{tpl_label}</b> | 🌐 {lang_label}\n\n"
            "⚙️ <b>3/3 — .pptx va .pdf fayllar yig'ilmoqda...</b>",
            parse_mode=ParseMode.HTML,
        )

        await asyncio.to_thread(
            build_pptx_with_template, pres_data, images, template_path, pptx_file
        )
        pdf_file = await asyncio.to_thread(convert_pptx_to_pdf, pptx_file)

        safe_name = "".join(
            c for c in data["display_title"] if c.isalnum() or c in (" ", "-", "_")
        ).rstrip()[:30] or "Prezentatsiya"

        found_imgs = sum(1 for img in images if img is not None)
        await callback.message.answer_document(
            document=FSInputFile(pptx_file, filename=f"{safe_name}.pptx"),
            caption=(
                f"🎉 <b>Prezentatsiya tayyor!</b>\n\n"
                f"📌 <b>Sarlavha:</b> {pres_data.main_title}\n"
                f"📑 <b>Slaydlar:</b> {len(pres_data.slides) + 1} ta\n"
                f"🖼 <b>Rasmlar:</b> {found_imgs} ta\n"
                f"Yana yaratish uchun /start bosing."
            ),
            parse_mode=ParseMode.HTML,
        )

        if pdf_file and os.path.exists(pdf_file):
            await callback.message.answer_document(
                document=FSInputFile(pdf_file, filename=f"{safe_name}.pdf"),
                caption="📱 Telefonda qulay ko'rish uchun PDF nusxasi",
            )

        await increment_usage(user_id)
        await status_msg.delete()

    except errors.ServerError as e:
        logger.exception("Gemini ServerError:")
        if e.code == 503:
            await status_msg.edit_text(
                "⏳ Sun'iy intellekt serverlarida hozir yuklama juda yuqori.\n"
                "Iltimos, 10-15 soniyadan so'ng qayta urinib ko'ring: /start"
            )
        else:
            await status_msg.edit_text("⚠️ Serverda vaqtinchalik texnik xatolik yuz berdi: /start")
    except Exception:
        logger.exception("Prezentatsiya yaratishda xatolik:")
        await status_msg.edit_text("❌ Xatolik yuz berdi. Qayta urinib ko'ring: /start")
    finally:
        for f in (pptx_file, pdf_file):
            if f and os.path.exists(f):
                os.remove(f)


# --- ADMIN PANEL ---
@router.message(Command("admin"))
async def cmd_admin(message: Message):
    if message.from_user.id != ADMIN_ID:
        return

    async with db_pool.acquire() as conn:
        total_users = await conn.fetchval("SELECT COUNT(*) FROM ppt_users")
        pro_users = await conn.fetchval("SELECT COUNT(*) FROM ppt_users WHERE is_pro = TRUE")
        total_ppts = await conn.fetchval("SELECT COALESCE(SUM(total_generated), 0) FROM ppt_users")

    kb = InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="📢 Barchaga xabar yuborish", callback_data="admin:broadcast")]
        ]
    )

    await message.answer(
        "🛠 <b>Admin Panel</b>\n\n"
        f"👥 Jami foydalanuvchilar: <b>{total_users} ta</b>\n"
        f"💎 PRO foydalanuvchilar: <b>{pro_users} ta</b>\n"
        f"📊 Yaratilgan prezentatsiyalar: <b>{total_ppts} ta</b>\n\n"
        "<i>Buyruqlar:</i>\n"
        "<code>/givepro USER_ID</code> — foydalanuvchiga PRO berish\n"
        "<code>/delpro USER_ID</code> — PRO statusini olish",
        parse_mode=ParseMode.HTML,
        reply_markup=kb,
    )


@router.message(Command("givepro"))
async def cmd_give_pro(message: Message):
    if message.from_user.id != ADMIN_ID:
        return
    parts = message.text.split()
    if len(parts) != 2 or not parts[1].isdigit():
        await message.answer("⚠️ Namuna: <code>/givepro 123456789</code>", parse_mode=ParseMode.HTML)
        return

    target_id = int(parts[1])
    async with db_pool.acquire() as conn:
        await conn.execute(
            """
            INSERT INTO ppt_users (user_id, is_pro) VALUES ($1, TRUE)
            ON CONFLICT (user_id) DO UPDATE SET is_pro = TRUE
            """,
            target_id,
        )
    await message.answer(f"✅ <code>{target_id}</code> foydalanuvchiga PRO berildi.", parse_mode=ParseMode.HTML)


@router.message(Command("delpro"))
async def cmd_del_pro(message: Message):
    if message.from_user.id != ADMIN_ID:
        return
    parts = message.text.split()
    if len(parts) != 2 or not parts[1].isdigit():
        await message.answer("⚠️ Namuna: <code>/delpro 123456789</code>", parse_mode=ParseMode.HTML)
        return

    target_id = int(parts[1])
    async with db_pool.acquire() as conn:
        await conn.execute("UPDATE ppt_users SET is_pro = FALSE WHERE user_id = $1", target_id)
    await message.answer(f"✅ <code>{target_id}</code> dan PRO olib tashlandi.", parse_mode=ParseMode.HTML)


@router.callback_query(F.data == "admin:broadcast")
async def start_broadcast(callback: CallbackQuery, state: FSMContext):
    if callback.from_user.id != ADMIN_ID:
        return
    await callback.answer()
    await callback.message.answer("📢 Barcha foydalanuvchilarga yuboriladigan xabarni yozing (bekor qilish uchun /start):")
    await state.set_state(PresStates.waiting_for_broadcast)


@router.message(PresStates.waiting_for_broadcast)
async def send_broadcast(message: Message, state: FSMContext, bot: Bot):
    if message.from_user.id != ADMIN_ID:
        return
    await state.clear()

    async with db_pool.acquire() as conn:
        rows = await conn.fetch("SELECT user_id FROM ppt_users")

    sent, failed = 0, 0
    status = await message.answer(f"⏳ Yuborilmoqda (0/{len(rows)})...")

    for row in rows:
        try:
            await message.copy_to(row["user_id"])
            sent += 1
        except Exception:
            failed += 1
        await asyncio.sleep(0.05)

    await status.edit_text(f"✅ Xabar yuborildi: <b>{sent} ta</b>\n❌ Yetib bormadi: <b>{failed} ta</b>", parse_mode=ParseMode.HTML)


# --- RENDER WEB SERVER ---
async def health_check(request):
    return web.Response(text="Bot is running!")


async def start_web_server():
    app = web.Application()
    app.router.add_get("/", health_check)
    runner = web.AppRunner(app)
    await runner.setup()
    site = web.TCPSite(runner, "0.0.0.0", PORT)
    await site.start()


async def main():
    await init_db()
    bot = Bot(token=BOT_TOKEN)
    dp = Dispatcher()
    dp.include_router(router)

    await start_web_server()
    await bot.delete_webhook(drop_pending_updates=True)
    logger.info("Prezentatsiya boti ishga tushdi...")
    await dp.start_polling(bot)


async def generate_with_fallback(prompt: str):
    last_error = None
    for model in MODELS:
        for idx, api_key in enumerate(API_KEYS, start=1):
            client = genai.Client(api_key=api_key)
            try:
                response = await client.aio.models.generate_content(
                    model=model,
                    contents=prompt,
                )
                return response.text
            except APIError as e:
                last_error = e
                if e.code == 404:
                    # Model deprecated/unavailable -> skip remaining keys for this model
                    break
                if e.code == 429:
                    # Quota hit -> try next project's key, or move to next model
                    await asyncio.sleep(1)
                    continue
                raise
    raise RuntimeError(f"All Gemini models/keys exhausted: {last_error}")

await bot.delete_webhook(drop_pending_updates=True)
await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
