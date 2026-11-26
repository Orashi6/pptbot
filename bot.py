import asyncio
import html
import io
import logging
import os
import re
import subprocess
import tempfile
import uuid
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from typing import Dict, List, Optional, Set, Tuple

import aiohttp
import asyncpg
import docx
from aiogram import Bot, Dispatcher, F, Router
from aiogram.enums import ParseMode
from aiogram.filters import Command, CommandStart, StateFilter
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import (
    BotCommand,
    CallbackQuery,
    FSInputFile,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    KeyboardButton,
    Message,
    ReplyKeyboardMarkup,
    ReplyKeyboardRemove,
)
from aiohttp import web
from google import genai
from google.genai import errors, types
from PIL import Image, ImageOps
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt
from pydantic import BaseModel, Field
from pypdf import PdfReader

# ==============================================================================
# SOZLAMALAR (ENV VARIABLES)
# ==============================================================================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
)
logger = logging.getLogger("ppt_bot")


def _parse_int_env(name: str, default: int) -> int:
    raw = os.getenv(name, "").strip()
    try:
        return int(raw) if raw else default
    except ValueError:
        logger.warning("Noto'g'ri %s=%r, standart qiymat (%d) ishlatiladi.", name, raw, default)
        return default


def _parse_api_keys() -> List[str]:
    """API_KEYS (vergul/nuqta-vergul/yangi qator bilan) va GEMINI_API_KEY ni yig'adi."""
    raw_sources = [
        os.getenv("API_KEYS", ""),
        os.getenv("GEMINI_API_KEY", ""),
    ]
    placeholders = {"YOUR_GEMINI_API_KEY", "YORU API KEYS", "YOUR_API_KEYS", ""}
    keys: List[str] = []
    seen: Set[str] = set()

    for raw in raw_sources:
        for item in re.split(r"[,;\n\r\s]+", raw.strip()):
            cleaned = item.strip().strip("'\"")
            if not cleaned or cleaned in placeholders or cleaned.upper().startswith("YOUR_"):
                continue
            if cleaned not in seen:
                seen.add(cleaned)
                keys.append(cleaned)
    return keys


def _parse_models() -> List[str]:
    default_models = [
        "gemini-3.8-flash",
        "gemini-3.6-flash",
        "gemini-3.5-flash-lite",
        "gemini-2.5-flash",
    ]
    raw = os.getenv("GEMINI_MODELS", "").strip()
    if not raw:
        return default_models
    parsed = [m.strip() for m in raw.split(",") if m.strip()]
    return parsed or default_models


BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()
GEMINI_API_KEYS = _parse_api_keys()
GEMINI_MODELS = _parse_models()
PIXABAY_API_KEY = os.getenv("PIXABAY_API_KEY", "").strip()
PEXELS_API_KEY = os.getenv("PEXELS_API_KEY", "").strip()
DATABASE_URL = os.getenv("DATABASE_URL", "").strip()
ADMIN_ID = _parse_int_env("ADMIN_ID", 705345765)
GIFT_BOT_USERNAME = os.getenv("GIFT_BOT_USERNAME", "giftolishbot").lstrip("@").strip()

WEEKLY_FREE_LIMIT = _parse_int_env("WEEKLY_FREE_LIMIT", 3)
FREE_MAX_SLIDES = _parse_int_env("FREE_MAX_SLIDES", 10)
PRO_MAX_SLIDES = _parse_int_env("PRO_MAX_SLIDES", 20)
MIN_SLIDES = 3
PORT = _parse_int_env("PORT", 8080)

# ==============================================================================
# DIZAYN SHABLONLARI, TILLAR VA USLUBLAR
# ==============================================================================
@dataclass(frozen=True)
class ThemePalette:
    label: str
    template_path: str
    bg_cover: Tuple[int, int, int]
    bg_slide: Tuple[int, int, int]
    card_bg: Tuple[int, int, int]
    primary: Tuple[int, int, int]
    secondary: Tuple[int, int, int]
    cover_title: Tuple[int, int, int]
    cover_subtitle: Tuple[int, int, int]
    slide_title: Tuple[int, int, int]
    body_text: Tuple[int, int, int]
    muted_text: Tuple[int, int, int]
    font_name: str = "DejaVu Sans"


THEMES: Dict[str, ThemePalette] = {
    "dark": ThemePalette(
        label="🌙 Dark Modern",
        template_path="templates/template_dark.pptx",
        bg_cover=(15, 23, 42),       # #0F172A
        bg_slide=(17, 24, 39),       # #111827
        card_bg=(30, 41, 59),        # #1E293B
        primary=(56, 189, 248),      # #38BDF8
        secondary=(129, 140, 248),   # #818CF8
        cover_title=(248, 250, 252),
        cover_subtitle=(148, 163, 184),
        slide_title=(248, 250, 252),
        body_text=(226, 232, 240),
        muted_text=(148, 163, 184),
    ),
    "business": ThemePalette(
        label="💼 Business",
        template_path="templates/template_business.pptx",
        bg_cover=(23, 37, 84),       # #172554
        bg_slide=(248, 250, 252),    # #F8FAFC
        card_bg=(255, 255, 255),
        primary=(30, 58, 138),       # #1E3A8A
        secondary=(217, 119, 6),     # #D97706
        cover_title=(255, 255, 255),
        cover_subtitle=(203, 213, 225),
        slide_title=(15, 23, 42),
        body_text=(51, 65, 85),
        muted_text=(100, 116, 139),
    ),
    "minimal": ThemePalette(
        label="⚪ Minimalist",
        template_path="templates/template_minimal.pptx",
        bg_cover=(250, 250, 250),
        bg_slide=(255, 255, 255),
        card_bg=(244, 244, 245),
        primary=(79, 70, 229),       # #4F46E5
        secondary=(161, 161, 170),
        cover_title=(17, 24, 39),
        cover_subtitle=(82, 82, 91),
        slide_title=(17, 24, 39),
        body_text=(39, 39, 42),
        muted_text=(113, 113, 122),
    ),
    "academic": ThemePalette(
        label="🌿 Akademik",
        template_path="templates/template_academic.pptx",
        bg_cover=(20, 83, 45),       # #14532D
        bg_slide=(250, 249, 245),    # Warm ivory
        card_bg=(255, 255, 255),
        primary=(21, 128, 61),       # #15803D
        secondary=(202, 138, 4),     # #CA8A04
        cover_title=(255, 255, 255),
        cover_subtitle=(220, 252, 231),
        slide_title=(20, 83, 45),
        body_text=(31, 41, 55),
        muted_text=(107, 114, 128),
    ),
    "creative": ThemePalette(
        label="🚀 Kreativ Tech",
        template_path="templates/template_creative.pptx",
        bg_cover=(49, 16, 104),      # Deep violet
        bg_slide=(249, 247, 255),
        card_bg=(255, 255, 255),
        primary=(124, 58, 237),      # #7C3AED
        secondary=(236, 72, 153),    # #EC4899
        cover_title=(255, 255, 255),
        cover_subtitle=(233, 213, 255),
        slide_title=(46, 16, 101),
        body_text=(55, 48, 74),
        muted_text=(124, 115, 150),
    ),
}

LANGUAGES = {
    "uz": (
        "🇺🇿 O'zbek tili",
        "O'zbek tilida (lotin alifbosida, ravon va savodli)",
        "Xulosa va asosiy fikrlar",
        "E'tiboringiz uchun rahmat!",
    ),
    "ru": (
        "🇷🇺 Русский язык",
        "На русском языке (грамотно и структурированно)",
        "Выводы и ключевые итоги",
        "Спасибо за внимание!",
    ),
    "en": (
        "🇬🇧 English",
        "In English (clear, natural, and professional)",
        "Key Takeaways & Conclusion",
        "Thank you for your attention!",
    ),
}

STYLES = {
    "academic": ("🎓 Akademik / Ilmiy", "Ilmiy, tizimli, dalillar va aniq ta'riflarga tayangan uslubda"),
    "business": ("💼 Biznes / Rasmiy", "Rasmiy biznes uslubida, natijaga yo'naltirilgan, aniq metrika va amaliy fikrlar bilan"),
    "creative": ("🎨 Kreativ / Qiziqarli", "Qiziqarli, zamonaviy, jonli misollar va esda qolarli iboralar bilan"),
    "simple": ("⚡ Oddiy / Tushunarli", "Sodda, ravon va har qanday tinglovchi oson tushunadigan uslubda"),
}

router = Router()
db_pool: Optional[asyncpg.Pool] = None
_gemini_clients: Dict[str, genai.Client] = {}
ACTIVE_GENERATIONS: Set[int] = set()


# ==============================================================================
# PYDANTIC SCHEMA
# ==============================================================================
class SlideContent(BaseModel):
    title: str = Field(description="Slayd sarlavhasi (qisqa, 3-7 so'z)")
    points: List[str] = Field(description="Slaydning asosiy fikrlari (3-5 ta aniq va qisqa punkt)")
    image_query: str = Field(
        description="Pixabay/Pexels/Wikimedia orqali sifatli rasm qidirish uchun 2-4 so'zdan iborat INGLIZCHA kalit so'z"
    )
    speaker_notes: str = Field(
        default="",
        description="Ma'ruzachi (taqdimotchi) ushbu slaydda aytishi uchun 2-3 jumlalik qisqa nutq eslatmasi",
    )


class PresentationData(BaseModel):
    main_title: str = Field(description="Prezentatsiyaning asosiy sarlavhasi")
    subtitle: str = Field(description="Muqova slaydi uchun qisqa tavsif (1 jumla)")
    slides: List[SlideContent] = Field(description="Asosiy mazmunli slaydlar ro'yxati")
    conclusion_points: List[str] = Field(
        default_factory=list,
        description="Yakuniy xulosa slaydi uchun 2-3 ta eng muhim xulosa punkti",
    )


# ==============================================================================
# FSM HOLATLARI
# ==============================================================================
class PresStates(StatesGroup):
    waiting_for_source = State()
    waiting_for_count = State()
    waiting_for_lang = State()
    waiting_for_style = State()
    waiting_for_template = State()
    waiting_for_broadcast = State()


# ==============================================================================
# POSTGRESQL BAZA
# ==============================================================================
async def init_db() -> None:
    global db_pool
    if not DATABASE_URL:
        raise RuntimeError("DATABASE_URL environment variable sozlanmagan!")
    dsn = DATABASE_URL.replace("postgres://", "postgresql://", 1)
    db_pool = await asyncpg.create_pool(
        dsn,
        min_size=1,
        max_size=5,
        statement_cache_size=0,
        max_inactive_connection_lifetime=240,
    )

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
    assert db_pool is not None
    val = await db_pool.fetchval("SELECT value FROM bot_settings WHERE key='ppt_pro'")
    return int(val) if val and int(val) > 0 else 50


async def set_ppt_pro_price(price: int) -> None:
    assert db_pool is not None
    async with db_pool.acquire() as conn:
        await conn.execute(
            """
            INSERT INTO bot_settings (key, value) VALUES ('ppt_pro', $1)
            ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value
            """,
            price,
        )


async def check_and_get_user(user_id: int, full_name: str) -> asyncpg.Record:
    assert db_pool is not None
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
        else:
            reset_at = user["week_reset_at"]
            if reset_at is None or now >= reset_at:
                user = await conn.fetchrow(
                    """
                    UPDATE ppt_users
                    SET weekly_used = 0, week_reset_at = $1, full_name = $2
                    WHERE user_id = $3
                    RETURNING *
                    """,
                    now + timedelta(days=7),
                    full_name,
                    user_id,
                )
            elif user["full_name"] != full_name:
                user = await conn.fetchrow(
                    "UPDATE ppt_users SET full_name = $1 WHERE user_id = $2 RETURNING *",
                    full_name,
                    user_id,
                )
        return user


async def increment_usage(user_id: int) -> None:
    assert db_pool is not None
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


def user_max_slides(user: asyncpg.Record, user_id: int) -> int:
    if user["is_pro"] or user_id == ADMIN_ID:
        return PRO_MAX_SLIDES
    return FREE_MAX_SLIDES


# ==============================================================================
# KLAVIATURALAR
# ==============================================================================
def get_count_keyboard(max_slides: int) -> ReplyKeyboardMarkup:
    if max_slides >= 20:
        buttons = [
            [KeyboardButton(text="5"), KeyboardButton(text="7"), KeyboardButton(text="10")],
            [KeyboardButton(text="12"), KeyboardButton(text="15"), KeyboardButton(text="20")],
            [KeyboardButton(text="❌ Bekor qilish")],
        ]
    else:
        buttons = [
            [KeyboardButton(text="5"), KeyboardButton(text="7"), KeyboardButton(text="10")],
            [KeyboardButton(text="❌ Bekor qilish")],
        ]
    return ReplyKeyboardMarkup(keyboard=buttons, resize_keyboard=True, one_time_keyboard=True)


def get_lang_keyboard() -> InlineKeyboardMarkup:
    buttons = [
        [InlineKeyboardButton(text=label, callback_data=f"lang:{key}")]
        for key, (label, _, _, _) in LANGUAGES.items()
    ]
    buttons.append([InlineKeyboardButton(text="❌ Bekor qilish", callback_data="nav:cancel")])
    return InlineKeyboardMarkup(inline_keyboard=buttons)


def get_style_keyboard() -> InlineKeyboardMarkup:
    buttons = [
        [InlineKeyboardButton(text=label, callback_data=f"style:{key}")]
        for key, (label, _) in STYLES.items()
    ]
    buttons.append(
        [
            InlineKeyboardButton(text="⬅️ Orqaga", callback_data="nav:back_to_lang"),
            InlineKeyboardButton(text="❌ Bekor qilish", callback_data="nav:cancel"),
        ]
    )
    return InlineKeyboardMarkup(inline_keyboard=buttons)


def get_template_keyboard() -> InlineKeyboardMarkup:
    buttons = [
        [InlineKeyboardButton(text=theme.label, callback_data=f"tpl:{key}")]
        for key, theme in THEMES.items()
    ]
    buttons.append(
        [
            InlineKeyboardButton(text="⬅️ Orqaga", callback_data="nav:back_to_style"),
            InlineKeyboardButton(text="❌ Bekor qilish", callback_data="nav:cancel"),
        ]
    )
    return InlineKeyboardMarkup(inline_keyboard=buttons)


async def get_pro_keyboard(user_id: int) -> InlineKeyboardMarkup:
    stars = await get_ppt_pro_price()
    pay_url = f"https://t.me/{GIFT_BOT_USERNAME}?start=pay_ppt_pro_{user_id}"
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text=f"⭐ PRO obuna olish — {stars} Stars", url=pay_url)]
        ]
    )


# ==============================================================================
# FAYLDAN MATN AJRATIB OLISH (.PDF / .DOCX / .TXT)
# ==============================================================================
def extract_text_from_bytes(file_bytes: bytes, filename: str) -> str:
    ext = filename.lower().rsplit(".", 1)[-1] if "." in filename else ""
    text_parts: List[str] = []

    if ext == "pdf":
        reader = PdfReader(io.BytesIO(file_bytes))
        for page in reader.pages[:20]:
            extracted = page.extract_text()
            if extracted and extracted.strip():
                text_parts.append(extracted.strip())
    elif ext == "docx":
        document = docx.Document(io.BytesIO(file_bytes))
        for para in document.paragraphs:
            if para.text.strip():
                text_parts.append(para.text.strip())
        for table in document.tables:
            for row in table.rows:
                row_cells = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                if row_cells:
                    text_parts.append(" | ".join(row_cells))
    elif ext == "txt":
        for encoding in ("utf-8", "cp1251", "latin-1"):
            try:
                text_parts.append(file_bytes.decode(encoding))
                break
            except UnicodeDecodeError:
                continue

    full_text = "\n".join(text_parts).strip()
    return full_text[:18000]


# ==============================================================================
# GEMINI AI GENERATSIYA (MULTI-KEY & MULTI-MODEL FALLBACK)
# ==============================================================================
def _get_gemini_client(api_key: str) -> genai.Client:
    client = _gemini_clients.get(api_key)
    if client is None:
        client = genai.Client(api_key=api_key)
        _gemini_clients[api_key] = client
    return client


def _is_invalid_key_error(err: errors.APIError) -> bool:
    code = getattr(err, "code", None)
    msg = str(err).upper()
    if code in (401, 403):
        return True
    if code == 400 and ("API_KEY" in msg or "API KEY" in msg or "PERMISSION" in msg):
        return True
    return False


async def generate_with_fallback(contents: str, config: types.GenerateContentConfig):
    if not GEMINI_API_KEYS:
        raise RuntimeError("Gemini API kalitlari topilmadi! API_KEYS yoki GEMINI_API_KEY ni sozlang.")

    last_error: Optional[Exception] = None
    bad_keys: Set[str] = set()

    for model_name in GEMINI_MODELS:
        model_not_found = False
        for key_idx, api_key in enumerate(GEMINI_API_KEYS, start=1):
            if api_key in bad_keys:
                continue

            client = _get_gemini_client(api_key)
            for attempt in range(2):
                try:
                    response = await client.aio.models.generate_content(
                        model=model_name,
                        contents=contents,
                        config=config,
                    )
                    if response and (getattr(response, "parsed", None) or getattr(response, "text", None)):
                        return response
                    raise RuntimeError(f"Gemini bo'sh javob qaytardi (model={model_name})")
                except errors.APIError as e:
                    last_error = e
                    code = getattr(e, "code", None)
                    logger.warning(
                        "Gemini APIError | model=%s | key=#%d | attempt=%d | code=%s | %s",
                        model_name,
                        key_idx,
                        attempt + 1,
                        code,
                        e,
                    )
                    if code == 404:
                        model_not_found = True
                        break
                    if _is_invalid_key_error(e):
                        bad_keys.add(api_key)
                        break
                    if code == 429:
                        await asyncio.sleep(1.0)
                        break  # Keyingi kalitga (yoki modelga) o'tamiz
                    if code in (500, 502, 503, 504):
                        await asyncio.sleep(2.0 * (attempt + 1))
                        continue
                    break
                except Exception as e:
                    last_error = e
                    logger.warning(
                        "Gemini kutilmagan xato | model=%s | key=#%d | %s",
                        model_name,
                        key_idx,
                        e,
                    )
                    await asyncio.sleep(1.0)
                    break

            if model_not_found:
                break

    if isinstance(last_error, errors.ServerError) and getattr(last_error, "code", None) == 503:
        raise last_error
    raise RuntimeError(f"Barcha Gemini modellar/kalitlarda xatolik yuz berdi: {last_error}")


def _extract_json_payload(raw_text: str) -> str:
    text = raw_text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.IGNORECASE)
        text = re.sub(r"\s*```$", "", text)
    return text.strip()


async def generate_presentation_content(
    source_text: str,
    slide_count: int,
    lang_instruction: str,
    style_instruction: str,
    is_document: bool = False,
) -> PresentationData:
    # Muqova (1) + Mazmunli slaydlar (slide_count - 2) + Xulosa slaydi (1) = slide_count
    content_slides_count = max(1, slide_count - 2)

    if is_document:
        task_desc = (
            "Quyidagi hujjat/matn mazmunini chuqur tahlil qilib, uning eng muhim g'oyalaridan "
            "professional prezentatsiya tuzib ber:\n\n"
            f"--- MATN BOSHI ---\n{source_text}\n--- MATN OXIRI ---"
        )
    else:
        task_desc = f"Mavzu: '{source_text}'. Ushbu mavzu bo'yicha professional prezentatsiya tayyorla."

    prompt = (
        f"{task_desc}\n\n"
        "Qat'iy talablar:\n"
        f"1. Til: {lang_instruction}.\n"
        f"2. Uslub: {style_instruction}.\n"
        f"3. 'slides' ro'yxatida AYNAN {content_slides_count} ta mazmunli slayd bo'lsin "
        f"(muqova va yakuniy xulosa slaydi bilan jami {slide_count} ta slayd chiqadi).\n"
        "4. Har bir slaydda 3-4 ta aniq, mazmunli va qisqa punkt bo'lsin (har bir punkt 10-16 so'z atrofida).\n"
        "5. Har bir slayd uchun 'image_query' maydoniga rasm topishga qulay 2-4 so'zdan iborat INGLIZCHA kalit so'z yoz.\n"
        "6. Har bir slayd uchun 'speaker_notes' maydoniga ma'ruzachi og'zaki aytishi uchun 2 ta qisqa jumla yoz.\n"
        "7. 'conclusion_points' maydoniga prezentatsiyani yakunlovchi 3 ta eng muhim xulosa punktini yoz."
    )

    config = types.GenerateContentConfig(
        response_mime_type="application/json",
        response_schema=PresentationData,
        temperature=0.6,
    )

    last_parse_err: Optional[Exception] = None
    for _ in range(2):
        response = await generate_with_fallback(prompt, config)
        parsed = getattr(response, "parsed", None)
        if isinstance(parsed, PresentationData):
            pres = parsed
        else:
            raw_text = getattr(response, "text", "") or ""
            try:
                pres = PresentationData.model_validate_json(_extract_json_payload(raw_text))
            except Exception as err:
                last_parse_err = err
                logger.warning("Gemini JSON parse xatosi, qayta urinilmoqda: %s", err)
                continue

        if pres.slides:
            # Slaydlar sonini kerakli chegarada ushlash
            pres.slides = pres.slides[:content_slides_count]
            return pres

    raise RuntimeError(f"AI javobini tahlil qilib bo'lmadi: {last_parse_err}")


# ==============================================================================
# RASMLARNI YUKLASH VA PILLOW BILAN OPTIMALLASHTIRISH
# ==============================================================================
def process_image_bytes(
    raw_bytes: bytes,
    target_width: int = 1040,
    target_height: int = 780,
) -> Optional[io.BytesIO]:
    """Rasmni tekshiradi, orientatsiyasini to'g'rilaydi va aniq 4:3 nisbatda JPEG ga o'giradi."""
    try:
        with Image.open(io.BytesIO(raw_bytes)) as img:
            img = ImageOps.exif_transpose(img)
            if img.width < 220 or img.height < 160:
                return None

            if img.mode in ("RGBA", "LA", "P"):
                rgba = img.convert("RGBA")
                bg = Image.new("RGB", rgba.size, (255, 255, 255))
                bg.paste(rgba, mask=rgba.split()[-1])
                img = bg
            elif img.mode != "RGB":
                img = img.convert("RGB")

            fitted = ImageOps.fit(
                img,
                (target_width, target_height),
                method=Image.Resampling.LANCZOS,
                centering=(0.5, 0.5),
            )
            out = io.BytesIO()
            fitted.save(out, format="JPEG", quality=88, optimize=True)
            out.seek(0)
            return out
    except Exception:
        return None


async def _download_valid_image(
    session: aiohttp.ClientSession,
    img_url: str,
    headers: Optional[Dict[str, str]] = None,
) -> Optional[io.BytesIO]:
    try:
        async with session.get(
            img_url,
            headers=headers,
            timeout=aiohttp.ClientTimeout(total=15),
        ) as resp:
            if resp.status != 200:
                return None
            content_type = resp.headers.get("Content-Type", "").lower()
            if "image" not in content_type or "svg" in content_type:
                return None
            raw = await resp.read()
            if len(raw) > 12 * 1024 * 1024:
                return None
            return await asyncio.to_thread(process_image_bytes, raw)
    except Exception:
        return None


async def fetch_pixabay_image(
    session: aiohttp.ClientSession,
    query: str,
    used_urls: Set[str],
) -> Optional[io.BytesIO]:
    if not PIXABAY_API_KEY:
        return None
    url = "https://pixabay.com/api/"
    params = {
        "key": PIXABAY_API_KEY,
        "q": query[:100],
        "image_type": "photo",
        "orientation": "horizontal",
        "per_page": 6,
        "safesearch": "true",
    }
    try:
        async with session.get(url, params=params, timeout=aiohttp.ClientTimeout(total=10)) as resp:
            if resp.status != 200:
                return None
            data = await resp.json()
            for hit in data.get("hits", []):
                img_url = hit.get("largeImageURL") or hit.get("webformatURL")
                if not img_url or img_url in used_urls:
                    continue
                img = await _download_valid_image(session, img_url)
                if img:
                    used_urls.add(img_url)
                    return img
    except Exception:
        return None
    return None


async def fetch_pexels_image(
    session: aiohttp.ClientSession,
    query: str,
    used_urls: Set[str],
) -> Optional[io.BytesIO]:
    if not PEXELS_API_KEY:
        return None
    url = "https://api.pexels.com/v1/search"
    headers = {"Authorization": PEXELS_API_KEY}
    params = {"query": query[:100], "orientation": "landscape", "per_page": 6}
    try:
        async with session.get(
            url,
            headers=headers,
            params=params,
            timeout=aiohttp.ClientTimeout(total=10),
        ) as resp:
            if resp.status != 200:
                return None
            data = await resp.json()
            for photo in data.get("photos", []):
                img_url = photo.get("src", {}).get("large") or photo.get("src", {}).get("medium")
                if not img_url or img_url in used_urls:
                    continue
                img = await _download_valid_image(session, img_url)
                if img:
                    used_urls.add(img_url)
                    return img
    except Exception:
        return None
    return None


async def fetch_wikimedia_image(
    session: aiohttp.ClientSession,
    query: str,
    used_urls: Set[str],
) -> Optional[io.BytesIO]:
    url = "https://commons.wikimedia.org/w/api.php"
    headers = {"User-Agent": "TelegramPresentationBot/2.0 (https://t.me/giftolishbot)"}
    params = {
        "action": "query",
        "format": "json",
        "generator": "search",
        "gsrsearch": f"filetype:bitmap {query}",
        "gsrnamespace": 6,
        "gsrlimit": 5,
        "prop": "imageinfo",
        "iiprop": "url|mime",
        "iiurlwidth": 1200,
    }
    try:
        async with session.get(
            url,
            headers=headers,
            params=params,
            timeout=aiohttp.ClientTimeout(total=10),
        ) as resp:
            if resp.status != 200:
                return None
            data = await resp.json()
            pages = data.get("query", {}).get("pages", {})
            for page in pages.values():
                info = (page.get("imageinfo") or [{}])[0]
                mime = info.get("mime", "").lower()
                if mime and not mime.startswith(("image/jpeg", "image/png", "image/webp")):
                    continue
                img_url = info.get("thumburl") or info.get("url")
                if not img_url or img_url in used_urls:
                    continue
                img = await _download_valid_image(session, img_url, headers=headers)
                if img:
                    used_urls.add(img_url)
                    return img
    except Exception:
        return None
    return None


async def fetch_slide_image(
    session: aiohttp.ClientSession,
    query: str,
    used_urls: Set[str],
    sem: asyncio.Semaphore,
) -> Optional[io.BytesIO]:
    clean_query = (query or "").strip()
    if not clean_query:
        return None
    async with sem:
        for fetcher in (fetch_pixabay_image, fetch_pexels_image, fetch_wikimedia_image):
            img = await fetcher(session, clean_query, used_urls)
            if img is not None:
                return img
        # Agar murakkab so'rov topilmasa, birinchi 2 ta so'z bilan qayta urinamiz
        short_q = " ".join(clean_query.split()[:2])
        if short_q and short_q != clean_query:
            for fetcher in (fetch_pixabay_image, fetch_pexels_image, fetch_wikimedia_image):
                img = await fetcher(session, short_q, used_urls)
                if img is not None:
                    return img
    return None


async def fetch_all_slide_images(slides: List[SlideContent]) -> List[Optional[io.BytesIO]]:
    used_urls: Set[str] = set()
    sem = asyncio.Semaphore(4)
    async with aiohttp.ClientSession() as session:
        tasks = [fetch_slide_image(session, slide.image_query, used_urls, sem) for slide in slides]
        return list(await asyncio.gather(*tasks))


# ==============================================================================
# .PPTX YIG'ISH (ZAMONAVIY DIZAYN + SHABLON QO'LLAB-QUVVATLASH)
# ==============================================================================
def calculate_optimal_font_size(points: List[str], has_image: bool) -> Pt:
    total_chars = sum(len(p) for p in points)
    count = len(points)
    if has_image:
        if total_chars > 380 or count >= 5:
            return Pt(14)
        if total_chars > 270:
            return Pt(16)
        if total_chars > 180:
            return Pt(17.5)
        return Pt(19)
    else:
        if total_chars > 460 or count >= 5:
            return Pt(16)
        if total_chars > 320:
            return Pt(18)
        return Pt(20)


def _set_slide_bg(slide, color: Tuple[int, int, int]) -> None:
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(*color)


def _add_rect(slide, left, top, width, height, color: Tuple[int, int, int], rounded: bool = False):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(shape_type, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(*color)
    shape.line.fill.background()
    if rounded and shape.adjustments:
        shape.adjustments[0] = 0.04
    return shape


def _set_notes(slide, notes_text: str) -> None:
    if not notes_text:
        return
    try:
        notes_slide = slide.notes_slide
        tf = notes_slide.notes_text_frame
        tf.text = notes_text.strip()
    except Exception:
        pass


def _build_styled_pptx(
    data: PresentationData,
    images: List[Optional[io.BytesIO]],
    theme: ThemePalette,
    lang_key: str,
    style_label: str,
    author_name: str,
    output_path: str,
) -> None:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    _, _, conclusion_heading, thank_you_text = LANGUAGES[lang_key]
    total_slides = len(data.slides) + 2  # Cover + Content + Closing
    today_str = datetime.now().strftime("%d.%m.%Y")

    # ------------------------------------------------------------------
    # 1. MUQOVA SLAYDI (COVER SLIDE)
    # ------------------------------------------------------------------
    cover = prs.slides.add_slide(blank_layout)
    _set_slide_bg(cover, theme.bg_cover)

    # Chap vertikal aksent chizig'i va o'ng bezak bloklari
    _add_rect(cover, Inches(0), Inches(0), Inches(0.32), Inches(7.5), theme.primary)
    _add_rect(cover, Inches(10.8), Inches(0), Inches(2.533), Inches(7.5), theme.primary)
    _add_rect(cover, Inches(10.45), Inches(0), Inches(0.35), Inches(7.5), theme.secondary)

    # Yuqori uslub nishoni (badge)
    badge = _add_rect(cover, Inches(0.95), Inches(1.0), Inches(3.2), Inches(0.48), theme.primary, rounded=True)
    badge_tf = badge.text_frame
    badge_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    badge_p = badge_tf.paragraphs[0]
    badge_p.alignment = PP_ALIGN.CENTER
    badge_p.text = re.sub(r"^[^\w]+", "", style_label).strip().upper()
    badge_p.font.name = theme.font_name
    badge_p.font.size = Pt(11)
    badge_p.font.bold = True
    badge_p.font.color.rgb = RGBColor(255, 255, 255)

    # Asosiy sarlavha
    title_box = cover.shapes.add_textbox(Inches(0.95), Inches(1.85), Inches(9.0), Inches(2.4))
    title_tf = title_box.text_frame
    title_tf.word_wrap = True
    tp = title_tf.paragraphs[0]
    tp.text = data.main_title
    tp.font.name = theme.font_name
    tp.font.bold = True
    tp.font.size = Pt(34 if len(data.main_title) > 55 else 42)
    tp.font.color.rgb = RGBColor(*theme.cover_title)

    # Ajratuvchi aksent chiziq
    _add_rect(cover, Inches(0.95), Inches(4.45), Inches(2.1), Inches(0.08), theme.secondary)

    # Tavsif (subtitle)
    sub_box = cover.shapes.add_textbox(Inches(0.95), Inches(4.75), Inches(8.8), Inches(1.3))
    sub_tf = sub_box.text_frame
    sub_tf.word_wrap = True
    sp = sub_tf.paragraphs[0]
    sp.text = data.subtitle
    sp.font.name = theme.font_name
    sp.font.size = Pt(19)
    sp.font.color.rgb = RGBColor(*theme.cover_subtitle)

    # Pastki muallif va sana
    footer_meta = f"{author_name}  •  {today_str}" if author_name else today_str
    meta_box = cover.shapes.add_textbox(Inches(0.95), Inches(6.55), Inches(8.5), Inches(0.45))
    meta_tf = meta_box.text_frame
    mp = meta_tf.paragraphs[0]
    mp.text = footer_meta
    mp.font.name = theme.font_name
    mp.font.size = Pt(13)
    mp.font.color.rgb = RGBColor(*theme.cover_subtitle)

    # ------------------------------------------------------------------
    # 2. MAZMUNLI SLAYDLAR (CONTENT SLIDES)
    # ------------------------------------------------------------------
    short_main_title = data.main_title[:48] + ("..." if len(data.main_title) > 48 else "")

    for idx, (slide_info, img_stream) in enumerate(zip(data.slides, images), start=2):
        slide = prs.slides.add_slide(blank_layout)
        _set_slide_bg(slide, theme.bg_slide)

        # Yuqori aksent chiziq
        _add_rect(slide, Inches(0), Inches(0), Inches(13.333), Inches(0.12), theme.primary)

        # Slayd raqami nishoni
        num_badge = _add_rect(slide, Inches(0.65), Inches(0.45), Inches(0.68), Inches(0.55), theme.primary, rounded=True)
        ntf = num_badge.text_frame
        ntf.vertical_anchor = MSO_ANCHOR.MIDDLE
        np = ntf.paragraphs[0]
        np.alignment = PP_ALIGN.CENTER
        np.text = f"{idx - 1:02d}"
        np.font.name = theme.font_name
        np.font.bold = True
        np.font.size = Pt(16)
        np.font.color.rgb = RGBColor(255, 255, 255)

        # Slayd sarlavhasi
        st_box = slide.shapes.add_textbox(Inches(1.5), Inches(0.4), Inches(11.1), Inches(0.75))
        st_tf = st_box.text_frame
        st_tf.word_wrap = True
        st_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        stp = st_tf.paragraphs[0]
        stp.text = slide_info.title
        stp.font.name = theme.font_name
        stp.font.bold = True
        stp.font.size = Pt(24 if len(slide_info.title) > 48 else 27)
        stp.font.color.rgb = RGBColor(*theme.slide_title)

        # Sarlavha ostidagi yupqa chiziq
        _add_rect(slide, Inches(0.65), Inches(1.22), Inches(12.03), Inches(0.03), theme.secondary)

        has_image = img_stream is not None
        image_on_left = has_image and (idx % 2 == 1)

        if has_image:
            if image_on_left:
                img_left, card_left = Inches(0.65), Inches(6.25)
                img_width, card_width = Inches(5.25), Inches(6.43)
            else:
                card_left, img_left = Inches(0.65), Inches(7.43)
                card_width, img_width = Inches(6.43), Inches(5.25)

            img_top = Inches(1.68)
            img_height = Inches(3.94)  # 4:3 nisbat (5.25 / 3.94 = 1.332)

            # Rasm orqasidagi nafis soya/ramka
            _add_rect(
                slide,
                img_left + Inches(0.1),
                img_top + Inches(0.1),
                img_width,
                img_height,
                theme.secondary,
                rounded=False,
            )
            try:
                img_stream.seek(0)
                slide.shapes.add_picture(
                    img_stream,
                    left=img_left,
                    top=img_top,
                    width=img_width,
                    height=img_height,
                )
            except Exception as img_err:
                logger.warning("Rasm joylashtirishda xatolik (slayd %d): %s", idx, img_err)
        else:
            card_left = Inches(0.65)
            card_width = Inches(12.03)

        # Matn kartasi
        card_top = Inches(1.58)
        card_height = Inches(5.15)
        _add_rect(slide, card_left, card_top, card_width, card_height, theme.card_bg, rounded=True)
        # Karta chap aksent chizig'i
        _add_rect(slide, card_left, card_top + Inches(0.25), Inches(0.09), card_height - Inches(0.5), theme.primary)

        # Punktlar matni
        body_box = slide.shapes.add_textbox(
            card_left + Inches(0.35),
            card_top + Inches(0.3),
            card_width - Inches(0.6),
            card_height - Inches(0.55),
        )
        tf = body_box.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        font_size = calculate_optimal_font_size(slide_info.points, has_image=has_image)

        for p_idx, point in enumerate(slide_info.points):
            p = tf.paragraphs[0] if p_idx == 0 else tf.add_paragraph()
            p.space_after = Pt(14 if len(slide_info.points) <= 4 else 10)
            p.line_spacing = 1.2

            bullet_run = p.add_run()
            bullet_run.text = "▪  "
            bullet_run.font.name = theme.font_name
            bullet_run.font.bold = True
            bullet_run.font.size = font_size
            bullet_run.font.color.rgb = RGBColor(*theme.primary)

            text_run = p.add_run()
            text_run.text = point.strip()
            text_run.font.name = theme.font_name
            text_run.font.size = font_size
            text_run.font.color.rgb = RGBColor(*theme.body_text)

        # Pastki footer (chapda sarlavha, o'ngda sahifa raqami)
        f_left = slide.shapes.add_textbox(Inches(0.65), Inches(6.95), Inches(9.5), Inches(0.35))
        flp = f_left.text_frame.paragraphs[0]
        flp.text = short_main_title
        flp.font.name = theme.font_name
        flp.font.size = Pt(10)
        flp.font.color.rgb = RGBColor(*theme.muted_text)

        f_right = slide.shapes.add_textbox(Inches(11.3), Inches(6.95), Inches(1.38), Inches(0.35))
        frp = f_right.text_frame.paragraphs[0]
        frp.alignment = PP_ALIGN.RIGHT
        frp.text = f"{idx} / {total_slides}"
        frp.font.name = theme.font_name
        frp.font.bold = True
        frp.font.size = Pt(10.5)
        frp.font.color.rgb = RGBColor(*theme.muted_text)

        # Ma'ruzachi eslatmalari (Speaker Notes)
        _set_notes(slide, slide_info.speaker_notes)

    # ------------------------------------------------------------------
    # 3. YAKUNIY XULOSA VA RAHMAT SLAYDI (CLOSING SLIDE)
    # ------------------------------------------------------------------
    closing = prs.slides.add_slide(blank_layout)
    _set_slide_bg(closing, theme.bg_cover)
    _add_rect(closing, Inches(0), Inches(0), Inches(13.333), Inches(0.16), theme.primary)
    _add_rect(closing, Inches(0), Inches(7.34), Inches(13.333), Inches(0.16), theme.secondary)

    # Sarlavha: Xulosa
    ch_box = closing.shapes.add_textbox(Inches(0.95), Inches(0.75), Inches(11.4), Inches(0.85))
    chp = ch_box.text_frame.paragraphs[0]
    chp.text = conclusion_heading
    chp.font.name = theme.font_name
    chp.font.bold = True
    chp.font.size = Pt(32)
    chp.font.color.rgb = RGBColor(*theme.cover_title)

    _add_rect(closing, Inches(0.95), Inches(1.62), Inches(2.2), Inches(0.07), theme.secondary)

    conclusions = data.conclusion_points or [
        s.title for s in data.slides[:3]
    ]
    c_box = closing.shapes.add_textbox(Inches(0.95), Inches(2.0), Inches(11.4), Inches(3.2))
    c_tf = c_box.text_frame
    c_tf.word_wrap = True
    for c_idx, point in enumerate(conclusions[:4]):
        p = c_tf.paragraphs[0] if c_idx == 0 else c_tf.add_paragraph()
        p.space_after = Pt(16)
        p.line_spacing = 1.25

        r1 = p.add_run()
        r1.text = "✓  "
        r1.font.name = theme.font_name
        r1.font.bold = True
        r1.font.size = Pt(20)
        r1.font.color.rgb = RGBColor(*theme.primary)

        r2 = p.add_run()
        r2.text = point.strip()
        r2.font.name = theme.font_name
        r2.font.size = Pt(19)
        r2.font.color.rgb = RGBColor(*theme.cover_title)

    # Pastki "Rahmat!" bloki
    ty_box = closing.shapes.add_textbox(Inches(0.95), Inches(5.65), Inches(11.4), Inches(0.9))
    typ = ty_box.text_frame.paragraphs[0]
    typ.text = thank_you_text
    typ.font.name = theme.font_name
    typ.font.bold = True
    typ.font.size = Pt(28)
    typ.font.color.rgb = RGBColor(*theme.secondary)

    prs.save(output_path)


def build_pptx_with_template(
    data: PresentationData,
    images: List[Optional[io.BytesIO]],
    tpl_key: str,
    lang_key: str,
    style_label: str,
    author_name: str,
    output_path: str,
) -> None:
    theme = THEMES[tpl_key]

    # Agar maxsus .pptx shablon fayli diskda mavjud bo'lsa, undan foydalanamiz
    if os.path.exists(theme.template_path):
        prs = Presentation(theme.template_path)
        for i in range(len(prs.slides) - 1, -1, -1):
            rId = prs.slides._sldIdLst[i].rId
            prs.part.drop_rel(rId)
            del prs.slides._sldIdLst[i]

        cover = prs.slides.add_slide(prs.slide_layouts[0])
        if cover.shapes.title:
            cover.shapes.title.text = data.main_title
        for shape in cover.placeholders:
            if shape.placeholder_format.idx == 1:
                shape.text = data.subtitle
                break

        content_layout = prs.slide_layouts[1] if len(prs.slide_layouts) > 1 else prs.slide_layouts[0]
        for slide_info, img_stream in zip(data.slides, images):
            slide = prs.slides.add_slide(content_layout)
            if slide.shapes.title:
                slide.shapes.title.text = slide_info.title

            body_shape = next(
                (s for s in slide.placeholders if s.placeholder_format.idx == 1),
                None,
            )
            if body_shape:
                if img_stream:
                    body_shape.left = Inches(0.7)
                    body_shape.width = Inches(6.3)
                tf = body_shape.text_frame
                tf.clear()
                tf.word_wrap = True
                font_size = calculate_optimal_font_size(slide_info.points, has_image=bool(img_stream))
                for i, point in enumerate(slide_info.points):
                    p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                    p.text = point
                    p.level = 0
                    p.font.size = font_size

            if img_stream:
                try:
                    img_stream.seek(0)
                    slide.shapes.add_picture(
                        img_stream,
                        left=Inches(7.4),
                        top=Inches(1.8),
                        width=Inches(5.2),
                        height=Inches(3.9),
                    )
                except Exception as img_err:
                    logger.warning("Shablonda rasm qo'yishda xato: %s", img_err)

            _set_notes(slide, slide_info.speaker_notes)

        # Xulosa slaydi
        _, _, conclusion_heading, thank_you_text = LANGUAGES[lang_key]
        closing = prs.slides.add_slide(content_layout)
        if closing.shapes.title:
            closing.shapes.title.text = conclusion_heading
        body_shape = next((s for s in closing.placeholders if s.placeholder_format.idx == 1), None)
        if body_shape:
            tf = body_shape.text_frame
            tf.clear()
            for i, point in enumerate((data.conclusion_points or []) + [thank_you_text]):
                p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                p.text = point
                p.font.size = Pt(18)

        prs.save(output_path)
        return

    # Aks holda — o'rnatilgan zamonaviy dasturiy dizayn bilan chizamiz
    _build_styled_pptx(
        data=data,
        images=images,
        theme=theme,
        lang_key=lang_key,
        style_label=style_label,
        author_name=author_name,
        output_path=output_path,
    )


# ==============================================================================
# .PPTX -> .PDF KONVERTATSIYA (PARALLEL ISHLASHGA XAVFSIZ)
# ==============================================================================
def convert_pptx_to_pdf(pptx_path: str, output_dir: str) -> Optional[str]:
    profile_dir = os.path.join(output_dir, "lo_profile")
    try:
        subprocess.run(
            [
                "soffice",
                f"-env:UserInstallation=file://{profile_dir}",
                "--headless",
                "--norestore",
                "--convert-to",
                "pdf",
                pptx_path,
                "--outdir",
                output_dir,
            ],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            timeout=45,
        )
        pdf_path = pptx_path.rsplit(".", 1)[0] + ".pdf"
        return pdf_path if os.path.exists(pdf_path) else None
    except Exception as e:
        logger.warning("PDF konvertatsiya amalga oshmadi: %s", e)
        return None


# ==============================================================================
# FOYDALANUVCHI KOMANDALARI
# ==============================================================================
@router.message(CommandStart())
async def cmd_start(message: Message, state: FSMContext):
    await state.clear()
    user = await check_and_get_user(message.from_user.id, message.from_user.full_name)
    first_name = html.escape(message.from_user.first_name or "Foydalanuvchi")

    if user["is_pro"] or message.from_user.id == ADMIN_ID:
        status_text = "💎 Tarif: <b>PRO (Cheksiz, 20 tagacha slayd)</b>"
    else:
        left = max(0, WEEKLY_FREE_LIMIT - user["weekly_used"])
        status_text = f"🆓 Haftalik bepul limit: <b>{left}/{WEEKLY_FREE_LIMIT} ta qoldi</b>"

    await message.answer(
        f"👋 Salom, <b>{first_name}</b>!\n"
        f"{status_text}\n\n"
        "📌 <b>Prezentatsiya mavzusini yozing</b>, uzun matn yuboring yoki "
        "<b>.pdf / .docx / .txt</b> fayl tashlang:\n\n"
        "✨ <i>Imkoniyatlar: 5 xil zamonaviy dizayn, avtomatik rasmlar, "
        "ma'ruzachi eslatmalari va PDF nusxa!</i>\n"
        "<i>💎 PRO obuna: /pro | 👤 Profil: /profile | ❓ Yordam: /help</i>",
        parse_mode=ParseMode.HTML,
        reply_markup=ReplyKeyboardRemove(),
    )
    await state.set_state(PresStates.waiting_for_source)


@router.message(Command("cancel"))
@router.message(F.text == "❌ Bekor qilish")
async def cmd_cancel(message: Message, state: FSMContext):
    await state.clear()
    await message.answer(
        "❌ Amal bekor qilindi.\n\nYangi prezentatsiya yaratish uchun mavzu yozing yoki /start bosing.",
        reply_markup=ReplyKeyboardRemove(),
    )


@router.message(Command("help"))
async def cmd_help(message: Message):
    await message.answer(
        "❓ <b>Botdan foydalanish yo'riqnomasi:</b>\n\n"
        "1️⃣ Mavzu yozing (masalan: <i>Sun'iy intellekt kelajagi</i>), uzun maqola tashlang "
        "yoki <b>.pdf / .docx / .txt</b> hujjat yuboring.\n"
        f"2️⃣ Slaydlar sonini tanlang (Bepul: {MIN_SLIDES}–{FREE_MAX_SLIDES}, PRO: {MIN_SLIDES}–{PRO_MAX_SLIDES}).\n"
        "3️⃣ Tilni tanlang (🇺🇿 O'zbek, 🇷🇺 Rus, 🇬🇧 Ingliz).\n"
        "4️⃣ Taqdimot uslubini tanlang (Akademik, Biznes, Kreativ, Oddiy).\n"
        "5️⃣ Dizayn shablonini tanlang — bot <b>.pptx</b> va <b>.pdf</b> fayllarni tayyorlab beradi!\n\n"
        "💡 <b>Maslahat:</b> PowerPoint fayl ichida har bir slayd uchun "
        "<i>Ma'ruzachi eslatmalari (Speaker Notes)</i> ham qo'shiladi.\n\n"
        "Komandalar: /start • /profile • /pro • /cancel",
        parse_mode=ParseMode.HTML,
    )


@router.message(Command("profile"))
async def cmd_profile(message: Message):
    user = await check_and_get_user(message.from_user.id, message.from_user.full_name)
    full_name = html.escape(message.from_user.full_name or "Foydalanuvchi")
    is_pro = user["is_pro"] or message.from_user.id == ADMIN_ID

    if is_pro:
        tariff_line = "💎 Tarif: <b>PRO (Cheksiz)</b>"
        limit_line = f"📑 Maksimal slayd: <b>{PRO_MAX_SLIDES} ta</b>"
    else:
        left = max(0, WEEKLY_FREE_LIMIT - user["weekly_used"])
        reset_dt = user["week_reset_at"]
        reset_str = reset_dt.strftime("%d.%m.%Y") if reset_dt else "7 kun ichida"
        tariff_line = "🆓 Tarif: <b>Bepul</b>"
        limit_line = (
            f"📊 Haftalik limit: <b>{left}/{WEEKLY_FREE_LIMIT} ta qoldi</b> "
            f"(yangilanish: {reset_str})\n"
            f"📑 Maksimal slayd: <b>{FREE_MAX_SLIDES} ta</b>"
        )

    kb = None if is_pro else await get_pro_keyboard(message.from_user.id)
    await message.answer(
        f"👤 <b>Sizning profilingiz</b>\n\n"
        f"🏷 Ism: <b>{full_name}</b>\n"
        f"🆔 ID: <code>{message.from_user.id}</code>\n"
        f"{tariff_line}\n"
        f"{limit_line}\n"
        f"🎯 Jami yaratilgan prezentatsiyalar: <b>{user['total_generated']} ta</b>",
        parse_mode=ParseMode.HTML,
        reply_markup=kb,
    )


@router.message(Command("pro"))
async def cmd_pro(message: Message):
    kb = await get_pro_keyboard(message.from_user.id)
    await message.answer(
        "⭐ <b>PRO obuna imkoniyatlari:</b>\n\n"
        f"• Haftalik {WEEKLY_FREE_LIMIT} talik cheklov olib tashlanadi (Cheksiz yaratish)\n"
        f"• {PRO_MAX_SLIDES} tagacha slayd yaratish imkoniyati (Bepulda {FREE_MAX_SLIDES} tagacha)\n"
        "• Navbatsiz tezkor generatsiya\n\n"
        f"💳 To'lov rasmiy @{GIFT_BOT_USERNAME} orqali Telegram Stars bilan amalga oshiriladi va avtomatik faollashadi:",
        parse_mode=ParseMode.HTML,
        reply_markup=kb,
    )


# ==============================================================================
# PREZENTATSIYA YARATISH OQIMI (SOURCE -> COUNT -> LANG -> STYLE -> TEMPLATE)
# ==============================================================================
async def _handle_document_input(message: Message, state: FSMContext, bot: Bot) -> None:
    doc = message.document
    filename = doc.file_name or ""
    if not filename.lower().endswith((".pdf", ".docx", ".txt")):
        await message.answer(
            "⚠️ Faqat <b>.pdf</b>, <b>.docx</b> yoki <b>.txt</b> formatdagi fayllarni yuboring.",
            parse_mode=ParseMode.HTML,
        )
        return

    if (doc.file_size or 0) > 10 * 1024 * 1024:
        await message.answer("⚠️ Fayl hajmi 10 MB dan kichik bo'lishi kerak.")
        return

    wait_msg = await message.answer("📄 Fayl o'qilmoqda...")
    try:
        file_io = io.BytesIO()
        await bot.download(doc, destination=file_io)
        extracted_text = await asyncio.to_thread(
            extract_text_from_bytes, file_io.getvalue(), filename
        )
    except Exception:
        logger.exception("Fayl o'qishda xatolik:")
        await wait_msg.edit_text("⚠️ Faylni o'qishda xatolik yuz berdi. Boshqa fayl yuborib ko'ring.")
        return
    finally:
        try:
            await wait_msg.delete()
        except Exception:
            pass

    if len(extracted_text) < 80:
        await message.answer(
            "⚠️ Fayl ichidan yetarli matn topilmadi (skaner qilingan rasm-PDF bo'lishi mumkin)."
        )
        return

    user = await check_and_get_user(message.from_user.id, message.from_user.full_name)
    max_slides = user_max_slides(user, message.from_user.id)

    await state.update_data(
        source_text=extracted_text,
        display_title=filename,
        is_document=True,
    )
    safe_filename = html.escape(filename)
    pro_hint = (
        f"\n<i>💎 PRO obunada {PRO_MAX_SLIDES} tagacha slayd yaratish mumkin (/pro)</i>"
        if max_slides < PRO_MAX_SLIDES
        else ""
    )
    await message.answer(
        f"✅ Fayl qabul qilindi: <b>{safe_filename}</b>\n\n"
        f"📊 <b>Nechta slayd kerak?</b> ({MIN_SLIDES} dan {max_slides} gacha):{pro_hint}",
        parse_mode=ParseMode.HTML,
        reply_markup=get_count_keyboard(max_slides),
    )
    await state.set_state(PresStates.waiting_for_count)


async def _handle_text_input(message: Message, state: FSMContext) -> None:
    text = (message.text or "").strip()
    if len(text) < 3:
        await message.answer("⚠️ Mavzu juda qisqa. Batafsilroq yozing:")
        return

    user = await check_and_get_user(message.from_user.id, message.from_user.full_name)
    max_slides = user_max_slides(user, message.from_user.id)

    is_long_text = len(text) > 150
    display_title = text[:45] + "..." if is_long_text else text

    await state.update_data(
        source_text=text,
        display_title=display_title,
        is_document=is_long_text,
    )
    safe_title = html.escape(display_title)
    pro_hint = (
        f"\n<i>💎 PRO obunada {PRO_MAX_SLIDES} tagacha slayd yaratish mumkin (/pro)</i>"
        if max_slides < PRO_MAX_SLIDES
        else ""
    )
    await message.answer(
        f"✅ Qabul qilindi: <b>{safe_title}</b>\n\n"
        f"📊 <b>Nechta slayd kerak?</b> ({MIN_SLIDES} dan {max_slides} gacha):{pro_hint}",
        parse_mode=ParseMode.HTML,
        reply_markup=get_count_keyboard(max_slides),
    )
    await state.set_state(PresStates.waiting_for_count)


@router.message(PresStates.waiting_for_source, F.document)
async def process_document_source(message: Message, state: FSMContext, bot: Bot):
    await _handle_document_input(message, state, bot)


@router.message(PresStates.waiting_for_source, F.text, ~F.text.startswith("/"))
async def process_text_source(message: Message, state: FSMContext):
    await _handle_text_input(message, state)


@router.message(PresStates.waiting_for_count, F.text)
async def process_count(message: Message, state: FSMContext):
    text = (message.text or "").strip()
    user = await check_and_get_user(message.from_user.id, message.from_user.full_name)
    max_slides = user_max_slides(user, message.from_user.id)

    if not text.isdigit():
        await message.answer(f"⚠️ Iltimos, {MIN_SLIDES} dan {max_slides} gacha raqam kiriting:")
        return

    count = int(text)
    if count < MIN_SLIDES:
        await message.answer(f"⚠️ Minimal slaydlar soni — {MIN_SLIDES} ta:")
        return
    if count > max_slides:
        if max_slides < PRO_MAX_SLIDES:
            kb = await get_pro_keyboard(message.from_user.id)
            await message.answer(
                f"🔒 Bepul tarifda maksimal <b>{FREE_MAX_SLIDES} ta</b> slayd yaratish mumkin.\n"
                f"<b>{PRO_MAX_SLIDES} tagacha</b> slayd uchun PRO obunani faollashtiring yoki "
                f"{MIN_SLIDES}–{FREE_MAX_SLIDES} oralig'ida raqam kiriting:",
                parse_mode=ParseMode.HTML,
                reply_markup=kb,
            )
        else:
            await message.answer(f"⚠️ Maksimal slaydlar soni — {max_slides} ta:")
        return

    await state.update_data(slide_count=count)
    await message.answer(
        f"📑 Slaydlar soni: <b>{count} ta</b>",
        parse_mode=ParseMode.HTML,
        reply_markup=ReplyKeyboardRemove(),
    )
    await message.answer(
        "🌐 <b>Prezentatsiya tilini tanlang:</b>",
        parse_mode=ParseMode.HTML,
        reply_markup=get_lang_keyboard(),
    )
    await state.set_state(PresStates.waiting_for_lang)


@router.callback_query(F.data == "nav:cancel")
async def nav_cancel(callback: CallbackQuery, state: FSMContext):
    await state.clear()
    await callback.answer("Bekor qilindi")
    await callback.message.edit_text(
        "❌ Prezentatsiya yaratish bekor qilindi.\n\nYana boshlash uchun mavzu yozing yoki /start bosing."
    )


@router.callback_query(F.data == "nav:back_to_lang")
async def nav_back_to_lang(callback: CallbackQuery, state: FSMContext):
    data = await state.get_data()
    if "source_text" not in data:
        await callback.answer("Sessiya eskirgan, /start bosing", show_alert=True)
        return
    await callback.answer()
    await callback.message.edit_text(
        "🌐 <b>Prezentatsiya tilini tanlang:</b>",
        parse_mode=ParseMode.HTML,
        reply_markup=get_lang_keyboard(),
    )
    await state.set_state(PresStates.waiting_for_lang)


@router.callback_query(F.data == "nav:back_to_style")
async def nav_back_to_style(callback: CallbackQuery, state: FSMContext):
    data = await state.get_data()
    lang_key = data.get("lang_key")
    if not lang_key or lang_key not in LANGUAGES:
        await callback.answer("Sessiya eskirgan, /start bosing", show_alert=True)
        return
    await callback.answer()
    await callback.message.edit_text(
        f"🌐 Til: <b>{LANGUAGES[lang_key][0]}</b>\n\n"
        "✍️ <b>Prezentatsiya uslubini tanlang:</b>",
        parse_mode=ParseMode.HTML,
        reply_markup=get_style_keyboard(),
    )
    await state.set_state(PresStates.waiting_for_style)


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
        "✍️ <b>Prezentatsiya uslubini tanlang:</b>",
        parse_mode=ParseMode.HTML,
        reply_markup=get_style_keyboard(),
    )
    await state.set_state(PresStates.waiting_for_style)


@router.callback_query(PresStates.waiting_for_style, F.data.startswith("style:"))
async def process_style(callback: CallbackQuery, state: FSMContext):
    style_key = callback.data.split(":", 1)[1]
    if style_key not in STYLES:
        await callback.answer("Xato tanlov", show_alert=True)
        return

    data = await state.get_data()
    lang_key = data.get("lang_key", "uz")
    await state.update_data(style_key=style_key)
    await callback.answer()
    await callback.message.edit_text(
        f"🌐 Til: <b>{LANGUAGES[lang_key][0]}</b> | ✍️ <b>{STYLES[style_key][0]}</b>\n\n"
        "🎨 <b>Dizayn shablonini tanlang:</b>",
        parse_mode=ParseMode.HTML,
        reply_markup=get_template_keyboard(),
    )
    await state.set_state(PresStates.waiting_for_template)


@router.callback_query(PresStates.waiting_for_template, F.data.startswith("tpl:"))
async def process_template(callback: CallbackQuery, state: FSMContext):
    tpl_key = callback.data.split(":", 1)[1]
    if tpl_key not in THEMES:
        await callback.answer("Xato shablon", show_alert=True)
        return

    user_id = callback.from_user.id
    if user_id in ACTIVE_GENERATIONS:
        await callback.answer("⏳ Oldingi prezentatsiyangiz hali tayyorlanmoqda...", show_alert=True)
        return

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

    data = await state.get_data()
    if "source_text" not in data or "lang_key" not in data:
        await callback.answer("Sessiya eskirgan, /start bosing", show_alert=True)
        await state.clear()
        return

    await state.clear()
    await callback.answer()

    theme = THEMES[tpl_key]
    lang_key = data["lang_key"]
    style_key = data.get("style_key", "academic")
    lang_label, lang_instruction, _, _ = LANGUAGES[lang_key]
    style_label, style_instruction = STYLES[style_key]

    header_line = f"🎨 <b>{theme.label}</b> | 🌐 {lang_label} | ✍️ {style_label}"
    status_msg = await callback.message.edit_text(
        f"{header_line}\n\n⏳ <b>1/3 — Gemini AI matn va slayd rejasini tuzmoqda...</b>",
        parse_mode=ParseMode.HTML,
    )

    ACTIVE_GENERATIONS.add(user_id)
    try:
        with tempfile.TemporaryDirectory(prefix="pptbot_") as tmpdir:
            uid = uuid.uuid4().hex[:8]
            pptx_file = os.path.join(tmpdir, f"pres_{uid}.pptx")

            pres_data = await generate_presentation_content(
                source_text=data["source_text"],
                slide_count=data["slide_count"],
                lang_instruction=lang_instruction,
                style_instruction=style_instruction,
                is_document=data["is_document"],
            )

            await status_msg.edit_text(
                f"{header_line}\n\n🖼 <b>2/3 — Slaydlar uchun mos rasmlar qidirilmoqda...</b>",
                parse_mode=ParseMode.HTML,
            )
            images = await fetch_all_slide_images(pres_data.slides)

            await status_msg.edit_text(
                f"{header_line}\n\n⚙️ <b>3/3 — .pptx va .pdf fayllar yig'ilmoqda...</b>",
                parse_mode=ParseMode.HTML,
            )
            await asyncio.to_thread(
                build_pptx_with_template,
                pres_data,
                images,
                tpl_key,
                lang_key,
                style_label,
                callback.from_user.full_name or "",
                pptx_file,
            )
            pdf_file = await asyncio.to_thread(convert_pptx_to_pdf, pptx_file, tmpdir)

            safe_name = (
                "".join(
                    c for c in data["display_title"] if c.isalnum() or c in (" ", "-", "_")
                ).strip()[:35]
                or "Prezentatsiya"
            )
            found_imgs = sum(1 for img in images if img is not None)
            total_slides_built = len(pres_data.slides) + 2
            safe_main_title = html.escape(pres_data.main_title)

            await callback.message.answer_document(
                document=FSInputFile(pptx_file, filename=f"{safe_name}.pptx"),
                caption=(
                    f"🎉 <b>Prezentatsiya tayyor!</b>\n\n"
                    f"📌 <b>Sarlavha:</b> {safe_main_title}\n"
                    f"📑 <b>Slaydlar:</b> {total_slides_built} ta (Muqova + {len(pres_data.slides)} asosiy + Xulosa)\n"
                    f"🖼 <b>Rasmlar:</b> {found_imgs} ta\n"
                    f"📝 <b>Ma'ruzachi eslatmalari:</b> qo'shildi\n\n"
                    f"Yana yaratish uchun yangi mavzu yuboring yoki /start bosing."
                ),
                parse_mode=ParseMode.HTML,
            )

            if pdf_file and os.path.exists(pdf_file):
                await callback.message.answer_document(
                    document=FSInputFile(pdf_file, filename=f"{safe_name}.pdf"),
                    caption="📱 Telefonda qulay ko'rish uchun PDF nusxasi",
                )

            await increment_usage(user_id)
            try:
                await status_msg.delete()
            except Exception:
                pass

    except errors.ServerError as e:
        logger.exception("Gemini ServerError:")
        if getattr(e, "code", None) == 503:
            await status_msg.edit_text(
                "⏳ Sun'iy intellekt serverlarida hozir yuklama yuqori.\n"
                "Iltimos, 10–15 soniyadan so'ng qayta urinib ko'ring: /start"
            )
        else:
            await status_msg.edit_text("⚠️ AI serverida vaqtinchalik xatolik yuz berdi. Qayta urinib ko'ring: /start")
    except Exception as e:
        logger.exception("Prezentatsiya yaratishda xatolik: %s", e)
        await status_msg.edit_text(
            "❌ Prezentatsiya yaratishda texnik xatolik yuz berdi.\n"
            "Iltimos, birozdan so'ng qayta urinib ko'ring: /start"
        )
    finally:
        ACTIVE_GENERATIONS.discard(user_id)


# ==============================================================================
# ADMIN PANEL
# ==============================================================================
@router.message(Command("admin"))
async def cmd_admin(message: Message):
    if message.from_user.id != ADMIN_ID:
        return

    assert db_pool is not None
    async with db_pool.acquire() as conn:
        total_users = await conn.fetchval("SELECT COUNT(*) FROM ppt_users")
        pro_users = await conn.fetchval("SELECT COUNT(*) FROM ppt_users WHERE is_pro = TRUE")
        total_ppts = await conn.fetchval("SELECT COALESCE(SUM(total_generated), 0) FROM ppt_users")
        recent_users = await conn.fetchval(
            "SELECT COUNT(*) FROM ppt_users WHERE created_at >= NOW() - INTERVAL '7 days'"
        )
    pro_price = await get_ppt_pro_price()

    kb = InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="📢 Barchaga xabar yuborish", callback_data="admin:broadcast")]
        ]
    )

    await message.answer(
        "🛠 <b>Admin Panel</b>\n\n"
        f"👥 Jami foydalanuvchilar: <b>{total_users} ta</b> (+{recent_users} oxirgi 7 kunda)\n"
        f"💎 PRO foydalanuvchilar: <b>{pro_users} ta</b>\n"
        f"📊 Yaratilgan prezentatsiyalar: <b>{total_ppts} ta</b>\n"
        f"⭐ PRO narxi: <b>{pro_price} Stars</b>\n"
        f"🔑 Gemini API kalitlar: <b>{len(GEMINI_API_KEYS)} ta</b>\n"
        f"🤖 Modellar: <code>{html.escape(', '.join(GEMINI_MODELS))}</code>\n\n"
        "<i>Buyruqlar:</i>\n"
        "<code>/givepro USER_ID</code> — PRO berish\n"
        "<code>/delpro USER_ID</code> — PRO olib tashlash\n"
        "<code>/userinfo USER_ID</code> — foydalanuvchi ma'lumoti\n"
        "<code>/setprice STARS</code> — PRO narxini o'zgartirish",
        parse_mode=ParseMode.HTML,
        reply_markup=kb,
    )


@router.message(Command("givepro"))
async def cmd_give_pro(message: Message):
    if message.from_user.id != ADMIN_ID:
        return
    parts = (message.text or "").split()
    if len(parts) != 2 or not parts[1].isdigit():
        await message.answer("⚠️ Namuna: <code>/givepro 123456789</code>", parse_mode=ParseMode.HTML)
        return

    target_id = int(parts[1])
    assert db_pool is not None
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
    parts = (message.text or "").split()
    if len(parts) != 2 or not parts[1].isdigit():
        await message.answer("⚠️ Namuna: <code>/delpro 123456789</code>", parse_mode=ParseMode.HTML)
        return

    target_id = int(parts[1])
    assert db_pool is not None
    async with db_pool.acquire() as conn:
        await conn.execute("UPDATE ppt_users SET is_pro = FALSE WHERE user_id = $1", target_id)
    await message.answer(f"✅ <code>{target_id}</code> dan PRO olib tashlandi.", parse_mode=ParseMode.HTML)


@router.message(Command("setprice"))
async def cmd_set_price(message: Message):
    if message.from_user.id != ADMIN_ID:
        return
    parts = (message.text or "").split()
    if len(parts) != 2 or not parts[1].isdigit() or int(parts[1]) <= 0:
        await message.answer("⚠️ Namuna: <code>/setprice 50</code>", parse_mode=ParseMode.HTML)
        return

    price = int(parts[1])
    await set_ppt_pro_price(price)
    await message.answer(f"✅ PRO narxi <b>{price} Stars</b> etib belgilandi.", parse_mode=ParseMode.HTML)


@router.message(Command("userinfo"))
async def cmd_user_info(message: Message):
    if message.from_user.id != ADMIN_ID:
        return
    parts = (message.text or "").split()
    if len(parts) != 2 or not parts[1].isdigit():
        await message.answer("⚠️ Namuna: <code>/userinfo 123456789</code>", parse_mode=ParseMode.HTML)
        return

    target_id = int(parts[1])
    assert db_pool is not None
    async with db_pool.acquire() as conn:
        user = await conn.fetchrow("SELECT * FROM ppt_users WHERE user_id = $1", target_id)
    if not user:
        await message.answer("❌ Bunday foydalanuvchi bazada topilmadi.")
        return

    safe_name = html.escape(user["full_name"] or "Noma'lum")
    await message.answer(
        f"👤 <b>Foydalanuvchi ma'lumoti</b>\n\n"
        f"🆔 ID: <code>{user['user_id']}</code>\n"
        f"🏷 Ism: <b>{safe_name}</b>\n"
        f"💎 PRO: <b>{'Ha' if user['is_pro'] else 'Yoʻq'}</b>\n"
        f"📊 Haftalik ishlatgan: <b>{user['weekly_used']}/{WEEKLY_FREE_LIMIT}</b>\n"
        f"🎯 Jami yaratgan: <b>{user['total_generated']} ta</b>",
        parse_mode=ParseMode.HTML,
    )


@router.callback_query(F.data == "admin:broadcast")
async def start_broadcast(callback: CallbackQuery, state: FSMContext):
    if callback.from_user.id != ADMIN_ID:
        return
    await callback.answer()
    await callback.message.answer(
        "📢 Barcha foydalanuvchilarga yuboriladigan xabarni yozing (bekor qilish uchun /cancel):"
    )
    await state.set_state(PresStates.waiting_for_broadcast)


@router.message(PresStates.waiting_for_broadcast)
async def send_broadcast(message: Message, state: FSMContext):
    if message.from_user.id != ADMIN_ID:
        return
    await state.clear()

    assert db_pool is not None
    async with db_pool.acquire() as conn:
        rows = await conn.fetch("SELECT user_id FROM ppt_users")

    total = len(rows)
    sent, failed = 0, 0
    status = await message.answer(f"⏳ Yuborilmoqda (0/{total})...")

    for i, row in enumerate(rows, start=1):
        try:
            await message.copy_to(row["user_id"])
            sent += 1
        except Exception:
            failed += 1
        if i % 25 == 0 or i == total:
            try:
                await status.edit_text(f"⏳ Yuborilmoqda ({i}/{total})...")
            except Exception:
                pass
        await asyncio.sleep(0.05)

    await status.edit_text(
        f"✅ Xabar yuborildi: <b>{sent} ta</b>\n❌ Yetib bormadi: <b>{failed} ta</b>",
        parse_mode=ParseMode.HTML,
    )


# ==============================================================================
# FALLBACK HANDLERLAR (/start BOSMASDAN HAM MAVZU YOKI FAYL YUBORSA ISHLAYDI)
# ==============================================================================
@router.message(StateFilter(None), F.document)
async def fallback_document(message: Message, state: FSMContext, bot: Bot):
    await _handle_document_input(message, state, bot)


@router.message(StateFilter(None), F.text, ~F.text.startswith("/"))
async def fallback_text(message: Message, state: FSMContext):
    await _handle_text_input(message, state)


# ==============================================================================
# RENDER WEB SERVER VA BOT ISHGA TUSHIRISH
# ==============================================================================
async def health_check(request: web.Request) -> web.Response:
    return web.Response(text="Bot is running!")


async def start_web_server() -> web.AppRunner:
    app = web.Application()
    app.router.add_get("/", health_check)
    app.router.add_get("/health", health_check)
    runner = web.AppRunner(app)
    await runner.setup()
    site = web.TCPSite(runner, "0.0.0.0", PORT)
    await site.start()
    logger.info("Health-check server 0.0.0.0:%d da ishga tushdi.", PORT)
    return runner


async def main() -> None:
    if not BOT_TOKEN or BOT_TOKEN == "YOUR_PPT_BOT_TOKEN":
        raise RuntimeError("BOT_TOKEN environment variable sozlanmagan!")
    if not GEMINI_API_KEYS:
        logger.warning(
            "DIQQAT: Gemini API kaliti topilmadi! API_KEYS yoki GEMINI_API_KEY ni sozlang."
        )

    await init_db()
    runner = await start_web_server()
    bot = Bot(token=BOT_TOKEN)
    dp = Dispatcher()
    dp.include_router(router)

    try:
        await bot.set_my_commands(
            [
                BotCommand(command="start", description="Yangi prezentatsiya yaratish"),
                BotCommand(command="profile", description="Profil va haftalik limit"),
                BotCommand(command="pro", description="PRO obuna haqida"),
                BotCommand(command="help", description="Foydalanish yo'riqnomasi"),
                BotCommand(command="cancel", description="Joriy amalni bekor qilish"),
            ]
        )
        await bot.delete_webhook(drop_pending_updates=True)
        logger.info(
            "Prezentatsiya boti ishga tushdi | API kalitlar: %d | Modellar: %s",
            len(GEMINI_API_KEYS),
            ", ".join(GEMINI_MODELS),
        )
        await dp.start_polling(bot)
    finally:
        await bot.session.close()
        await runner.cleanup()
        if db_pool is not None:
            await db_pool.close()


if __name__ == "__main__":
    asyncio.run(main())
