"""All Persian user-facing strings and data labels (single source of truth)."""

from __future__ import annotations

# ---------------------------------------------------------------------------
# Row-level labels (written into the data; must stay exactly as specified)
# ---------------------------------------------------------------------------
COMMENT_WIN = "بای باکس هست"
COMMENT_LOSS = "بای باکس نیست"
COMMENT_NO_WEIGHT = "وزن ثبت نشده - قابل ارزیابی نیست"
COMMENT_BAD_PRICE = "قیمت نامعتبر - قابل ارزیابی نیست"
COMMENT_MISSING_ID = "شناسه ناقص - قابل ارزیابی نیست"

COMP_COMPETITIVE = "رقابتی"
COMP_SINGLE_DUP = "تک‌فروشنده - DKPC تکراری"
COMP_SINGLE = "تک‌فروشنده"
COMP_NO_WEIGHT = "وزن ثبت نشده"
COMP_BAD_PRICE = "قیمت نامعتبر"
COMP_MISSING_ID = "شناسه ناقص"

CHANGE_NEW = "تنوع جدید"
CHANGE_NO_WEIGHT = "وزن ثبت نشده"
CHANGE_BAD_PRICE = "خارج از ارزیابی"
CHANGE_SAME = "بدون تغییر"
CHANGE_GAINED = "بای باکس گرفته"
CHANGE_LOST = "بای باکس از دست رفته"

MARGIN_FLAG = "هشدار حاشیه"

DIGIKALA_RETAIL = "دیجی‌کالا (رتیل)"
DIGIKALA_SELLER_ID = 1
KAM_NOT_IN_LIST = "در لیست KAM نیست"
KAM_DUPLICATE_FLAG = "تکراری در لیست KAM"

JOIN_SEP = " / "

# ---------------------------------------------------------------------------
# Seller table headers (internal key -> Persian header), in display order
# ---------------------------------------------------------------------------
SELLER_COLUMNS: dict[str, str] = {
    "Seller_ID": "Seller_ID",
    "Seller_Name": "Seller Name",
    "KAM": "KAM",
    "Main_Brand": "برند اصلی",
    "Variants": "تعداد تنوع فعال (DKPC)",
    "Not_Evaluated": "تنوع بدون وزن (خارج از ارزیابی)",
    "Evaluated": "تنوع ارزیابی‌شده",
    "Wins": "تعداد بای باکس",
    "Win_Rate": "وین‌ریت کل",
    "Comp_Variants": "تنوع در رقابت",
    "Comp_Wins": "بای باکس در رقابت",
    "Comp_Win_Rate": "وین‌ریت رقابتی",
    "Prev_Comp_Win_Rate": "وین‌ریت رقابتی دوره قبل",
    "Comp_Win_Rate_Change": "تغییر (واحد درصد)",
    "Gained": "بای باکس گرفته‌شده",
    "Lost": "بای باکس از دست رفته",
    "Single_Seller_Variants": "تنوع تک‌فروشنده",
    "No_BuyBox": "تنوع بدون بای باکس",
    "Near_Miss": "باخت نزدیک",
    "Median_Loss_Gap": "میانه فاصله باخت‌ها",
    "Share": "سهم از کل تنوع‌ها",
}
TOTAL_ROW_LABEL = "جمع کل"
PREV_ONLY_COLUMNS = ("Prev_Comp_Win_Rate", "Comp_Win_Rate_Change", "Gained", "Lost")

# ---------------------------------------------------------------------------
# UI strings
# ---------------------------------------------------------------------------
APP_TITLE = "تحلیل بای باکس شمش و سکه"
APP_SUBTITLE = "شمش طلا · سکه پارسیان (گرمی) · پک شمش و پلاک طلا"

SIDEBAR_FILES = "فایل‌ها"
UPLOAD_CURRENT = "اسورتمنت فعلی (xlsx)"
UPLOAD_KAM = "لیست KAM (xlsx، شیت KA Selection)"
UPLOAD_PREV = "اسورتمنت دوره قبل (اختیاری)"
PREV_SOURCE = "منبع دوره قبل برای مقایسه"
PREV_NONE = "بدون مقایسه"
PREV_FILE = "آپلود فایل"
PREV_RUN = "انتخاب اجرای ذخیره‌شده"
SELECT_RUN = "اجرای ذخیره‌شده"
KAM_SAVED_INFO = "از آخرین لیست KAM ذخیره‌شده استفاده می‌شود."
KAM_SAVED_FORGET = "حذف لیست KAM ذخیره‌شده"

SIDEBAR_PARAMS = "پارامترها"
PARAM_NEAR_MISS = "آستانه باخت نزدیک (٪)"
PARAM_UNDERCUT = "مقدار زیر قیمت زدن (ریال)"
PARAM_UNDERCUT_HELP = (
    "۰ یعنی هم‌قیمت شدن با بای باکس. قانون تساوی (tie-break) الگوریتم واقعی بای باکس مشخص نیست؛ "
    "هم‌قیمت شدن ممکن است برای گرفتن بای باکس کافی نباشد."
)
PARAM_MAX_CUT = "حداکثر کاهش قیمت قابل قبول (٪)"

SIDEBAR_FILTERS = "فیلترها"
FILTER_KAM = "KAM"
FILTER_CATEGORY = "دسته‌بندی"
FILTER_BRAND = "برند"
FILTER_SELLER = "فروشنده"

SIDEBAR_SETTINGS = "تنظیمات"
SETTINGS_MY_KAM = "نام KAM من (فیلتر پیش‌فرض)"
SETTINGS_AUTOSAVE = "ذخیره خودکار اجرا در تاریخچه"

TAB_SUMMARY = "خلاصه مدیریتی"
TAB_SELLERS = "فروشندگان"
TAB_NEAR_MISS = "باخت‌های نزدیک (Quick Win)"
TAB_CHANGES = "تغییرات نسبت به دوره قبل"
TAB_CATEGORY = "دسته‌بندی و برند"
TAB_QUALITY = "کیفیت داده"

KPI_EVALUATED = "تنوع ارزیابی‌شده"
KPI_COMPETITIVE = "تنوع در رقابت"
KPI_COMP_WIN_RATE = "وین‌ریت رقابتی"
KPI_NEAR_MISS = "باخت نزدیک"
KPI_GAINED = "بای باکس گرفته"
KPI_LOST = "بای باکس از دست رفته"
KPI_NEW = "تنوع جدید"
KPI_REMOVED = "تنوع حذف‌شده"

INFO_COMPETITIVE_KPI = (
    "چرا «وین‌ریت رقابتی» شاخص اصلی است؟ بخش بزرگی از گروه‌های DKP+وزن فقط یک فروشنده دارند "
    "(در این فایل: {single_pct}). این گروه‌ها همیشه بای باکس را می‌برند و وین‌ریت کل را بالا می‌برند "
    "(وین‌ریت کل: {overall}). وین‌ریت واقعی در رقابت: {competitive}."
)
INFO_PRICE_ONLY = (
    "بای باکس فقط بر اساس قیمت سنجیده می‌شود؛ زمان ارسال، امتیاز فروشنده و موجودی در فایل ورودی نیست."
)
INFO_NO_PREV = "برای دیدن تغییرات، در نوار کناری فایل دوره قبل را آپلود کنید یا یک اجرای ذخیره‌شده را انتخاب کنید."
INFO_UPLOAD_FIRST = "برای شروع، فایل اسورتمنت فعلی را از نوار کناری آپلود کنید."

MAPPING_TITLE = "تطبیق ستون‌ها"
MAPPING_HELP = "برخی ستون‌های لازم پیدا نشدند. ستون معادل هر مورد را انتخاب کنید."
MAPPING_NONE = "— انتخاب نشده —"
MAPPING_ACTUAL_BB = "وضعیت واقعی بای باکس (اختیاری، برای اعتبارسنجی قانون قیمت)"

ERR_FILE_LOCKED = "فایل باز است یا قفل شده است. فایل را در اکسل ببندید و دوباره آپلود کنید."
ERR_FILE_MALFORMED = "فایل اکسل خراب است یا قالب آن قابل خواندن نیست ({detail})."
ERR_FILE_EMPTY = "فایل خالی است یا هیچ ردیف داده‌ای ندارد."
ERR_KAM_COLUMNS = "ستون‌های لازم در لیست KAM پیدا نشدند: {missing}"
WARN_KAM_SHEET = "شیت «KA Selection» پیدا نشد؛ از شیت «{sheet}» استفاده شد."

EXPORT_MAIN = "دانلود فایل اکسل کامل"
EXPORT_PREPARE = "آماده‌سازی فایل اکسل کامل"
EXPORT_QUICKWIN = "دانلود Quick Win"

# Excel notes block (Persian method explanation)
EXCEL_TITLE = "گزارش بای باکس فروشندگان - شمش و سکه"
EXCEL_NOTES: tuple[str, ...] = (
    "روش محاسبه:",
    "۱. هر DKPC در گروه «DKP + وزن» (وزن گرد شده تا ۴ رقم اعشار) با بقیه مقایسه می‌شود.",
    "۲. کمترین Selling_Price هر گروه بای باکس است؛ در حالت تساوی همه ردیف‌های با کمترین قیمت برنده‌اند.",
    "۳. ردیف‌های بدون وزن یا با قیمت نامعتبر ارزیابی نمی‌شوند و در مخرج هیچ وین‌ریتی نیستند.",
    "۴. «رقابتی» یعنی حداقل یک فروشنده دیگر در همان گروه حضور دارد. وین‌ریت رقابتی شاخص اصلی است.",
    "۵. وین‌ریت رقابتی سطر جمع کل به صورت تجمیعی (مجموع برد رقابتی / مجموع تنوع رقابتی) حساب شده است، نه میانگین درصدها.",
    "۶. باخت نزدیک: فاصله قیمت تا بای باکس بزرگ‌تر از صفر و حداکثر برابر آستانه.",
    "۷. قیمت هدف = قیمت بای باکس − مقدار زیر قیمت زدن. قانون تساوی الگوریتم واقعی بای باکس مشخص نیست.",
    "۸. بای باکس فقط بر اساس قیمت سنجیده شده؛ زمان ارسال، امتیاز فروشنده و موجودی در داده ورودی نیست.",
    "۹. همه مقادیر به صورت عدد ذخیره شده‌اند (بدون فرمول).",
)
