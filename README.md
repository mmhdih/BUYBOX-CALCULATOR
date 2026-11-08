# تحلیل بای باکس شمش و سکه — Buy Box Analyzer

اپ Streamlit برای تحلیل بای باکس دسته‌های «شمش طلا»، «سکه پارسیان (گرمی)» و «پک شمش و پلاک طلا».
جایگزین تحلیل دستی اکسل: تعیین بای باکس هر DKPC، وین‌ریت فروشندگان (با **وین‌ریت رقابتی** به‌عنوان شاخص اصلی)،
مقایسه با دوره قبل، فهرست باخت‌های نزدیک (Quick Win) و خروجی اکسل.

> داده‌ها فقط روی همان کامپیوتر پردازش می‌شوند و به هیچ سرویس بیرونی ارسال نمی‌شوند
> (آمار استفاده Streamlit خاموش است و سرور فقط روی `localhost` گوش می‌دهد).

---

## فارسی

### ۱. نسخه پرتابل ویندوز (بدون نصب)

1. فایل `BuyBoxAnalyzer_portable_win64.zip` را بگیرید (از بخش **Actions → آخرین اجرای موفق → Artifacts** در گیت‌هاب،
   یا از **Releases** برای نسخه‌های تگ‌خورده، یا خودتان بسازید — بند ۳).
2. کل zip را Extract کنید (روی دسکتاپ، درایو دیگر یا فلش).
3. روی `BuyBox.exe` دوبار کلیک کنید. مرورگر خودکار باز می‌شود و پنجره کوچک برنامه به نوار وظیفه می‌رود.
4. برای بستن: دکمه «بستن برنامه» در نوار کناری اپ (یا بستن پنجره Buy Box Analyzer در نوار وظیفه).
   اگر برنامه باز باشد، اجرای دوباره `BuyBox.exe` نسخه دوم باز نمی‌کند و فقط همان را در مرورگر نشان می‌دهد.

`BuyBox.exe` یک لانچر بومی ویندوز است (با آیکون)؛ پایتونِ داخل بسته را اجرا می‌کند و با بستن آن، برنامه کامل
بسته می‌شود و چیزی در پس‌زمینه نمی‌ماند. فایل امضای دیجیتال ندارد؛ اگر SmartScreen هشدار داد:
More info ← Run anyway.

نیازی به نصب پایتون یا دسترسی ادمین نیست. برنامه روی پورت **8642** اجرا می‌شود (نه 8501 که پیش‌فرض Streamlit
و اپ‌هایی مثل DK ATP است)، پس هر دو هم‌زمان باز می‌مانند؛ اگر 8642 هم اشغال باشد، پورت آزاد بعدی انتخاب می‌شود.
برای پورت دلخواه، متغیر محیطی `BUYBOX_PORT` را تنظیم کنید. تاریخچه اجراها، تنظیمات و آخرین لیست KAM در
`app\data` ذخیره می‌شود؛ برای رفتن به نسخه جدید همین پوشه را کپی کنید.

### ۲. اجرا با پایتون (برای توسعه)

```bash
pip install -r requirements.txt
streamlit run app.py
```

یا `python launch.py` (از پورت 8642 شروع می‌کند، پورت آزاد را خودش پیدا می‌کند و مرورگر را باز می‌کند).

### ۳. ساخت نسخه پرتابل

روی هر سیستمی (ویندوز، لینوکس یا مک) با پایتون ۳.۱۰+ و اینترنت:

```bash
python portable/build_portable.py
```

خروجی: پوشه `dist/BuyBoxAnalyzer` و فایل `dist/BuyBoxAnalyzer_portable_win64.zip` (حدود ۹۰ مگابایت، شامل BuyBox.exe).
`BuyBox.exe` از `portable/launcher/launcher.c` با کامپایلر C ابزار zig (`pip install ziglang`، خودکار نصب می‌شود)
ساخته می‌شود، پس روی لینوکس و مک هم فایل exe ویندوز ساخته می‌شود. هر push به گیت‌هاب هم همین بسته را روی یک ماشین
ویندوز واقعی می‌سازد، `BuyBox.exe` را اجرا و تست می‌کند و به‌عنوان Artifact قرار می‌دهد. برای انتشار نسخه:
Actions ← «Tests & portable Windows build» ← Run workflow و در `release_tag` مثلاً `v1.3` را بنویسید.

### ۴. روال به‌روزرسانی (هر چند روز)

1. اسورتمنت جدید را در نوار کناری آپلود کنید — **xlsx یا csv تا ۵۰۰ مگابایت** (لیست KAM از دفعه قبل به خاطر
   سپرده شده است). در CSV، رمزگذاری (UTF-8 یا Windows-1256)، جداکننده (`,` `;` تب `|`)، ارقام فارسی و
   جداکننده هزارگان خودکار تشخیص داده می‌شود.
2. «منبع دوره قبل» را روی «انتخاب اجرای ذخیره‌شده» بگذارید و اجرای قبلی را انتخاب کنید
   (هر اجرا به‌طور خودکار در `data/runs` ذخیره می‌شود؛ یا فایل دوره قبل را آپلود کنید).
3. اول پنل «کیفیت داده» را ببینید، بعد تب‌ها را.
4. از تب «باخت‌های نزدیک» برای هر فروشنده فایل `QuickWin_<نام>_<تاریخ>.xlsx` را دانلود و ارسال کنید.
5. خروجی اکسل (نوار بالای تب‌ها):
   - **«خروجی اکسل کامل — همه تب‌ها»**: یک فایل مرتب و فارسی با یک شیت برای هر تب (خلاصه با نمودار، فروشندگان،
     باخت به چه کسی، بای باکس‌های از دست رفته، باخت‌های نزدیک، تغییرات، دسته‌بندی و برند، کیفیت داده، همه ردیف‌ها)
     و شیت «راهنمای ستون‌ها». فیلترهای فعلی اعمال می‌شوند و روی هر شیت نوشته می‌شوند.
   - **«اکسل این تب»** در بالای هر تب: فقط محتوای همان تب.
   - **«خروجی استاندارد»** `BuyBox_<نام فایل>.xlsx`: قالب ثابت (همه ستون‌های فایل اصلی + ستون‌های محاسبه‌شده)، بدون فیلتر.
   همه فایل‌ها: جدول‌های اکسل با دکمه فیلتر و ردیف‌های راه‌راه، سرستون‌های فارسی با توضیح (روی سرستون مکث کنید)،
   رنگ‌بندی قرمز/زرد/سبز وین‌ریت، فرمت عدد و درصد، و فقط مقدار (بدون فرمول). بیش از ۱٬۰۰۰٬۰۰۰ ردیف در چند شیت
   پشت سر هم نوشته می‌شود.

### ۵. منطق محاسبه (خلاصه)

- گروه‌بندی بر اساس **DKP + وزن** (وزن گرد شده تا ۴ رقم اعشار). کمترین `Selling_Price` گروه = بای باکس؛ در تساوی همه برنده‌اند.
- ردیف‌های **بدون وزن** ارزیابی نمی‌شوند و در مخرج هیچ وین‌ریتی نیستند (فقط با DKP گروه‌بندی نمی‌شوند).
- «رقابتی» = حداقل یک فروشنده دیگر در گروه. **وین‌ریت رقابتی** شاخص اصلی است؛ سطر جمع کل به‌صورت تجمیعی حساب می‌شود.
- باخت نزدیک: `0 < Gap ≤ آستانه` (خود آستانه هم حساب است). قیمت هدف = قیمت بای باکس − زیر قیمت زدن.
- هشدار حاشیه وقتی کاهش لازم از «حداکثر کاهش قابل قبول» بیشتر باشد.

### ۶. محدودیت‌های شناخته‌شده

- **بای باکس فقط با قیمت سنجیده می‌شود**؛ زمان ارسال، امتیاز فروشنده و موجودی در فایل ورودی نیست.
  برای اعتبارسنجی، در «تطبیق ستون‌ها» می‌توانید ستون اختیاری «وضعیت واقعی بای باکس» را انتخاب کنید
  تا درصد انطباق قانون قیمت با واقعیت در پنل کیفیت داده نمایش داده شود.
- قانون تساوی (tie-break) الگوریتم واقعی مشخص نیست؛ «زیر قیمت زدن = ۰» یعنی هم‌قیمت شدن که ممکن است کافی نباشد.
- روند (Trend) فقط از اجراهای ذخیره‌شده ساخته می‌شود و فیلتر KAM/دسته/برند روی آن اعمال نمی‌شود (فیلتر فروشنده اعمال می‌شود).

---

## English

### Run

- **Portable Windows build:** extract `BuyBoxAnalyzer_portable_win64.zip`, double-click `BuyBox.exe`
  (native launcher with icon; second launch reuses the running app; the in-app «بستن برنامه» button or
  closing the launcher stops everything — the Python child runs in a kill-on-close job object).
  Contains the official embeddable Python 3.12 + all dependencies; no install, no admin rights.
  Runs on port **8642** (not Streamlit's default 8501, used by other internal apps such as DK ATP);
  falls back to the next free port, or set `BUYBOX_PORT`.
- **From source:** `pip install -r requirements.txt` then `streamlit run app.py` (or `python launch.py`).
- **Build the portable zip:** `python portable/build_portable.py` (any OS; fetches win_amd64 wheels and
  cross-compiles `BuyBox.exe` from `portable/launcher/` with `zig cc` from the `ziglang` wheel).
- **Inputs:** `.xlsx` or `.csv` up to 500 MB (CSV encoding, delimiter, Persian digits and separators detected).
  Large files are cached by content hash with `st.cache_resource`, so reruns neither re-hash nor copy them.
- **Exports:** full styled report (all tabs) / one tab (`buybox/report.py`), and the standard
  `BuyBox_<file>.xlsx` (`buybox/export.py`); shared styling in `buybox/xlstyle.py`.
  CI (`.github/workflows/portable.yml`) runs the tests, builds the package on `windows-latest`,
  smoke-tests it with the embedded interpreter, runs `BuyBox.exe` with port 8501 taken and checks the app
  on 8642, that a second launch reuses it, and that closing the launcher leaves no Python process,
  then uploads it as an artifact (and attaches the zip to the release for `v*` tags).

### Tests

```bash
python -m pytest -q
```

Synthetic unit tests cover win/loss, ties, same-seller duplicates, blank weight, single-seller groups,
the near-miss boundary, float weight mismatch, KAM merge, all comparison labels, removed DKPCs and pooled
vs averaged win rate, plus a 25k-row performance test (< 5 s). `tests/test_regression.py` checks the real
exports when `اسورتمنت شمش 14050802.xlsx` and `اسورتمنت شمش.xlsx` are placed in `tests/fixtures/`
(not committed; skipped otherwise).

### Project layout

```
app.py              Streamlit UI only
launch.py           free port from 8642, single instance, open browser (started by BuyBox.exe)
buybox/
  io.py             read/validate/normalize inputs, column mapping, KAM merge
  core.py           buy-box logic (pure pandas, vectorized) and seller/market aggregations
  compare.py        comparison with the previous period
  quality.py        data-quality checks
  export.py         standard BuyBox_<file>.xlsx + QuickWin files (xlsxwriter, values only)
  report.py         styled Persian report: all tabs or one tab
  xlstyle.py        shared Excel styling (tables, formats, colours, tooltips, sheet splitting)
  history.py        saved runs (parquet + JSON)
  pipeline.py       orchestration and filters (no Streamlit)
  labels.py         every Persian string (single source of truth)
portable/           Windows portable builder, pinned requirements, smoke test
portable/launcher/  BuyBox.exe source (C), icon and version resources
static/fonts/       Vazirmatn (OFL) for offline use; Google Fonts is also requested
scripts/make_sample_data.py   fake sample files for trying the app
data/runs/          saved runs (gitignored)
```

### Decisions taken where the spec was ambiguous (conservative choices)

- **Invalid price** (`Selling_Price` blank or ≤ 0) and **missing DKP/Seller_ID** rows are excluded like blank
  weight, with their own labels (`قیمت نامعتبر - قابل ارزیابی نیست`, `شناسه ناقص - قابل ارزیابی نیست`). The seller
  column «تنوع بدون وزن (خارج از ارزیابی)» counts all non-evaluated rows; for the real files this equals blank weight.
- **BuyBox_Change** for a row that was *not evaluable* before and *loses* now is `بدون تغییر` (it did not lose a buy box
  it held). A current row that is non-evaluable for a reason other than blank weight gets `خارج از ارزیابی`.
- **Cut_vs_Wage** (`Required_Cut_% / True_Wage`) is written next to `Required_Cut_%` when `True_Wage` exists; the margin
  flag itself uses only `Required_Cut_% > max_acceptable_cut`, as specified. `True_Wage` is used as given (no unit conversion).
- **Saved runs** are de-duplicated by the input file's content hash (buy-box results do not depend on the parameters).
- **Filters** restrict which rows are summarized; the buy box itself is always computed on the full file.
  The main Excel export always contains the full, unfiltered data.
- **Duplicate DKPC rows** keep the first occurrence; dropped rows are listed in the data-quality panel.
- Original columns whose names collide with computed ones (e.g. an old `Comment`) are exported with an `_input` suffix.
- Change in competitive win rate is shown in percentage points (`pp`); Excel highlights ≥ +5 pp green and ≤ −5 pp red.

### Known limitations

- Buy box is judged on **price only** (no shipping time, seller score or stock in the input). Map the optional
  "actual buy-box status" column to measure how often the price-only rule matches reality.
- The real tie-break rule is unknown; an undercut of 0 means "match the price".
- The trend chart uses saved runs only; KAM/category/brand filters do not apply to it (seller filter does).
- The QuickWin file name keeps Persian characters (sent via RFC 5987 `filename*`); all current browsers handle it.
