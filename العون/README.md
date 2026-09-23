# Odoo MCP Server — محلات العون لمواد البناء

سيرفر [MCP](https://modelcontextprotocol.io) بيربط **كلود** بقاعدة بيانات أودو
مباشرة، فبتقدر تسأل بالعربي: "شو مبيعات الشهر؟" أو "ضيف زبون جديد" وكلود بينفذ
على أودو فورًا — قراءة **وكتابة**.

مجرّب على `alawn.odoo.com` — أودو **19.2 Enterprise** (`saas~19.2+e`).

---

## المتطلبات

- Python 3.10 أو أحدث
- [`uv`](https://docs.astral.sh/uv/) (أو `pip` عادي)
- مستخدم أودو + **API Key**

---

## 1. جيب الـ API Key من أودو

1. افتح `https://alawn.odoo.com` وسجّل دخول.
2. اضغط على اسمك فوق على اليمين → **My Profile**.
3. تبويب **Account Security** → زر **New API Key**.
4. اكتب اسم للمفتاح (مثلًا `claude`) → انسخ المفتاح.

> المفتاح بينعرض **مرة وحدة بس**. احفظه بمكان آمن.
>
> الـ API Key بيستعمل مستخدمك الحالي، **ما بياخد مقعد (seat) إضافي** — فرسالة
> "All seats have been used" ما بتمنعك.

---

## 2. التنصيب

```bash
git clone <repo-url> odoo-mcp && cd odoo-mcp
uv venv && uv pip install -e .
cp .env.example .env      # وبعدين عبّي ODOO_API_KEY جواه
```

تأكد إنه الاتصال شغّال قبل ما توصله بكلود:

```bash
uv run python scripts/check_connection.py
```

المفروض يطلعلك:

```
Server   : https://alawn.odoo.com
Database : alawn
Version  : saas~19.2+e
User     : ... (uid 2)
Connection OK.
```

---

## 3. وصّله بكلود

### Claude Desktop

عدّل ملف الإعدادات:
`~/Library/Application Support/Claude/claude_desktop_config.json` (ماك) أو
`%APPDATA%\Claude\claude_desktop_config.json` (ويندوز):

```json
{
  "mcpServers": {
    "odoo": {
      "command": "uv",
      "args": ["run", "--directory", "/absolute/path/to/odoo-mcp", "odoo-mcp"],
      "env": {
        "ODOO_URL": "https://alawn.odoo.com",
        "ODOO_DB": "alawn",
        "ODOO_USERNAME": "you@example.com",
        "ODOO_API_KEY": "المفتاح-هون"
      }
    }
  }
}
```

بعدها سكّر كلود وافتحه من جديد.

### Claude Code

```bash
claude mcp add odoo \
  --env ODOO_URL=https://alawn.odoo.com \
  --env ODOO_DB=alawn \
  --env ODOO_USERNAME=you@example.com \
  --env ODOO_API_KEY=المفتاح-هون \
  -- uv run --directory /absolute/path/to/odoo-mcp odoo-mcp
```

أو انسخ `.mcp.json.example` لـ `.mcp.json` وعبّيه.

---

## الأدوات المتاحة (12)

| الأداة | الوظيفة |
|---|---|
| `odoo_status` | فحص الاتصال: النسخة، المستخدم، الشركة، وصلاحيات الكتابة |
| `odoo_list_models` | استعراض الموديلات الموجودة بقاعدة البيانات |
| `odoo_fields` | حقول أي موديل: الاسم التقني، النوع، العلاقة |
| `odoo_search_read` | البحث والقراءة — الأداة الأساسية |
| `odoo_count` | عدّ السجلات بدون جلبها |
| `odoo_read` | قراءة سجلات محددة بالـ id |
| `odoo_name_search` | تحويل اسم (عربي أو إنجليزي) لـ id |
| `odoo_aggregate` | تجميع وحسابات — للتقارير والمجاميع |
| `odoo_create` | إنشاء سجل جديد |
| `odoo_write` | تعديل سجلات موجودة |
| `odoo_delete` | حذف نهائي — **مقفول افتراضيًا** |
| `odoo_call` | نداء أي ميثود بأودو (تأكيد أمر بيع، إلخ) |

---

## الحماية

| المتغير | الافتراضي | الوظيفة |
|---|---|---|
| `ODOO_ALLOW_WRITE` | `1` | الإنشاء والتعديل ونداء الميثودات |
| `ODOO_ALLOW_DELETE` | `0` | الحذف النهائي — لازم تفعّله يدويًا |
| `ODOO_MAX_LIMIT` | `1000` | سقف عدد السجلات بأي استعلام |
| `ODOO_ALLOW_INSECURE` | `0` | السماح بـ `http://` (لسيرفر محلي بس) |

`odoo_call` بيعتبر أي ميثود مش معروفة كقراءة **عملية كتابة**، فبينحكم عليها بنفس
المفتاح. الحذف مقفول لحاله لأن الأفضل بأودو تعمل أرشفة (`active = false`) بدل
الحذف — وأودو أصلًا بيرفض حذف أي سجل مربوط بقيود محاسبية أو حركات مخزون.

للتشغيل بوضع قراءة فقط: `ODOO_ALLOW_WRITE=0`.

---

## أمثلة أسئلة

- "شو مبيعات هالشهر مقارنة بالشهر يلي قبله؟"
- "أعطيني أكثر ١٠ منتجات مبيعًا من الـ POS هالأسبوع"
- "أي صنف مخزونه أقل من ١٠؟"
- "ضيف زبون جديد اسمه أبو محمد، تلفون 0791234567"
- "شو الفواتير يلي عليها ذمم من أكثر من ٣٠ يوم؟"

---

## حل المشاكل

| الرسالة | السبب |
|---|---|
| `Missing required environment variable(s)` | ما انقرأ الـ `.env` أو ناقصه قيم |
| `Login failed for ...` | إيميل غلط أو الـ API Key غلط/ملغي |
| `database "..." does not exist` | اسم الداتابيز غلط — عنا هو `alawn` مش `alawn.odoo.com` |
| `... is blocked: this server runs read-only` | `ODOO_ALLOW_WRITE=0` |
| `Deleting records is blocked` | فعّل `ODOO_ALLOW_DELETE=1` عن قصد |

---

## التطوير

```bash
uv pip install -e ".[dev]"
uv run pytest -q

# مع فحص السيرفر الحقيقي كمان (بدون ما تلزم credentials صحيحة):
ODOO_SMOKE_URL=https://alawn.odoo.com ODOO_SMOKE_DB=alawn uv run pytest -q
```
