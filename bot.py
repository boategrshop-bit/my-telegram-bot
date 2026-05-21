import os
import json
import re
from datetime import datetime, timedelta
import anthropic
import gspread
from google.oauth2.service_account import Credentials
from telegram import Update
from telegram.ext import Application, MessageHandler, filters, ContextTypes, JobQueue
import pytz

# ========== CONFIG ==========
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")
GOOGLE_CREDENTIALS = os.environ.get("GOOGLE_CREDENTIALS")
SHEET_ID = os.environ.get("SHEET_ID")
OWNER_CHAT_ID = os.environ.get("OWNER_CHAT_ID")  # chat_id ของเจ้าของ
TZ = pytz.timezone("Asia/Bangkok")

client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

PRODUCTS = {
    "5000":  {"name": "5000 เครดิต",  "cost": 1000, "price": 1490},
    "10000": {"name": "10000 เครดิต", "cost": 1300, "price": 1690},
    "25000": {"name": "25000 เครดิต", "cost": 1500, "price": 2190},
    "grok":  {"name": "Grok",         "cost": 250,  "price": 599},
}

# ========== GOOGLE SHEETS ==========

def get_gc():
    creds_dict = json.loads(GOOGLE_CREDENTIALS)
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)

def init_sheets():
    gc = get_gc()
    sh = gc.open_by_key(SHEET_ID)
    try:
        sh.worksheet("ลูกค้า")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet("ลูกค้า", rows=1000, cols=10)
        ws.append_row(["ID", "ชื่อ", "อีเมล", "เบอร์โทร", "แพ็กเกจ",
                       "ราคาขาย", "ต้นทุน", "วันสมัคร", "วันหมดอายุ", "สถานะ"])
        ws.format("A1:J1", {
            "backgroundColor": {"red": 0.18, "green": 0.46, "blue": 0.71},
            "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
            "horizontalAlignment": "CENTER"
        })
    try:
        sh.worksheet("บัญชี")
    except gspread.WorksheetNotFound:
        ws2 = sh.add_worksheet("บัญชี", rows=1000, cols=7)
        ws2.append_row(["วันที่", "อีเมลลูกค้า", "แพ็กเกจ",
                        "ราคาขาย (บาท)", "ต้นทุน (บาท)", "กำไร (บาท)", "% กำไร"])
        ws2.format("A1:G1", {
            "backgroundColor": {"red": 0.22, "green": 0.34, "blue": 0.14},
            "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
            "horizontalAlignment": "CENTER"
        })
    return sh

def add_orders_batch(orders_list, order_date=None):
    sh = init_sheets()
    today = datetime.now(TZ)
    order_dt = datetime.strptime(order_date, "%Y-%m-%d").replace(tzinfo=TZ) if order_date else today
    today_str = order_dt.strftime("%Y-%m-%d")
    expire = (order_dt + timedelta(days=30)).strftime("%Y-%m-%d")

    ws_c = sh.worksheet("ลูกค้า")
    ws_a = sh.worksheet("บัญชี")

    existing_ids = ws_c.col_values(1)[1:]
    numeric = [int(i) for i in existing_ids if str(i).isdigit()]
    next_id = max(numeric) + 1 if numeric else 1
    next_row_a = len(ws_a.col_values(1)) + 1

    rows_c, rows_a, results, errors = [], [], [], []

    for o in orders_list:
        pkg = PRODUCTS.get(str(o.get("package_key", "")).lower())
        if not pkg:
            errors.append({"email": o.get("email"), "error": "ไม่พบแพ็กเกจ"})
            continue
        email = o.get("email", "")
        profit = pkg["price"] - pkg["cost"]
        margin = profit / pkg["price"] * 100

        rows_c.append([next_id, "", email, "", pkg["name"],
                       pkg["price"], pkg["cost"], today_str, expire, "ใช้งานได้"])

        profit_f = f"=D{next_row_a}-E{next_row_a}"
        margin_f = f"=IF(D{next_row_a}=0,0,F{next_row_a}/D{next_row_a})"
        rows_a.append([today_str, email, pkg["name"],
                       pkg["price"], pkg["cost"], profit_f, margin_f])

        results.append({
            "id": next_id, "email": email, "package": pkg["name"],
            "price": pkg["price"], "cost": pkg["cost"],
            "profit": profit, "margin": margin,
            "expire": expire, "date": today_str
        })
        next_id += 1
        next_row_a += 1

    if rows_c:
        ws_c.append_rows(rows_c)
    if rows_a:
        ws_a.append_rows(rows_a, value_input_option="USER_ENTERED")

    return results, errors

def add_order(email, package_key, name="", phone="", order_date=None):
    results, errors = add_orders_batch(
        [{"email": email, "package_key": package_key}], order_date=order_date)
    if errors:
        return None, errors[0]["error"]
    return results[0], None

def daily_summary(date_str=None):
    sh = init_sheets()
    ws = sh.worksheet("บัญชี")
    rows = ws.get_all_records()
    if not date_str:
        date_str = datetime.now(TZ).strftime("%Y-%m-%d")
    day_rows = [r for r in rows if str(r.get("วันที่", ""))[:10] == date_str]
    total_price = sum(float(r.get("ราคาขาย (บาท)", 0) or 0) for r in day_rows)
    total_cost  = sum(float(r.get("ต้นทุน (บาท)", 0) or 0) for r in day_rows)
    total_profit = total_price - total_cost
    margin = (total_profit / total_price * 100) if total_price > 0 else 0
    return {"date": date_str, "orders": len(day_rows),
            "total_price": total_price, "total_cost": total_cost,
            "total_profit": total_profit, "margin": margin}

def overall_summary():
    sh = init_sheets()
    ws_c = sh.worksheet("ลูกค้า")
    ws_a = sh.worksheet("บัญชี")
    today = datetime.now(TZ).date()

    customers = ws_c.get_all_records()
    total = len(customers)
    active = 0
    expired = 0
    for r in customers:
        exp_str = str(r.get("วันหมดอายุ", ""))[:10]
        try:
            exp_date = datetime.strptime(exp_str, "%Y-%m-%d").date()
            if exp_date >= today:
                active += 1
            else:
                expired += 1
        except:
            expired += 1

    transactions = ws_a.get_all_records()
    total_price = sum(float(r.get("ราคาขาย (บาท)", 0) or 0) for r in transactions)
    total_cost  = sum(float(r.get("ต้นทุน (บาท)", 0) or 0) for r in transactions)
    profit = total_price - total_cost
    margin = (profit / total_price * 100) if total_price > 0 else 0
    return {"total_customers": total, "active": active, "expired": expired,
            "total_price": total_price, "total_cost": total_cost,
            "profit": profit, "margin": margin}

def check_expiring(days=7, target_date=None):
    sh = init_sheets()
    ws = sh.worksheet("ลูกค้า")
    rows = ws.get_all_records()
    today = datetime.now(TZ).date()
    expiring = []

    for row in rows:
        exp_str = str(row.get("วันหมดอายุ", ""))[:10]
        try:
            exp_date = datetime.strptime(exp_str, "%Y-%m-%d").date()

            if target_date:
                # ถามเฉพาะวันที่ระบุ
                target = datetime.strptime(target_date, "%Y-%m-%d").date()
                if exp_date == target:
                    diff = (exp_date - today).days
                    expiring.append({
                        "email": row.get("อีเมล"),
                        "package": row.get("แพ็กเกจ"),
                        "expire": exp_str,
                        "days_left": diff
                    })
            else:
                # ถามช่วงกี่วันข้างหน้า (รวมที่หมดไปแล้วด้วยถ้า days < 0)
                diff = (exp_date - today).days
                if 0 <= diff <= days:
                    expiring.append({
                        "email": row.get("อีเมล"),
                        "package": row.get("แพ็กเกจ"),
                        "expire": exp_str,
                        "days_left": diff
                    })
        except:
            pass
    return expiring

def get_expired_today():
    sh = init_sheets()
    ws = sh.worksheet("ลูกค้า")
    rows = ws.get_all_records()
    today_str = datetime.now(TZ).strftime("%Y-%m-%d")
    return [r for r in rows if str(r.get("วันหมดอายุ", ""))[:10] == today_str]

def search_customer(keyword):
    sh = init_sheets()
    ws = sh.worksheet("ลูกค้า")
    rows = ws.get_all_records()
    results = []
    for row in rows:
        if (keyword.lower() in str(row.get("ชื่อ", "")).lower() or
            keyword.lower() in str(row.get("อีเมล", "")).lower() or
            keyword in str(row.get("เบอร์โทร", ""))):
            results.append(row)
    return results

# ========== CLAUDE AI ==========

PRODUCT_LIST = "\n".join([f'- "{k}": {v["name"]} ราคาขาย {v["price"]} ต้นทุน {v["cost"]} บาท'
                           for k, v in PRODUCTS.items()])

def get_system_prompt():
    today = datetime.now(TZ).strftime("%Y-%m-%d")
    return f"""คุณคือผู้ช่วยธุรกิจ ตอบกลับเป็น JSON เท่านั้น ห้ามมีข้อความอื่น

FORMAT:
{{"action": "...", "data": {{}}, "reply": "..."}}

ACTION:
- "add_order": ออเดอร์เดียว (data: email, package_key, order_date YYYY-MM-DD หรือ "")
- "add_orders_bulk": หลายออเดอร์ (data: order_date, orders: [{{email, package_key}}])
- "daily_summary": รวมยอดวัน (data: date YYYY-MM-DD หรือ "")
- "summary": รวมยอดทั้งหมด
- "check_expiring": ดูใกล้หมดอายุ (data: days จำนวนวัน) หรือระบุวันที่ (data: target_date "YYYY-MM-DD")
  ตัวอย่าง: "วันที่ 21 พ.ค. มีใครหมดอายุ" → check_expiring, target_date: "{today}"
  ตัวอย่าง: "ใครหมดอายุสัปดาห์นี้" → check_expiring, days: 7
- "search": ค้นหา (data: keyword)
- "chat": แค่คุย

แพ็กเกจ:
{PRODUCT_LIST}

กฎอ่านออเดอร์:
- O1-5000 old done ✅ → package_key: "5000"
- O2-grok new → package_key: "grok"
- บรรทัดมี @ = อีเมลของออเดอร์บรรทัดก่อน
- คำว่า old/new/done/✅/❌ ละเว้นทั้งหมด
- หลายออเดอร์ → add_orders_bulk
- ย้อนหลัง เช่น "วันที่ 17 พ.ค." → order_date: "2026-05-17"

วันนี้คือ {today} (เวลาไทย Asia/Bangkok)
ห้ามใช้วันอื่นนอกจากนี้ ถ้าไม่มีการระบุวันย้อนหลัง
"""

def ask_claude(user_message):
    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=800,
        system=get_system_prompt(),
        messages=[{"role": "user", "content": user_message}]
    )
    text = re.sub(r"```json|```", "", response.content[0].text.strip()).strip()
    return json.loads(text)

# ========== แจ้งกำไรต่อวันทุกออเดอร์ ==========

async def get_daily_profit_line(order_date=None):
    s = daily_summary(order_date)
    return f"📊 กำไรวันนี้รวม {s['orders']} ออเดอร์: {s['total_profit']:,.0f} บาท ({s['margin']:.1f}%)"

# ========== MORNING ALERT 08:00 ==========

async def morning_alert(context):
    if not OWNER_CHAT_ID:
        return
    today_str = datetime.now(TZ).strftime("%Y-%m-%d")
    expired_today = get_expired_today()

    if not expired_today:
        msg = f"🌅 {today_str}\n✅ วันนี้ไม่มีลูกค้าหมดอายุ"
    else:
        lines = [f"🌅 {today_str} — หมดอายุวันนี้ {len(expired_today)} ราย:\n"]
        for r in expired_today:
            lines.append(f"• 📧 {r.get('อีเมล')} | 📦 {r.get('แพ็กเกจ')}")
        msg = "\n".join(lines)

    await context.bot.send_message(chat_id=int(OWNER_CHAT_ID), text=msg)

# ========== TELEGRAM HANDLER ==========

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_msg = update.message.text
    await update.message.chat.send_action("typing")

    try:
        result = ask_claude(user_msg)
        action = result.get("action")
        data = result.get("data", {})
        reply = ""

        if action == "add_order":
            order, err = add_order(
                email=data.get("email", ""),
                package_key=data.get("package_key", ""),
                order_date=data.get("order_date") or None,
            )
            if err:
                reply = f"❌ {err}"
            else:
                daily_line = await get_daily_profit_line(order.get("date"))
                date_label = f" (ย้อนหลัง {order['date']})" if data.get("order_date") else ""
                reply = (
                    f"✅ บันทึกออเดอร์แล้ว{date_label}!\n"
                    f"📧 {order['email']}\n"
                    f"📦 {order['package']}\n"
                    f"💰 ราคาขาย: {order['price']:,} บาท\n"
                    f"💸 ต้นทุน: {order['cost']:,} บาท\n"
                    f"📈 กำไรออเดอร์นี้: {order['profit']:,} บาท ({order['margin']:.1f}%)\n"
                    f"📅 หมดอายุ: {order['expire']}\n\n"
                    f"{daily_line}"
                )

        elif action == "add_orders_bulk":
            orders = data.get("orders", [])
            order_date = data.get("order_date") or None
            results, errors = add_orders_batch(orders, order_date=order_date)

            total_profit = sum(o["profit"] for o in results)
            date_label = f" (ย้อนหลัง {order_date})" if order_date else ""
            lines = [f"📋 บันทึก {len(results)}/{len(orders)} ออเดอร์{date_label}\n"]
            for o in results:
                lines.append(f"✅ {o['email']} | {o['package']} | กำไร {o['profit']:,} บาท")
            for e in errors:
                lines.append(f"❌ {e['email']} - {e['error']}")

            daily_line = await get_daily_profit_line(order_date)
            lines.append(f"\n📈 กำไรชุดนี้: {total_profit:,} บาท")
            lines.append(daily_line)
            reply = "\n".join(lines)

        elif action == "daily_summary":
            date_str = data.get("date", "") or datetime.now(TZ).strftime("%Y-%m-%d")
            s = daily_summary(date_str)
            reply = (
                f"📊 ยอดวันที่ {s['date']}\n"
                f"🛒 ออเดอร์: {s['orders']} รายการ\n"
                f"💰 รายรับรวม: {s['total_price']:,.0f} บาท\n"
                f"💸 ต้นทุนรวม: {s['total_cost']:,.0f} บาท\n"
                f"📈 กำไรสุทธิ: {s['total_profit']:,.0f} บาท ({s['margin']:.1f}%)"
            )

        elif action == "summary":
            s = overall_summary()
            reply = (
                f"📊 สรุปยอดรวมทั้งหมด\n"
                f"👥 ลูกค้า: {s['total_customers']} ราย "
                f"(ใช้งาน {s['active']} / หมดอายุ {s['expired']})\n"
                f"💰 รายรับรวม: {s['total_price']:,.0f} บาท\n"
                f"💸 ต้นทุนรวม: {s['total_cost']:,.0f} บาท\n"
                f"📈 กำไรสุทธิ: {s['profit']:,.0f} บาท ({s['margin']:.1f}%)"
            )

        elif action == "check_expiring":
            days = data.get("days", 7)
            target_date = data.get("target_date") or None
            expiring = check_expiring(days=days, target_date=target_date)

            label = f"วันที่ {target_date}" if target_date else f"ใน {days} วันข้างหน้า"
            if expiring:
                lines = [f"⚠️ หมดอายุ{label} ({len(expiring)} ราย):\n"]
                for c in expiring:
                    days_label = f"(อีก {c['days_left']} วัน)" if c['days_left'] >= 0 else f"(หมดไปแล้ว {abs(c['days_left'])} วัน)"
                    lines.append(f"• 📧 {c['email']} | 📦 {c['package']} {days_label}")
                reply = "\n".join(lines)
            else:
                reply = f"✅ ไม่มีลูกค้าหมดอายุ{label}"

        elif action == "search":
            keyword = data.get("keyword", "")
            results = search_customer(keyword)
            if results:
                lines = [f"🔍 พบ {len(results)} ราย:\n"]
                for r in results:
                    lines.append(
                        f"• {r.get('อีเมล')} | {r.get('แพ็กเกจ')}\n"
                        f"  หมด: {str(r.get('วันหมดอายุ',''))[:10]} | {r.get('สถานะ')}"
                    )
                reply = "\n".join(lines)
            else:
                reply = f"❌ ไม่พบลูกค้า '{keyword}'"

        else:
            reply = result.get("reply", "✅ เสร็จแล้วครับ")

        await update.message.reply_text(reply or "✅ เสร็จแล้วครับ")

    except json.JSONDecodeError:
        await update.message.reply_text("❌ ระบบไม่เข้าใจคำสั่ง ลองพิมพ์ใหม่นะครับ")
    except Exception as e:
        await update.message.reply_text(f"❌ เกิดข้อผิดพลาด: {str(e)}")

def main():
    init_sheets()
    print("🤖 Bot กำลังทำงาน...")
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

    # แจ้งเตือนทุกเช้า 08:00 เวลาไทย
    if OWNER_CHAT_ID:
        app.job_queue.run_daily(
            morning_alert,
            time=datetime.strptime("08:00", "%H:%M").replace(tzinfo=TZ).timetz(),
            name="morning_alert"
        )

    app.run_polling()

if __name__ == "__main__":
    main()
