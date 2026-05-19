import os
import json
import re
from datetime import datetime, timedelta
import anthropic
import gspread
from google.oauth2.service_account import Credentials
from telegram import Update
from telegram.ext import Application, MessageHandler, filters, ContextTypes

# ========== CONFIG ==========
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")
GOOGLE_CREDENTIALS = os.environ.get("GOOGLE_CREDENTIALS")
SHEET_ID = os.environ.get("SHEET_ID")

client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# ========== แคตตาล็อกสินค้า ==========
PRODUCTS = {
    "5000": {"name": "5000 เครดิต", "cost": 1000, "price": 1490},
    "10000": {"name": "10000 เครดิต", "cost": 1200, "price": 1690},
    "25000": {"name": "25000 เครดิต", "cost": 1500, "price": 2190},
    "grok": {"name": "Grok", "cost": 250, "price": 599},
}

# ========== GOOGLE SHEETS ==========

def get_gc():
    creds_dict = json.loads(GOOGLE_CREDENTIALS)
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)

def init_sheets():
    gc = get_gc()
    sh = gc.open_by_key(SHEET_ID)

    # Sheet ลูกค้า
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

    # Sheet บัญชี
    try:
        sh.worksheet("บัญชี")
    except gspread.WorksheetNotFound:
        ws2 = sh.add_worksheet("บัญชี", rows=1000, cols=7)
        ws2.append_row(["วันที่", "อีเมลลูกค้า", "แพ็กเกจ", "ราคาขาย", "ต้นทุน", "กำไร", "% กำไร"])
        ws2.format("A1:G1", {
            "backgroundColor": {"red": 0.22, "green": 0.34, "blue": 0.14},
            "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
            "horizontalAlignment": "CENTER"
        })

    return sh

def get_next_id(sh):
    ws = sh.worksheet("ลูกค้า")
    ids = ws.col_values(1)[1:]
    numeric = [int(i) for i in ids if str(i).isdigit()]
    return max(numeric) + 1 if numeric else 1

def add_order(email, package_key, name="", phone="", note=""):
    """เพิ่มออเดอร์ใหม่ — บันทึกทั้ง sheet ลูกค้า และ บัญชี"""
    sh = init_sheets()
    today = datetime.now()
    today_str = today.strftime("%Y-%m-%d")

    # หาข้อมูลสินค้า
    pkg = PRODUCTS.get(package_key.lower())
    if not pkg:
        return None, "ไม่พบแพ็กเกจนี้ในระบบ"

    # คำนวณวันหมดอายุ (30 วัน)
    expire = (today + timedelta(days=30)).strftime("%Y-%m-%d")

    # === บันทึก Sheet ลูกค้า ===
    ws_c = sh.worksheet("ลูกค้า")
    cid = get_next_id(sh)
    row_c = [cid, name, email, phone, pkg["name"],
             pkg["price"], pkg["cost"], today_str, expire, "ใช้งานได้"]
    ws_c.append_row(row_c)
    last_c = len(ws_c.col_values(1))
    color = {"red": 0.87, "green": 0.92, "blue": 0.95} if last_c % 2 == 0 else {"red": 1, "green": 1, "blue": 1}
    ws_c.format(f"A{last_c}:J{last_c}", {"backgroundColor": color})

    # === บันทึก Sheet บัญชี ===
    ws_a = sh.worksheet("บัญชี")
    last_a = len(ws_a.col_values(1)) + 1
    profit_formula = f"=D{last_a}-E{last_a}"
    margin_formula = f"=IF(D{last_a}=0,0,F{last_a}/D{last_a})"
    row_a = [today_str, email, pkg["name"], pkg["price"], pkg["cost"], profit_formula, margin_formula]
    ws_a.append_row(row_a, value_input_option="USER_ENTERED")
    ws_a.format(f"G{last_a}", {"numberFormat": {"type": "PERCENT", "pattern": "0.0%"}})
    color2 = {"red": 0.89, "green": 0.94, "blue": 0.85} if last_a % 2 == 0 else {"red": 1, "green": 1, "blue": 1}
    ws_a.format(f"A{last_a}:G{last_a}", {"backgroundColor": color2})

    profit = pkg["price"] - pkg["cost"]
    margin = profit / pkg["price"] * 100

    return {
        "id": cid, "email": email, "package": pkg["name"],
        "price": pkg["price"], "cost": pkg["cost"],
        "profit": profit, "margin": margin,
        "expire": expire
    }, None

def daily_summary(date_str=None):
    """รวมยอดกำไรตามวัน"""
    sh = init_sheets()
    ws = sh.worksheet("บัญชี")
    rows = ws.get_all_records()

    if not date_str:
        date_str = datetime.now().strftime("%Y-%m-%d")

    day_rows = [r for r in rows if str(r.get("วันที่", ""))[:10] == date_str]

    total_price = sum(float(r.get("ราคาขาย", 0) or 0) for r in day_rows)
    total_cost = sum(float(r.get("ต้นทุน", 0) or 0) for r in day_rows)
    total_profit = total_price - total_cost
    margin = (total_profit / total_price * 100) if total_price > 0 else 0

    return {
        "date": date_str,
        "orders": len(day_rows),
        "total_price": total_price,
        "total_cost": total_cost,
        "total_profit": total_profit,
        "margin": margin
    }

def overall_summary():
    """รวมยอดทั้งหมด"""
    sh = init_sheets()
    ws_c = sh.worksheet("ลูกค้า")
    ws_a = sh.worksheet("บัญชี")
    today = datetime.now()

    customers = ws_c.get_all_records()
    total = len(customers)
    active = sum(1 for r in customers if str(r.get("สถานะ", "")) == "ใช้งานได้")
    expired = total - active

    transactions = ws_a.get_all_records()
    total_price = sum(float(r.get("ราคาขาย", 0) or 0) for r in transactions)
    total_cost = sum(float(r.get("ต้นทุน", 0) or 0) for r in transactions)
    profit = total_price - total_cost
    margin = (profit / total_price * 100) if total_price > 0 else 0

    return {
        "total_customers": total, "active": active, "expired": expired,
        "total_price": total_price, "total_cost": total_cost,
        "profit": profit, "margin": margin
    }

def check_expiring(days=7):
    sh = init_sheets()
    ws = sh.worksheet("ลูกค้า")
    rows = ws.get_all_records()
    today = datetime.now()
    expiring = []
    for row in rows:
        try:
            exp = datetime.strptime(str(row.get("วันหมดอายุ", ""))[:10], "%Y-%m-%d")
            diff = (exp - today).days
            if 0 <= diff <= days:
                expiring.append({
                    "name": row.get("ชื่อ"), "email": row.get("อีเมล"),
                    "phone": row.get("เบอร์โทร"), "package": row.get("แพ็กเกจ"),
                    "expire": str(row.get("วันหมดอายุ"))[:10], "days_left": diff
                })
        except:
            pass
    return expiring

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

PRODUCT_LIST = "\n".join([f'- "{k}": {v["name"]} ราคาขาย {v["price"]} บาท ต้นทุน {v["cost"]} บาท'
                           for k, v in PRODUCTS.items()])

SYSTEM_PROMPT = f"""คุณคือผู้ช่วยธุรกิจ ตอบกลับเป็น JSON เท่านั้น ห้ามมีข้อความอื่น

FORMAT:
{{"action": "...", "data": {{}}, "reply": "..."}}

ACTION:
- "add_order": รับออเดอร์ใหม่ (data: email, package_key, name?, phone?)
- "daily_summary": รวมยอดกำไรวันที่ระบุ (data: date YYYY-MM-DD หรือ "" สำหรับวันนี้)
- "summary": รวมยอดทั้งหมด
- "check_expiring": ใกล้หมดอายุ (data: days)
- "search": ค้นหาลูกค้า (data: keyword)
- "chat": แค่คุย

แพ็กเกจที่มี (ใช้ package_key ตามนี้):
{PRODUCT_LIST}

ตัวอย่าง:
- "somchai@gmail.com แพ็ก 5000" → add_order, email: somchai@gmail.com, package_key: "5000"
- "somchai@gmail.com grok" → add_order, package_key: "grok"
- "รวมยอดวันนี้" → daily_summary, date: ""
- "รวมยอดวันที่ 18 พ.ค." → daily_summary, date: "2026-05-18"
- "สรุปยอดรวม" → summary
- "ใครหมดอายุสัปดาห์นี้" → check_expiring, days: 7

วันนี้คือ {datetime.now().strftime("%Y-%m-%d")}
"""

def ask_claude(user_message):
    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=500,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_message}]
    )
    text = re.sub(r"```json|```", "", response.content[0].text.strip()).strip()
    return json.loads(text)

# ========== TELEGRAM ==========

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
                name=data.get("name", ""),
                phone=data.get("phone", ""),
            )
            if err:
                reply = f"❌ {err}"
            else:
                reply = (
                    f"✅ บันทึกออเดอร์แล้ว!\n"
                    f"📧 {order['email']}\n"
                    f"📦 {order['package']}\n"
                    f"💰 ราคาขาย: {order['price']:,} บาท\n"
                    f"💸 ต้นทุน: {order['cost']:,} บาท\n"
                    f"📈 กำไรออเดอร์นี้: {order['profit']:,} บาท ({order['margin']:.1f}%)\n"
                    f"📅 หมดอายุ: {order['expire']}"
                )

        elif action == "daily_summary":
            date_str = data.get("date", "") or datetime.now().strftime("%Y-%m-%d")
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
                f"👥 ลูกค้า: {s['total_customers']} ราย (ใช้งาน {s['active']} / หมดอายุ {s['expired']})\n"
                f"💰 รายรับรวม: {s['total_price']:,.0f} บาท\n"
                f"💸 ต้นทุนรวม: {s['total_cost']:,.0f} บาท\n"
                f"📈 กำไรสุทธิ: {s['profit']:,.0f} บาท ({s['margin']:.1f}%)"
            )

        elif action == "check_expiring":
            days = data.get("days", 7)
            expiring = check_expiring(days)
            if expiring:
                lines = [f"⚠️ ใกล้หมดอายุใน {days} วัน ({len(expiring)} ราย):\n"]
                for c in expiring:
                    lines.append(f"• {c['name'] or c['email']} | {c['email']} - หมด {c['expire']} (อีก {c['days_left']} วัน)")
                reply = "\n".join(lines)
            else:
                reply = f"✅ ไม่มีลูกค้าหมดอายุใน {days} วันข้างหน้า"

        elif action == "search":
            keyword = data.get("keyword", "")
            results = search_customer(keyword)
            if results:
                lines = [f"🔍 พบ {len(results)} ราย:\n"]
                for r in results:
                    lines.append(
                        f"• {r.get('ชื่อ') or '-'} | {r.get('อีเมล')}\n"
                        f"  แพ็ก: {r.get('แพ็กเกจ')} | หมด: {str(r.get('วันหมดอายุ',''))[:10]} | {r.get('สถานะ')}"
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
    app.run_polling()

if __name__ == "__main__":
    main()
