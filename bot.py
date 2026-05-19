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
GOOGLE_CREDENTIALS = os.environ.get("GOOGLE_CREDENTIALS")  # JSON string
SHEET_ID = os.environ.get("SHEET_ID")  # Google Sheet ID

client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# ========== GOOGLE SHEETS SETUP ==========

def get_sheets_client():
    creds_dict = json.loads(GOOGLE_CREDENTIALS)
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)

def init_sheets():
    gc = get_sheets_client()
    sh = gc.open_by_key(SHEET_ID)

    # Sheet ลูกค้า
    try:
        ws1 = sh.worksheet("ลูกค้า")
    except gspread.WorksheetNotFound:
        ws1 = sh.add_worksheet("ลูกค้า", rows=1000, cols=9)
        ws1.append_row(["ID", "ชื่อ", "เบอร์โทร", "แพ็กเกจ", "ยอดเงิน (บาท)",
                        "วันสมัคร", "วันหมดอายุ", "สถานะ", "หมายเหตุ"])
        ws1.format("A1:I1", {
            "backgroundColor": {"red": 0.18, "green": 0.46, "blue": 0.71},
            "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
            "horizontalAlignment": "CENTER"
        })

    # Sheet บัญชี
    try:
        ws2 = sh.worksheet("บัญชี")
    except gspread.WorksheetNotFound:
        ws2 = sh.add_worksheet("บัญชี", rows=1000, cols=8)
        ws2.append_row(["วันที่", "รายการ", "ประเภท", "ราคาขาย (บาท)",
                        "ต้นทุน (บาท)", "กำไร (บาท)", "% กำไร", "หมายเหตุ"])
        ws2.format("A1:H1", {
            "backgroundColor": {"red": 0.22, "green": 0.34, "blue": 0.14},
            "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
            "horizontalAlignment": "CENTER"
        })

    return sh

def get_next_id(sh):
    ws = sh.worksheet("ลูกค้า")
    ids = ws.col_values(1)[1:]
    if not ids:
        return 1
    numeric = [int(i) for i in ids if str(i).isdigit()]
    return max(numeric) + 1 if numeric else 1

def add_customer(name, phone, package, amount, start_date, expire_date, note=""):
    sh = init_sheets()
    ws = sh.worksheet("ลูกค้า")
    cid = get_next_id(sh)
    today = datetime.now()

    try:
        exp = datetime.strptime(expire_date, "%Y-%m-%d")
    except:
        exp = today + timedelta(days=30)

    status = "ใช้งานได้" if exp >= today else "หมดอายุ"
    row = [cid, name, phone, package, float(amount),
           start_date, exp.strftime("%Y-%m-%d"), status, note]
    ws.append_row(row)

    last_row = len(ws.col_values(1))
    color = {"red": 0.87, "green": 0.92, "blue": 0.95} if last_row % 2 == 0 else {"red": 1, "green": 1, "blue": 1}
    ws.format(f"A{last_row}:I{last_row}", {"backgroundColor": color})

    return cid

def add_transaction(date, name, ttype, selling_price, cost, note=""):
    sh = init_sheets()
    ws = sh.worksheet("บัญชี")
    last_row = len(ws.col_values(1)) + 1

    profit_formula = f"=D{last_row}-E{last_row}"
    margin_formula = f"=IF(D{last_row}=0,0,F{last_row}/D{last_row})"

    ws.append_row([date, name, ttype, float(selling_price), float(cost),
                   profit_formula, margin_formula, note],
                  value_input_option="USER_ENTERED")

    ws.format(f"G{last_row}", {"numberFormat": {"type": "PERCENT", "pattern": "0.0%"}})
    color = {"red": 0.89, "green": 0.94, "blue": 0.85} if last_row % 2 == 0 else {"red": 1, "green": 1, "blue": 1}
    ws.format(f"A{last_row}:H{last_row}", {"backgroundColor": color})

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
                    "id": row.get("ID"), "name": row.get("ชื่อ"),
                    "phone": row.get("เบอร์โทร"), "package": row.get("แพ็กเกจ"),
                    "expire": str(row.get("วันหมดอายุ"))[:10], "days_left": diff
                })
        except:
            pass
    return expiring

def get_summary():
    sh = init_sheets()
    ws_c = sh.worksheet("ลูกค้า")
    ws_a = sh.worksheet("บัญชี")
    today = datetime.now()

    customers = ws_c.get_all_records()
    total_customers = len(customers)
    active = expired = 0
    for row in customers:
        try:
            exp = datetime.strptime(str(row.get("วันหมดอายุ", ""))[:10], "%Y-%m-%d")
            if exp >= today:
                active += 1
            else:
                expired += 1
        except:
            pass

    transactions = ws_a.get_all_records()
    total_income = sum(float(r.get("ราคาขาย (บาท)", 0) or 0) for r in transactions)
    total_cost = sum(float(r.get("ต้นทุน (บาท)", 0) or 0) for r in transactions)
    profit = total_income - total_cost
    margin = (profit / total_income * 100) if total_income > 0 else 0

    return {
        "total_customers": total_customers, "active": active, "expired": expired,
        "total_income": total_income, "total_cost": total_cost,
        "profit": profit, "margin": margin
    }

def search_customer(keyword):
    sh = init_sheets()
    ws = sh.worksheet("ลูกค้า")
    rows = ws.get_all_records()
    results = []
    for row in rows:
        name = str(row.get("ชื่อ", "")).lower()
        phone = str(row.get("เบอร์โทร", ""))
        if keyword.lower() in name or keyword in phone:
            results.append({
                "id": row.get("ID"), "name": row.get("ชื่อ"),
                "phone": row.get("เบอร์โทร"), "package": row.get("แพ็กเกจ"),
                "amount": row.get("ยอดเงิน (บาท)"),
                "expire": str(row.get("วันหมดอายุ", ""))[:10],
                "status": row.get("สถานะ"), "note": row.get("หมายเหตุ")
            })
    return results

# ========== CLAUDE AI ==========

SYSTEM_PROMPT = """คุณคือผู้ช่วยธุรกิจที่ช่วยจัดการข้อมูลลูกค้า บัญชีรายรับ-รายจ่าย และวันหมดอายุบัญชี
คุณต้องวิเคราะห์ข้อความของผู้ใช้และตอบกลับเป็น JSON เพื่อให้ระบบดำเนินการต่อ

FORMAT การตอบกลับ (JSON เท่านั้น ห้ามมีข้อความอื่น):
{
  "action": "...",
  "data": {},
  "reply": "..."
}

ACTION ที่รองรับ:
- "add_customer": เพิ่มลูกค้าใหม่ (data: name, phone, package, amount, start_date YYYY-MM-DD, expire_date YYYY-MM-DD, note)
- "add_transaction": บันทึกราคาขาย/ต้นทุน (data: date YYYY-MM-DD, name, type, selling_price, cost, note)
- "check_expiring": ดูลูกค้าที่ใกล้หมดอายุ (data: days)
- "summary": ดูสรุปภาพรวม
- "search": ค้นหาลูกค้า (data: keyword)
- "chat": แค่คุย

ตัวอย่าง:
- "เพิ่มลูกค้าใหม่ ชื่อ สมชาย เบอร์ 0812345678 แพ็ก 30 วัน ราคา 500 บาท"
  → action: add_customer, expire_date = start_date + 30 วัน
- "ขายสินค้า A ได้ 2000 ต้นทุน 1200"
  → action: add_transaction, selling_price: 2000, cost: 1200
- "ใครหมดอายุสัปดาห์นี้" → check_expiring, days: 7
- "สรุปยอด" → summary
- "หา สมชาย" → search, keyword: สมชาย

วันนี้คือ """ + datetime.now().strftime("%Y-%m-%d") + """
ถ้าไม่ระบุวันสมัคร ให้ใช้วันนี้
ถ้าบอกจำนวนวันให้คำนวณ expire_date เอง
"""

def ask_claude(user_message):
    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=1000,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_message}]
    )
    text = re.sub(r"```json|```", "", response.content[0].text.strip()).strip()
    return json.loads(text)

# ========== TELEGRAM HANDLER ==========

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_msg = update.message.text
    await update.message.chat.send_action("typing")

    try:
        result = ask_claude(user_msg)
        action = result.get("action")
        data = result.get("data", {})
        reply = result.get("reply", "")

        if action == "add_customer":
            cid = add_customer(
                name=data.get("name", ""),
                phone=data.get("phone", ""),
                package=data.get("package", ""),
                amount=data.get("amount", 0),
                start_date=data.get("start_date", datetime.now().strftime("%Y-%m-%d")),
                expire_date=data.get("expire_date", ""),
                note=data.get("note", "")
            )
            reply = f"✅ เพิ่มลูกค้าสำเร็จ! (ID: {cid})\n" + reply

        elif action == "add_transaction":
            selling_price = data.get("selling_price", data.get("income", 0))
            cost = data.get("cost", 0)
            profit = float(selling_price) - float(cost)
            margin = (profit / float(selling_price) * 100) if float(selling_price) > 0 else 0
            add_transaction(
                date=data.get("date", datetime.now().strftime("%Y-%m-%d")),
                name=data.get("name", ""),
                ttype=data.get("type", ""),
                selling_price=selling_price,
                cost=cost,
                note=data.get("note", "")
            )
            reply = (
                f"✅ บันทึกรายการแล้ว\n"
                f"💰 ราคาขาย: {float(selling_price):,.0f} บาท\n"
                f"💸 ต้นทุน: {float(cost):,.0f} บาท\n"
                f"📈 กำไร: {profit:,.0f} บาท ({margin:.1f}%)"
            )

        elif action == "check_expiring":
            days = data.get("days", 7)
            expiring = check_expiring(days)
            if expiring:
                lines = [f"⚠️ ลูกค้าใกล้หมดอายุภายใน {days} วัน ({len(expiring)} ราย):\n"]
                for c in expiring:
                    lines.append(f"• {c['name']} ({c['phone']}) - หมด {c['expire']} (อีก {c['days_left']} วัน)")
                reply = "\n".join(lines)
            else:
                reply = f"✅ ไม่มีลูกค้าหมดอายุใน {days} วันข้างหน้า"

        elif action == "summary":
            s = get_summary()
            reply = (
                f"📊 สรุปภาพรวม\n"
                f"👥 ลูกค้าทั้งหมด: {s['total_customers']} ราย\n"
                f"   ✅ ใช้งานได้: {s['active']} ราย\n"
                f"   ❌ หมดอายุ: {s['expired']} ราย\n\n"
                f"💰 ราคาขายรวม: {s['total_income']:,.0f} บาท\n"
                f"💸 ต้นทุนรวม: {s['total_cost']:,.0f} บาท\n"
                f"📈 กำไรสุทธิ: {s['profit']:,.0f} บาท\n"
                f"📉 อัตรากำไร: {s['margin']:.1f}%"
            )

        elif action == "search":
            keyword = data.get("keyword", "")
            results = search_customer(keyword)
            if results:
                lines = [f"🔍 ผลการค้นหา '{keyword}' ({len(results)} ราย):\n"]
                for c in results:
                    lines.append(
                        f"• [{c['id']}] {c['name']} | {c['phone']}\n"
                        f"  แพ็ก: {c['package']} | หมด: {c['expire']} | {c['status']}"
                    )
                reply = "\n".join(lines)
            else:
                reply = f"❌ ไม่พบลูกค้าที่ค้นหา '{keyword}'"

        await update.message.reply_text(reply or "✅ เสร็จแล้วครับ")

    except json.JSONDecodeError:
        await update.message.reply_text("❌ ระบบไม่เข้าใจคำสั่ง ลองพิมพ์ใหม่อีกครั้งนะครับ")
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
