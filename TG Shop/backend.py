import os, json
from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
import httpx
from dotenv import load_dotenv

from tg_auth import validate_init_data, TelegramAuthError

load_dotenv()
BOT_TOKEN = os.environ["BOT_TOKEN"]
ADMIN_CHAT_ID = int(os.environ.get("ADMIN_CHAT_ID", "0"))

app = FastAPI()
app.mount("/web", StaticFiles(directory="web"), name="web")

@app.get("/")
def root():
    return FileResponse("web/index.html")

PRODUCTS = [
    {"id": "p1", "name": "Футболка", "price": 1990},
    {"id": "p2", "name": "Кепка", "price": 1490},
    {"id": "p3", "name": "Худи", "price": 4990},
]

@app.get("/api/products")
def products():
    return PRODUCTS

class OrderItem(BaseModel):
    id: str
    qty: int

class OrderRequest(BaseModel):
    initData: str
    items: list[OrderItem]

@app.post("/api/order")
async def create_order(req: OrderRequest):
    try:
        data = validate_init_data(req.initData, BOT_TOKEN)
    except TelegramAuthError as e:
        raise HTTPException(status_code=401, detail=str(e))

    user_json = data.get("user")
    user = json.loads(user_json) if user_json else {}
    user_id = user.get("id")

    price_map = {p["id"]: p["price"] for p in PRODUCTS}
    lines = []
    total = 0
    for it in req.items:
        if it.id not in price_map:
            raise HTTPException(400, detail=f"Unknown product: {it.id}")
        if it.qty <= 0:
            raise HTTPException(400, detail="qty must be > 0")
        cost = price_map[it.id] * it.qty
        total += cost
        lines.append(f"- {it.id} × {it.qty} = {cost}₽")

    text = (
        f"🧾 Новый заказ\n"
        f"От: {user.get('first_name','')} (id={user_id})\n"
        f"{chr(10).join(lines)}\n"
        f"Итого: {total}₽"
    )

    async with httpx.AsyncClient(timeout=10) as client:
        if user_id:
            await client.post(
                f"https://api.telegram.org/bot{BOT_TOKEN}/sendMessage",
                json={"chat_id": user_id, "text": f"✅ Заказ принят!\nИтого: {total}₽"}
            )
        if ADMIN_CHAT_ID:
            await client.post(
                f"https://api.telegram.org/bot{BOT_TOKEN}/sendMessage",
                json={"chat_id": ADMIN_CHAT_ID, "text": text}
            )

    return {"ok": True, "total": total}
