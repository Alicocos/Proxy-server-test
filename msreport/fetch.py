"""Выгрузка сырых данных из МойСклад в один словарь."""
from datetime import datetime, timedelta

from .api import MoySklad

FMT = "%Y-%m-%d %H:%M:%S"


def fetch_all(ms: MoySklad, cfg: dict, today: datetime, log=print) -> dict:
    d = cfg["documents"]
    t_to = today.strftime(FMT)
    t30 = (today - timedelta(days=30)).strftime(FMT)
    t90 = (today - timedelta(days=90)).strftime(FMT)

    log("Склады…")
    stores = ms.rows("entity/store")
    log("Остатки…")
    stock_all = ms.rows("report/stock/all", {"stockMode": "all"})
    stock_bystore = ms.rows("report/stock/bystore")
    log("Продажи за 30 и 90 дней…")
    sales30 = ms.rows("report/profit/byvariant", {"momentFrom": t30, "momentTo": t_to})
    sales90 = ms.rows("report/profit/byvariant", {"momentFrom": t90, "momentTo": t_to})
    sales90_by_store = {}
    for st in stores:
        log(f"  продажи: {st['name']}")
        sales90_by_store[st["id"]] = ms.rows(
            "report/profit/byvariant",
            {"momentFrom": t90, "momentTo": t_to, "filter": f"store={st['meta']['href']}"},
        )

    log("Заказы поставщикам…")
    since = (today - timedelta(days=d["orders_lookback_days"])).strftime(FMT)
    orders = ms.rows("entity/purchaseorder", {
        "filter": f"moment>={since}", "expand": "agent,state,rate.currency", "order": "moment,desc",
    })
    log("Платежи…")
    pay_since = (today - timedelta(days=d["payments_days"])).strftime(FMT)
    payments = []
    for kind in ("paymentout", "cashout"):
        payments += ms.rows(f"entity/{kind}", {
            "filter": f"moment>={pay_since}", "expand": "agent,rate.currency,operations", "order": "moment,desc",
        })

    return {
        "stores": stores, "stock_all": stock_all, "stock_bystore": stock_bystore,
        "sales30": sales30, "sales90": sales90, "sales90_by_store": sales90_by_store,
        "orders": orders, "payments": payments,
    }
