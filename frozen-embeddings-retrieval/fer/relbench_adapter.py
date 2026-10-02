"""RelBench recommendation tasks as the published benchmark (UNTESTED here: relbench data is on Hugging Face).

Plan (see docs/benchmark-candidates-2026-10-02.md):
  dataset = relbench.load_dataset("rel-hm"); task = dataset.load_task("user-item-purchase")   # MAP@12, 7-day window
  db = task.get_db(upto_test_timestamp=False)  -> customer / article / transactions tables
  for split in ("val", "test"): tbl = task.get_table(split)  (src ids + timestamps; test has no dst)
Our rows: embed customers (and articles) from data strictly before each split timestamp with the embedders in
``fer.embedders`` (TabPFN hidden state over a flattened aggregate table; KumoRelational / OpenRFM hidden states),
rank articles by (a) two-tower cosine, (b) user-kNN CF (neighbours' past purchases), (c) PastVisit + kNN fill,
and score with ``task.evaluate(pred_df, split)`` which returns link_prediction_map / precision / recall @K.

This module adapts a RelBench database to the ``ShopDB`` views used by the rest of the package, so the existing
embedders and ``future_purchase_retrieval`` run unchanged; MAP@K is then computed by RelBench's own evaluator.
"""
from __future__ import annotations

import pandas as pd

from .db import ShopDB


def shopdb_from_relbench_hm(db, cutoff) -> ShopDB:
    """Map rel-hm (customer / article / transactions) onto ShopDB. ``db`` is ``task.get_db(...)``."""
    t = db.table_dict
    cust = t["customer"].df.rename(columns={"customer_id": "customer_id"})
    art = t["article"].df.rename(columns={"article_id": "product_id", "product_group_name": "category", "price": "price"})
    tx = t["transactions"].df.rename(columns={"t_dat": "order_date", "article_id": "product_id"})
    tx = tx.reset_index(drop=True)
    tx["order_id"] = tx.index + 1
    orders = tx[["order_id", "customer_id", "order_date"]].copy()
    orders["order_date"] = pd.to_datetime(orders.order_date).dt.strftime("%Y-%m-%d")
    orders["status"] = "paid"
    items = tx[["order_id", "product_id"]].copy()
    items["item_id"] = items.index + 1
    items["quantity"] = 1
    if "price" in tx.columns:
        art = art.drop(columns=[c for c in ("price",) if c in art.columns]).merge(
            tx.groupby("product_id").price.mean().rename("price").reset_index(), on="product_id", how="left")
    customers = pd.DataFrame({"customer_id": cust.customer_id, "signup_date": orders.groupby("customer_id").order_date.min()
                              .reindex(cust.customer_id).fillna(orders.order_date.min()).to_numpy(),
                              "region": cust.get("postal_code", pd.Series(["?"] * len(cust))).astype(str).str[:3],
                              "age": cust.get("age", pd.Series([0] * len(cust))).fillna(0).astype(int),
                              "plan": cust.get("club_member_status", pd.Series(["?"] * len(cust))).astype(str)})
    products = art[["product_id", "category", "price"]].fillna({"price": 0.0})
    tickets = pd.DataFrame(columns=["ticket_id", "customer_id", "created_date", "severity", "status"])
    return ShopDB(customers, products, orders, items, tickets, pd.Timestamp(cutoff), None, name="rel-hm")
