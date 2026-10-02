"""Load a relational SQLite database into pandas tables with a temporal cutoff, exposing the views
every embedder needs: the entity (customer) table, pre-cutoff interaction rows, and post-cutoff
interactions used only as retrieval ground truth."""
from __future__ import annotations

import sqlite3
from dataclasses import dataclass, field
from pathlib import Path

import numpy as np
import pandas as pd


@dataclass
class ShopDB:
    customers: pd.DataFrame
    products: pd.DataFrame
    orders: pd.DataFrame
    items: pd.DataFrame
    tickets: pd.DataFrame
    cutoff: pd.Timestamp
    hidden: pd.DataFrame | None = None  # ground truth (segment, churned); never used for embeddings
    name: str = "synth_shop"
    extras: dict = field(default_factory=dict)

    # -- derived views ------------------------------------------------------------------------
    def items_joined(self) -> pd.DataFrame:
        """One row per order line with order date/status and product category/price."""
        df = self.items.merge(self.orders, on="order_id").merge(self.products, on="product_id")
        df["order_date"] = pd.to_datetime(df["order_date"])
        return df

    def past_items(self) -> pd.DataFrame:
        j = self.items_joined()
        return j[j.order_date <= self.cutoff]

    def future_items(self) -> pd.DataFrame:
        j = self.items_joined()
        return j[j.order_date > self.cutoff]

    def customer_ids(self) -> np.ndarray:
        return self.customers.customer_id.to_numpy()


def load_shop(path: str | Path, cutoff: str | None = None) -> ShopDB:
    con = sqlite3.connect(str(path))
    t = {n: pd.read_sql(f"select * from {n}", con) for n in
         ["customers", "products", "orders", "order_items", "support_tickets"]}
    hidden = None
    try:
        hidden = pd.read_sql("select * from _hidden_ground_truth", con)
    except Exception:
        pass
    if cutoff is None:
        try:
            cutoff = pd.read_sql("select value from _meta where key='cutoff'", con).value[0]
        except Exception:
            cutoff = str(pd.to_datetime(t["orders"].order_date).quantile(0.7).date())
    con.close()
    return ShopDB(t["customers"], t["products"], t["orders"], t["order_items"], t["support_tickets"],
                  pd.Timestamp(cutoff), hidden, name=Path(path).stem)


def load_northwind(path: str | Path, cutoff: str | None = None) -> ShopDB:
    """Map the Northwind sample DB onto the same ShopDB views (Customers / Products / Orders / Order Details)."""
    con = sqlite3.connect(str(path))
    cust = pd.read_sql("select CustomerID as customer_id, Country as region, City as city, Region as state from Customers", con)
    prod = pd.read_sql("select ProductID as product_id, CategoryID as category, UnitPrice as price from Products", con)
    orders = pd.read_sql("select OrderID as order_id, CustomerID as customer_id, OrderDate as order_date, "
                         "ShipVia as status from Orders", con)
    items = pd.read_sql('select rowid as item_id, OrderID as order_id, ProductID as product_id, Quantity as quantity '
                        'from "Order Details"', con)
    con.close()
    orders["order_date"] = pd.to_datetime(orders.order_date, format="mixed").dt.strftime("%Y-%m-%d")
    orders["status"] = "paid"
    prod["category"] = prod.category.astype(str)
    cust["signup_date"] = orders.groupby("customer_id").order_date.min().reindex(cust.customer_id).fillna(
        orders.order_date.min()).to_numpy()
    cust["age"] = 0
    cust["plan"] = cust.pop("city").astype(str)
    cust = cust.drop(columns=["state"])
    tickets = pd.DataFrame(columns=["ticket_id", "customer_id", "created_date", "severity", "status"])
    if cutoff is None:
        cutoff = str(pd.to_datetime(orders.order_date).quantile(0.75).date())
    return ShopDB(cust, prod, orders, items, tickets, pd.Timestamp(cutoff), None, name="northwind")
