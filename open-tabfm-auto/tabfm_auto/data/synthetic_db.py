"""A synthetic relational database (SQLite) with planted structure, used by the DB agent
experiment and by the relational-embedding retrieval benchmark.

Tables: customers, products, orders, order_items, support_tickets.
Planted structure:
  * every customer has a latent ``segment`` (K communities). Segment drives WHICH products a
    customer buys (per-segment product-category preference) -- a purely relational signal that is
    NOT visible in the customers table itself (segment is stored only in a hidden side table).
  * ``churned`` (label for the predictive task) depends on recent order count, open high-severity
    tickets, plan and tenure, with noise.
"""
from __future__ import annotations

import sqlite3
from pathlib import Path

import numpy as np
import pandas as pd

from .registry import TabularTask

DEFAULT_DB = Path(__file__).resolve().parents[2] / "data_cache" / "synth_shop.sqlite"


def generate(path: Path = DEFAULT_DB, n_customers: int = 2000, n_products: int = 300, n_segments: int = 6,
             seed: int = 0, cutoff: str = "2025-07-01", overwrite: bool = False) -> Path:
    path = Path(path)
    if path.exists() and not overwrite:
        return path
    path.parent.mkdir(parents=True, exist_ok=True)
    rng = np.random.default_rng(seed)
    cutoff_ts = pd.Timestamp(cutoff)
    start = cutoff_ts - pd.Timedelta(days=540)

    regions = ["north", "south", "east", "west", "central"]
    plans = ["free", "basic", "pro"]
    categories = ["books", "garden", "electronics", "toys", "kitchen", "sports", "beauty", "office"]

    seg = rng.integers(0, n_segments, n_customers)
    customers = pd.DataFrame({
        "customer_id": np.arange(1, n_customers + 1),
        "signup_date": [str((start + pd.Timedelta(days=int(d))).date()) for d in rng.integers(0, 500, n_customers)],
        "region": rng.choice(regions, n_customers),
        "age": rng.integers(18, 75, n_customers),
        "plan": rng.choice(plans, n_customers, p=[0.5, 0.3, 0.2]),
    })
    products = pd.DataFrame({
        "product_id": np.arange(1, n_products + 1),
        "category": rng.choice(categories, n_products),
        "price": np.round(rng.lognormal(3.0, 0.6, n_products), 2),
    })
    # per-segment category preference (peaked Dirichlet) -> relational signal
    pref = rng.dirichlet(np.ones(len(categories)) * 0.25, size=n_segments)
    cat_idx = {c: np.where(products.category.values == c)[0] for c in categories}

    orders, items, tickets = [], [], []
    oid = iid = tid = 0
    signup = pd.to_datetime(customers.signup_date)
    for i in range(n_customers):
        n_ord = rng.poisson(6)
        for _ in range(n_ord):
            oid += 1
            day = rng.uniform(0, max(1, (cutoff_ts + pd.Timedelta(days=120) - signup[i]).days))
            od = signup[i] + pd.Timedelta(days=float(day))
            orders.append((oid, i + 1, str(od.date()), rng.choice(["paid", "paid", "paid", "refunded"])))
            for _ in range(rng.integers(1, 4)):
                iid += 1
                c = categories[rng.choice(len(categories), p=pref[seg[i]])]
                pool = cat_idx[c]
                if len(pool) == 0:
                    pool = np.arange(n_products)
                items.append((iid, oid, int(products.product_id.values[rng.choice(pool)]), int(rng.integers(1, 4))))
        for _ in range(rng.poisson(0.7)):
            tid += 1
            td = signup[i] + pd.Timedelta(days=float(rng.uniform(0, 540)))
            tickets.append((tid, i + 1, str(td.date()), rng.choice(["low", "medium", "high"], p=[0.5, 0.3, 0.2]),
                            rng.choice(["open", "closed"], p=[0.3, 0.7])))
    orders_df = pd.DataFrame(orders, columns=["order_id", "customer_id", "order_date", "status"])
    items_df = pd.DataFrame(items, columns=["item_id", "order_id", "product_id", "quantity"])
    tickets_df = pd.DataFrame(tickets, columns=["ticket_id", "customer_id", "created_date", "severity", "status"])

    # churn label: no order in (cutoff, cutoff+90d]; driven by behaviour before cutoff
    od = pd.to_datetime(orders_df.order_date)
    recent = orders_df[(od > cutoff_ts - pd.Timedelta(days=90)) & (od <= cutoff_ts)].groupby("customer_id").size()
    recent = customers.customer_id.map(recent).fillna(0).values
    high_open = tickets_df[(tickets_df.severity == "high") & (tickets_df.status == "open")
                           & (pd.to_datetime(tickets_df.created_date) <= cutoff_ts)].groupby("customer_id").size()
    high_open = customers.customer_id.map(high_open).fillna(0).values
    tenure = (cutoff_ts - signup).dt.days.values / 365
    plan_w = customers.plan.map({"free": 0.6, "basic": 0.0, "pro": -0.6}).values
    logit = -0.9 * recent + 1.2 * high_open + plan_w - 0.4 * tenure + 0.2 + 0.5 * rng.standard_normal(n_customers)
    churn_prob = 1 / (1 + np.exp(-logit))
    churned = (rng.random(n_customers) < churn_prob).astype(int)
    # make the future consistent with the label: churned customers get no orders after cutoff
    future = od > cutoff_ts
    drop = future & orders_df.customer_id.map(dict(zip(customers.customer_id, churned))).astype(bool).values
    orders_df = orders_df[~drop]
    items_df = items_df[items_df.order_id.isin(orders_df.order_id)]
    # customers who never ordered after cutoff are churned by definition
    has_future = orders_df[pd.to_datetime(orders_df.order_date) > cutoff_ts].customer_id.unique()
    churned = (~customers.customer_id.isin(has_future)).astype(int).values

    hidden = pd.DataFrame({"customer_id": customers.customer_id, "segment": seg, "churned": churned,
                           "churn_prob": np.round(churn_prob, 4)})
    con = sqlite3.connect(path)
    customers.to_sql("customers", con, index=False, if_exists="replace")
    products.to_sql("products", con, index=False, if_exists="replace")
    orders_df.to_sql("orders", con, index=False, if_exists="replace")
    items_df.to_sql("order_items", con, index=False, if_exists="replace")
    tickets_df.to_sql("support_tickets", con, index=False, if_exists="replace")
    hidden.to_sql("_hidden_ground_truth", con, index=False, if_exists="replace")
    meta = pd.DataFrame({"key": ["cutoff", "label_window_days", "n_segments", "seed"],
                         "value": [cutoff, "90", str(n_segments), str(seed)]})
    meta.to_sql("_meta", con, index=False, if_exists="replace")
    con.close()
    return path


DB_TASK_DESCRIPTION = (
    "Online shop database. Predict whether each customer will CHURN, defined as placing no paid order in the "
    "90 days after the cutoff date {cutoff}. Only information dated on or before the cutoff may be used as "
    "features. The label must be derived from orders after the cutoff."
)


def load_db_task(name: str = "churn", path: Path = DEFAULT_DB) -> TabularTask:
    """Reference (hand-written) feature table for the churn task, used as the 'human' baseline."""
    path = generate(path)
    con = sqlite3.connect(path)
    cutoff = pd.read_sql("select value from _meta where key='cutoff'", con).value[0]
    q = f"""
    with ords as (select customer_id, count(*) n_orders,
                         sum(case when order_date > date('{cutoff}', '-90 days') then 1 else 0 end) n_recent,
                         max(order_date) last_order
                  from orders where order_date <= '{cutoff}' and status='paid' group by customer_id),
         tix as (select customer_id, count(*) n_tickets,
                        sum(case when severity='high' and status='open' then 1 else 0 end) n_high_open
                 from support_tickets where created_date <= '{cutoff}' group by customer_id),
         fut as (select distinct customer_id from orders where order_date > '{cutoff}')
    select c.customer_id, c.region, c.age, c.plan, julianday('{cutoff}') - julianday(c.signup_date) as tenure_days,
           coalesce(o.n_orders,0) n_orders, coalesce(o.n_recent,0) n_recent,
           julianday('{cutoff}') - julianday(o.last_order) days_since_last_order,
           coalesce(t.n_tickets,0) n_tickets, coalesce(t.n_high_open,0) n_high_open,
           case when f.customer_id is null then 1 else 0 end as churned
    from customers c left join ords o using(customer_id) left join tix t using(customer_id)
         left join fut f using(customer_id)
    """
    df = pd.read_sql(q, con)
    con.close()
    y = df.pop("churned")
    X = df.drop(columns=["customer_id"])
    meta = {"description": DB_TASK_DESCRIPTION.format(cutoff=cutoff) + " (reference hand-built feature table)",
            "source": f"sqlite:{path}", "columns": {c: "" for c in X.columns}}
    return TabularTask("db_churn_reference", X, pd.Series(y.values, name="churned"), "binary", meta)


if __name__ == "__main__":  # pragma: no cover
    p = generate(overwrite=True)
    print("wrote", p)
    t = load_db_task()
    print(t.summary(), "churn rate", round(float(t.y.mean()), 3))
