WITH o AS (
  SELECT customer_id, order_id, order_date, status,
         julianday('2025-07-01') - julianday(order_date) AS age_d
  FROM orders WHERE order_date <= '2025-07-01'
),
oa AS (
  SELECT customer_id,
    COUNT(*) AS n_orders,
    SUM(status='paid') AS n_paid,
    SUM(status='refunded') AS n_refunded,
    1.0*SUM(status='refunded')/COUNT(*) AS refund_rate,
    SUM(status='paid' AND age_d<=30) AS paid_30d,
    SUM(status='paid' AND age_d<=90) AS paid_90d,
    SUM(status='paid' AND age_d<=180) AS paid_180d,
    SUM(status='paid' AND age_d>90 AND age_d<=180) AS paid_90_180d,
    SUM(status='refunded' AND age_d<=90) AS refunded_90d,
    MIN(CASE WHEN status='paid' THEN age_d END) AS days_since_last_paid,
    MAX(CASE WHEN status='paid' THEN age_d END) AS days_since_first_paid,
    MIN(age_d) AS days_since_last_order
  FROM o GROUP BY customer_id
),
oi AS (
  SELECT o.customer_id,
    SUM(i.quantity) AS units_total,
    SUM(i.quantity*p.price) AS spend_total,
    SUM(CASE WHEN o.age_d<=90 THEN i.quantity*p.price ELSE 0 END) AS spend_90d,
    SUM(CASE WHEN o.age_d<=30 THEN i.quantity*p.price ELSE 0 END) AS spend_30d,
    AVG(p.price) AS avg_item_price,
    COUNT(DISTINCT i.product_id) AS n_distinct_products,
    COUNT(DISTINCT p.category) AS n_categories,
    SUM(CASE WHEN p.category='electronics' THEN i.quantity ELSE 0 END)*1.0/SUM(i.quantity) AS share_electronics,
    SUM(CASE WHEN p.category='sports' THEN i.quantity ELSE 0 END)*1.0/SUM(i.quantity) AS share_sports,
    SUM(CASE WHEN p.category='office' THEN i.quantity ELSE 0 END)*1.0/SUM(i.quantity) AS share_office,
    SUM(CASE WHEN p.category='beauty' THEN i.quantity ELSE 0 END)*1.0/SUM(i.quantity) AS share_beauty,
    SUM(CASE WHEN p.category='kitchen' THEN i.quantity ELSE 0 END)*1.0/SUM(i.quantity) AS share_kitchen,
    SUM(CASE WHEN p.category='books' THEN i.quantity ELSE 0 END)*1.0/SUM(i.quantity) AS share_books,
    SUM(CASE WHEN p.category='garden' THEN i.quantity ELSE 0 END)*1.0/SUM(i.quantity) AS share_garden,
    SUM(CASE WHEN p.category='toys' THEN i.quantity ELSE 0 END)*1.0/SUM(i.quantity) AS share_toys
  FROM o JOIN order_items i ON i.order_id=o.order_id
         JOIN products p ON p.product_id=i.product_id
  WHERE o.status='paid'
  GROUP BY o.customer_id
),
t AS (
  SELECT customer_id,
    COUNT(*) AS n_tickets,
    SUM(julianday('2025-07-01')-julianday(created_date)<=90) AS tickets_90d,
    SUM(severity='high') AS tickets_high,
    SUM(severity='medium') AS tickets_medium,
    SUM(severity='low') AS tickets_low,
    MIN(julianday('2025-07-01')-julianday(created_date)) AS days_since_last_ticket
  FROM support_tickets WHERE created_date <= '2025-07-01'
  GROUP BY customer_id
),
lab AS (
  SELECT customer_id, 1 AS has_paid
  FROM orders
  WHERE status='paid' AND order_date > '2025-07-01' AND order_date <= date('2025-07-01','+90 day')
  GROUP BY customer_id
)
SELECT c.customer_id,
  c.region, c.plan, c.age,
  julianday('2025-07-01')-julianday(c.signup_date) AS tenure_days,
  COALESCE(oa.n_orders,0) AS n_orders,
  COALESCE(oa.n_paid,0) AS n_paid,
  COALESCE(oa.n_refunded,0) AS n_refunded,
  oa.refund_rate,
  COALESCE(oa.paid_30d,0) AS paid_30d,
  COALESCE(oa.paid_90d,0) AS paid_90d,
  COALESCE(oa.paid_180d,0) AS paid_180d,
  COALESCE(oa.paid_90_180d,0) AS paid_90_180d,
  COALESCE(oa.refunded_90d,0) AS refunded_90d,
  oa.days_since_last_paid, oa.days_since_first_paid, oa.days_since_last_order,
  1.0*COALESCE(oa.n_paid,0)/MAX(1, julianday('2025-07-01')-julianday(c.signup_date))*30 AS paid_per_month,
  COALESCE(oi.units_total,0) AS units_total,
  COALESCE(oi.spend_total,0) AS spend_total,
  COALESCE(oi.spend_90d,0) AS spend_90d,
  COALESCE(oi.spend_30d,0) AS spend_30d,
  oi.avg_item_price,
  COALESCE(oi.n_distinct_products,0) AS n_distinct_products,
  COALESCE(oi.n_categories,0) AS n_categories,
  oi.share_electronics, oi.share_sports, oi.share_office, oi.share_beauty,
  oi.share_kitchen, oi.share_books, oi.share_garden, oi.share_toys,
  COALESCE(t.n_tickets,0) AS n_tickets,
  COALESCE(t.tickets_90d,0) AS tickets_90d,
  COALESCE(t.tickets_high,0) AS tickets_high,
  COALESCE(t.tickets_medium,0) AS tickets_medium,
  COALESCE(t.tickets_low,0) AS tickets_low,
  t.days_since_last_ticket,
  CASE WHEN lab.has_paid IS NULL THEN 1 ELSE 0 END AS churned
FROM customers c
LEFT JOIN oa ON oa.customer_id=c.customer_id
LEFT JOIN oi ON oi.customer_id=c.customer_id
LEFT JOIN t ON t.customer_id=c.customer_id
LEFT JOIN lab ON lab.customer_id=c.customer_id
WHERE c.signup_date <= '2025-07-01'
