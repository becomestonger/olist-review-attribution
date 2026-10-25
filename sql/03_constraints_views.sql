-- ============================================================================
-- 外键约束 + 分析视图
-- 用法：psql -h 127.0.0.1 -p 5432 -U analyst -d analysis -f 03_constraints_views.sql
-- ============================================================================
SET search_path TO olist, public;
\pset pager off

\echo ''
\echo '=== 1. 孤儿记录检查（为 0 才能加外键）==='
SELECT 'orders.customer_id' AS 检查项, count(*) AS 孤儿数
  FROM orders o LEFT JOIN customers c ON c.customer_id = o.customer_id
 WHERE c.customer_id IS NULL
UNION ALL
SELECT 'order_items.order_id', count(*)
  FROM order_items i LEFT JOIN orders o ON o.order_id = i.order_id
 WHERE o.order_id IS NULL
UNION ALL
SELECT 'order_items.product_id', count(*)
  FROM order_items i LEFT JOIN products p ON p.product_id = i.product_id
 WHERE p.product_id IS NULL
UNION ALL
SELECT 'order_items.seller_id', count(*)
  FROM order_items i LEFT JOIN sellers s ON s.seller_id = i.seller_id
 WHERE s.seller_id IS NULL
UNION ALL
SELECT 'order_payments.order_id', count(*)
  FROM order_payments p LEFT JOIN orders o ON o.order_id = p.order_id
 WHERE o.order_id IS NULL
UNION ALL
SELECT 'order_reviews.order_id', count(*)
  FROM order_reviews r LEFT JOIN orders o ON o.order_id = r.order_id
 WHERE o.order_id IS NULL
ORDER BY 1;

\echo ''
\echo '=== 2. 添加外键 ==='
BEGIN;
ALTER TABLE orders
    ADD CONSTRAINT fk_orders_customer
    FOREIGN KEY (customer_id) REFERENCES customers (customer_id);
ALTER TABLE order_items
    ADD CONSTRAINT fk_items_order   FOREIGN KEY (order_id)   REFERENCES orders (order_id),
    ADD CONSTRAINT fk_items_product FOREIGN KEY (product_id) REFERENCES products (product_id),
    ADD CONSTRAINT fk_items_seller  FOREIGN KEY (seller_id)  REFERENCES sellers (seller_id);
ALTER TABLE order_payments
    ADD CONSTRAINT fk_payments_order FOREIGN KEY (order_id) REFERENCES orders (order_id);
ALTER TABLE order_reviews
    ADD CONSTRAINT fk_reviews_order FOREIGN KEY (order_id) REFERENCES orders (order_id);
COMMIT;

\echo ''
\echo '=== 3. 创建分析视图 ==='

DROP VIEW IF EXISTS v_orders_full;
CREATE VIEW v_orders_full AS
SELECT
    o.order_id,
    o.order_status,
    o.order_purchase_timestamp                     AS purchase_at,
    date_trunc('month', o.order_purchase_timestamp) AS purchase_month,
    o.order_delivered_customer_date                AS delivered_at,
    o.order_estimated_delivery_date                AS estimated_at,
    c.customer_unique_id,
    c.customer_city,
    c.customer_state,
    SUM(i.price)                                   AS item_total,
    SUM(i.freight_value)                           AS freight_total,
    SUM(i.price + i.freight_value)                 AS order_total,
    count(i.order_item_id)                         AS item_count,
    p.payment_total,
    r.review_score
FROM orders o
JOIN customers c        ON c.customer_id = o.customer_id
LEFT JOIN order_items i ON i.order_id    = o.order_id
LEFT JOIN (SELECT order_id, SUM(payment_value) AS payment_total
             FROM order_payments GROUP BY order_id) p ON p.order_id = o.order_id
LEFT JOIN (SELECT order_id, max(review_score) AS review_score
             FROM order_reviews GROUP BY order_id) r ON r.order_id = o.order_id
GROUP BY o.order_id, o.order_status, o.order_purchase_timestamp,
         o.order_delivered_customer_date, o.order_estimated_delivery_date,
         c.customer_unique_id, c.customer_city, c.customer_state,
         p.payment_total, r.review_score;

DROP VIEW IF EXISTS v_product_sales;
CREATE VIEW v_product_sales AS
SELECT
    p.product_id,
    p.product_category_name         AS category_pt,
    t.product_category_name_english AS category_en,
    p.product_weight_g,
    count(*)                        AS units_sold,
    sum(i.price)                    AS revenue,
    sum(i.freight_value)            AS freight
FROM order_items i
JOIN products p ON p.product_id = i.product_id
LEFT JOIN product_category_translation t
       ON t.product_category_name = p.product_category_name
GROUP BY p.product_id, p.product_category_name,
         t.product_category_name_english, p.product_weight_g;

DROP VIEW IF EXISTS v_seller_performance;
CREATE VIEW v_seller_performance AS
SELECT
    s.seller_id,
    s.seller_city,
    s.seller_state,
    count(DISTINCT i.order_id) AS orders_count,
    sum(i.price)               AS revenue,
    sum(i.freight_value)       AS freight,
    avg(i.price)               AS avg_item_price
FROM sellers s
LEFT JOIN order_items i ON i.seller_id = s.seller_id
GROUP BY s.seller_id, s.seller_city, s.seller_state;

\echo ''
\echo '=== 4. 最终对象清单 ==='
SELECT table_name AS 对象名,
       CASE table_type WHEN 'VIEW' THEN '视图' ELSE '表' END AS 类型
FROM information_schema.tables
WHERE table_schema = 'olist'
ORDER BY table_type, table_name;

\echo ''
\echo '=== 5. 外键清单 ==='
SELECT conrelid::regclass AS 表, conname AS 约束
FROM pg_constraint
WHERE contype = 'f' AND connamespace = 'olist'::regnamespace
ORDER BY 1, 2;
