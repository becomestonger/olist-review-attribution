-- ============================================================================
-- Olist 巴西电商数据集 —— 建表脚本
-- 目标：PostgreSQL 18 / 数据库 analysis / 用户 analyst
-- 用法：psql -h 127.0.0.1 -p 5432 -U analyst -d analysis -f 01_create_schema.sql
-- ============================================================================

DROP SCHEMA IF EXISTS olist CASCADE;
CREATE SCHEMA olist;
SET search_path TO olist, public;

-- 1. 客户
CREATE TABLE customers (
    customer_id              CHAR(32) PRIMARY KEY,
    customer_unique_id       CHAR(32) NOT NULL,
    customer_zip_code_prefix INTEGER,
    customer_city            TEXT,
    customer_state           CHAR(2)
);
CREATE INDEX idx_customers_unique ON customers (customer_unique_id);
CREATE INDEX idx_customers_state  ON customers (customer_state);

-- 2. 地理位置（同一邮编多条，无自然主键）
CREATE TABLE geolocation (
    geolocation_zip_code_prefix INTEGER,
    geolocation_lat             DOUBLE PRECISION,
    geolocation_lng             DOUBLE PRECISION,
    geolocation_city            TEXT,
    geolocation_state           CHAR(2)
);
CREATE INDEX idx_geo_zip   ON geolocation (geolocation_zip_code_prefix);
CREATE INDEX idx_geo_state ON geolocation (geolocation_state);

-- 3. 订单
CREATE TABLE orders (
    order_id                      CHAR(32) PRIMARY KEY,
    customer_id                   CHAR(32) NOT NULL,
    order_status                  TEXT,
    order_purchase_timestamp      TIMESTAMP,
    order_approved_at             TIMESTAMP,
    order_delivered_carrier_date  TIMESTAMP,
    order_delivered_customer_date TIMESTAMP,
    order_estimated_delivery_date TIMESTAMP
);
CREATE INDEX idx_orders_customer ON orders (customer_id);
CREATE INDEX idx_orders_status   ON orders (order_status);
CREATE INDEX idx_orders_purchase ON orders (order_purchase_timestamp);

-- 4. 订单明细
CREATE TABLE order_items (
    order_id            CHAR(32) NOT NULL,
    order_item_id       INTEGER  NOT NULL,
    product_id          CHAR(32) NOT NULL,
    seller_id           CHAR(32) NOT NULL,
    shipping_limit_date TIMESTAMP,
    price               NUMERIC(12,2),
    freight_value       NUMERIC(12,2),
    PRIMARY KEY (order_id, order_item_id)
);
CREATE INDEX idx_items_product ON order_items (product_id);
CREATE INDEX idx_items_seller  ON order_items (seller_id);

-- 5. 支付
CREATE TABLE order_payments (
    order_id             CHAR(32) NOT NULL,
    payment_sequential   INTEGER  NOT NULL,
    payment_type         TEXT,
    payment_installments INTEGER,
    payment_value        NUMERIC(12,2),
    PRIMARY KEY (order_id, payment_sequential)
);
CREATE INDEX idx_pay_type ON order_payments (payment_type);

-- 6. 评论
CREATE TABLE order_reviews (
    review_id               CHAR(32) NOT NULL,
    order_id                CHAR(32) NOT NULL,
    review_score            SMALLINT,
    review_comment_title    TEXT,
    review_comment_message  TEXT,
    review_creation_date    TIMESTAMP,
    review_answer_timestamp TIMESTAMP,
    PRIMARY KEY (review_id, order_id)
);
CREATE INDEX idx_rev_order ON order_reviews (order_id);
CREATE INDEX idx_rev_score ON order_reviews (review_score);

-- 7. 商品
CREATE TABLE products (
    product_id                 CHAR(32) PRIMARY KEY,
    product_category_name      TEXT,
    product_name_lenght        INTEGER,
    product_description_lenght INTEGER,
    product_photos_qty         INTEGER,
    product_weight_g           NUMERIC(12,2),
    product_length_cm          NUMERIC(12,2),
    product_height_cm          NUMERIC(12,2),
    product_width_cm           NUMERIC(12,2)
);
CREATE INDEX idx_prod_cat ON products (product_category_name);

-- 8. 卖家
CREATE TABLE sellers (
    seller_id              CHAR(32) PRIMARY KEY,
    seller_zip_code_prefix INTEGER,
    seller_city            TEXT,
    seller_state           CHAR(2)
);
CREATE INDEX idx_sellers_state ON sellers (seller_state);

-- 9. 品类翻译
CREATE TABLE product_category_translation (
    product_category_name         TEXT PRIMARY KEY,
    product_category_name_english TEXT
);

\echo '=== 9 张表创建完成 ==='
\dt olist.*
