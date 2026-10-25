-- ============================================================================
-- Olist 数据集导入脚本（PostgreSQL 18 / analysis 库）
-- 用 psql 的 \copy（客户端读取，普通角色即可，无需超级用户权限）
-- 用法：psql -h 127.0.0.1 -p 5432 -U analyst -d analysis -f 02_import_data.sql
-- ============================================================================

SET search_path TO olist, public;
\encoding UTF8
\timing on

\copy customers (customer_id, customer_unique_id, customer_zip_code_prefix, customer_city, customer_state) FROM 'C:/Users/28223/Desktop/巴西olist数据分析项目/olist_customers_dataset.csv' WITH (FORMAT csv, HEADER true, ENCODING 'UTF8', NULL '')

\copy geolocation (geolocation_zip_code_prefix, geolocation_lat, geolocation_lng, geolocation_city, geolocation_state) FROM 'C:/Users/28223/Desktop/巴西olist数据分析项目/olist_geolocation_dataset.csv' WITH (FORMAT csv, HEADER true, ENCODING 'UTF8', NULL '')

\copy orders (order_id, customer_id, order_status, order_purchase_timestamp, order_approved_at, order_delivered_carrier_date, order_delivered_customer_date, order_estimated_delivery_date) FROM 'C:/Users/28223/Desktop/巴西olist数据分析项目/olist_orders_dataset.csv' WITH (FORMAT csv, HEADER true, ENCODING 'UTF8', NULL '')

\copy order_items (order_id, order_item_id, product_id, seller_id, shipping_limit_date, price, freight_value) FROM 'C:/Users/28223/Desktop/巴西olist数据分析项目/olist_order_items_dataset.csv' WITH (FORMAT csv, HEADER true, ENCODING 'UTF8', NULL '')

\copy order_payments (order_id, payment_sequential, payment_type, payment_installments, payment_value) FROM 'C:/Users/28223/Desktop/巴西olist数据分析项目/olist_order_payments_dataset.csv' WITH (FORMAT csv, HEADER true, ENCODING 'UTF8', NULL '')

\copy order_reviews (review_id, order_id, review_score, review_comment_title, review_comment_message, review_creation_date, review_answer_timestamp) FROM 'C:/Users/28223/Desktop/巴西olist数据分析项目/olist_order_reviews_dataset.csv' WITH (FORMAT csv, HEADER true, ENCODING 'UTF8', NULL '')

\copy products (product_id, product_category_name, product_name_lenght, product_description_lenght, product_photos_qty, product_weight_g, product_length_cm, product_height_cm, product_width_cm) FROM 'C:/Users/28223/Desktop/巴西olist数据分析项目/olist_products_dataset.csv' WITH (FORMAT csv, HEADER true, ENCODING 'UTF8', NULL '')

\copy sellers (seller_id, seller_zip_code_prefix, seller_city, seller_state) FROM 'C:/Users/28223/Desktop/巴西olist数据分析项目/olist_sellers_dataset.csv' WITH (FORMAT csv, HEADER true, ENCODING 'UTF8', NULL '')

\copy product_category_translation (product_category_name, product_category_name_english) FROM 'C:/Users/28223/Desktop/巴西olist数据分析项目/product_category_name_translation.csv' WITH (FORMAT csv, HEADER true, ENCODING 'UTF8', NULL '')

\timing off

ANALYZE olist.customers;
ANALYZE olist.geolocation;
ANALYZE olist.orders;
ANALYZE olist.order_items;
ANALYZE olist.order_payments;
ANALYZE olist.order_reviews;
ANALYZE olist.products;
ANALYZE olist.sellers;
ANALYZE olist.product_category_translation;

\echo '=== 导入完成，各表行数 ==='
SELECT 'customers' AS 表, count(*) AS 行数 FROM customers
UNION ALL SELECT 'geolocation',                  count(*) FROM geolocation
UNION ALL SELECT 'orders',                       count(*) FROM orders
UNION ALL SELECT 'order_items',                  count(*) FROM order_items
UNION ALL SELECT 'order_payments',               count(*) FROM order_payments
UNION ALL SELECT 'order_reviews',                count(*) FROM order_reviews
UNION ALL SELECT 'products',                     count(*) FROM products
UNION ALL SELECT 'sellers',                      count(*) FROM sellers
UNION ALL SELECT 'product_category_translation', count(*) FROM product_category_translation
ORDER BY 1;
