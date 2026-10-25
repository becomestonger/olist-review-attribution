-- ============================================================================
-- Olist 巴西电商数据集 —— 数据分析探索 SQL
--
-- 连接信息（VS Code 底部状态栏应显示 Olist | analysis）：
--   主机 127.0.0.1 · 端口 5432 · 数据库 analysis · schema olist
--
-- 用法：选中某段 SQL 执行（Ctrl+Enter，或右键 → Run Query）
-- 说明：所有表名都带 olist. 前缀，所以每条查询都能独立执行，
--       不需要先设 search_path。
-- ============================================================================


-- ══════════════════════════════════════════════════════════════════════════
-- 第 1 部分：数据概览
-- ══════════════════════════════════════════════════════════════════════════

-- 1.1 各表行数，确认数据规模
SELECT 'customers' AS 表, count(*) AS 行数 FROM olist.customers
UNION ALL SELECT 'geolocation',                  count(*) FROM olist.geolocation
UNION ALL SELECT 'orders',                       count(*) FROM olist.orders
UNION ALL SELECT 'order_items',                  count(*) FROM olist.order_items
UNION ALL SELECT 'order_payments',               count(*) FROM olist.order_payments
UNION ALL SELECT 'order_reviews',                count(*) FROM olist.order_reviews
UNION ALL SELECT 'products',                     count(*) FROM olist.products
UNION ALL SELECT 'sellers',                      count(*) FROM olist.sellers
UNION ALL SELECT 'product_category_translation', count(*) FROM olist.product_category_translation
ORDER BY 1;

-- 1.2 整体经营指标（一行看完）
SELECT count(*)                                          AS 订单数,
       count(DISTINCT customer_unique_id)                AS 去重客户数,
       round(sum(order_total)::numeric, 2)               AS 商品总额,
       round(avg(order_total)::numeric, 2)               AS 客单价,
       round(avg(review_score), 2)                       AS 平均评分,
       round(avg(EXTRACT(EPOCH FROM (delivered_at - purchase_at)) / 86400), 1) AS 平均送达天数
FROM olist.v_orders_full;

-- 1.3 时间范围
SELECT min(order_purchase_timestamp) AS 最早下单,
       max(order_purchase_timestamp) AS 最晚下单
FROM olist.orders;


-- ══════════════════════════════════════════════════════════════════════════
-- 第 2 部分：销售趋势
-- ══════════════════════════════════════════════════════════════════════════

-- 2.1 按月订单量与销售额
SELECT to_char(purchase_month, 'YYYY-MM') AS 月份,
       count(*)                            AS 订单数,
       round(sum(order_total)::numeric, 2)  AS 销售额,
       round(avg(order_total)::numeric, 2)  AS 客单价
FROM olist.v_orders_full
GROUP BY 1
ORDER BY 1;

-- 2.2 按周几看下单习惯（0=周日，1=周一…）
SELECT EXTRACT(DOW FROM o.order_purchase_timestamp)::int AS 星期几,
       to_char(o.order_purchase_timestamp, 'Dy')          AS 名称,
       count(*)                                           AS 订单数,
       round(sum(i.total)::numeric, 2)                    AS 销售额
FROM olist.orders o
LEFT JOIN (SELECT order_id, sum(price + freight_value) AS total
             FROM olist.order_items GROUP BY order_id) i ON i.order_id = o.order_id
GROUP BY 1, 2
ORDER BY 1;

-- 2.3 分时段下单（0-23 点）
SELECT EXTRACT(HOUR FROM order_purchase_timestamp)::int AS 小时,
       count(*) AS 订单数
FROM olist.orders
GROUP BY 1
ORDER BY 1;


-- ══════════════════════════════════════════════════════════════════════════
-- 第 3 部分：地域分析
-- ══════════════════════════════════════════════════════════════════════════

-- 3.1 各州销售额排名
SELECT customer_state                     AS 州,
       count(*)                           AS 订单数,
       round(sum(order_total)::numeric, 2) AS 销售额,
       round(avg(order_total)::numeric, 2) AS 客单价,
       round(avg(review_score), 2)         AS 平均评分
FROM olist.v_orders_full
GROUP BY 1
ORDER BY 销售额 DESC;

-- 3.2 城市 Top 15
SELECT customer_city                      AS 城市,
       customer_state                     AS 州,
       count(*)                           AS 订单数,
       round(sum(order_total)::numeric, 2) AS 销售额
FROM olist.v_orders_full
GROUP BY 1, 2
ORDER BY 销售额 DESC
LIMIT 15;

-- 3.3 各州平均经纬度（可用于地图可视化）
SELECT geolocation_state                            AS 州,
       round(avg(geolocation_lat)::numeric, 4)      AS 平均纬度,
       round(avg(geolocation_lng)::numeric, 4)      AS 平均经度,
       count(DISTINCT geolocation_zip_code_prefix)  AS 邮编数
FROM olist.geolocation
GROUP BY 1
ORDER BY 1;


-- ══════════════════════════════════════════════════════════════════════════
-- 第 4 部分：品类与商品
-- ══════════════════════════════════════════════════════════════════════════

-- 4.1 品类销售排名（英文名已翻译）
SELECT category_en                      AS 品类,
       sum(units_sold)                  AS 销量,
       round(sum(revenue)::numeric, 2)  AS 销售额,
       round(avg(revenue)::numeric, 2)  AS 单品均收
FROM olist.v_product_sales
GROUP BY 1
ORDER BY 销售额 DESC NULLS LAST;

-- 4.2 只有 1 笔销量的"长尾"商品占比
WITH s AS (SELECT product_id, units_sold FROM olist.v_product_sales)
SELECT count(*) FILTER (WHERE units_sold = 1)        AS 只卖1件的商品,
       count(*)                                      AS 商品总数,
       round(100.0 * count(*) FILTER (WHERE units_sold = 1) / count(*), 1) AS 占比百分比
FROM s;

-- 4.3 运费占销售额比例（按品类，仅统计销量>100 的品类）
SELECT category_en                                                  AS 品类,
       round(avg(freight / nullif(revenue, 0) * 100)::numeric, 2)   AS 运费占比百分比
FROM olist.v_product_sales
GROUP BY 1
HAVING sum(units_sold) > 100
ORDER BY 2 DESC
LIMIT 15;


-- ══════════════════════════════════════════════════════════════════════════
-- 第 5 部分：支付与评价
-- ══════════════════════════════════════════════════════════════════════════

-- 5.1 支付方式分布
SELECT payment_type                                     AS 支付方式,
       count(*)                                         AS 笔数,
       round(sum(payment_value)::numeric, 2)            AS 金额,
       round(100.0 * count(*) / sum(count(*)) OVER (), 1) AS 占比百分比
FROM olist.order_payments
GROUP BY 1
ORDER BY 金额 DESC;

-- 5.2 信用卡分期数分布
SELECT payment_installments                     AS 分期数,
       count(*)                                 AS 笔数,
       round(avg(payment_value)::numeric, 2)    AS 平均金额
FROM olist.order_payments
WHERE payment_type = 'credit_card'
GROUP BY 1
ORDER BY 1;

-- 5.3 评分分布
SELECT review_score                                     AS 评分,
       count(*)                                         AS 评论数,
       round(100.0 * count(*) / sum(count(*)) OVER (), 1) AS 占比百分比
FROM olist.order_reviews
GROUP BY 1
ORDER BY 1;

-- 5.4 评分与配送时效的关系（关键业务洞察）
SELECT CASE
         WHEN delivered_at IS NULL THEN '未送达'
         WHEN EXTRACT(EPOCH FROM (delivered_at - purchase_at)) / 86400 <= 7  THEN '7天内'
         WHEN EXTRACT(EPOCH FROM (delivered_at - purchase_at)) / 86400 <= 14 THEN '8-14天'
         WHEN EXTRACT(EPOCH FROM (delivered_at - purchase_at)) / 86400 <= 21 THEN '15-21天'
         ELSE '21天以上'
       END                              AS 送达时长,
       count(*)                         AS 订单数,
       round(avg(review_score), 2)      AS 平均评分
FROM olist.v_orders_full
GROUP BY 1
ORDER BY 平均评分 DESC NULLS LAST;


-- ══════════════════════════════════════════════════════════════════════════
-- 第 6 部分：配送效率
-- ══════════════════════════════════════════════════════════════════════════

-- 6.1 迟到率（实际签收晚于预估送达）
SELECT count(*) FILTER (WHERE delivered_at > estimated_at)  AS 迟到订单,
       count(*) FILTER (WHERE delivered_at IS NOT NULL)     AS 已送达订单,
       round(100.0 * count(*) FILTER (WHERE delivered_at > estimated_at)
             / nullif(count(*) FILTER (WHERE delivered_at IS NOT NULL), 0), 2) AS 迟到率百分比
FROM olist.v_orders_full;

-- 6.2 各州配送表现（仅统计已送达>200 单的州）
SELECT customer_state AS 州,
       count(*) FILTER (WHERE delivered_at IS NOT NULL) AS 已送达,
       round(avg(EXTRACT(EPOCH FROM (delivered_at - purchase_at)) / 86400)::numeric, 1) AS 平均送达天数,
       round(100.0 * count(*) FILTER (WHERE delivered_at > estimated_at)
             / nullif(count(*) FILTER (WHERE delivered_at IS NOT NULL), 0), 2) AS 迟到率
FROM olist.v_orders_full
GROUP BY 1
HAVING count(*) FILTER (WHERE delivered_at IS NOT NULL) > 200
ORDER BY 迟到率 DESC;

-- 6.3 订单状态分布
SELECT order_status                                        AS 状态,
       count(*)                                            AS 订单数,
       round(100.0 * count(*) / sum(count(*)) OVER (), 2)  AS 占比
FROM olist.orders
GROUP BY 1
ORDER BY 订单数 DESC;


-- ══════════════════════════════════════════════════════════════════════════
-- 第 7 部分：卖家与客户
-- ══════════════════════════════════════════════════════════════════════════

-- 7.1 卖家销售额 Top 20
SELECT seller_id,
       seller_state                    AS 州,
       seller_city                     AS 城市,
       orders_count                    AS 订单数,
       round(revenue::numeric, 2)      AS 销售额
FROM olist.v_seller_performance
ORDER BY revenue DESC NULLS LAST
LIMIT 20;

-- 7.2 复购率（同一客户下过几单）
SELECT 下单次数,
       count(*)                                            AS 客户数,
       round(100.0 * count(*) / sum(count(*)) OVER (), 2)  AS 占比
FROM (
    SELECT customer_unique_id, count(*) AS 下单次数
    FROM olist.v_orders_full
    GROUP BY 1
) t
GROUP BY 1
ORDER BY 1;

-- 7.3 RFM 客户分层（R=最近购买, F=购买频次, M=消费金额）
WITH rfm AS (
    SELECT customer_unique_id,
           max(purchase_at)::date   AS last_purchase,
           count(*)                 AS frequency,
           sum(order_total)         AS monetary
    FROM olist.v_orders_full
    GROUP BY 1
),
scored AS (
    SELECT *,
           ntile(5) OVER (ORDER BY last_purchase DESC) AS r_score,
           ntile(5) OVER (ORDER BY frequency DESC)     AS f_score,
           ntile(5) OVER (ORDER BY monetary DESC)      AS m_score
    FROM rfm
)
SELECT CASE WHEN r_score >= 4 AND f_score >= 4 AND m_score >= 4 THEN '重要价值客户'
            WHEN r_score >= 4 AND m_score >= 4 THEN '重要发展客户'
            WHEN r_score <= 2 AND f_score >= 4 THEN '重要挽留客户'
            WHEN r_score <= 2 AND f_score <= 2 THEN '一般流失客户'
            ELSE '一般客户' END           AS 客户分层,
       count(*)                           AS 客户数,
       round(sum(monetary)::numeric, 2)   AS 总消费,
       round(avg(frequency)::numeric, 2)  AS 平均频次
FROM scored
GROUP BY 1
ORDER BY 总消费 DESC;
