-- ============================================================================
-- Olist 巴西电商 —— 差评归因分析 完整 SQL 汇编
--
-- 研究问题（已收敛）：
--   配送速度是否影响客户满意度？哪些因素导致差评？
--   各因素的影响权重如何？业务该优先抓哪个因素？
--
-- 差评定义：review_score <= 2（1-2 星）
--
-- 数据：PostgreSQL 18 · analysis 库 · olist 模式
-- 执行方式：逐段选中执行（Ctrl+Enter），或整段 Run
--
-- 阅读顺序即分析顺序：
--   第 0 部分 → 建立基线（差评率是多少）
--   第 1 部分 → 定位瓶颈（哪一环节最耗时）
--   第 2 部分 → 地理因素
--   第 3 部分 → 卖家因素（备货环节 + 违约）
--   第 4 部分 → 商品因素（品类/重量/件数）
--   第 5 部分 → 时间因素（趋势/旺季/异常）
--   第 6 部分 → 汇总对比（各因素对差评的影响权重）★ 核心结论
--   第 7 部分 → 数据质量与局限（写报告用）
-- ============================================================================


-- ████████████████████████████████████████████████████████████████████████████
-- 第 0 部分  基础信息与基线
-- ████████████████████████████████████████████████████████████████████████████

-- 0.1 数据规模（确认 9 张表行数）
SELECT 'customers' AS 表, count(*) AS 行数 FROM olist.customers
UNION ALL SELECT 'geolocation',         count(*) FROM olist.geolocation
UNION ALL SELECT 'orders',              count(*) FROM olist.orders
UNION ALL SELECT 'order_items',         count(*) FROM olist.order_items
UNION ALL SELECT 'order_payments',      count(*) FROM olist.order_payments
UNION ALL SELECT 'order_reviews',       count(*) FROM olist.order_reviews
UNION ALL SELECT 'products',            count(*) FROM olist.products
UNION ALL SELECT 'sellers',             count(*) FROM olist.sellers
UNION ALL SELECT 'category_translation',count(*) FROM olist.product_category_translation
ORDER BY 1;

-- 0.2 评分分布（看清差评的规模）
SELECT review_score AS 星级,
       count(*)     AS 评论数,
       round(100.0 * count(*) / sum(count(*)) OVER (), 2) AS 占比
FROM olist.order_reviews
GROUP BY 1 ORDER BY 1;

-- 0.3 差评基线 ★ 后面所有分组都要和这个数比
SELECT count(*) AS 已评订单数,
       count(*) FILTER (WHERE review_score <= 2) AS 差评订单数,
       round(100.0 * count(*) FILTER (WHERE review_score <= 2) / count(*), 2) AS 差评率基线
FROM (SELECT order_id, max(review_score) AS review_score
      FROM olist.order_reviews GROUP BY order_id) t;


-- ████████████████████████████████████████████████████████████████████████████
-- 第 1 部分  流程拆解：瓶颈在哪一环节
-- ████████████████████████████████████████████████████████████████████████████

-- 1.1 四段耗时拆解 ★ 这一步决定后面查哪里
--     结果解读：占比最大的那一段就是瓶颈
SELECT round(avg(EXTRACT(EPOCH FROM (order_approved_at - order_purchase_timestamp))/86400)::numeric, 2)          AS 下单到审核_天,
       round(avg(EXTRACT(EPOCH FROM (order_delivered_carrier_date - order_approved_at))/86400)::numeric, 2)      AS 审核到发货_天,
       round(avg(EXTRACT(EPOCH FROM (order_delivered_customer_date - order_delivered_carrier_date))/86400)::numeric, 2) AS 发货到签收_天,
       round(avg(EXTRACT(EPOCH FROM (order_delivered_customer_date - order_purchase_timestamp))/86400)::numeric, 2)     AS 全程_天,
       count(*) AS 样本数
FROM olist.orders
WHERE order_delivered_customer_date IS NOT NULL
  AND order_approved_at IS NOT NULL
  AND order_delivered_carrier_date IS NOT NULL;

-- 1.2 运输时长分布（验证平均值是否被极端值拉高）
SELECT round(avg(d)::numeric, 2)                                             AS 平均,
       round(percentile_cont(0.5)  WITHIN GROUP (ORDER BY d)::numeric, 2)    AS 中位数,
       round(percentile_cont(0.9)  WITHIN GROUP (ORDER BY d)::numeric, 2)    AS p90,
       round(percentile_cont(0.99) WITHIN GROUP (ORDER BY d)::numeric, 2)    AS p99,
       round(max(d)::numeric, 1)                                             AS 最大,
       count(*) FILTER (WHERE d < 0)                                         AS 负值异常条数
FROM (SELECT EXTRACT(EPOCH FROM (order_delivered_customer_date - order_delivered_carrier_date))/86400 AS d
      FROM olist.orders
      WHERE order_delivered_customer_date IS NOT NULL
        AND order_delivered_carrier_date IS NOT NULL) t;

-- 1.3 各环节耗时 → 差评率 ★ 把过程指标和结果指标连起来
SELECT '备货 1天内'   AS 环节, count(*) AS 订单数,
       round(100.0*count(*) FILTER (WHERE r.review_score <= 2)/count(r.review_score), 2) AS 差评率
FROM olist.orders o
LEFT JOIN (SELECT order_id, max(review_score) review_score FROM olist.order_reviews GROUP BY 1) r ON r.order_id=o.order_id
WHERE o.order_delivered_carrier_date IS NOT NULL AND o.order_approved_at IS NOT NULL
  AND EXTRACT(EPOCH FROM (o.order_delivered_carrier_date - o.order_approved_at))/86400 <= 1
UNION ALL
SELECT '备货 1-3天', count(*),
       round(100.0*count(*) FILTER (WHERE r.review_score <= 2)/count(r.review_score), 2)
FROM olist.orders o
LEFT JOIN (SELECT order_id, max(review_score) review_score FROM olist.order_reviews GROUP BY 1) r ON r.order_id=o.order_id
WHERE o.order_delivered_carrier_date IS NOT NULL AND o.order_approved_at IS NOT NULL
  AND EXTRACT(EPOCH FROM (o.order_delivered_carrier_date - o.order_approved_at))/86400 > 1
  AND EXTRACT(EPOCH FROM (o.order_delivered_carrier_date - o.order_approved_at))/86400 <= 3
UNION ALL
SELECT '备货 3-7天', count(*),
       round(100.0*count(*) FILTER (WHERE r.review_score <= 2)/count(r.review_score), 2)
FROM olist.orders o
LEFT JOIN (SELECT order_id, max(review_score) review_score FROM olist.order_reviews GROUP BY 1) r ON r.order_id=o.order_id
WHERE o.order_delivered_carrier_date IS NOT NULL AND o.order_approved_at IS NOT NULL
  AND EXTRACT(EPOCH FROM (o.order_delivered_carrier_date - o.order_approved_at))/86400 > 3
  AND EXTRACT(EPOCH FROM (o.order_delivered_carrier_date - o.order_approved_at))/86400 <= 7
UNION ALL
SELECT '备货 7-14天', count(*),
       round(100.0*count(*) FILTER (WHERE r.review_score <= 2)/count(r.review_score), 2)
FROM olist.orders o
LEFT JOIN (SELECT order_id, max(review_score) review_score FROM olist.order_reviews GROUP BY 1) r ON r.order_id=o.order_id
WHERE o.order_delivered_carrier_date IS NOT NULL AND o.order_approved_at IS NOT NULL
  AND EXTRACT(EPOCH FROM (o.order_delivered_carrier_date - o.order_approved_at))/86400 > 7
  AND EXTRACT(EPOCH FROM (o.order_delivered_carrier_date - o.order_approved_at))/86400 <= 14
UNION ALL
SELECT '备货 14天以上', count(*),
       round(100.0*count(*) FILTER (WHERE r.review_score <= 2)/count(r.review_score), 2)
FROM olist.orders o
LEFT JOIN (SELECT order_id, max(review_score) review_score FROM olist.order_reviews GROUP BY 1) r ON r.order_id=o.order_id
WHERE o.order_delivered_carrier_date IS NOT NULL AND o.order_approved_at IS NOT NULL
  AND EXTRACT(EPOCH FROM (o.order_delivered_carrier_date - o.order_approved_at))/86400 > 14
ORDER BY 1;


-- ████████████████████████████████████████████████████████████████████████████
-- 第 2 部分  地理因素
-- ████████████████████████████████████████████████████████████████████████████

-- 2.1 各州运输时长与迟到率（样本 >= 200）
SELECT c.customer_state AS 客户州,
       count(*) AS 订单数,
       round(avg(EXTRACT(EPOCH FROM (o.order_delivered_customer_date - o.order_delivered_carrier_date))/86400)::numeric, 1) AS 平均运输天数,
       round(100.0*count(*) FILTER (WHERE o.order_delivered_customer_date > o.order_estimated_delivery_date)/count(*), 1) AS 迟到率
FROM olist.orders o
JOIN olist.customers c ON c.customer_id = o.customer_id
WHERE o.order_delivered_customer_date IS NOT NULL
  AND o.order_delivered_carrier_date IS NOT NULL
GROUP BY 1 HAVING count(*) >= 200
ORDER BY 平均运输天数 DESC;

-- 2.2 同州 vs 跨州（隔离"距离"这个变量）★
WITH first_seller AS (
    SELECT order_id, min(seller_id) AS seller_id
    FROM olist.order_items GROUP BY 1
)
SELECT CASE WHEN c.customer_state = s.seller_state THEN '同州' ELSE '跨州' END AS 类型,
       count(*) AS 订单数,
       round(avg(EXTRACT(EPOCH FROM (o.order_delivered_customer_date - o.order_delivered_carrier_date))/86400)::numeric, 2) AS 平均运输天数,
       round(100.0*count(*) FILTER (WHERE o.order_delivered_customer_date > o.order_estimated_delivery_date)/count(*), 2) AS 迟到率,
       round(100.0*count(*) FILTER (WHERE r.review_score <= 2)/count(r.review_score), 2) AS 差评率
FROM olist.orders o
JOIN olist.customers c   ON c.customer_id = o.customer_id
JOIN first_seller fs     ON fs.order_id   = o.order_id
JOIN olist.sellers s     ON s.seller_id   = fs.seller_id
LEFT JOIN (SELECT order_id, max(review_score) review_score FROM olist.order_reviews GROUP BY 1) r ON r.order_id = o.order_id
WHERE o.order_delivered_customer_date IS NOT NULL
  AND o.order_delivered_carrier_date IS NOT NULL
GROUP BY 1 ORDER BY 1;

-- 2.3 平台预估窗口是否按距离动态调整 ★
--     看"平均富余"这一列：偏远州富余越大，说明平台越大方
SELECT c.customer_state AS 州,
       count(*) AS 订单数,
       round(avg(EXTRACT(EPOCH FROM (o.order_estimated_delivery_date - o.order_purchase_timestamp))/86400)::numeric, 1) AS 平台预估天数,
       round(avg(EXTRACT(EPOCH FROM (o.order_delivered_customer_date - o.order_purchase_timestamp))/86400)::numeric, 1) AS 实际全程天数,
       round(avg(EXTRACT(EPOCH FROM (o.order_estimated_delivery_date - o.order_delivered_customer_date))/86400)::numeric, 1) AS 平均富余天数
FROM olist.orders o
JOIN olist.customers c ON c.customer_id = o.customer_id
WHERE o.order_delivered_customer_date IS NOT NULL
GROUP BY 1 HAVING count(*) >= 200
ORDER BY 平台预估天数 DESC;

-- 2.4 圣保罗 vs 其他州（对比标杆）
SELECT CASE WHEN c.customer_state = 'SP' THEN 'SP 圣保罗' ELSE '其他州' END AS 分组,
       count(*) AS 订单数,
       round(avg(EXTRACT(EPOCH FROM (o.order_delivered_customer_date - o.order_purchase_timestamp))/86400)::numeric, 1) AS 全程天数,
       round(100.0*count(*) FILTER (WHERE o.order_delivered_customer_date > o.order_estimated_delivery_date)/count(*), 1) AS 迟到率,
       round(avg(r.review_score), 2) AS 平均评分,
       round(100.0*count(*) FILTER (WHERE r.review_score <= 2)/count(r.review_score), 2) AS 差评率
FROM olist.orders o
JOIN olist.customers c ON c.customer_id = o.customer_id
LEFT JOIN (SELECT order_id, max(review_score) review_score FROM olist.order_reviews GROUP BY 1) r ON r.order_id = o.order_id
WHERE o.order_delivered_customer_date IS NOT NULL
GROUP BY 1 ORDER BY 1;


-- ████████████████████████████████████████████████████████████████████████████
-- 第 3 部分  卖家因素（备货环节 + 违约）
-- ████████████████████████████████████████████████████████████████████████████

-- 3.1 备货最慢的卖家（样本 >= 50，避免小样本噪音）
SELECT i.seller_id,
       count(DISTINCT o.order_id) AS 订单数,
       round(avg(EXTRACT(EPOCH FROM (o.order_delivered_carrier_date - o.order_approved_at))/86400)::numeric, 1) AS 平均备货天数
FROM olist.orders o
JOIN olist.order_items i ON i.order_id = o.order_id
WHERE o.order_delivered_carrier_date IS NOT NULL
  AND o.order_approved_at IS NOT NULL
GROUP BY 1
HAVING count(DISTINCT o.order_id) >= 50
ORDER BY 平均备货天数 DESC
LIMIT 20;

-- 3.2 卖家是否超出发货期限（shipping_limit_date）★ 硬违约证据
SELECT CASE WHEN o.order_delivered_carrier_date > i.shipping_limit_date
            THEN '超期发货' ELSE '按期发货' END AS 状态,
       count(*) AS 明细行数,
       round(100.0*count(*) / sum(count(*)) OVER (), 2) AS 占比
FROM olist.order_items i
JOIN olist.orders o ON o.order_id = i.order_id
WHERE o.order_delivered_carrier_date IS NOT NULL
  AND i.shipping_limit_date IS NOT NULL
GROUP BY 1 ORDER BY 2 DESC;

-- 3.3 超期发货的代价 ★★ 全项目最强发现
WITH base AS (
    SELECT o.order_id,
           CASE WHEN max(o.order_delivered_carrier_date) > max(i.shipping_limit_date)
                THEN '超期发货' ELSE '按期发货' END AS 状态,
           max(o.order_estimated_delivery_date) AS est,
           max(o.order_delivered_customer_date) AS deliv,
           EXTRACT(EPOCH FROM (max(o.order_delivered_customer_date) - max(o.order_purchase_timestamp)))/86400 AS 全程
    FROM olist.orders o
    JOIN olist.order_items i ON i.order_id = o.order_id
    WHERE o.order_delivered_customer_date IS NOT NULL
      AND o.order_delivered_carrier_date IS NOT NULL
      AND i.shipping_limit_date IS NOT NULL
    GROUP BY o.order_id
)
SELECT b.状态,
       count(*) AS 订单数,
       round(avg(b.全程)::numeric, 1) AS 全程天数,
       round(100.0*count(*) FILTER (WHERE b.deliv > b.est)/count(*), 1) AS 迟到率,
       round(avg(r.review_score), 2) AS 平均评分,
       round(100.0*count(*) FILTER (WHERE r.review_score <= 2)/count(r.review_score), 2) AS 差评率
FROM base b
LEFT JOIN (SELECT order_id, max(review_score) review_score FROM olist.order_reviews GROUP BY 1) r ON r.order_id = b.order_id
GROUP BY 1 ORDER BY 1;

-- 3.4 按"影响面"排序（超额延误总量 = 超出均值的天数 × 订单量）
--     说明：解决问题要看影响总量，不是看谁最严重
WITH sp AS (
    SELECT i.seller_id,
           count(DISTINCT o.order_id) AS orders,
           avg(EXTRACT(EPOCH FROM (o.order_delivered_carrier_date - o.order_approved_at))/86400) AS prep
    FROM olist.orders o
    JOIN olist.order_items i ON i.order_id = o.order_id
    WHERE o.order_delivered_carrier_date IS NOT NULL AND o.order_approved_at IS NOT NULL
    GROUP BY 1 HAVING count(DISTINCT o.order_id) >= 50
)
SELECT left(sp.seller_id, 10) AS 卖家,
       sp.orders AS 订单数,
       round(sp.prep::numeric, 1) AS 备货天数,
       s.seller_state AS 所在州,
       round(((sp.prep - 2.8) * sp.orders)::numeric, 0) AS 超额延误总量
FROM sp
JOIN olist.sellers s ON s.seller_id = sp.seller_id
ORDER BY 超额延误总量 DESC
LIMIT 15;

-- 3.5 单卖家 vs 多卖家订单（多卖家需等所有卖家发货）
--     注意：多卖家订单的一星率异常高，但时间戳可能失真（见第 7 部分局限）
WITH oc AS (SELECT order_id, count(DISTINCT seller_id) AS sellers FROM olist.order_items GROUP BY 1)
SELECT CASE WHEN oc.sellers = 1 THEN '单卖家'
            WHEN oc.sellers = 2 THEN '2个卖家'
            ELSE '3个及以上' END AS 类型,
       count(*) AS 订单数,
       round(avg(EXTRACT(EPOCH FROM (o.order_delivered_customer_date - o.order_purchase_timestamp))/86400)::numeric, 1) AS 全程天数,
       round(100.0*count(*) FILTER (WHERE o.order_delivered_customer_date > o.order_estimated_delivery_date)/count(*), 1) AS 迟到率,
       round(100.0*count(*) FILTER (WHERE r.review_score <= 2)/count(r.review_score), 2) AS 差评率
FROM olist.orders o
JOIN oc ON oc.order_id = o.order_id
LEFT JOIN (SELECT order_id, max(review_score) review_score FROM olist.order_reviews GROUP BY 1) r ON r.order_id = o.order_id
WHERE o.order_delivered_customer_date IS NOT NULL
GROUP BY 1 ORDER BY 1;


-- ████████████████████████████████████████████████████████████████████████████
-- 第 4 部分  商品因素（品类 / 重量 / 件数）
-- ████████████████████████████████████████████████████████████████████████████

-- 4.1 商品重量 → 备货时长
WITH base AS (
    SELECT o.order_id,
           max(p.product_weight_g) AS w,
           EXTRACT(EPOCH FROM (max(o.order_delivered_carrier_date) - max(o.order_approved_at)))/86400 AS prep
    FROM olist.orders o
    JOIN olist.order_items i ON i.order_id = o.order_id
    JOIN olist.products p ON p.product_id = i.product_id
    WHERE o.order_delivered_carrier_date IS NOT NULL AND o.order_approved_at IS NOT NULL
    GROUP BY o.order_id
)
SELECT CASE WHEN w IS NULL    THEN '未知'
            WHEN w < 500      THEN '1) 小于0.5kg'
            WHEN w < 2000     THEN '2) 0.5-2kg'
            WHEN w < 5000     THEN '3) 2-5kg'
            WHEN w < 15000    THEN '4) 5-15kg'
            ELSE                   '5) 大于15kg' END AS 重量档,
       count(*) AS 订单数,
       round(avg(prep)::numeric, 2) AS 平均备货天数
FROM base GROUP BY 1 ORDER BY 1;

-- 4.2 品类 → 备货时长（样本 >= 300）
WITH base AS (
    SELECT o.order_id,
           t.product_category_name_english AS cat,
           EXTRACT(EPOCH FROM (max(o.order_delivered_carrier_date) - max(o.order_approved_at)))/86400 AS prep
    FROM olist.orders o
    JOIN olist.order_items i ON i.order_id = o.order_id
    JOIN olist.products p ON p.product_id = i.product_id
    LEFT JOIN olist.product_category_translation t ON t.product_category_name = p.product_category_name
    WHERE o.order_delivered_carrier_date IS NOT NULL
      AND o.order_approved_at IS NOT NULL
      AND p.product_category_name IS NOT NULL
    GROUP BY o.order_id, t.product_category_name_english
)
SELECT cat AS 品类, count(*) AS 订单数,
       round(avg(prep)::numeric, 2) AS 平均备货天数
FROM base GROUP BY 1 HAVING count(*) >= 300
ORDER BY 平均备货天数 DESC
LIMIT 15;

-- 4.3 一件 vs 多件商品
WITH oc AS (SELECT order_id, count(*) AS items FROM olist.order_items GROUP BY 1)
SELECT CASE WHEN oc.items = 1 THEN '1件'
            WHEN oc.items <= 3 THEN '2-3件'
            ELSE '4件以上' END AS 件数,
       count(*) AS 订单数,
       round(avg(EXTRACT(EPOCH FROM (o.order_delivered_carrier_date - o.order_approved_at))/86400)::numeric, 2) AS 平均备货天数
FROM olist.orders o
JOIN oc ON oc.order_id = o.order_id
WHERE o.order_delivered_carrier_date IS NOT NULL AND o.order_approved_at IS NOT NULL
GROUP BY 1 ORDER BY 1;

-- 4.4 排除"办公家具"后，卖家备货分化还剩多少 ★ 混杂因素分层验证
WITH per_item AS (
    SELECT o.order_id, i.seller_id
    FROM olist.orders o
    JOIN olist.order_items i ON i.order_id = o.order_id
    JOIN olist.products p ON p.product_id = i.product_id
    LEFT JOIN olist.product_category_translation t ON t.product_category_name = p.product_category_name
    WHERE o.order_delivered_carrier_date IS NOT NULL
      AND o.order_approved_at IS NOT NULL
      AND coalesce(t.product_category_name_english, '') <> 'office_furniture'
    GROUP BY o.order_id, i.seller_id
),
sp AS (
    SELECT pi.seller_id,
           count(DISTINCT pi.order_id) AS orders,
           avg(EXTRACT(EPOCH FROM (o.order_delivered_carrier_date - o.order_approved_at))/86400) AS prep
    FROM per_item pi
    JOIN olist.orders o ON o.order_id = pi.order_id
    GROUP BY 1
    HAVING count(DISTINCT pi.order_id) >= 50
)
SELECT '排除办公家具后' AS 口径,
       count(*) AS 卖家数,
       round(avg(prep)::numeric, 2) AS 平均备货天数,
       round(max(prep)::numeric, 1) AS 最慢卖家备货天数
FROM sp;


-- ████████████████████████████████████████████████████████████████████████████
-- 第 5 部分  时间因素（趋势 / 旺季 / 异常）
-- ████████████████████████████████████████████████████████████████████████████

-- 5.1 月度趋势
--     注意：2016-09(4单) 2016-12(1单) 2018-09(16单) 2018-10(4单) 数据不完整，分析时剔除
SELECT to_char(o.order_purchase_timestamp, 'YYYY-MM') AS 月份,
       count(*) AS 订单数,
       round(avg(EXTRACT(EPOCH FROM (o.order_delivered_customer_date - o.order_purchase_timestamp))/86400)::numeric, 1) AS 全程天数,
       round(100.0*count(*) FILTER (WHERE o.order_delivered_customer_date > o.order_estimated_delivery_date)
             / nullif(count(*) FILTER (WHERE o.order_delivered_customer_date IS NOT NULL), 0), 1) AS 迟到率,
       round(avg(r.review_score), 2) AS 平均评分,
       round(100.0*count(*) FILTER (WHERE r.review_score <= 2)/count(r.review_score), 2) AS 差评率
FROM olist.orders o
LEFT JOIN (SELECT order_id, max(review_score) review_score FROM olist.order_reviews GROUP BY 1) r ON r.order_id = o.order_id
GROUP BY 1 ORDER BY 1;

-- 5.2 迟到率由"平台承诺"决定 ★ 反直觉的关键洞察
--     看"平均富余"：富余越小的月份，迟到率越高
SELECT to_char(order_purchase_timestamp, 'YYYY-MM') AS 月份,
       count(*) AS 订单数,
       round(avg(EXTRACT(EPOCH FROM (order_estimated_delivery_date - order_purchase_timestamp))/86400)::numeric, 1) AS 平台预估天数,
       round(avg(EXTRACT(EPOCH FROM (order_delivered_customer_date - order_purchase_timestamp))/86400)::numeric, 1) AS 实际全程天数,
       round(avg(EXTRACT(EPOCH FROM (order_estimated_delivery_date - order_delivered_customer_date))/86400)::numeric, 1) AS 平均富余天数,
       round(100.0*count(*) FILTER (WHERE order_delivered_customer_date > order_estimated_delivery_date)
             / nullif(count(*) FILTER (WHERE order_delivered_customer_date IS NOT NULL), 0), 1) AS 迟到率
FROM olist.orders
WHERE order_purchase_timestamp >= '2017-01-01' AND order_purchase_timestamp < '2018-09-01'
GROUP BY 1 ORDER BY 1;

-- 5.3 周几下单的影响
SELECT to_char(order_purchase_timestamp, 'Dy') AS 星期,
       extract(dow from order_purchase_timestamp)::int AS 序号,
       count(*) AS 订单数,
       round(avg(EXTRACT(EPOCH FROM (order_delivered_customer_date - order_purchase_timestamp))/86400)::numeric, 1) AS 全程天数,
       round(100.0*count(*) FILTER (WHERE order_delivered_customer_date > order_estimated_delivery_date)
             / nullif(count(*) FILTER (WHERE order_delivered_customer_date IS NOT NULL), 0), 1) AS 迟到率
FROM olist.orders
WHERE order_delivered_customer_date IS NOT NULL
GROUP BY 1, 2 ORDER BY 2;

-- 5.4 2018-02/03 异常归因：是哪一环节出的问题 ★
--     结果：备货稳定 3.0-3.2 天，运输从 10.5 涨到 13.4 天 → 运输事故
SELECT to_char(order_purchase_timestamp, 'YYYY-MM') AS 月份,
       count(*) AS 订单数,
       round(avg(EXTRACT(EPOCH FROM (order_estimated_delivery_date - order_purchase_timestamp))/86400)::numeric, 1) AS 预估天数,
       round(avg(EXTRACT(EPOCH FROM (order_delivered_customer_date - order_purchase_timestamp))/86400)::numeric, 1) AS 实际天数,
       round(avg(EXTRACT(EPOCH FROM (order_delivered_customer_date - order_delivered_carrier_date))/86400)::numeric, 1) AS 运输天数,
       round(avg(EXTRACT(EPOCH FROM (order_delivered_carrier_date - order_approved_at))/86400)::numeric, 1) AS 备货天数
FROM olist.orders
WHERE order_purchase_timestamp >= '2018-01-01' AND order_purchase_timestamp < '2018-07-01'
  AND order_delivered_customer_date IS NOT NULL
  AND order_delivered_carrier_date IS NOT NULL
  AND order_approved_at IS NOT NULL
GROUP BY 1 ORDER BY 1;

-- 5.5 2018-03 各州迟到率（判断是地域性还是全国性）
SELECT c.customer_state AS 客户州,
       count(*) AS 订单数,
       round(100.0*count(*) FILTER (WHERE o.order_delivered_customer_date > o.order_estimated_delivery_date)/count(*), 1) AS 迟到率,
       round(avg(EXTRACT(EPOCH FROM (o.order_estimated_delivery_date - o.order_purchase_timestamp))/86400)::numeric, 1) AS 预估天数,
       round(avg(EXTRACT(EPOCH FROM (o.order_delivered_customer_date - o.order_purchase_timestamp))/86400)::numeric, 1) AS 实际天数
FROM olist.orders o
JOIN olist.customers c ON c.customer_id = o.customer_id
WHERE o.order_purchase_timestamp >= '2018-03-01' AND o.order_purchase_timestamp < '2018-04-01'
  AND o.order_delivered_customer_date IS NOT NULL
GROUP BY 1 HAVING count(*) >= 100
ORDER BY 迟到率 DESC;


-- ████████████████████████████████████████████████████████████████████████████
-- 第 6 部分  汇总对比：各因素对差评的影响权重   ★★★ 核心结论
-- ████████████████████████████████████████████████████████████████████████████

-- 6.1 五个因素并列对比（统一口径：差评率 = 1-2星占已评订单比例）
--     基线：见 0.3 的"差评率基线"
--     阅读方法：看"差评率"列与基线的差距，差距越大影响越强
WITH base AS (
    SELECT o.order_id,
           EXTRACT(EPOCH FROM (o.order_delivered_carrier_date - o.order_approved_at))/86400          AS prep,
           EXTRACT(EPOCH FROM (o.order_delivered_customer_date - o.order_delivered_carrier_date))/86400 AS ship,
           EXTRACT(EPOCH FROM (o.order_estimated_delivery_date - o.order_delivered_customer_date))/86400 AS buffer,
           (o.order_delivered_customer_date > o.order_estimated_delivery_date) AS is_late,
           c.customer_state,
           s.seller_state,
           CASE WHEN max(i.shipping_limit_date) IS NOT NULL
                 AND max(o.order_delivered_carrier_date) > max(i.shipping_limit_date)
                THEN 1 ELSE 0 END AS is_overdue,
           r.review_score
    FROM olist.orders o
    JOIN olist.customers c ON c.customer_id = o.customer_id
    LEFT JOIN olist.order_items i ON i.order_id = o.order_id
    LEFT JOIN olist.sellers s ON s.seller_id = (
        SELECT min(seller_id) FROM olist.order_items WHERE order_id = o.order_id)
    LEFT JOIN (SELECT order_id, max(review_score) review_score FROM olist.order_reviews GROUP BY 1) r ON r.order_id = o.order_id
    WHERE o.order_delivered_customer_date IS NOT NULL
      AND o.order_delivered_carrier_date IS NOT NULL
      AND o.order_approved_at IS NOT NULL
    GROUP BY o.order_id, c.customer_state, s.seller_state, r.review_score,
             o.order_estimated_delivery_date, o.order_delivered_customer_date,
             o.order_delivered_carrier_date, o.order_approved_at
)
SELECT '① 卖家超期发货' AS 因素,
       CASE WHEN is_overdue = 1 THEN '超期发货' ELSE '按期发货' END AS 分组,
       count(*) AS 订单数,
       round(100.0*count(*) FILTER (WHERE review_score <= 2)/nullif(count(review_score),0), 2) AS 差评率
FROM base GROUP BY 1, 2

UNION ALL
SELECT '② 是否迟到',
       CASE WHEN is_late THEN '迟到' ELSE '准时' END,
       count(*),
       round(100.0*count(*) FILTER (WHERE review_score <= 2)/nullif(count(review_score),0), 2)
FROM base GROUP BY 1, 2

UNION ALL
SELECT '③ 是否跨州',
       CASE WHEN customer_state = seller_state THEN '同州' ELSE '跨州' END,
       count(*),
       round(100.0*count(*) FILTER (WHERE review_score <= 2)/nullif(count(review_score),0), 2)
FROM base WHERE seller_state IS NOT NULL GROUP BY 1, 2

UNION ALL
SELECT '④ 备货时长',
       CASE WHEN prep <= 1 THEN '1天内'
            WHEN prep <= 3 THEN '1-3天'
            WHEN prep <= 7 THEN '3-7天'
            WHEN prep <= 14 THEN '7-14天'
            ELSE '14天以上' END,
       count(*),
       round(100.0*count(*) FILTER (WHERE review_score <= 2)/nullif(count(review_score),0), 2)
FROM base WHERE prep IS NOT NULL GROUP BY 1, 2

UNION ALL
SELECT '⑤ 运输时长',
       CASE WHEN ship <= 3 THEN '3天内'
            WHEN ship <= 7 THEN '3-7天'
            WHEN ship <= 14 THEN '7-14天'
            WHEN ship <= 21 THEN '14-21天'
            ELSE '21天以上' END,
       count(*),
       round(100.0*count(*) FILTER (WHERE review_score <= 2)/nullif(count(review_score),0), 2)
FROM base WHERE ship IS NOT NULL AND ship >= 0 GROUP BY 1, 2

ORDER BY 因素, 分组;

-- 6.2 预估富余天数 → 差评率（验证"承诺"本身是否影响满意度）
WITH base AS (
    SELECT EXTRACT(EPOCH FROM (o.order_estimated_delivery_date - o.order_delivered_customer_date))/86400 AS buffer,
           r.review_score
    FROM olist.orders o
    LEFT JOIN (SELECT order_id, max(review_score) review_score FROM olist.order_reviews GROUP BY 1) r ON r.order_id = o.order_id
    WHERE o.order_delivered_customer_date IS NOT NULL
      AND o.order_estimated_delivery_date IS NOT NULL
)
SELECT CASE WHEN buffer < 0  THEN '① 已超期(负富余)'
            WHEN buffer < 3  THEN '② 富余 0-3天'
            WHEN buffer < 7  THEN '③ 富余 3-7天'
            WHEN buffer < 15 THEN '④ 富余 7-15天'
            ELSE                  '⑤ 富余 15天以上' END AS 预估富余档,
       count(*) AS 订单数,
       round(100.0*count(*) FILTER (WHERE review_score <= 2)/nullif(count(review_score),0), 2) AS 差评率,
       round(avg(review_score), 2) AS 平均评分
FROM base
GROUP BY 1 ORDER BY 1;

-- 6.3 因素交叉：超期发货 × 跨州（看是否叠加）
WITH base AS (
    SELECT o.order_id,
           (max(o.order_delivered_carrier_date) > max(i.shipping_limit_date)) AS is_overdue,
           (c.customer_state <> s.seller_state) AS is_cross,
           r.review_score
    FROM olist.orders o
    JOIN olist.customers c ON c.customer_id = o.customer_id
    JOIN olist.order_items i ON i.order_id = o.order_id
    JOIN olist.sellers s ON s.seller_id = i.seller_id
    LEFT JOIN (SELECT order_id, max(review_score) review_score FROM olist.order_reviews GROUP BY 1) r ON r.order_id = o.order_id
    WHERE o.order_delivered_customer_date IS NOT NULL
      AND o.order_delivered_carrier_date IS NOT NULL
      AND i.shipping_limit_date IS NOT NULL
    GROUP BY o.order_id, c.customer_state, s.seller_state, r.review_score
)
SELECT CASE WHEN is_overdue THEN '超期' ELSE '按期' END AS 发货,
       CASE WHEN is_cross THEN '跨州' ELSE '同州' END AS 距离,
       count(*) AS 订单数,
       round(100.0*count(*) FILTER (WHERE review_score <= 2)/nullif(count(review_score),0), 2) AS 差评率
FROM base
GROUP BY 1, 2 ORDER BY 1, 2;


-- ████████████████████████████████████████████████████████████████████████████
-- 第 7 部分  数据质量与局限（写报告的"数据说明"部分用）
-- ████████████████████████████████████████████████████████████████████████████

-- 7.1 关键字段缺失率
SELECT '订单-审核时间'   AS 字段, count(*) FILTER (WHERE order_approved_at IS NULL) AS 缺失数, count(*) AS 总数,
       round(100.0*count(*) FILTER (WHERE order_approved_at IS NULL)/count(*), 2) AS 缺失率
FROM olist.orders
UNION ALL
SELECT '订单-签收时间', count(*) FILTER (WHERE order_delivered_customer_date IS NULL), count(*),
       round(100.0*count(*) FILTER (WHERE order_delivered_customer_date IS NULL)/count(*), 2)
FROM olist.orders
UNION ALL
SELECT '商品-品类', count(*) FILTER (WHERE product_category_name IS NULL), count(*),
       round(100.0*count(*) FILTER (WHERE product_category_name IS NULL)/count(*), 2)
FROM olist.products
UNION ALL
SELECT '评价-评论内容', count(*) FILTER (WHERE review_comment_message IS NULL), count(*),
       round(100.0*count(*) FILTER (WHERE review_comment_message IS NULL)/count(*), 2)
FROM olist.order_reviews;

-- 7.2 时间戳异常（负值：签收早于发货）
SELECT count(*) AS 负运输天数条数,
       round(100.0*count(*)/(SELECT count(*) FROM olist.orders
            WHERE order_delivered_customer_date IS NOT NULL
              AND order_delivered_carrier_date IS NOT NULL), 4) AS 占比
FROM olist.orders
WHERE order_delivered_customer_date < order_delivered_carrier_date
  AND order_delivered_customer_date IS NOT NULL
  AND order_delivered_carrier_date IS NOT NULL;

-- 7.3 订单状态分布（明确分析口径：是否只分析 delivered）
SELECT order_status AS 状态,
       count(*) AS 订单数,
       round(100.0*count(*)/sum(count(*)) OVER (), 2) AS 占比
FROM olist.orders
GROUP BY 1 ORDER BY 订单数 DESC;

-- 7.4 顾客标识的唯一性（说明为何必须用 customer_unique_id）
SELECT count(*) AS 行数,
       count(DISTINCT customer_id)        AS 唯一customer_id,
       count(DISTINCT customer_unique_id) AS 唯一customer_unique_id
FROM olist.customers;
