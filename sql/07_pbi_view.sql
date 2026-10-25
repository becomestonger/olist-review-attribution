-- ============================================================================
-- Power BI 专用宽表（订单级）
--
-- 作用：把散在 7 张表里的字段合成一张「一行一单」的宽表，
--       让 Power BI 里只需要导入 1 个对象就能做完整报表，
--       不用建关系、不用在 DAX 里写 JOIN。
--
-- 设计要点：
--   ① 用 LEFT JOIN 保留全部 99,441 笔订单（775 笔无商品明细的订单也要留住，
--      否则差评率会从 14.63% 变成 14.13%，和报告对不上）
--   ② 比率类标记位在「无法判断」时返回 NULL 而不是 0，
--      这样 Power BI 里用 DIVIDE(分子, 分母) 会自动排除它们
--
-- 输出：olist.v_pbi_orders
--
-- 运行方式（任选其一）：
--   ① PowerShell（推荐，最省事）：
--      & "C:\Users\28223\pgsql\bin\psql.exe" -h 127.0.0.1 -p 5432 -U postgres -d analysis -f "C:\Users\28223\Desktop\巴西olist数据分析项目\sql\07_pbi_view.sql"
--   ② VS Code：打开本文件，全选 → Ctrl+Enter（或右键 Run Query）
--   ③ pgAdmin：右键 analysis → Query Tool → 粘贴内容 → F5
--
-- 密码：123456（postgres 用户）
-- 本文件可重复执行（先 DROP 再 CREATE）
-- ============================================================================

DROP VIEW IF EXISTS olist.v_pbi_orders;

CREATE VIEW olist.v_pbi_orders AS
WITH item_agg AS (
    -- 先把明细聚合成「一行一单」，避免一对多 JOIN 造成重复计数
    SELECT order_id,
           sum(price)                 AS items_total,      -- 商品金额合计
           sum(freight_value)         AS freight_total,    -- 运费合计
           sum(price + freight_value) AS order_total,      -- 订单总额
           count(*)                   AS item_count,       -- 商品件数
           max(shipping_limit_date)   AS ship_limit,       -- 卖家承诺发货期限
           count(DISTINCT seller_id)  AS seller_count,     -- 该订单涉及几个卖家
           min(seller_id)             AS main_seller_id,   -- 主卖家（取第一个）
           min(product_id)            AS main_product_id   -- 主商品（取第一个）
    FROM olist.order_items
    GROUP BY order_id
),
pay AS (
    -- 支付合计（一个订单可能分多笔支付）
    SELECT order_id, sum(payment_value) AS payment_total
    FROM olist.order_payments
    GROUP BY order_id
),
rev AS (
    -- 评分：一单可能有多条评价，取最高分作为该单评分
    SELECT order_id, max(review_score) AS review_score
    FROM olist.order_reviews
    GROUP BY order_id
)
SELECT
    -- ---------- 主键与状态 ----------
    o.order_id,
    o.order_status,

    -- ---------- 时间维度 ----------
    o.order_purchase_timestamp::date                      AS purchase_date,
    date_trunc('month', o.order_purchase_timestamp)::date AS purchase_month,

    -- ---------- 客户维度 ----------
    c.customer_unique_id,
    c.customer_city,
    c.customer_state,

    -- ---------- 卖家 / 品类维度 ----------
    s.seller_state                                        AS main_seller_state,
    coalesce(t.product_category_name_english, 'unknown')  AS main_category,

    -- ---------- 数量与金额 ----------
    a.item_count,
    a.seller_count,
    round(a.items_total::numeric,   2) AS items_total,
    round(a.freight_total::numeric, 2) AS freight_total,
    round(a.order_total::numeric,   2) AS order_total,
    round(p.payment_total::numeric, 2) AS payment_total,

    -- ---------- 结果指标（被解释变量）----------
    r.review_score,
    -- 差评标记：1=差评，0=非差评，NULL=没有评分（不参与差评率分母）
    CASE WHEN r.review_score IS NULL THEN NULL
         WHEN r.review_score <= 2 THEN 1 ELSE 0 END AS is_bad_review,

    -- ---------- 过程指标（解释变量）----------

    -- 是否已送达：1/0，始终有值
    CASE WHEN o.order_delivered_customer_date IS NOT NULL THEN 1 ELSE 0 END AS is_delivered,

    -- 是否迟到：签收时间 > 平台预估送达时间。未送达 = NULL（不参与迟到率分母）
    CASE WHEN o.order_delivered_customer_date IS NULL THEN NULL
         WHEN o.order_estimated_delivery_date IS NULL THEN NULL
         WHEN o.order_delivered_customer_date > o.order_estimated_delivery_date THEN 1
         ELSE 0 END AS is_late,

    -- 是否超期发货：交给物流时间 > 卖家承诺发货期限。无明细/无期限 = NULL
    CASE WHEN a.ship_limit IS NULL THEN NULL
         WHEN o.order_delivered_carrier_date IS NULL THEN NULL
         WHEN o.order_delivered_carrier_date > a.ship_limit THEN 1
         ELSE 0 END AS is_overdue_ship,

    -- 是否跨州：客户州 ≠ 卖家州。任一方缺失 = NULL
    CASE WHEN c.customer_state IS NULL OR s.seller_state IS NULL THEN NULL
         WHEN c.customer_state <> s.seller_state THEN 1
         ELSE 0 END AS is_cross_state,

    -- 四段耗时（天）：审核 / 备货 / 运输 / 全程（时间戳缺失则为 NULL）
    round((EXTRACT(EPOCH FROM (o.order_approved_at
                               - o.order_purchase_timestamp)) / 86400)::numeric, 2) AS approve_days,
    round((EXTRACT(EPOCH FROM (o.order_delivered_carrier_date
                               - o.order_approved_at)) / 86400)::numeric, 2)        AS prep_days,
    round((EXTRACT(EPOCH FROM (o.order_delivered_customer_date
                               - o.order_delivered_carrier_date)) / 86400)::numeric, 2) AS ship_days,
    round((EXTRACT(EPOCH FROM (o.order_delivered_customer_date
                               - o.order_purchase_timestamp)) / 86400)::numeric, 2) AS total_days,

    -- 预估富余：预估送达 - 实际签收。负数 = 已超期，正数 = 提前到
    round((EXTRACT(EPOCH FROM (o.order_estimated_delivery_date
                               - o.order_delivered_customer_date)) / 86400)::numeric, 2) AS buffer_days
FROM olist.orders o
JOIN olist.customers c            ON c.customer_id = o.customer_id
LEFT JOIN item_agg a              ON a.order_id    = o.order_id
LEFT JOIN olist.sellers s         ON s.seller_id   = a.main_seller_id
LEFT JOIN olist.products pr       ON pr.product_id = a.main_product_id
LEFT JOIN olist.product_category_translation t
                                  ON t.product_category_name = pr.product_category_name
LEFT JOIN pay p                   ON p.order_id    = o.order_id
LEFT JOIN rev r                   ON r.order_id    = o.order_id;

COMMENT ON VIEW olist.v_pbi_orders IS
    'Power BI 专用订单级宽表：一行一单（99,441 行），含四段耗时、是否迟到/超期/跨州、评分与差评标记';


-- ============================================================================
-- 验证：跑完下面 4 条 SELECT，确认视图建好且数字和报告一致
-- （故意不用 psql 的 \echo，因为 VS Code 和 pgAdmin 不支持元命令）
-- ============================================================================

-- 1. 视图是否创建成功（应返回 1 行：v_pbi_orders）
SELECT table_name AS 视图名
FROM information_schema.views
WHERE table_schema = 'olist' AND table_name = 'v_pbi_orders';

-- 2. 行数（应为 99,441）
SELECT count(*) AS 订单数_应99441 FROM olist.v_pbi_orders;

-- 3. 关键指标核对（应和 59 页分析报告完全一致）
SELECT round(100.0 * count(*) FILTER (WHERE is_bad_review = 1)
             / nullif(count(is_bad_review), 0), 2)  AS 差评率_应14_63,
       round(100.0 * count(*) FILTER (WHERE is_late = 1)
             / nullif(count(is_late), 0), 2)        AS 迟到率_应8_11,
       round(100.0 * count(*) FILTER (WHERE is_overdue_ship = 1)
             / nullif(count(is_overdue_ship), 0), 2) AS 超期发货率_应9_35,
       round(avg(prep_days)::numeric, 2)            AS 平均备货_应2_80,
       round(avg(ship_days)::numeric, 2)            AS 平均运输_应9_33,
       round(avg(review_score), 2)                  AS 平均评分_应4_17
FROM olist.v_pbi_orders;

-- 4. 前 3 行抽样（看字段是否正常，NULL 是否合理）
SELECT order_id, purchase_month, customer_state, main_category,
       review_score, is_bad_review, is_late, prep_days, ship_days, total_days
FROM olist.v_pbi_orders
ORDER BY purchase_date
LIMIT 3;
