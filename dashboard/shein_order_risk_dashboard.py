# -*- coding: utf-8 -*-
"""
SHEIN 订单超时风险网页版看板

用法：
1. 安装依赖：
   pip install streamlit pandas openpyxl

2. 启动：
   streamlit run shein_order_risk_dashboard.py

3. 浏览器打开页面后，上传你当前脚本导出的 Excel 表格。

说明：
- 表格时间默认已经是洛杉矶时间。
- 当前时间按洛杉矶时间计算。
- 不再依赖导出地址接口。
"""

import re
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st


LA_TZ = ZoneInfo("America/Los_Angeles")


# =========================
# 字段名称
# =========================

COL_STORE = "店铺"
COL_ORDER_NO = "订单编号"
COL_STATUS = "平台原始状态"
COL_CREATE_TIME = "订单创建时间"
COL_EXPORT_ADDRESS_TIME = "导出地址时间"
COL_PRINT_TIME = "打印面单时间"
COL_PENDING_PROCESS_TIMEOUT = "待处理超时时间"
COL_PENDING_SHIP_TIMEOUT = "待发货超时时间"
COL_PENDING_COLLECT_TIMEOUT = "待揽收超时时间"
COL_REQUIRE_SIGN_TIME = "要求签收时间"
COL_REFUND_TIME = "订单退款时间"
COL_WAYBILL = "运单号"
COL_LOGISTICS = "物流信息"


st.set_page_config(
    page_title="SHEIN 订单超时风险看板",
    page_icon="📦",
    layout="wide",
)


# =========================
# 工具函数
# =========================

def now_la() -> datetime:
    return datetime.now(LA_TZ).replace(tzinfo=None)


def safe_str(v) -> str:
    if pd.isna(v):
        return ""
    return str(v).strip()


def parse_dt(v):
    """表格时间默认已经是洛杉矶时间，返回 naive datetime。"""
    if pd.isna(v) or v == "":
        return None

    if isinstance(v, datetime):
        return v.replace(tzinfo=None)

    s = str(v).strip()
    if not s:
        return None

    s = s.replace("/", "-")
    s = re.sub(r"\.0$", "", s)

    for fmt in ["%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"]:
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            pass

    try:
        dt = pd.to_datetime(s, errors="coerce")
        if pd.isna(dt):
            return None
        return dt.to_pydatetime().replace(tzinfo=None)
    except Exception:
        return None


def hours_passed(start_dt, current_dt):
    if not start_dt:
        return None
    return (current_dt - start_dt).total_seconds() / 3600


def hours_left(target_dt, current_dt):
    if not target_dt:
        return None
    return (target_dt - current_dt).total_seconds() / 3600


def is_empty(v) -> bool:
    s = safe_str(v)
    return s == "" or s.lower() in ["nan", "none", "null"]


def normalize_status(status: str) -> str:
    s = safe_str(status)

    if "待处理" in s:
        return "待处理"
    if "待发货" in s:
        return "待发货"
    if "待揽收" in s:
        return "待揽收"
    if "已发货" in s or "发货中" in s or "SHEIN发货中" in s:
        return "已发货"
    if "已签收" in s or "已送达" in s:
        return "已签收"
    if "取消" in s or "已取消" in s:
        return "取消"
    if "退款" in s:
        return "退款"

    return s


def logistics_has_delivered(logistics: str) -> bool:
    s = safe_str(logistics).lower()
    keywords = ["已签收", "已送达", "delivered", "delivery complete"]
    return any(k.lower() in s for k in keywords)


def logistics_has_pickup_or_progress(logistics: str) -> bool:
    """
    判断物流是否已经从“已创建面单/已创建发货标签”进入下一阶段。
    """
    s = safe_str(logistics).lower()
    if not s:
        return False

    positive_keywords = [
        "已揽收",
        "揽收",
        "运输中",
        "已抵达",
        "已到达",
        "已离开",
        "正在运往",
        "处理中",
        "usps 处理中心",
        "accepted",
        "acceptance",
        "in transit",
        "arrived",
        "departed",
        "processed",
        "origin facility",
        "regional facility",
        "distribution center",
        "post office",
        "picked up",
        "received by",
        "possession",
    ]

    label_only_keywords = [
        "已创建面单",
        "已创建发货标签",
        "shipping label created",
        "label created",
        "pre-shipment",
        "pre shipment",
        "usps awaits item",
    ]

    if any(k in s for k in positive_keywords):
        return True

    if any(k in s for k in label_only_keywords):
        return False

    return False


def risk_level_sort(level: str) -> int:
    order = {"红色": 0, "黄色": 1, "灰色": 2, "绿色": 3}
    return order.get(level, 9)


def format_hours(h):
    if h is None:
        return ""
    sign = "-" if h < 0 else ""
    h = abs(h)
    return f"{sign}{h:.1f}小时"


def add_risk(risks, risk_type, level, metric, action):
    risks.append({
        "风险类型": risk_type,
        "风险等级": level,
        "时间指标": metric,
        "建议操作": action,
    })


# =========================
# 核心风险判断
# =========================

def analyze_one_order(row, current_dt):
    status = normalize_status(row.get(COL_STATUS, ""))
    create_dt = parse_dt(row.get(COL_CREATE_TIME, ""))
    pending_collect_timeout = parse_dt(row.get(COL_PENDING_COLLECT_TIMEOUT, ""))
    require_sign_time = parse_dt(row.get(COL_REQUIRE_SIGN_TIME, ""))

    waybill = safe_str(row.get(COL_WAYBILL, ""))
    logistics = safe_str(row.get(COL_LOGISTICS, ""))

    age_h = hours_passed(create_dt, current_dt)
    has_waybill = not is_empty(waybill)
    has_pickup = logistics_has_pickup_or_progress(logistics)
    delivered = logistics_has_delivered(logistics)

    risks = []

    if status == "已签收":
        return [{"风险类型": "已完成", "风险等级": "灰色", "时间指标": "", "建议操作": "订单已签收，无需处理。"}]

    if status in ["取消", "退款"] or not is_empty(row.get(COL_REFUND_TIME, "")):
        return [{"风险类型": "已取消/已退款", "风险等级": "灰色", "时间指标": "", "建议操作": "订单已取消或退款，无需处理。"}]

    # 1. 即将处理超时：待处理，按订单创建时间正向计算
    if status == "待处理" and age_h is not None:
        if age_h >= 24:
            add_risk(risks, "已处理超时", "红色", f"已过去 {format_hours(age_h)}", "立即检查是否已导出地址、在线下单或打印面单。")
        elif 18 <= age_h < 24:
            add_risk(risks, "即将处理超时", "红色", f"已过去 {format_hours(age_h)}", "优先处理：确认是否完成导出地址、在线下单或打印面单。")
        elif 12 <= age_h < 18:
            add_risk(risks, "即将处理超时", "黄色", f"已过去 {format_hours(age_h)}", "关注处理进度，避免进入 18 小时红色风险。")

    # 2. 即将发货超时：待处理/待发货，按订单创建时间正向计算，主要看有没有运单号
    if status in ["待处理", "待发货"] and age_h is not None and not has_waybill:
        if age_h >= 48:
            add_risk(risks, "已发货超时", "红色", f"已过去 {format_hours(age_h)}", "立即上传运单号或检查在线下单/物流单号是否生成。")
        elif 36 <= age_h < 48:
            add_risk(risks, "即将发货超时", "红色", f"已过去 {format_hours(age_h)}", "优先处理：还没有运单号，距离 48 小时发货风险很近。")
        elif 24 <= age_h < 36:
            add_risk(risks, "即将发货超时", "黄色", f"已过去 {format_hours(age_h)}", "关注是否生成/上传运单号。")

    # 3. 即将延迟交付：待揽收，按订单创建时间正向 44-48h，且物流未进入已揽收/下一阶段
    if status == "待揽收" and age_h is not None and not has_pickup:
        if age_h >= 48:
            add_risk(risks, "已延迟交付", "红色", f"已过去 {format_hours(age_h)}", "48小时内未出现第一次揽收/后续轨迹，检查包裹是否交给承运商。")
        elif 44 <= age_h < 48:
            add_risk(risks, "即将延迟交付", "黄色", f"已过去 {format_hours(age_h)}", "物流仍停留在已创建面单/无后续轨迹，需尽快确认是否交运。")

    # 4. 即将揽收超时：待揽收，按待揽收超时时间逆向，且未出现第一次揽收/后续轨迹
    if status == "待揽收" and pending_collect_timeout and not has_pickup:
        left_h = hours_left(pending_collect_timeout, current_dt)

        if left_h <= 0:
            add_risk(risks, "已揽收超时", "红色", f"已超时 {format_hours(left_h)}", "已经超过待揽收超时时间，立即联系物流商核查。")
        elif left_h <= 12:
            add_risk(risks, "即将揽收超时", "红色", f"剩余 {format_hours(left_h)}", "12小时内即将揽收超时，优先核查包裹揽收状态。")
        elif left_h <= 24:
            add_risk(risks, "即将揽收超时", "黄色", f"剩余 {format_hours(left_h)}", "24小时内即将揽收超时，关注是否出现第一次揽收状态。")

    # 5. 即将到货超时：已发货，按要求签收时间逆向，已签收则不算
    if status == "已发货" and require_sign_time and not delivered:
        left_h = hours_left(require_sign_time, current_dt)

        if left_h <= 0:
            add_risk(risks, "已到货超时", "红色", f"已超时 {format_hours(left_h)}", "已经超过要求签收时间，检查是否物流异常、地址问题、天气延误或驿站滞留。")
        elif left_h <= 12:
            add_risk(risks, "即将到货超时", "红色", f"剩余 {format_hours(left_h)}", "12小时内即将到货超时，优先关注运输异常。")
        elif left_h <= 24:
            add_risk(risks, "即将到货超时", "黄色", f"剩余 {format_hours(left_h)}", "24小时内即将到货超时，关注物流是否正常派送。")

    if not risks:
        return [{"风险类型": "正常", "风险等级": "绿色", "时间指标": "", "建议操作": "暂无需要优先处理的风险。"}]

    return risks


def analyze_dataframe(df: pd.DataFrame, current_dt: datetime):
    rows = []

    for idx, row in df.iterrows():
        risks = analyze_one_order(row, current_dt)

        for risk in risks:
            rows.append({
                "原始行号": idx + 1,
                "店铺": row.get(COL_STORE, ""),
                "订单编号": row.get(COL_ORDER_NO, ""),
                "平台原始状态": row.get(COL_STATUS, ""),
                "订单创建时间": row.get(COL_CREATE_TIME, ""),
                "打印面单时间": row.get(COL_PRINT_TIME, ""),
                "待处理超时时间": row.get(COL_PENDING_PROCESS_TIMEOUT, ""),
                "待发货超时时间": row.get(COL_PENDING_SHIP_TIMEOUT, ""),
                "待揽收超时时间": row.get(COL_PENDING_COLLECT_TIMEOUT, ""),
                "要求签收时间": row.get(COL_REQUIRE_SIGN_TIME, ""),
                "运单号": row.get(COL_WAYBILL, ""),
                "物流信息": row.get(COL_LOGISTICS, ""),
                **risk,
            })

    result = pd.DataFrame(rows)

    if not result.empty:
        result["风险排序"] = result["风险等级"].apply(risk_level_sort)
        result = result.sort_values(["风险排序", "风险类型", "店铺", "订单编号"], ascending=True)

    return result


def read_excel_safely(uploaded_file):
    """兼容导出表格前面有标题行的情况。"""
    for header in [0, 1, 2, 3]:
        try:
            uploaded_file.seek(0)
            df = pd.read_excel(uploaded_file, header=header)
            df.columns = [str(c).strip() for c in df.columns]
            if COL_ORDER_NO in df.columns and COL_STATUS in df.columns:
                return df
        except Exception:
            pass

    uploaded_file.seek(0)
    df = pd.read_excel(uploaded_file)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def validate_columns(df):
    required = [
        COL_ORDER_NO,
        COL_STATUS,
        COL_CREATE_TIME,
        COL_PENDING_COLLECT_TIMEOUT,
        COL_REQUIRE_SIGN_TIME,
        COL_WAYBILL,
        COL_LOGISTICS,
    ]
    return [c for c in required if c not in df.columns]


def color_level(val):
    if val == "红色":
        return "background-color: #ffdddd; color: #9b0000; font-weight: bold;"
    if val == "黄色":
        return "background-color: #fff3cd; color: #8a6d00; font-weight: bold;"
    if val == "绿色":
        return "background-color: #ddffdd; color: #006b21;"
    if val == "灰色":
        return "background-color: #eeeeee; color: #666666;"
    return ""


# =========================
# 页面
# =========================

st.title("📦 SHEIN 订单超时风险看板")

current_dt = now_la()

with st.sidebar:
    st.header("设置")
    st.write("当前洛杉矶时间：")
    st.code(current_dt.strftime("%Y-%m-%d %H:%M:%S"))

    uploaded = st.file_uploader("上传订单物流状态跟踪表 Excel", type=["xlsx", "xls"])

    st.caption("规则：表格时间按洛杉矶时间处理。")


if not uploaded:
    st.info("请先在左侧上传 Excel 表格。")
    st.stop()


df = read_excel_safely(uploaded)
missing = validate_columns(df)

if missing:
    st.error("表格缺少必要字段：")
    st.write(missing)
    st.write("当前识别到的字段：")
    st.write(list(df.columns))
    st.stop()


risk_df = analyze_dataframe(df, current_dt)

# =========================
# 风险明细筛选
# =========================

st.subheader("风险明细")

# 第一项改成与顶部统计卡片一致的业务分类
category_options = [
    "订单总数",
    "即将处理超时",
    "即将发货超时",
    "即将延迟交付",
    "即将揽收超时",
    "即将到货超时",
    "已超时/已延迟",
]

# 平台原始状态动态读取
status_values = sorted(
    [safe_str(v) for v in df[COL_STATUS].dropna().unique().tolist() if safe_str(v)]
)
status_options = ["全部"] + status_values

# 风险等级保留
level_options = ["全部"] + sorted(
    risk_df["风险等级"].dropna().unique().tolist(),
    key=risk_level_sort
)

f1, f2, f3, f4 = st.columns([1.15, 1.15, 1.0, 1.7])

with f1:
    category_filter = st.selectbox(
        "统计/风险类型",
        category_options,
        index=0,
    )

with f2:
    status_filter = st.selectbox(
        "平台原始状态",
        status_options,
        index=0,
    )

with f3:
    level_filter = st.selectbox(
        "风险等级",
        level_options,
        index=0,
    )

with f4:
    keyword = st.text_input(
        "搜索订单号 / 运单号 / 店铺 / 物流信息"
    )


def apply_base_filters(frame):
    """
    先按平台原始状态、风险等级、关键字过滤。
    顶部统计卡片会基于这些筛选条件重新计数。
    """
    out = frame.copy()

    if status_filter != "全部":
        out = out[out["平台原始状态"].astype(str) == status_filter]

    if level_filter != "全部":
        out = out[out["风险等级"] == level_filter]

    if keyword.strip():
        kw = keyword.strip()
        mask = (
            out["订单编号"].astype(str).str.contains(kw, case=False, na=False)
            | out["运单号"].astype(str).str.contains(kw, case=False, na=False)
            | out["店铺"].astype(str).str.contains(kw, case=False, na=False)
            | out["物流信息"].astype(str).str.contains(kw, case=False, na=False)
        )
        out = out[mask]

    return out


base_filtered_df = apply_base_filters(risk_df)


def count_risk_in(frame, name):
    return int((frame["风险类型"] == name).sum())


def count_timeout_group(frame):
    return int(frame["风险类型"].isin([
        "已处理超时",
        "已发货超时",
        "已延迟交付",
        "已揽收超时",
        "已到货超时",
    ]).sum())


# “订单总数”按唯一订单号重新计数，避免一个订单多个风险被重复统计
filtered_order_count = int(
    base_filtered_df["订单编号"].astype(str).nunique()
) if not base_filtered_df.empty else 0


# =========================
# 风险统计：每次筛选后重新计数
# =========================

st.subheader("风险统计")

cards = [
    ("订单总数", filtered_order_count),
    ("即将处理超时", count_risk_in(base_filtered_df, "即将处理超时")),
    ("即将发货超时", count_risk_in(base_filtered_df, "即将发货超时")),
    ("即将延迟交付", count_risk_in(base_filtered_df, "即将延迟交付")),
    ("即将揽收超时", count_risk_in(base_filtered_df, "即将揽收超时")),
    ("即将到货超时", count_risk_in(base_filtered_df, "即将到货超时")),
    ("已超时/已延迟", count_timeout_group(base_filtered_df)),
]

cols = st.columns(len(cards))
for col, (name, value) in zip(cols, cards):
    col.metric(name, value)

st.divider()


# =========================
# 根据“统计/风险类型”重新筛选并重新排序
# =========================

show_df = base_filtered_df.copy()

if category_filter == "即将处理超时":
    show_df = show_df[show_df["风险类型"] == "即将处理超时"]

elif category_filter == "即将发货超时":
    show_df = show_df[show_df["风险类型"] == "即将发货超时"]

elif category_filter == "即将延迟交付":
    show_df = show_df[show_df["风险类型"] == "即将延迟交付"]

elif category_filter == "即将揽收超时":
    show_df = show_df[show_df["风险类型"] == "即将揽收超时"]

elif category_filter == "即将到货超时":
    show_df = show_df[show_df["风险类型"] == "即将到货超时"]

elif category_filter == "已超时/已延迟":
    show_df = show_df[show_df["风险类型"].isin([
        "已处理超时",
        "已发货超时",
        "已延迟交付",
        "已揽收超时",
        "已到货超时",
    ])]

# 订单总数 = 展示全部筛选后的订单风险明细
# 每次切换选项后，重新按风险等级和时间指标排列，而不是依赖表格手工排序
if not show_df.empty:
    show_df = show_df.copy()
    show_df["_风险等级排序"] = show_df["风险等级"].map(risk_level_sort)
    show_df = show_df.sort_values(
        by=["_风险等级排序", "风险类型", "订单创建时间", "订单编号"],
        ascending=[True, True, True, True],
        kind="stable",
    )
    show_df = show_df.drop(columns=["_风险等级排序"], errors="ignore")

display_cols = [
    "风险类型",
    "风险等级",
    "时间指标",
    "建议操作",
    "店铺",
    "订单编号",
    "平台原始状态",
    "订单创建时间",
    "打印面单时间",
    "待处理超时时间",
    "待发货超时时间",
    "待揽收超时时间",
    "要求签收时间",
    "运单号",
    "物流信息",
]

show_df = show_df[[c for c in display_cols if c in show_df.columns]]

st.dataframe(
    show_df.style.map(color_level, subset=["风险等级"]),
    use_container_width=True,
    height=620,
)

csv = show_df.to_csv(index=False, encoding="utf-8-sig")
st.download_button(
    "下载当前筛选结果 CSV",
    data=csv,
    file_name=f"shein_order_risk_result_{current_dt.strftime('%Y%m%d_%H%M%S')}.csv",
    mime="text/csv",
)

st.divider()

with st.expander("查看规则说明"):
    st.markdown(
        """
### 当前规则

1. **待处理状态**
- 当前洛杉矶时间 - 订单创建时间
- 12-18小时：黄色，即将处理超时
- 18-24小时：红色，即将处理超时
- 超过24小时：已处理超时

2. **待处理 / 待发货状态**
- 当前洛杉矶时间 - 订单创建时间
- 运单号为空才判断
- 24-36小时：黄色，即将发货超时
- 36-48小时：红色，即将发货超时
- 超过48小时：已发货超时

3. **待揽收状态：第一次超时未揽收**
- 当前洛杉矶时间 - 订单创建时间
- 物流未从“已创建面单/已创建发货标签”进入下一阶段
- 44-48小时：黄色，即将延迟交付
- 超过48小时：红色，已延迟交付

4. **待揽收状态：待揽收超时**
- 待揽收超时时间 - 当前洛杉矶时间
- 物流未出现已揽收/运输中/到达/离开等后续轨迹
- 剩余24小时内：黄色，即将揽收超时
- 剩余12小时内：红色，即将揽收超时
- 小于等于0：已揽收超时

5. **已发货状态：到货超时**
- 要求签收时间 - 当前洛杉矶时间
- 物流未签收
- 剩余24小时内：黄色，即将到货超时
- 剩余12小时内：红色，即将到货超时
- 小于等于0：已到货超时
        """
    )
