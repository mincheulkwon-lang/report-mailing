import os
import re
import pandas as pd
from datetime import timedelta

# =========================
# 1. 공통 유틸
# =========================
def extract_order(text, idx):
    match = re.search(r'(\d+)', str(text))
    return int(match.group(1)) if match else (idx + 1000)

def to_number(val):
    try:
        txt = str(val).replace(",", "").replace("%", "").strip()
        return float(txt)
    except:
        return 0.0

def format_korean(val):
    val = int(round(to_number(val)))

    if abs(val) >= 100000000:
        return f"{val/100000000:.1f}억"
    elif abs(val) >= 10000:
        return f"{val//10000}만"
    else:
        return f"{val:,}"

def format_won(val):
    return f"{int(round(to_number(val))):,}원"

def clean_text(x):
    return str(x).strip().lower() if pd.notna(x) else ""

def normalize_ctr(val):
    txt = str(val).strip()
    num = to_number(txt)

    if "%" in txt:
        return num / 100
    if num > 1:
        return num / 100
    return num

def normalize_roas(val):
    txt = str(val).strip()
    num = to_number(txt)

    if "%" in txt:
        return num / 100
    if num >= 100:
        return num / 100
    return num

def trend_word(v):
    if v > 0:
        return "증가"
    elif v < 0:
        return "감소"
    else:
        return "유지"

def is_summary_text(x):
    txt = str(x).strip()
    return "총합계" in txt or txt == "총합계" or txt == "총계" or txt == "합계"

def format_value(col, val):
    if pd.isna(val):
        return ""

    col = str(col).strip()

    if "일자" in col or "날짜" in col:
        try:
            return str(pd.to_datetime(val).date())
        except:
            return str(val).split(" ")[0]

    if col == "월":
        try:
            return str(int(float(val)))
        except:
            return val

    try:
        val = float(str(val).replace(",", ""))
    except:
        return val

    if col in ["CTR", "CVR", "도달율"]:
        return f"{val*100:.2f}%"

    if col == "ROAS":
        return f"{val*100:.0f}%"

    if col in ["광고비", "매출"]:
        return f"{int(round(val)):,}"

    if col in ["CPC", "CPM", "CPA"]:
        return f"{int(round(val)):,}"

    if col in ["노출수", "도달수", "클릭수", "전환수"]:
        return f"{int(val):,}"

    return f"{val:.2f}"

def make_mail_subject(file_name):
    return os.path.splitext(os.path.basename(file_name))[0]

# =========================
# 2. 파일 1개 처리 함수
# =========================
def build_html_from_file(file_name):
    df = pd.read_excel(file_name, sheet_name="SUMMARY", header=None)

    # =========================
    # REPORT / COMMENT 위치 찾기
    # =========================
    marker_positions = []

    for i in range(df.shape[0]):
        for j in range(df.shape[1]):
            val = df.iloc[i, j]
            if pd.notna(val):
                cell = str(val).strip().upper()
                if cell.startswith("REPORT") or cell.startswith("COMMENT"):
                    marker_positions.append((i, j))
                    break

    display_tables = []
    comment_source_tables = []

    # =========================
    # 피벗 추출
    # - REPORT  : 메일 본문 표 출력용
    # - COMMENT : 코멘트 계산용
    # =========================
    for idx, (r, c) in enumerate(marker_positions):
        marker_type = str(df.iloc[r, c]).strip().upper()
        order = extract_order(df.iloc[r, c], idx)
        start_col = c + 1

        title = None
        for j in range(start_col, min(start_col + 6, df.shape[1])):
            title_val = df.iloc[r, j]
            if pd.notna(title_val) and str(title_val).strip() != "":
                title = str(title_val).strip()
                break

        if title is None:
            title = f"REPORT_{order}"

        header_row = r + 1

        if marker_type.startswith("COMMENT"):
            for candidate_row in range(r + 1, min(r + 4, df.shape[0])):
                non_empty_count = 0
                for j in range(start_col, min(start_col + 10, df.shape[1])):
                    v = df.iloc[candidate_row, j]
                    if pd.notna(v) and str(v).strip() != "":
                        non_empty_count += 1
                if non_empty_count >= 3:
                    header_row = candidate_row
                    break

        header = []
        end_col = start_col

        for j in range(start_col, df.shape[1]):
            val = df.iloc[header_row, j]
            if pd.isna(val) or str(val).strip() == "":
                break
            header.append(str(val).strip())
            end_col += 1

        if len(header) == 0:
            continue

        data = []

        for i in range(header_row + 1, df.shape[0]):
            next_val = df.iloc[i, c]

            if pd.notna(next_val):
                next_cell = str(next_val).strip().upper()
                if next_cell.startswith("REPORT") or next_cell.startswith("COMMENT"):
                    break

            row = df.iloc[i, start_col:end_col]

            if row.isnull().all():
                if marker_type.startswith("COMMENT"):
                    next_i = i + 1
                    if next_i < df.shape[0]:
                        next_row = df.iloc[next_i, start_col:end_col]
                        if next_row.isnull().all():
                            break
                        else:
                            continue
                else:
                    break

            data.append(row.values)

        if len(data) == 0:
            continue

        table = pd.DataFrame(data, columns=header)

        if table.shape[1] > 0:
            table.iloc[:, 0] = table.iloc[:, 0].ffill()

        if "캠페인별 성과" in str(title) and all(col in table.columns for col in ["프로모션", "매체", "목표"]):
            group_cols = ["프로모션", "매체", "목표"]
            for col in group_cols:
                table[col] = table[col].mask(table[col].eq(table[col].shift()))

        if marker_type.startswith("COMMENT"):
            comment_source_tables.append((order, title, table.copy()))
        else:
            date_col = None
            for col in table.columns:
                col_str = str(col).strip().lower()
                if "일자" in col_str or "날짜" in col_str or "date" in col_str:
                    date_col = col
                    break

            if date_col:
                try:
                    table[date_col] = pd.to_datetime(table[date_col], errors='coerce')
                    table = table.dropna(subset=[date_col])
                    table = table.sort_values(date_col, ascending=False).head(7)
                    table = table.sort_values(date_col)
                except:
                    pass

            display_tables.append((order, title, table))

    # =========================
    # 정렬
    # =========================
    display_tables.sort(key=lambda x: x[0])
    comment_source_tables.sort(key=lambda x: x[0])

    # =========================
    # HTML 생성
    # =========================
    html = ""

    for _, title, table in display_tables:
        html += f"<h3 style='margin:6px 0 4px;font-size:12px;font-weight:bold;'>[{title}]</h3>"

        html += """
        <table style="
            border-collapse:separate;
            border-spacing:0;
            width:auto;
            table-layout:auto;
            font-size:11px;
            font-family:Arial;
            margin-bottom:12px;
            border:1px solid #e5e7eb;
            border-radius:6px;
            overflow:hidden;
        ">
        """

        html += "<tr>"
        for col in table.columns:
            html += f"""
            <th style="
                background:#2f3b52;
                color:white;
                padding:4px 6px;
                border-bottom:1px solid #d1d5db;
                text-align:center;
                font-size:12px;
            ">{col}</th>
            """
        html += "</tr>"

        for _, row in table.iterrows():
            first_val = str(row.iloc[0])

            if "총" in first_val:
                row_bg = "#e8eef7"
                font_weight = "bold"
            else:
                row_bg = "#ffffff"
                font_weight = "normal"

            html += "<tr>"

            for idx_col, (col, val) in enumerate(row.items()):
                align = "left" if idx_col == 0 else "right"
                cell_value = "" if pd.isna(val) else format_value(col, val)

                html += f"""
                <td style="
                    padding:4px 6px;
                    border-bottom:1px solid #eef2f7;
                    text-align:{align};
                    font-size:12px;
                    line-height:1.1;
                    white-space:nowrap;
                    background:{row_bg};
                    font-weight:{font_weight};
                ">
                {cell_value}
                </td>
                """

            html += "</tr>"

        html += "</table>"

    # =========================
    # 코멘트 생성
    # =========================
    report_date = pd.Timestamp.now(tz="Asia/Seoul").normalize().tz_localize(None)
    base_date = report_date - timedelta(days=1)

    comment_map_exact = {}
    comment_map_fallback = {}

    # =========================
    # COMMENT 원본만 사용
    # =========================
    for _, title, table in comment_source_tables:
        title_str = str(title).replace(" ", "").replace("\n", "")
        if "캠페인" not in title_str or "일자별" not in title_str:
            continue

        temp = table.copy()

        date_col = None
        for col in temp.columns:
            col_str = str(col).strip().lower()
            if "기간" in col_str or "일자" in col_str or "날짜" in col_str or "date" in col_str:
                date_col = col
                break

        if date_col is None:
            continue

        required_cols = ["프로모션", "매체", "목표", "광고비", "클릭수", "전환수", "매출", "ROAS", "CTR", "CPC"]
        if not all(col in temp.columns for col in required_cols):
            continue

        for col in ["프로모션", "매체", "목표", "타겟"]:
            if col in temp.columns:
                temp[col] = temp[col].ffill()

        temp[date_col] = pd.to_datetime(temp[date_col], errors="coerce")
        temp = temp.dropna(subset=[date_col]).sort_values(date_col)

        for col in ["노출수", "광고비", "클릭수", "전환수", "매출", "ROAS", "CTR", "CPC"]:
            if col in temp.columns:
                temp[col] = pd.to_numeric(temp[col], errors="coerce")

        # exact key
        if "타겟" in temp.columns:
            grouped_exact = temp.groupby(["프로모션", "매체", "목표", "타겟"], dropna=False)

            for (promo, media, goal, target), g in grouped_exact:
                if is_summary_text(promo):
                    continue

                daily = g.groupby(date_col, as_index=False).agg({
                    "노출수": "sum",
                    "광고비": "sum",
                    "클릭수": "sum",
                    "전환수": "sum",
                    "매출": "sum",
                    "ROAS": "mean",
                    "CTR": "mean",
                    "CPC": "mean"
                }).sort_values(date_col)

                if len(daily) == 0:
                    continue

                exact_key = (
                    clean_text(promo),
                    clean_text(media),
                    clean_text(goal),
                    clean_text(target)
                )

                if report_date.weekday() == 0:
                    friday = base_date - timedelta(days=2)
                    saturday = base_date - timedelta(days=1)
                    sunday = base_date

                    weekend = daily[daily[date_col].isin([friday, saturday, sunday])]

                    if len(weekend) == 0:
                        continue

                    imp_sum = weekend["노출수"].sum() if "노출수" in weekend.columns else 0
                    cost_sum = weekend["광고비"].sum()
                    click_sum = weekend["클릭수"].sum()
                    conv_sum = weekend["전환수"].sum()
                    rev_sum = weekend["매출"].sum()

                    ctr_weekend = (click_sum / imp_sum) if imp_sum else 0
                    cpc_weekend = (cost_sum / click_sum) if click_sum else 0
                    roas_weekend = (rev_sum / cost_sum) if cost_sum else 0

                    if "트래픽" in str(goal):
                        comment_map_exact[exact_key] = (
                            f"- 주말(금~일) 기준 유입 {format_korean(click_sum)}건 / "
                            f"CTR {ctr_weekend * 100:.2f}% / "
                            f"CPC {int(round(cpc_weekend)):,}원"
                        )
                    else:
                        comment_map_exact[exact_key] = (
                            f"- 주말(금~일) 기준 구매 {format_korean(conv_sum)}건 / "
                            f"매출 {format_won(rev_sum)} / "
                            f"ROAS {int(round(roas_weekend * 100))}% 기록"
                        )

                else:
                    prev_date = base_date - timedelta(days=1)

                    curr_df = daily[daily[date_col] == base_date]
                    prev_df = daily[daily[date_col] == prev_date]

                    if len(curr_df) == 0 or len(prev_df) == 0:
                        continue

                    curr = curr_df.iloc[0]
                    prev = prev_df.iloc[0]

                    if "트래픽" in str(goal):
                        click_diff = to_number(curr["클릭수"]) - to_number(prev["클릭수"])
                        ctr_diff = (normalize_ctr(curr["CTR"]) - normalize_ctr(prev["CTR"])) * 100
                        cpc_diff = to_number(curr["CPC"]) - to_number(prev["CPC"])

                        if ctr_diff > 0:
                            ctr_txt = f"CTR {abs(ctr_diff):.2f}%p 증가"
                        elif ctr_diff < 0:
                            ctr_txt = f"CTR {abs(ctr_diff):.2f}%p 감소"
                        else:
                            ctr_txt = "CTR 변동 없음"

                        if cpc_diff > 0:
                            cpc_txt = f"CPC {int(round(abs(cpc_diff))):,}원 증가"
                        elif cpc_diff < 0:
                            cpc_txt = f"CPC {int(round(abs(cpc_diff))):,}원 감소"
                        else:
                            cpc_txt = "CPC 변동 없음"

                        comment_map_exact[exact_key] = (
                            f"- 전일 대비 유입 {format_korean(abs(click_diff))}건 {trend_word(click_diff)} / "
                            f"{ctr_txt} / {cpc_txt}"
                        )
                    else:
                        conv_diff = to_number(curr["전환수"]) - to_number(prev["전환수"])
                        rev_diff = to_number(curr["매출"]) - to_number(prev["매출"])
                        roas_diff = (normalize_roas(curr["ROAS"]) - normalize_roas(prev["ROAS"])) * 100

                        if roas_diff > 0:
                            roas_txt = f"ROAS {abs(roas_diff):.0f}%p 증가"
                        elif roas_diff < 0:
                            roas_txt = f"ROAS {abs(roas_diff):.0f}%p 감소"
                        else:
                            roas_txt = "ROAS 변동 없음"

                        comment_map_exact[exact_key] = (
                            f"- 전일 대비 구매 {format_korean(abs(conv_diff))}건 {trend_word(conv_diff)} / "
                            f"매출 {format_won(abs(rev_diff))} {trend_word(rev_diff)} / {roas_txt}"
                        )

        # fallback key
        grouped_fb = temp.groupby(["프로모션", "매체", "목표", date_col], dropna=False, as_index=False).agg({
            "노출수": "sum",
            "광고비": "sum",
            "클릭수": "sum",
            "전환수": "sum",
            "매출": "sum",
            "ROAS": "mean",
            "CTR": "mean",
            "CPC": "mean"
        })

        grouped_fb2 = grouped_fb.groupby(["프로모션", "매체", "목표"], dropna=False)

        for (promo, media, goal), g in grouped_fb2:
            if is_summary_text(promo):
                continue

            daily = g.sort_values(date_col)

            if len(daily) == 0:
                continue

            fb_key = (
                clean_text(promo),
                clean_text(media),
                clean_text(goal)
            )

            if report_date.weekday() == 0:
                friday = base_date - timedelta(days=2)
                saturday = base_date - timedelta(days=1)
                sunday = base_date

                weekend = daily[daily[date_col].isin([friday, saturday, sunday])]

                if len(weekend) == 0:
                    continue

                imp_sum = weekend["노출수"].sum() if "노출수" in weekend.columns else 0
                cost_sum = weekend["광고비"].sum()
                click_sum = weekend["클릭수"].sum()
                conv_sum = weekend["전환수"].sum()
                rev_sum = weekend["매출"].sum()

                ctr_weekend = (click_sum / imp_sum) if imp_sum else 0
                cpc_weekend = (cost_sum / click_sum) if click_sum else 0
                roas_weekend = (rev_sum / cost_sum) if cost_sum else 0

                if "트래픽" in str(goal):
                    comment_map_fallback[fb_key] = (
                        f"- 주말(금~일) 기준 유입 {format_korean(click_sum)}건 / "
                        f"CTR {ctr_weekend * 100:.2f}% / "
                        f"CPC {int(round(cpc_weekend)):,}원"
                    )
                else:
                    comment_map_fallback[fb_key] = (
                        f"- 주말(금~일) 기준 구매 {format_korean(conv_sum)}건 / "
                        f"매출 {format_won(rev_sum)} / "
                        f"ROAS {int(round(roas_weekend * 100))}% 기록"
                    )

            else:
                prev_date = base_date - timedelta(days=1)

                curr_df = daily[daily[date_col] == base_date]
                prev_df = daily[daily[date_col] == prev_date]

                if len(curr_df) == 0 or len(prev_df) == 0:
                    continue

                curr = curr_df.iloc[0]
                prev = prev_df.iloc[0]

                if "트래픽" in str(goal):
                    click_diff = to_number(curr["클릭수"]) - to_number(prev["클릭수"])
                    ctr_diff = (normalize_ctr(curr["CTR"]) - normalize_ctr(prev["CTR"])) * 100
                    cpc_diff = to_number(curr["CPC"]) - to_number(prev["CPC"])

                    if ctr_diff > 0:
                        ctr_txt = f"CTR {abs(ctr_diff):.2f}%p 증가"
                    elif ctr_diff < 0:
                        ctr_txt = f"CTR {abs(ctr_diff):.2f}%p 감소"
                    else:
                        ctr_txt = "CTR 변동 없음"

                    if cpc_diff > 0:
                        cpc_txt = f"CPC {int(round(abs(cpc_diff))):,}원 증가"
                    elif cpc_diff < 0:
                        cpc_txt = f"CPC {int(round(abs(cpc_diff))):,}원 감소"
                    else:
                        cpc_txt = "CPC 변동 없음"

                    comment_map_fallback[fb_key] = (
                        f"- 전일 대비 유입 {format_korean(abs(click_diff))}건 {trend_word(click_diff)} / "
                        f"{ctr_txt} / {cpc_txt}"
                    )
                else:
                    conv_diff = to_number(curr["전환수"]) - to_number(prev["전환수"])
                    rev_diff = to_number(curr["매출"]) - to_number(prev["매출"])
                    roas_diff = (normalize_roas(curr["ROAS"]) - normalize_roas(prev["ROAS"])) * 100

                    if roas_diff > 0:
                        roas_txt = f"ROAS {abs(roas_diff):.0f}%p 증가"
                    elif roas_diff < 0:
                        roas_txt = f"ROAS {abs(roas_diff):.0f}%p 감소"
                    else:
                        roas_txt = "ROAS 변동 없음"

                    comment_map_fallback[fb_key] = (
                        f"- 전일 대비 구매 {format_korean(abs(conv_diff))}건 {trend_word(conv_diff)} / "
                        f"매출 {format_won(abs(rev_diff))} {trend_word(rev_diff)} / {roas_txt}"
                    )

    # =========================
    # 코멘트 출력 (최종 안정화)
    # =========================
    for _, title, table in display_tables:

        if "TOTAL" in str(title).upper():
            total_rows = table[table.iloc[:, 0].astype(str).str.contains("총", na=False)]

            if len(total_rows) > 0:
                total_row = total_rows.iloc[-1]

                html += f"""
                <div style="font-size:12px; line-height:1.45; margin-bottom:12px;">
                <b>■ 전체 성과</b><br>
                - 누적 광고비 {format_won(total_row['광고비'])} 소진 / 노출 {format_korean(total_row['노출수'])}건, 유입 {format_korean(total_row['클릭수'])}건 / 구매 {format_korean(total_row['전환수'])}건, 매출 {format_won(total_row['매출'])} / ROAS {int(round(normalize_roas(total_row['ROAS']) * 100))}% 형성
                </div>
                """

        if "캠페인별 성과" in str(title):
            df = table.copy()

            for col in ["프로모션", "매체", "목표", "타겟"]:
                if col in df.columns:
                    df[col] = df[col].ffill()

            for col in ["노출수", "클릭수", "전환수", "광고비", "매출", "CTR", "CPC", "ROAS"]:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors="coerce")

            df = df[~df["프로모션"].astype(str).apply(is_summary_text)]

            promo_group = df.groupby("프로모션", sort=False)

            for promo, promo_df in promo_group:
                promo = "" if pd.isna(promo) else str(promo).strip()

                html += f"""
                <div style="font-size:12px; line-height:1.4; margin-bottom:8px;">
                <b>■ {promo}</b><br>
                """

                sub_group = promo_df.groupby(["매체", "목표", "타겟"], dropna=False, sort=False)

                for (media, goal, target), g in sub_group:
                    media = "" if pd.isna(media) else str(media).strip()
                    goal = "" if pd.isna(goal) else str(goal).strip()
                    target = "" if pd.isna(target) else str(target).strip()

                    row = g.sum(numeric_only=True)

                    imp = to_number(row.get("노출수", 0))
                    reach = to_number(row.get("도달수", 0))
                    click = to_number(row.get("클릭수", 0))
                    cost = to_number(row.get("광고비", 0))
                    conv = to_number(row.get("전환수", 0))
                    revenue = to_number(row.get("매출", 0))

                    ctr_calc = click / imp if imp else 0
                    cpc_calc = cost / click if click else 0
                    roas_calc = revenue / cost if cost else 0
                    cvr_calc = conv / click if click else 0

                    html += f"&nbsp;&nbsp;ㄴ <b>{media} / {goal}{' / ' + target if target else ''}</b><br>"

                    if "트래픽" in goal:
                        html += (
                            f"&nbsp;&nbsp;&nbsp;&nbsp;- 누적 광고비 {format_won(row['광고비'])} 소진 / "
                            f"노출 {format_korean(row['노출수'])}건, 유입 {format_korean(row['클릭수'])}건 / "
                            f"CTR {ctr_calc * 100:.2f}% / CPC {int(round(cpc_calc)):,}원<br>"
                        )
                    else:
                        html += (
                            f"&nbsp;&nbsp;&nbsp;&nbsp;- 누적 광고비 {format_won(row['광고비'])} 소진 / "
                            f"노출 {format_korean(row['노출수'])}건, 유입 {format_korean(row['클릭수'])}건 / "
                            f"구매 {format_korean(row['전환수'])}건, 매출 {format_won(row['매출'])} / "
                            f"ROAS {int(round(roas_calc * 100))}% 형성<br>"
                        )

                    exact_key = (
                        clean_text(promo),
                        clean_text(media),
                        clean_text(goal),
                        clean_text(target)
                    )

                    fb_key = (
                        clean_text(promo),
                        clean_text(media),
                        clean_text(goal)
                    )

                    if exact_key in comment_map_exact:
                        html += f"&nbsp;&nbsp;&nbsp;&nbsp;{comment_map_exact[exact_key]}<br>"
                    elif fb_key in comment_map_fallback:
                        html += f"&nbsp;&nbsp;&nbsp;&nbsp;{comment_map_fallback[fb_key]}<br>"

                    html += "&nbsp;<br>"

                html += "</div>"

    return html, display_tables, comment_source_tables
