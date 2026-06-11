import os
import json
import datetime
import requests
from google.oauth2.service_account import Credentials
import gspread

# ── 設定區（從環境變數讀取，不要直接寫在這裡）──────────────────────────
META_ACCESS_TOKEN = os.environ["META_ACCESS_TOKEN"]
META_AD_ACCOUNT_ID = os.environ["META_AD_ACCOUNT_ID"]   # 格式：act_123456789
GOOGLE_CREDS_JSON = os.environ["GOOGLE_CREDS_JSON"]      # Service Account JSON 字串
SPREADSHEET_ID = os.environ["SPREADSHEET_ID"]            # Google Sheets ID
SENDGRID_API_KEY = os.environ["SENDGRID_API_KEY"]
SENDER_EMAIL = os.environ.get("SENDER_EMAIL", "vivien.lin@beanne.com.tw")
REPORT_RECIPIENTS = os.environ.get("REPORT_RECIPIENTS", SENDER_EMAIL)  # 逗號分隔

# ── 日期設定 ─────────────────────────────────────────────────────────────
yesterday = (datetime.date.today() - datetime.timedelta(days=1)).isoformat()


# ── 1. 從 Meta Marketing API 拉資料 ─────────────────────────────────────
def learning_label(days: int) -> str:
    if days <= 7:
        return "🔵 學習中"
    elif days <= 14:
        return "🟡 剛出學習"
    else:
        return "🟢 已穩定"


def fetch_meta_data():
    # 先拿廣告清單（含上檔時間）
    url = f"https://graph.facebook.com/v19.0/{META_AD_ACCOUNT_ID}/ads"
    params = {
        "fields": "id,name,created_time",
        "access_token": META_ACCESS_TOKEN,
        "limit": 100,
    }
    res = requests.get(url, params=params)
    if not res.ok:
        print(f"API error: {res.status_code} {res.text}")
    res.raise_for_status()
    ads = res.json().get("data", [])

    # 建立 ad_id → 上檔日期 的對照表
    today = datetime.date.today()
    ad_start = {}
    for ad in ads:
        raw_time = ad.get("created_time", "")
        if raw_time:
            start_date = datetime.date.fromisoformat(raw_time[:10])
            ad_start[ad["id"]] = start_date

    # 再用 insights endpoint 拿成效（以 ad 為單位）
    ins_url = f"https://graph.facebook.com/v19.0/{META_AD_ACCOUNT_ID}/insights"
    ins_params = {
        "fields": "ad_id,ad_name,spend,impressions,clicks,ctr,cpm,purchase_roas",
        "level": "ad",
        "date_preset": "yesterday",
        "access_token": META_ACCESS_TOKEN,
        "limit": 200,
    }
    ins_res = requests.get(ins_url, params=ins_params)
    ins_res.raise_for_status()
    insights = ins_res.json().get("data", [])

    rows = []
    for ins in insights:
        roas_list = ins.get("purchase_roas", [])
        roas = float(roas_list[0]["value"]) if roas_list else 0.0
        ad_id = ins.get("ad_id", "")
        start_date = ad_start.get(ad_id)
        days_running = (today - start_date).days if start_date else None
        rows.append({
            "date": yesterday,
            "ad_id": ad_id,
            "name": ins.get("ad_name", ""),
            "thumbnail": "",
            "spend": float(ins.get("spend", 0)),
            "impressions": int(ins.get("impressions", 0)),
            "clicks": int(ins.get("clicks", 0)),
            "ctr": round(float(ins.get("ctr", 0)), 2),
            "cpm": round(float(ins.get("cpm", 0)), 2),
            "roas": round(roas, 2),
            "start_date": start_date.isoformat() if start_date else "",
            "days_running": days_running if days_running is not None else "",
            "learning_status": learning_label(days_running) if days_running is not None else "",
        })

    # 依版面拆分（Facebook / Instagram / Threads）
    placements = fetch_placement_breakdown()
    return rows, placements


def fetch_placement_breakdown():
    url = f"https://graph.facebook.com/v19.0/{META_AD_ACCOUNT_ID}/insights"
    params = {
        "fields": "ad_name,spend,impressions,clicks,ctr,cpm,purchase_roas",
        "breakdowns": "publisher_platform,platform_position",
        "date_preset": "yesterday",
        "access_token": META_ACCESS_TOKEN,
        "limit": 200,
    }
    res = requests.get(url, params=params)
    res.raise_for_status()
    raw = res.json().get("data", [])

    result = []
    platform_map = {
        "facebook": "fb",
        "instagram": "ig",
        "threads": "th",
    }
    for row in raw:
        plat_raw = row.get("publisher_platform", "").lower()
        plat = platform_map.get(plat_raw, plat_raw)
        roas_list = row.get("purchase_roas", [])
        roas = float(roas_list[0]["value"]) if roas_list else 0.0
        result.append({
            "date": yesterday,
            "platform": plat,
            "position": row.get("platform_position", ""),
            "ad_name": row.get("ad_name", ""),
            "spend": float(row.get("spend", 0)),
            "impressions": int(row.get("impressions", 0)),
            "clicks": int(row.get("clicks", 0)),
            "ctr": round(float(row.get("ctr", 0)), 2),
            "cpm": round(float(row.get("cpm", 0)), 2),
            "roas": round(roas, 2),
        })
    return result


# ── 2. 寫入 Google Sheets ────────────────────────────────────────────────
def write_to_sheets(rows, placements):
    creds_dict = json.loads(GOOGLE_CREDS_JSON)
    creds = Credentials.from_service_account_info(
        creds_dict,
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SPREADSHEET_ID)

    # 工作表：每日素材成效
    daily_headers = ["日期","素材名稱","花費","曝光","點擊","CTR(%)","CPM","ROAS","上檔日期","已投天數","學習狀態"]
    try:
        ws_daily = sh.worksheet("每日素材成效")
        # 若是舊版標題（缺少新欄位），直接更新第一列
        existing_headers = ws_daily.row_values(1)
        if existing_headers != daily_headers:
            ws_daily.update("A1", [daily_headers])
    except gspread.WorksheetNotFound:
        ws_daily = sh.add_worksheet("每日素材成效", rows=5000, cols=11)
        ws_daily.append_row(daily_headers)

    for r in rows:
        ws_daily.append_row([
            r["date"], r["name"],
            r["spend"], r["impressions"], r["clicks"],
            r["ctr"], r["cpm"], r["roas"],
            r["start_date"], r["days_running"], r["learning_status"],
        ])

    # 工作表：版面拆分
    plat_headers = ["日期","版面","位置","素材名稱","花費","曝光","點擊","CTR(%)","CPM","ROAS"]
    try:
        ws_plat = sh.worksheet("版面拆分")
        existing_plat_headers = ws_plat.row_values(1)
        if existing_plat_headers != plat_headers:
            ws_plat.update("A1", [plat_headers])
    except gspread.WorksheetNotFound:
        ws_plat = sh.add_worksheet("版面拆分", rows=5000, cols=10)
        ws_plat.append_row(plat_headers)

    for r in placements:
        ws_plat.append_row([
            r["date"], r["platform"], r["position"], r["ad_name"],
            r["spend"], r["impressions"], r["clicks"],
            r["ctr"], r["cpm"], r["roas"],
        ])

    # 工作表：彙總（放在最前面，每次全部重寫）
    try:
        ws_sum = sh.worksheet("📊 彙總")
    except gspread.WorksheetNotFound:
        ws_sum = sh.add_worksheet("📊 彙總", rows=200, cols=10)
        sh.reorder_worksheets([ws_sum, ws_daily, ws_plat])

    # 從所有歷史資料彙總（讀取整個每日素材成效）
    all_rows = ws_daily.get_all_records()

    # 依素材名稱彙總
    today = datetime.date.today()
    summary = {}
    for r in all_rows:
        name = r.get("素材名稱", "")
        if not name:
            continue
        if name not in summary:
            summary[name] = {"花費": 0, "曝光": 0, "點擊": 0, "roas_sum": 0, "count": 0, "start_date": ""}
        summary[name]["花費"] += float(r.get("花費", 0))
        summary[name]["曝光"] += int(r.get("曝光", 0))
        summary[name]["點擊"] += int(r.get("點擊", 0))
        summary[name]["roas_sum"] += float(r.get("ROAS", 0))
        summary[name]["count"] += 1
        # 取最早的上檔日期
        sd = r.get("上檔日期", "")
        if sd and (not summary[name]["start_date"] or sd < summary[name]["start_date"]):
            summary[name]["start_date"] = sd

    # 版面彙總
    all_plat = ws_plat.get_all_records()
    plat_summary = {}
    for r in all_plat:
        plat = r.get("版面", "")
        if not plat:
            continue
        if plat not in plat_summary:
            plat_summary[plat] = {"花費": 0, "點擊": 0, "ctr_sum": 0, "roas_sum": 0, "count": 0}
        plat_summary[plat]["花費"] += float(r.get("花費", 0))
        plat_summary[plat]["點擊"] += int(r.get("點擊", 0))
        plat_summary[plat]["ctr_sum"] += float(r.get("CTR(%)", 0))
        plat_summary[plat]["roas_sum"] += float(r.get("ROAS", 0))
        plat_summary[plat]["count"] += 1

    # 整理彙總工作表內容
    summary_data = [
        [f"PSK Meta 廣告彙總報表", "", "", "", "", "", "", ""],
        [f"最後更新：{yesterday}", "", "", "", "", "", "", ""],
        ["", "", "", "", "", "", "", ""],
        ["【各版面彙總】", "", "", "", "", "", "", ""],
        ["版面", "總花費 (NT$)", "總點擊", "平均 CTR (%)", "平均 ROAS", "", "", ""],
    ]
    for plat, v in sorted(plat_summary.items(), key=lambda x: -x[1]["花費"]):
        avg_ctr = round(v["ctr_sum"] / v["count"], 2) if v["count"] else 0
        avg_roas = round(v["roas_sum"] / v["count"], 2) if v["count"] else 0
        summary_data.append([plat, round(v["花費"]), v["點擊"], avg_ctr, avg_roas, "", "", ""])

    summary_data += [
        ["", "", "", "", "", "", "", "", ""],
        ["【素材彙總】", "", "", "", "", "", "", "", ""],
        ["素材名稱", "上檔日期", "已投天數", "總花費 (NT$)", "總曝光", "總點擊", "平均 ROAS", "學習狀態", "建議"],
    ]
    for name, v in sorted(summary.items(), key=lambda x: -(x[1]["roas_sum"] / x[1]["count"] if x[1]["count"] else 0)):
        avg_roas = round(v["roas_sum"] / v["count"], 2) if v["count"] else 0
        suggest = "擴大投放" if avg_roas >= 4 else ("優化測試" if avg_roas >= 2 else "停止/調整")
        start_date = v["start_date"]
        if start_date:
            days = (today - datetime.date.fromisoformat(start_date)).days
            status = learning_label(days)
        else:
            days, status = "", ""
        summary_data.append([name, start_date, days, round(v["花費"]), v["曝光"], v["點擊"], avg_roas, status, suggest])

    ws_sum.clear()
    ws_sum.update("A1", summary_data)

    # 確保彙總在第一個位置
    all_ws = sh.worksheets()
    ws_order = [ws_sum] + [w for w in all_ws if w.id != ws_sum.id]
    sh.reorder_worksheets(ws_order)

    print(f"✓ Google Sheets 已更新：{len(rows)} 筆素材，{len(placements)} 筆版面")


# ── 3. 更新 data/report.json（給網站用）──────────────────────────────────
def write_json(rows, placements):
    os.makedirs("data", exist_ok=True)

    # 讀取既有歷史（最多保留 30 天）
    json_path = "data/report.json"
    if os.path.exists(json_path):
        with open(json_path) as f:
            existing = json.load(f)
    else:
        existing = {"creatives": [], "placements": [], "updated_at": ""}

    # 移除今天同日期的舊資料（避免重跑時重複）
    existing["creatives"] = [r for r in existing["creatives"] if r["date"] != yesterday]
    existing["placements"] = [r for r in existing["placements"] if r["date"] != yesterday]

    existing["creatives"].extend(rows)
    existing["placements"].extend(placements)

    # 只保留最近 30 天
    cutoff = (datetime.date.today() - datetime.timedelta(days=30)).isoformat()
    existing["creatives"] = [r for r in existing["creatives"] if r["date"] >= cutoff]
    existing["placements"] = [r for r in existing["placements"] if r["date"] >= cutoff]
    existing["updated_at"] = datetime.datetime.now().isoformat()

    with open(json_path, "w") as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)

    print(f"✓ data/report.json 已更新")


# ── 4. 發送 Email 日報 ───────────────────────────────────────────────────
def send_email(rows, placements):
    total_spend = sum(r["spend"] for r in rows)
    avg_roas = (
        sum(r["roas"] * r["spend"] for r in rows) / total_spend
        if total_spend else 0
    )
    avg_ctr = sum(r["ctr"] for r in rows) / len(rows) if rows else 0

    # 版面小結
    plat_summary = {}
    for p in placements:
        k = p["platform"]
        if k not in plat_summary:
            plat_summary[k] = {"spend": 0, "roas_sum": 0, "ctr_sum": 0, "count": 0}
        plat_summary[k]["spend"] += p["spend"]
        plat_summary[k]["roas_sum"] += p["roas"]
        plat_summary[k]["ctr_sum"] += p["ctr"]
        plat_summary[k]["count"] += 1

    plat_rows_html = ""
    plat_labels = {"fb": "Facebook", "ig": "Instagram", "th": "Threads"}
    plat_colors = {"fb": "#E6F1FB", "ig": "#FBEAF0", "th": "#F1EFE8"}
    for k, v in plat_summary.items():
        avg_r = v["roas_sum"] / v["count"]
        avg_c = v["ctr_sum"] / v["count"]
        color = "#27500A" if avg_r >= 4 else ("#633806" if avg_r >= 2.5 else "#791F1F")
        plat_rows_html += f"""
        <tr>
          <td style="padding:10px 14px;border-bottom:1px solid #eee;">
            <span style="background:{plat_colors.get(k,'#eee')};padding:2px 8px;border-radius:6px;font-size:12px;">{plat_labels.get(k, k)}</span>
          </td>
          <td style="padding:10px 14px;border-bottom:1px solid #eee;">NT${v['spend']:,.0f}</td>
          <td style="padding:10px 14px;border-bottom:1px solid #eee;">{avg_c:.2f}%</td>
          <td style="padding:10px 14px;border-bottom:1px solid #eee;font-weight:600;color:{color};">{avg_r:.2f}</td>
        </tr>"""

    # 前5名素材
    top5 = sorted(rows, key=lambda x: x["roas"], reverse=True)[:5]
    top5_html = ""
    for r in top5:
        color = "#27500A" if r["roas"] >= 4 else ("#633806" if r["roas"] >= 2.5 else "#791F1F")
        days_str = f"{r['days_running']} 天" if r['days_running'] != "" else "-"
        top5_html += f"""
        <tr>
          <td style="padding:8px 14px;border-bottom:1px solid #eee;font-size:13px;">{r['name']}</td>
          <td style="padding:8px 14px;border-bottom:1px solid #eee;font-size:13px;">{r['learning_status'] or '-'}<br><span style="color:#aaa;font-size:11px;">{days_str}</span></td>
          <td style="padding:8px 14px;border-bottom:1px solid #eee;font-size:13px;">NT${r['spend']:,.0f}</td>
          <td style="padding:8px 14px;border-bottom:1px solid #eee;font-size:13px;">{r['ctr']:.2f}%</td>
          <td style="padding:8px 14px;border-bottom:1px solid #eee;font-size:13px;font-weight:600;color:{color};">{r['roas']:.2f}</td>
        </tr>"""

    roas_color = "#27500A" if avg_roas >= 3 else "#791F1F"

    html_body = f"""
<!DOCTYPE html>
<html>
<body style="margin:0;padding:0;background:#f5f5f3;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0">
<tr><td align="center" style="padding:32px 16px;">
<table width="600" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:12px;overflow:hidden;">

  <tr><td style="padding:28px 32px 20px;border-bottom:1px solid #eee;">
    <p style="margin:0;font-size:11px;color:#888;letter-spacing:.08em;text-transform:uppercase;">PSK · Meta 廣告日報</p>
    <h1 style="margin:6px 0 0;font-size:22px;font-weight:500;">{yesterday} 成效摘要</h1>
  </td></tr>

  <tr><td style="padding:20px 32px;">
    <table width="100%" cellpadding="0" cellspacing="0">
      <tr>
        <td width="33%" style="text-align:center;padding:12px;">
          <p style="margin:0;font-size:11px;color:#888;">總花費</p>
          <p style="margin:4px 0 0;font-size:26px;font-weight:500;">NT${total_spend:,.0f}</p>
        </td>
        <td width="33%" style="text-align:center;padding:12px;border-left:1px solid #eee;border-right:1px solid #eee;">
          <p style="margin:0;font-size:11px;color:#888;">加權 ROAS</p>
          <p style="margin:4px 0 0;font-size:26px;font-weight:500;color:{roas_color};">{avg_roas:.2f}</p>
        </td>
        <td width="33%" style="text-align:center;padding:12px;">
          <p style="margin:0;font-size:11px;color:#888;">平均 CTR</p>
          <p style="margin:4px 0 0;font-size:26px;font-weight:500;">{avg_ctr:.2f}%</p>
        </td>
      </tr>
    </table>
  </td></tr>

  <tr><td style="padding:0 32px 20px;">
    <p style="font-size:12px;font-weight:500;color:#555;margin-bottom:8px;">各版面成效</p>
    <table width="100%" cellpadding="0" cellspacing="0" style="font-size:13px;">
      <tr style="background:#f9f9f7;">
        <th style="padding:8px 14px;text-align:left;font-weight:500;color:#888;font-size:11px;">版面</th>
        <th style="padding:8px 14px;text-align:left;font-weight:500;color:#888;font-size:11px;">花費</th>
        <th style="padding:8px 14px;text-align:left;font-weight:500;color:#888;font-size:11px;">CTR</th>
        <th style="padding:8px 14px;text-align:left;font-weight:500;color:#888;font-size:11px;">ROAS</th>
      </tr>
      {plat_rows_html}
    </table>
  </td></tr>

  <tr><td style="padding:0 32px 24px;">
    <p style="font-size:12px;font-weight:500;color:#555;margin-bottom:8px;">ROAS 前 5 名素材</p>
    <table width="100%" cellpadding="0" cellspacing="0" style="font-size:13px;">
      <tr style="background:#f9f9f7;">
        <th style="padding:8px 14px;text-align:left;font-weight:500;color:#888;font-size:11px;">素材名稱</th>
        <th style="padding:8px 14px;text-align:left;font-weight:500;color:#888;font-size:11px;">學習狀態</th>
        <th style="padding:8px 14px;text-align:left;font-weight:500;color:#888;font-size:11px;">花費</th>
        <th style="padding:8px 14px;text-align:left;font-weight:500;color:#888;font-size:11px;">CTR</th>
        <th style="padding:8px 14px;text-align:left;font-weight:500;color:#888;font-size:11px;">ROAS</th>
      </tr>
      {top5_html}
    </table>
  </td></tr>

  <tr><td style="padding:16px 32px 28px;border-top:1px solid #eee;text-align:center;">
    <a href="https://vivienlin-git.github.io/psk-ads-report/" style="display:inline-block;padding:10px 24px;background:#000;color:#fff;border-radius:8px;font-size:13px;text-decoration:none;">查看完整儀表板 →</a>
  </td></tr>

  <tr><td style="padding:12px 32px;background:#f9f9f7;text-align:center;">
    <p style="margin:0;font-size:11px;color:#aaa;">PSK · 自動產生 · {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}</p>
  </td></tr>

</table>
</td></tr>
</table>
</body>
</html>"""

    recipients = [{"email": r.strip()} for r in REPORT_RECIPIENTS.split(",")]
    payload = {
        "personalizations": [{"to": recipients}],
        "from": {"email": SENDER_EMAIL, "name": "PSK 廣告報表"},
        "subject": f"📊 {yesterday} Meta 廣告日報 · ROAS {avg_roas:.2f} · 花費 NT${total_spend:,.0f}",
        "content": [{"type": "text/html", "value": html_body}],
    }
    res = requests.post(
        "https://api.sendgrid.com/v3/mail/send",
        headers={"Authorization": f"Bearer {SENDGRID_API_KEY}", "Content-Type": "application/json"},
        json=payload,
    )
    res.raise_for_status()
    print(f"✓ Email 已發送至 {REPORT_RECIPIENTS}")


# ── 主程式 ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print(f"▶ 開始拉取 {yesterday} 的資料...")
    rows, placements = fetch_meta_data()
    print(f"  拿到 {len(rows)} 筆素材，{len(placements)} 筆版面資料")

    write_to_sheets(rows, placements)
    write_json(rows, placements)
    send_email(rows, placements)

    print("✓ 完成")
