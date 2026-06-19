# -*- coding: utf-8 -*-
"""
부가가치세 안내문 생성기
================================================================
'근처횟집' 양식(A4 1page HTML)을 그대로 사용해, 입력 엑셀의 업체별 자료로
업체마다 안내문 HTML 을 자동 생성한다.

사용법
  py report_gen.py --init     # 입력양식(안내문_입력.xlsx) 샘플 생성
  py report_gen.py            # 안내문_입력.xlsx 의 모든 업체 → 안내문_출력/*.html
"""
import argparse
import math
import os
import re
import sys
from pathlib import Path

import pandas as pd

try:
    sys.stdout.reconfigure(encoding="utf-8")
except (AttributeError, ValueError):
    pass

HERE = Path(__file__).resolve().parent
INPUT_XLSX = HERE / "안내문_입력.xlsx"
OUT_DIR = HERE / "안내문_출력"


# ============================================================
# HTML 템플릿 (근처횟집 양식 / CSS 동일, 데이터부만 {{토큰}} 치환)
# ============================================================
TEMPLATE = r"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>부가가치세 안내문 · {{BIZ_NAME}}</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;500;700&display=swap" rel="stylesheet">
<style>
:root{
  --bg:#FFFFFF; --bg-soft:#F5F5F5; --bg-line:#ECECEC;
  --final-bg:#595959; --tx-strong:#1A1A1A; --tx-body:#5C5C5C;
  --tx-muted:#8A8A8A; --tx-faint:#B5B5B5; --warn:#C53939; --line:#ECECEC;
  --font:'Noto Sans KR','Pretendard',-apple-system,sans-serif;
  --mark:url("data:image/svg+xml,%3Csvg%20xmlns='http://www.w3.org/2000/svg'%20viewBox='0%200%2060%2068'%3E%3Cpolygon%20points='13,4%2030,12%2047,4%2047,14%2030,22%2013,14'%20fill='%2329ABE2'/%3E%3Cpolygon%20points='4,19%2030,27%2056,19%2056,29%2030,37%204,29'%20fill='%232BB6A8'/%3E%3Cpolygon%20points='4,34%2030,42%2056,34%2056,44%2030,52%204,44'%20fill='%230E3D6B'/%3E%3Cpolygon%20points='13,49%2030,57%2047,49%2047,59%2030,67%2013,59'%20fill='%231B75BC'/%3E%3C/svg%3E");
}
*{box-sizing:border-box;}
body{margin:0;background:#E9E9E9;font-family:var(--font);-webkit-font-smoothing:antialiased;}
.page{width:595px;height:842px;padding:40px 46px 32px;position:relative;background:var(--bg);
      margin:24px auto;box-shadow:0 2px 18px rgba(0,0,0,.10);display:flex;flex-direction:column;}
.hd{display:flex;align-items:center;justify-content:space-between;
    border-bottom:2px solid var(--tx-strong);padding-bottom:10px;margin-bottom:12px;}
.hd-l{display:flex;align-items:center;gap:11px;}
.hd-logo{width:31px;height:35px;background:var(--mark) center/contain no-repeat;}
.hd-eyebrow{font-size:7px;letter-spacing:.22em;color:var(--tx-faint);}
.hd-title{font-size:18px;font-weight:700;color:var(--tx-strong);margin:3px 0 0;letter-spacing:-.01em;}
.hd-biz{background:var(--final-bg);border-radius:8px;padding:9px 16px;text-align:right;min-width:140px;}
.hd-biz .lbl{font-size:7px;letter-spacing:.28em;color:rgba(255,255,255,.55);}
.hd-biz .nm{font-size:18px;font-weight:700;color:#fff;letter-spacing:-.01em;line-height:1.2;margin-top:3px;}
.hd-biz .dt{font-size:7px;color:rgba(255,255,255,.6);margin-top:5px;letter-spacing:.04em;}
.sec-h{display:flex;align-items:baseline;gap:7px;margin:10px 0 5px;}
.sec-h .n{font-size:7.5px;letter-spacing:.12em;color:var(--tx-faint);font-weight:700;}
.sec-h .t{font-size:11px;font-weight:700;color:var(--tx-strong);}
.cols{display:grid;grid-template-columns:1fr 1fr;gap:12px;}
.cardcols{display:grid;gap:12px;}
.tbl{width:100%;border-collapse:collapse;}
.tbl caption{font-size:7.5px;letter-spacing:.12em;color:var(--tx-muted);text-align:left;
             padding:0 0 3px;font-weight:700;}
.tbl th{font-size:7px;letter-spacing:.08em;color:var(--tx-muted);font-weight:400;
        text-align:left;padding:4px 7px;background:var(--bg-line);}
.tbl th.r{text-align:right;}
.tbl td{font-size:8.5px;color:var(--tx-body);padding:3.5px 7px;border-bottom:1px solid var(--line);}
.tbl td.r{text-align:right;}
.tbl td.num{font-weight:700;color:var(--tx-strong);}
.tbl td.fill{color:var(--tx-faint);text-align:right;}
.tbl td.card{font-weight:500;color:var(--tx-strong);letter-spacing:.04em;font-variant-numeric:tabular-nums;}
.tbl td.rmk{font-size:7px;color:var(--tx-faint);}
.tbl tr.sum td{background:var(--bg-soft);font-weight:700;}
.tbl td.lbl2{white-space:nowrap;font-size:7px;letter-spacing:-.02em;}
.final{background:var(--final-bg);border-radius:7px;padding:10px 16px;margin-top:8px;
       display:flex;justify-content:space-between;align-items:center;}
.final-en{font-size:7px;letter-spacing:.2em;color:rgba(255,255,255,.55);}
.final-ko{font-size:8.5px;color:rgba(255,255,255,.82);margin-top:3px;}
.final-v{font-size:17px;font-weight:700;color:#fff;}
.final-v span{font-size:9.5px;font-weight:400;color:rgba(255,255,255,.65);margin-left:3px;}
.chips{display:grid;grid-template-columns:repeat(4,1fr);gap:6px;}
.chip{border:1px solid var(--line);border-radius:6px;padding:6px 4px;text-align:center;}
.chip .c-ko{font-size:8.5px;font-weight:700;color:var(--tx-strong);}
.chip .c-en{font-size:6.5px;letter-spacing:.1em;color:var(--tx-faint);margin-top:2px;}
.note{background:var(--bg-soft);border-radius:7px;padding:8px 12px;margin-top:7px;}
.note p{font-size:8px;color:var(--tx-body);line-height:1.55;margin:0 0 3px;}
.note p:last-child{margin-bottom:0;}
.note p b{color:var(--tx-strong);font-weight:700;}
.warn-box{margin:7px 0 0;background:#FBF1F1;border:1px solid #EAD0D0;border-left:3px solid var(--warn);
          border-radius:6px;padding:7px 11px;}
.warn-box .wt{font-size:7px;letter-spacing:.14em;color:var(--warn);font-weight:700;margin-bottom:3px;}
.warn-box p{font-size:7.5px;color:var(--tx-body);line-height:1.55;margin:0;}
.warn-box p b{color:var(--warn);font-weight:700;}
.foot{margin-top:auto;display:flex;align-items:center;justify-content:space-between;
      font-size:7px;color:var(--tx-faint);border-top:1px solid var(--line);padding-top:9px;}
.foot-l{display:inline-flex;align-items:center;gap:7px;}
.foot-mark{display:inline-block;width:11px;height:13px;background:var(--mark) center/contain no-repeat;}
.foot b{color:var(--warn);font-weight:700;}
@page{ size:A4; margin:0; }
@media print{
  :root{ --tx-body:#2B2B2B; --tx-muted:#565656; --tx-faint:#7E7E7E; --line:#D8D8D8; }
  html,body{background:#fff;-webkit-print-color-adjust:exact;print-color-adjust:exact;}
  .page{width:595px;height:842px;zoom:1.3333;margin:0 auto;box-shadow:none;overflow:hidden;
        -webkit-print-color-adjust:exact;print-color-adjust:exact;}
}
</style>
</head>
<body>
<div class="page">
  <div class="hd">
    <div class="hd-l">
      <div class="hd-logo"></div>
      <div>
        <div class="hd-eyebrow">{{EYEBROW}}</div>
        <h1 class="hd-title">부가가치세 안내문</h1>
      </div>
    </div>
    <div class="hd-biz">
      <div class="nm">{{BIZ_NAME}}</div>
      <div class="dt">작성일자 {{WRITE_DATE}}</div>
    </div>
  </div>

  <div class="sec-h"><span class="n">01</span><span class="t">매입매출 분석</span></div>
  <div class="cols">
    <table class="tbl">
      <caption>매출</caption>
      <thead><tr><th>구분</th><th class="r">공급가액</th><th class="r">세액</th></tr></thead>
      <tbody>{{SALES_ROWS}}</tbody>
    </table>
    <table class="tbl">
      <caption>매입</caption>
      <thead><tr><th>구분</th><th class="r">공급가액</th><th class="r">세액</th></tr></thead>
      <tbody>{{PURCHASE_ROWS}}</tbody>
    </table>
  </div>
  <div class="warn-box">
    <div class="wt">{{WARN_TITLE}}</div>
    <p>{{WARN_NOTE}}</p>
  </div>

  <div class="cols" style="margin-top:4px;">
    <div>
      <div class="sec-h"><span class="n">02</span><span class="t">세액공제 안내</span></div>
      <table class="tbl">
        <thead><tr><th>공제 항목</th><th class="r">공제금액</th></tr></thead>
        <tbody>{{CREDIT_ROWS}}</tbody>
      </table>
    </div>
    <div>
      <div class="sec-h"><span class="n">03</span><span class="t">예상 납부(환급)세액</span></div>
      <table class="tbl">
        <thead><tr><th>구분</th><th class="r">금액</th></tr></thead>
        <tbody>{{ESTIMATE_ROWS}}</tbody>
      </table>
    </div>
  </div>

  <div class="final">
    <div><div class="final-en">ESTIMATED PAYABLE · 예상 납부세액</div><div class="final-ko">{{FINAL_DESC}}</div></div>
    <div class="final-v">{{FINAL_VALUE}}<span>원</span></div>
  </div>

  <div class="sec-h"><span class="n">04</span><span class="t">사업용카드 목록</span></div>
  {{CARDS_BLOCK}}

  <div class="sec-h"><span class="n">＊</span><span class="t">안내메모</span></div>
  <div class="note">{{MEMO_ROWS}}</div>

  <div class="foot">
    <span class="foot-l"><span class="foot-mark"></span>{{FOOTER}}</span>
    <span>무단배포 금지</span>
  </div>
</div>
</body>
</html>
"""


# ============================================================
# 유틸
# ============================================================
def to_num(v) -> float:
    if v is None or (isinstance(v, float) and pd.isna(v)) or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r"[^\d.\-]", "", str(v))
    return float(s) if s else 0.0


def won(v) -> str:
    return f"{int(round(to_num(v))):,}"


def row3(label, supply, tax, is_sum=False):
    cls = ' class="sum"' if is_sum else ""
    if not is_sum and to_num(supply) == 0 and to_num(tax) == 0:
        return f'<tr><td>{label}</td><td class="r fill">-</td><td class="r fill">-</td></tr>'
    return (f'<tr{cls}><td>{label}</td>'
            f'<td class="r num">{won(supply)}</td>'
            f'<td class="r num">{won(tax)}</td></tr>')


def row3c(label, supply, tax, label_cls=""):
    """라벨 칸에 클래스를 줄 수 있는 매출/매입 행 (긴 라벨 한 줄 처리용)."""
    lc = f' class="{label_cls}"' if label_cls else ""
    if to_num(supply) == 0 and to_num(tax) == 0:
        return f'<tr><td{lc}>{label}</td><td class="r fill">-</td><td class="r fill">-</td></tr>'
    return (f'<tr><td{lc}>{label}</td>'
            f'<td class="r num">{won(supply)}</td>'
            f'<td class="r num">{won(tax)}</td></tr>')


def row2(label, amt, is_sum=False):
    cls = ' class="sum"' if is_sum else ""
    cell = (f'<td class="r fill">-</td>' if to_num(amt) == 0 and not is_sum
            else f'<td class="r num">{won(amt)}</td>')
    return f'<tr{cls}><td>{label}</td>{cell}</tr>'


def card_rows(cards):
    return "".join(f'<tr><td>{사}</td><td class="r card">{번호}</td></tr>'
                   for 사, 번호 in cards)


def build_cards_block(cards):
    """카드 수에 따라 2→3→4열 자동 확장하여 한 페이지에 담는다."""
    head = '<thead><tr><th>카드사</th><th class="r">카드번호</th></tr></thead>'
    n = len(cards)
    if n == 0:
        body = '<tbody><tr><td class="rmk" colspan="2">등록된 사업용카드 없음</td></tr></tbody>'
        return f'<div class="cardcols" style="grid-template-columns:1fr 1fr;"><table class="tbl">{head}{body}</table></div>'
    cols = 2 if n <= 8 else (3 if n <= 12 else 4)
    per = math.ceil(n / cols)
    groups = [cards[i * per:(i + 1) * per] for i in range(cols)]
    tables = "".join(f'<table class="tbl">{head}<tbody>{card_rows(g)}</tbody></table>'
                     for g in groups)
    return f'<div class="cardcols" style="grid-template-columns:repeat({cols},1fr);">{tables}</div>'


# ============================================================
# 렌더링
# ============================================================
def render_company(row: dict, cfg: dict) -> str:
    # 매출: 세금계산서 매출 / 신카 및 현영 매출 / 그 외 매출
    s_tax_inv = (to_num(row.get("세금계산서매출_공급")), to_num(row.get("세금계산서매출_세액")))
    s_card = (to_num(row.get("신카및현영매출_공급")), to_num(row.get("신카및현영매출_세액")))
    s_etc = (to_num(row.get("그외매출_공급")), to_num(row.get("그외매출_세액")))
    s_sup = s_tax_inv[0] + s_card[0] + s_etc[0]
    s_tax = s_tax_inv[1] + s_card[1] + s_etc[1]
    sales_rows = (row3("세금계산서 매출", *s_tax_inv) + row3("신카 및 현영 매출", *s_card)
                  + row3c("그 외 매출 (판매대행사 등)", *s_etc, label_cls="lbl2")
                  + row3("소계", s_sup, s_tax, is_sum=True))

    # 매입
    p_tax_inv = (to_num(row.get("세금계산서매입_공급")), to_num(row.get("세금계산서매입_세액")))
    p_card = (to_num(row.get("신용카드매입_공급")), to_num(row.get("신용카드매입_세액")))
    p_deem = (to_num(row.get("의제매입_공급")), to_num(row.get("의제매입_세액")))
    p_sup = p_tax_inv[0] + p_card[0] + p_deem[0]
    p_tax = p_tax_inv[1] + p_card[1] + p_deem[1]
    purchase_rows = (row3("세금계산서 합계표", *p_tax_inv)
                     + row3("신용카드매입세액공제", *p_card)
                     + row3("의제매입세액공제", *p_deem)
                     + row3("소계", p_sup, p_tax, is_sum=True))

    # 세액공제: 신용카드매출세액공제 = (신카및현영 + 그외) 발행대가 × 공제율, 한도 적용
    rate = (to_num(cfg.get("공제율")) / 100) if to_num(cfg.get("공제율")) else 0.013
    limit = to_num(cfg.get("공제한도")) or 5_000_000
    발행대가 = (s_card[0] + s_card[1]) + (s_etc[0] + s_etc[1])
    cr_card = round(min(발행대가 * rate, limit))
    cr_etc = to_num(row.get("기타세액공제"))
    cr_sum = cr_card + cr_etc
    credit_rows = (row2("신용카드매출세액공제", cr_card) + row2("기타세액공제", cr_etc)
                   + row2("세액공제 합계", cr_sum, is_sum=True))

    # 예상 납부(환급)세액 (예정고지세액 차감 반영)
    예정고지 = to_num(row.get("예정고지세액"))
    payable = s_tax - p_tax - cr_sum - 예정고지
    estimate_rows = (row2("매출세액", s_tax) + row2("매입세액 (−)", p_tax)
                     + row2("세액공제 (−)", cr_sum)
                     + row2("예정고지세액 (−)", 예정고지)
                     + row2("예상 납부세액", payable, is_sum=True))

    # 사업용카드 (형식: "롯데카드:5342-..|비씨카드:6556-..")
    cards = []
    raw = str(row.get("사업용카드") or "").strip()
    if raw:
        for item in re.split(r"[|\n]", raw):
            item = item.strip()
            if not item:
                continue
            if ":" in item:
                사, 번호 = item.split(":", 1)
            else:
                사, 번호 = item, ""
            cards.append((사.strip(), 번호.strip()))
    cards_block = build_cards_block(cards)

    html = TEMPLATE
    repl = {
        "EYEBROW": cfg.get("귀속안내", ""),
        "BIZ_NAME": str(row.get("업체명", "")),
        "WRITE_DATE": cfg.get("작성일자", ""),
        "SALES_ROWS": sales_rows,
        "PURCHASE_ROWS": purchase_rows,
        "WARN_TITLE": cfg.get("주의제목", "⚠ 추정치 안내 (가정값)"),
        "WARN_NOTE": cfg.get("주의문구", ""),
        "CREDIT_ROWS": credit_rows,
        "ESTIMATE_ROWS": estimate_rows,
        "FINAL_DESC": cfg.get("예상세액설명", ""),
        "FINAL_VALUE": won(payable),
        "CARDS_BLOCK": cards_block,
        "MEMO_ROWS": cfg.get("_memo_html", ""),
        "FOOTER": cfg.get("푸터", "부가가치세 안내문 · 세무법인 재경"),
    }
    for k, v in repl.items():
        html = html.replace("{{" + k + "}}", str(v))
    return html


def load_config(path) -> dict:
    cfg = dict(zip(*[pd.read_excel(path, sheet_name="기본설정").fillna("")[c]
                     for c in ("항목", "값")]))
    memo = pd.read_excel(path, sheet_name="안내메모").fillna("")
    cfg["_memo_html"] = "".join(f"<p>{t}</p>" for t in memo["문구"] if str(t).strip())
    return cfg


def safe_name(s):
    return re.sub(r'[\\/:*?"<>|]', "_", str(s)).strip()


# ============================================================
# 입력양식 샘플 (근처횟집)
# ============================================================
def make_sample():
    cfg = pd.DataFrame({
        "항목": ["귀속안내", "작성일자", "예상세액설명", "주의제목", "주의문구", "푸터", "공제율", "공제한도"],
        "값": [
            "2026년 제1기 부가가치세 · 예정 안내",
            "2026. 06. 15.",
            "2026년 제1기 부가가치세 (매출세액 − 매입세액 − 세액공제 − 예정고지세액)",
            "⚠ 추정치 안내 (가정값)",
            "6월의 경우 데이터 확정 전(7월 중순 확정)이므로 <b>5월과 동일하다고 가정한 값</b>입니다. 따라서 위 금액은 확정치가 아닌 <b>추정치</b>이며, 실제 납부할 세액과 차이가 있을 수 있습니다.",
            "부가가치세 안내문 · 세무법인 재경",
            "1.3",
            "5000000",
        ],
    })
    memo = pd.DataFrame({"문구": [
        "① 세금계산서 및 계산서 수취내역 확인이 필요하신 경우 당사에 요청 바랍니다.",
        "② 미수취 세금계산서 및 계산서가 있는 경우 <b>7월 10일까지</b> 필히 수취해 주시길 바랍니다.",
        "③ 위 등록된 사업용카드 외 추가카드를 사용하시는 경우 신용카드 사용내역 제출 바랍니다. <b>(26.01.01 ~ 26.06.30 사용분)</b>",
        "④ 사업용카드 목록 중 신규카드 및 분실·해제카드가 있는 경우 상시 등록 및 해제 요청 바랍니다.",
    ]})
    companies = pd.DataFrame([{
        "업체명": "근처횟집",
        "세금계산서매출_공급": 0, "세금계산서매출_세액": 0,
        "신카및현영매출_공급": 122905111, "신카및현영매출_세액": 12290510,
        "그외매출_공급": 0, "그외매출_세액": 0,
        "세금계산서매입_공급": 4104940, "세금계산서매입_세액": 410494,
        "신용카드매입_공급": 15000000, "신용카드매입_세액": 1363636,
        "의제매입_공급": 55000000, "의제매입_세액": 4541284,
        "기타세액공제": 0, "예정고지세액": 0,
        "사업용카드": "롯데카드:5342-92**-****-7431 | 비씨카드:6556-32**-****-4724 | 신한카드:5421-58**-****-9944 | 현대카드:5472-27**-****-3824",
    }])
    with pd.ExcelWriter(INPUT_XLSX, engine="openpyxl") as w:
        companies.to_excel(w, sheet_name="업체별", index=False)
        cfg.to_excel(w, sheet_name="기본설정", index=False)
        memo.to_excel(w, sheet_name="안내메모", index=False)
    print(f"입력양식 생성: {INPUT_XLSX}")
    print("  - '업체별' 시트에 업체를 추가하면 한 번에 여러 장 생성됩니다.")


# ============================================================
# 메인
# ============================================================
def main():
    ap = argparse.ArgumentParser(description="부가가치세 안내문 생성기")
    ap.add_argument("--init", action="store_true", help="입력양식 샘플 생성")
    ap.add_argument("--input", default=str(INPUT_XLSX))
    args = ap.parse_args()

    if args.init:
        make_sample()
        _open(INPUT_XLSX)
        return

    path = Path(args.input)
    if not path.exists():
        make_sample()
        print("\n입력양식(안내문_입력.xlsx)을 새로 만들었습니다.")
        print("업체 정보를 채워 저장한 뒤, 다시 실행하세요.")
        _open(INPUT_XLSX)
        return

    cfg = load_config(path)
    companies = pd.read_excel(path, sheet_name="업체별").fillna("")
    companies = companies[companies["업체명"].astype(str).str.strip() != ""]

    OUT_DIR.mkdir(exist_ok=True)
    made = []
    for _, row in companies.iterrows():
        html = render_company(row.to_dict(), cfg)
        fn = OUT_DIR / f"안내문_{safe_name(row['업체명'])}.html"
        fn.write_text(html, encoding="utf-8")
        made.append(fn.name)

    print(f"\n안내문 {len(made)}건 생성 완료 -> {OUT_DIR}")
    for n in made:
        print(f"   · {n}")
    _open(OUT_DIR)


def _open(target):
    try:
        os.startfile(str(target))
    except Exception:
        pass


if __name__ == "__main__":
    main()
