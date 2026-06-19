# -*- coding: utf-8 -*-
"""
부가가치세 예상 안내문 작성 프로그램 (브라우저 방식 · Python 연동)
================================================================
- 파이썬 내장 웹서버가 화면(ui.html)을 띄우고 기본 브라우저로 엽니다.
- 엑셀/저장/안내문 출력은 Python(/api/...)이 처리합니다.
- pywebview(.NET) 미사용 → 네이티브 창 오류 없음. 추가 설치 불필요.

실행: 안내문작성.bat 더블클릭  (검은 서버 창을 닫으면 종료)
"""
import json
import os
import sys
import threading
import time
import webbrowser
from datetime import date, datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

import report_gen as rg

HERE = Path(__file__).resolve().parent
OUT_DIR = HERE / "안내문_출력"
STATE = HERE / "작업데이터.json"


# ============================================================
# 계산 (5개월 → 6개월 환산)
# ============================================================
def conv(x):
    return round((x or 0) * 6 / 5)


def tax10(supply):
    return round(supply * 0.1)


def rate_val(r):
    """'9/109' → 0.0826... 분수 문자열을 비율로."""
    p = str(r or "").split("/")
    try:
        return float(p[0]) / float(p[1]) if len(p) == 2 and float(p[1]) else 0.0
    except (ValueError, ZeroDivisionError):
        return 0.0


DEFAULT_MEMO = [
    "① 세금계산서 및 계산서 수취내역 확인이 필요하신 경우 당사에 요청 바랍니다.",
    "② 미수취 세금계산서 및 계산서가 있는 경우 <b>기한 내</b> 필히 수취해 주시길 바랍니다.",
    "③ 위 등록된 사업용카드 외 추가카드를 사용하시는 경우 신용카드 사용내역 제출 바랍니다.",
    "④ 사업용카드 목록 중 신규카드 및 분실·해제카드가 있는 경우 상시 등록 및 해제 요청 바랍니다.",
]


def build_report_html(client: dict, period: dict) -> str:
    year = int(period.get("year", date.today().year))
    half = int(period.get("half", 1))
    s = client.get("sales", {}) or {}
    m = client.get("manual", {}) or {}
    cr = client.get("credit", {}) or {}

    s_ti, s_cd, s_cs = conv(s.get("taxInv")), conv(s.get("card")), conv(s.get("cash"))
    s_sup = s_ti + s_cd + s_cs
    s_tax = tax10(s_sup)
    sales_rows = (rg.row3("세금계산서 매출", s_ti, tax10(s_ti))
                  + rg.row3c("신용카드·현금영수증 매출", s_cd, tax10(s_cd), label_cls="lbl2")
                  + rg.row3c("그 외 매출(결제대행사 등)", s_cs, tax10(s_cs), label_cls="lbl2")
                  + rg.row3("소계", s_sup, s_tax, is_sum=True))

    p_ti = conv(m.get("buyTaxInv"))
    p_cd = conv(m.get("buyCard"))
    deemed_amt = conv(m.get("deemedAmt"))
    deemed_tax = round(deemed_amt * rate_val(m.get("deemedRate")))
    p_tax = tax10(p_ti) + tax10(p_cd) + deemed_tax
    purchase_rows = (rg.row3("세금계산서 매입", p_ti, tax10(p_ti))
                     + rg.row3("신용카드 등 매입", p_cd, tax10(p_cd))
                     + rg.row3("의제매입세액공제", deemed_amt, deemed_tax)
                     + rg.row3("소계", p_ti + p_cd, p_tax, is_sum=True))

    cr_card = rg.to_num(cr.get("card"))
    cr_etc = rg.to_num(cr.get("etc"))
    cr_sum = cr_card + cr_etc
    credit_rows = (rg.row2("신용카드매출전표 발행세액공제", cr_card)
                   + rg.row2("기타세액공제", cr_etc)
                   + rg.row2("세액공제 합계", cr_sum, is_sum=True))

    notice = rg.to_num(client.get("notice"))
    payable = s_tax - p_tax - cr_sum - notice
    estimate_rows = (rg.row2("매출세액", s_tax) + rg.row2("매입세액 (−)", p_tax)
                     + rg.row2("세액공제 (−)", cr_sum)
                     + rg.row2("예정고지세액 (−)", notice)
                     + rg.row2("예상 납부세액", payable, is_sum=True))

    # 정상등록(ok) 카드만, 카드사+번호로 출력
    cards = [(cd.get("issuer", ""), cd.get("no", ""))
             for cd in client.get("cards", []) if cd.get("status") == "ok"]
    cards_block = rg.build_cards_block(cards)

    repl = {
        "BIZ_NAME": str(client.get("name", "")),
        "EYEBROW": f"{year}년 제{half}기 부가가치세 · 예정 안내",
        "WRITE_DATE": date.today().strftime("%Y. %m. %d."),
        "SALES_ROWS": sales_rows, "PURCHASE_ROWS": purchase_rows,
        "WARN_TITLE": "⚠ 추정치 안내 (5개월 → 6개월 환산)",
        "WARN_NOTE": "현재 5개월 실적을 6개월로 환산(×1.2)한 <b>추정치</b>입니다. 실제 확정 신고 시 납부세액과 차이가 있을 수 있습니다.",
        "CREDIT_ROWS": credit_rows, "ESTIMATE_ROWS": estimate_rows,
        "FINAL_DESC": f"{year}년 제{half}기 부가가치세 (매출세액 − 매입세액 − 세액공제 − 예정고지세액)",
        "FINAL_VALUE": rg.won(payable),
        "CARDS_BLOCK": cards_block,
        "MEMO_ROWS": "".join(f"<p>{t}</p>" for t in DEFAULT_MEMO),
        "FOOTER": "부가가치세 안내문 · 세무법인 재경",
    }
    html = rg.TEMPLATE
    for k, v in repl.items():
        html = html.replace("{{" + k + "}}", str(v))
    return html


# ============================================================
# API 처리
# ============================================================
def api_generate(payload):
    # 미리보기: 파일을 만들지 않고 HTML만 반환
    client = payload.get("client", {})
    html = build_report_html(client, payload.get("period", {}))
    return {"ok": True, "html": html}


def _render_one(page, html, base):
    """한 건을 PDF + PNG로 저장 (잘림 없음)."""
    page.set_content(html, wait_until="load")
    try:
        page.evaluate("document.fonts && document.fonts.ready")
    except Exception:
        pass
    page.wait_for_timeout(220)
    page.locator(".page").screenshot(path=str(base.with_suffix(".png")))   # 요소 경계 캡처
    page.pdf(path=str(base.with_suffix(".pdf")), prefer_css_page_size=True, print_background=True,
             margin={"top": "0", "bottom": "0", "left": "0", "right": "0"})


def render_files(html, base):
    """실제 Chromium 으로 안내문을 PDF + PNG(이미지)로 저장."""
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        browser = p.chromium.launch()
        try:
            page = browser.new_page(device_scale_factor=2, viewport={"width": 820, "height": 1160})
            _render_one(page, html, base)
        finally:
            browser.close()
    return base.with_suffix(".pdf"), base.with_suffix(".png")


def api_confirm(payload):
    """확정·출력: 작성기간별 하위 폴더에 안내문을 PDF + 이미지(PNG)로 저장."""
    client = payload.get("client", {})
    period = payload.get("period", {})
    html = build_report_html(client, period)
    year = int(period.get("year", date.today().year))
    half = int(period.get("half", 1))
    sub = OUT_DIR / f"{year}년 제{half}기"      # 작성기간별 폴더(없으면 생성, 있으면 재사용)
    sub.mkdir(parents=True, exist_ok=True)
    base = sub / f"안내문_{rg.safe_name(client.get('name', '업체'))}_{date.today().strftime('%Y%m%d')}"
    base.with_suffix(".html").write_text(html, encoding="utf-8")
    pdf_path, png_path = render_files(html, base)
    try:
        os.startfile(str(sub))
    except Exception:
        pass
    return {"ok": True, "folder": sub.name, "pdf": pdf_path.name, "png": png_path.name}


def api_confirm_batch(payload):
    """완료 거래처 일괄 출력: Chromium 한 번만 띄워 여러 건 PDF + PNG 저장."""
    clients = payload.get("clients", [])
    period = payload.get("period", {})
    year = int(period.get("year", date.today().year))
    half = int(period.get("half", 1))
    sub = OUT_DIR / f"{year}년 제{half}기"
    sub.mkdir(parents=True, exist_ok=True)
    ymd = date.today().strftime("%Y%m%d")
    names = []
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        browser = p.chromium.launch()
        try:
            for client in clients:
                html = build_report_html(client, period)
                base = sub / f"안내문_{rg.safe_name(client.get('name', '업체'))}_{ymd}"
                base.with_suffix(".html").write_text(html, encoding="utf-8")
                page = browser.new_page(device_scale_factor=2, viewport={"width": 820, "height": 1160})
                _render_one(page, html, base)
                page.close()
                names.append(client.get("name", ""))
        finally:
            browser.close()
    try:
        os.startfile(str(sub))
    except Exception:
        pass
    return {"ok": True, "count": len(names), "folder": sub.name}


def api_open_folder(payload):
    """저장 폴더(안내문_출력) 열기."""
    OUT_DIR.mkdir(exist_ok=True)
    try:
        os.startfile(str(OUT_DIR))
    except Exception:
        pass
    return {"ok": True}


def api_backup(payload):
    """전체 데이터를 날짜·시각이 붙은 파일로 '백업' 폴더에 저장.
    payload.label: 파일명 접두어, payload.open: 저장 후 폴더 열기 여부."""
    bk = HERE / "백업"
    bk.mkdir(exist_ok=True)
    label = rg.safe_name(str(payload.pop("label", "") or "작업데이터_백업")) or "작업데이터_백업"
    show = payload.pop("open", True)
    fn = bk / f"{label}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    fn.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    if show:
        try:
            os.startfile(str(bk))
        except Exception:
            pass
    return {"ok": True, "file": fn.name}


def api_save_state(payload):
    STATE.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return {"ok": True}


def api_pick_excel(payload):
    # 실제 '5개월 실적' 엑셀 형식 확정 후 연동 예정
    return {"ok": False, "note": "엑셀 불러오기는 실제 파일 형식 확정 후 연동됩니다."}


def api_shutdown(payload):
    # 응답을 보낸 뒤 잠시 후 프로세스 종료
    def _stop():
        time.sleep(0.3)
        os._exit(0)
    threading.Thread(target=_stop, daemon=True).start()
    return {"ok": True}


# ============================================================
# 웹서버
# ============================================================
class Handler(BaseHTTPRequestHandler):
    def _send(self, code, body, ctype="application/json"):
        data = body if isinstance(body, bytes) else str(body).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", f"{ctype}; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate")
        self.end_headers()
        try:
            self.wfile.write(data)
        except (BrokenPipeError, ConnectionResetError):
            pass

    def do_GET(self):
        if self.path in ("/", "/index.html", "/ui.html"):
            self._send(200, (HERE / "ui.html").read_bytes(), "text/html")
        elif self.path == "/api/load_state":
            self._send(200, STATE.read_text(encoding="utf-8") if STATE.exists() else "null")
        else:
            self._send(404, "{}")

    def do_POST(self):
        ln = int(self.headers.get("Content-Length", 0) or 0)
        raw = self.rfile.read(ln).decode("utf-8") if ln else ""
        try:
            payload = json.loads(raw) if raw else {}
        except json.JSONDecodeError:
            payload = {}
        routes = {"/api/generate": api_generate, "/api/confirm": api_confirm,
                  "/api/confirm_batch": api_confirm_batch,
                  "/api/open_folder": api_open_folder, "/api/backup": api_backup,
                  "/api/save_state": api_save_state,
                  "/api/pick_excel": api_pick_excel, "/api/shutdown": api_shutdown}
        fn = routes.get(self.path)
        if not fn:
            self._send(404, "{}")
            return
        try:
            result = fn(payload)
        except Exception as e:
            result = {"ok": False, "error": str(e)}
        self._send(200, json.dumps(result, ensure_ascii=False))

    def log_message(self, *args):
        pass  # 콘솔 조용히


def find_server(start=8765, tries=15):
    for port in range(start, start + tries):
        try:
            return ThreadingHTTPServer(("127.0.0.1", port), Handler), port
        except OSError:
            continue
    raise OSError("사용 가능한 포트를 찾지 못했습니다.")


def main():
    server, port = find_server()
    try:
        (HERE / "server.pid").write_text(str(os.getpid()), encoding="utf-8")
    except Exception:
        pass
    url = f"http://127.0.0.1:{port}/"
    print("=" * 50)
    print("  부가가치세 예상 안내문 작성 프로그램")
    print("=" * 50)
    print(f"  화면 주소: {url}")
    print("  브라우저가 자동으로 열립니다.")
    print("  ※ 종료하려면 이 검은 창을 닫으세요.")
    print("=" * 50)
    if "--noopen" not in sys.argv:
        webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        # 숨김 실행(콘솔 없음) 중 오류 시 메시지 창으로 알림
        try:
            import tkinter
            from tkinter import messagebox
            r = tkinter.Tk()
            r.withdraw()
            messagebox.showerror("부가가치세 안내문 프로그램 오류", str(e))
        except Exception:
            (HERE / "오류로그.txt").write_text(str(e), encoding="utf-8")
