# -*- coding: utf-8 -*-
"""
위킵(FBW) 재고 자동 동기화 스크립트
=====================================
하는 일 (순서대로):
  1) 위킵 사이트에 자동으로 로그인
  2) 재고 엑셀을 다운로드
  3) 엑셀 내용을 읽어서 정리
  4) clindex.html이 사용하는 Firestore(클라우드 DB)에 그대로 업로드

★ 완전 초보자를 위한 안내 ★
- 이 파일은 절대 혼자 실행하는 게 아니고, GitHub Actions가 매일 자동으로 실행해줍니다.
- 아이디/비밀번호는 이 파일 안에 적지 않습니다. GitHub Secrets에서 안전하게 불러옵니다.
- "# TODO" 라고 적힌 줄만 나중에 위킵 사이트에 맞게 살짝 고치면 됩니다. (제가 안내해드릴게요)
"""

import os
import re
import json
from datetime import datetime, timezone

from playwright.sync_api import sync_playwright
import openpyxl
import firebase_admin
from firebase_admin import credentials, firestore


# =========================================================
# 1. 환경설정 값들 (GitHub Secrets에서 자동으로 채워짐 — 여기 직접 쓰지 마세요!)
# =========================================================
WEEKEEP_ID = os.environ["WEEKEEP_ID"]
WEEKEEP_PW = os.environ["WEEKEEP_PW"]
FIREBASE_SERVICE_ACCOUNT_JSON = os.environ["FIREBASE_SERVICE_ACCOUNT"]

# clindex.html과 반드시 동일해야 하는 값들 (이미 확인 완료, 수정 불필요)
FIREBASE_PROJECT_ID = "clstpock"
APP_ID = "inventory-tracker-01"

DOWNLOAD_DIR = "./downloads"

# TODO: 위킵 로그인 페이지의 실제 주소로 바꿔주세요.
WEEKEEP_LOGIN_URL = "https://www.wekeep.co.kr/login"

# TODO: 로그인 성공 후, 재고 목록이 있는 페이지 주소로 바꿔주세요. (필요 없으면 비워두세요)
WEEKEEP_INVENTORY_URL = ""


# =========================================================
# 2. 위킵 로그인 후 재고 엑셀 다운로드
# =========================================================
def download_weekeep_excel() -> str:
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        try:
            print("① 위킵 로그인 페이지로 이동합니다...")
            page.goto(WEEKEEP_LOGIN_URL, wait_until="networkidle")

            # v3: "input[type='text']"는 실제 HTML에 type 속성이 없는 입력창은 못 찾음 (많은 로그인폼이 이런 구조).
            # → 비밀번호/체크박스/숨김/버튼류가 "아닌" 모든 입력창을 아이디 칸으로 간주하도록 조건을 넓힘.
            password_field = page.locator("input[type='password']").first
            id_field = page.locator(
                "input:not([type='password']):not([type='checkbox']):not([type='radio'])"
                ":not([type='hidden']):not([type='submit']):not([type='button'])"
            ).first

            id_field.fill(WEEKEEP_ID)
            password_field.fill(WEEKEEP_PW)

            # v2: 로그인 버튼 문구가 "시작하기"로 확인됨 (실제 화면 캡처 기준)
            page.get_by_role("button", name=re.compile("시작하기|로그인")).click()
            page.wait_for_load_state("networkidle")
            print("② 로그인 완료 (아마도) — 재고 페이지로 이동합니다...")

            if WEEKEEP_INVENTORY_URL:
                page.goto(WEEKEEP_INVENTORY_URL, wait_until="networkidle")

            # TODO: "엑셀 다운로드" 버튼의 실제 문구/위치에 맞게 수정하세요.
            with page.expect_download(timeout=60000) as download_info:
                page.get_by_text(re.compile("엑셀\\s*다운로드")).first.click()
            download = download_info.value

            filepath = os.path.join(DOWNLOAD_DIR, download.suggested_filename or "wk_stock.xlsx")
            download.save_as(filepath)
            print(f"③ 엑셀 다운로드 완료: {filepath}")
            return filepath

        except Exception as e:
            # v2: 실패한 화면을 그대로 사진(스크린샷)과 html로 저장.
            # GitHub Actions가 이 파일들을 "Artifacts"로 올려주므로, 어디서 왜 막혔는지 눈으로 바로 확인 가능.
            os.makedirs("debug", exist_ok=True)
            page.screenshot(path="debug/failure_screenshot.png", full_page=True)
            with open("debug/failure_page.html", "w", encoding="utf-8") as f:
                f.write(page.content())
            print(f"❌ 실패한 지점: {e}")
            print("→ debug 폴더에 실패 시점 화면 캡처를 저장했습니다 (Actions 결과 화면 하단 Artifacts에서 다운로드 가능)")
            raise

        finally:
            browser.close()


# =========================================================
# 3. 엑셀 내용을 읽어서 items 리스트로 정리
#    (이 부분은 위킵 엑셀 컬럼 구조를 이미 알고 있으므로 수정하실 필요 없습니다)
# =========================================================
def parse_excel(filepath: str):
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise ValueError("엑셀에 데이터가 없습니다.")

    header = [str(h).strip() if h is not None else "" for h in rows[0]]

    def find_idx(keywords):
        for kw in keywords:
            for i, col in enumerate(header):
                if col == kw:
                    return i
        return -1

    name_idx = find_idx(["상품명"])
    avail_idx = find_idx(["가용재고"])
    safe_idx = find_idx(["안전재고"])
    reserved_idx = find_idx(["유보재고"])
    defective_idx = find_idx(["하자재고"])
    turnover_idx = find_idx(["재고회전률", "재고회전율"])
    monthly_idx = find_idx(["월평균출고량"])
    daily_idx = find_idx(["일평균출고량"])

    if name_idx == -1 or avail_idx == -1:
        raise ValueError(f"엑셀 헤더에서 '상품명' 또는 '가용재고' 컬럼을 못 찾았습니다. 헤더: {header}")

    def to_num(v):
        if v in (None, ""):
            return 0
        try:
            return float(str(v).replace(",", ""))
        except ValueError:
            return 0

    items = []
    for row in rows[1:]:
        if not row or name_idx >= len(row):
            continue
        name = str(row[name_idx]).strip() if row[name_idx] else ""
        if not name or name == "상품명":
            continue
        items.append({
            "itemName": name,
            "stock": to_num(row[avail_idx]) if avail_idx != -1 else 0,
            "safeStock": to_num(row[safe_idx]) if safe_idx != -1 and safe_idx < len(row) else 0,
            "reservedStock": to_num(row[reserved_idx]) if reserved_idx != -1 and reserved_idx < len(row) else 0,
            "defectiveStock": to_num(row[defective_idx]) if defective_idx != -1 and defective_idx < len(row) else 0,
            "turnoverRate": to_num(row[turnover_idx]) if turnover_idx != -1 and turnover_idx < len(row) else 0,
            "avgMonthlyOut": to_num(row[monthly_idx]) if monthly_idx != -1 and monthly_idx < len(row) else 0,
            "avgDailyOut": to_num(row[daily_idx]) if daily_idx != -1 and daily_idx < len(row) else 0,
        })

    print(f"④ 엑셀 파싱 완료: {len(items)}개 품목")
    return items


# =========================================================
# 4. Firestore(clindex.html이 읽는 클라우드 DB)에 업로드
# =========================================================
def upload_to_firestore(items):
    cred_dict = json.loads(FIREBASE_SERVICE_ACCOUNT_JSON)
    cred = credentials.Certificate(cred_dict)
    firebase_admin.initialize_app(cred, {"projectId": FIREBASE_PROJECT_ID})
    db = firestore.client()

    coll_ref = (
        db.collection("artifacts")
        .document(APP_ID)
        .collection("public")
        .document("data")
        .collection("weekeep_inventory")
    )

    existing_count = len(list(coll_ref.stream()))
    now_iso = datetime.now(timezone.utc).isoformat()

    doc_data = {
        "version": f"V{existing_count + 1}",
        "createdAt": now_iso,
        "items": items,
    }

    coll_ref.add(doc_data)
    print(f"⑤ 업로드 완료! (버전: {doc_data['version']}, {len(items)}개 품목)")


# =========================================================
# 실행
# =========================================================
if __name__ == "__main__":
    excel_path = download_weekeep_excel()
    parsed_items = parse_excel(excel_path)
    upload_to_firestore(parsed_items)
    print("🎉 모든 작업이 끝났습니다. clindex.html을 열어서 확인해보세요.")
