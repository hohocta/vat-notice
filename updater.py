# -*- coding: utf-8 -*-
"""
자동 업데이트 모듈 — 프로그램을 켤 때 GitHub에서 최신 파일을 받아 적용합니다.
- 인터넷이 안 되면 조용히 건너뛰고 기존 버전으로 실행합니다(멈추지 않음).
- 표준 라이브러리만 사용(urllib/json) → 추가 설치 불필요.
- 고객 데이터는 건드리지 않습니다('files' 목록의 프로그램 파일만 교체).
- 개발(편집) 폴더는 '개발모드.flag' 파일로 보호되어 업데이트하지 않습니다.

동작:
  1) GitHub API로 최신 커밋 번호(SHA)를 확인 — 캐시를 타지 않아 '즉시' 최신을 봅니다.
  2) 저장해 둔 번호와 같으면 최신이므로 그대로 실행.
  3) 다르면 그 커밋에 고정된 주소로 파일을 받아 한꺼번에 교체.

흐름: 안내문작성.bat 이 app.py 보다 먼저 이 파일을 실행합니다.
"""
import json
import urllib.parse
import urllib.request
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
CONFIG = HERE / "update_config.json"
VERSION_FILE = HERE / "VERSION"
SHA_FILE = HERE / ".update_sha"   # 마지막으로 적용한 커밋 번호(직원 PC 기록용)
LOG = HERE / "업데이트로그.txt"
TIMEOUT = 8  # 초 (인터넷 응답 대기 한도)


def log(msg):
    try:
        with LOG.open("a", encoding="utf-8") as f:
            f.write(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}\n")
    except Exception:
        pass


def fetch(url, binary=False, accept=None):
    headers = {"User-Agent": "vat-notice-updater"}
    if accept:
        headers["Accept"] = accept
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, timeout=TIMEOUT) as r:
        data = r.read()
    return data if binary else data.decode("utf-8")


def main():
    # 개발(편집) 폴더 보호: 이 표시 파일이 있으면 자동 업데이트를 건너뜀.
    # → 배포 전에 프로그램을 켜도 수정 중인 코드가 옛 버전으로 덮어써지지 않음.
    if (HERE / "개발모드.flag").exists():
        return
    if not CONFIG.exists():
        return
    try:
        cfg = json.loads(CONFIG.read_text(encoding="utf-8-sig"))
    except Exception as e:
        log(f"설정 읽기 실패: {e}")
        return

    owner = str(cfg.get("owner", "")).strip()
    repo = str(cfg.get("repo", "")).strip()
    branch = str(cfg.get("branch", "main")).strip() or "main"
    if not owner or not repo or owner.startswith("<"):
        return  # 아직 배포처(GitHub 계정)가 설정되지 않음 → 조용히 통과

    # 1) 최신 커밋 번호(SHA) 확인 — API라서 캐시 없이 즉시 최신을 봄
    try:
        meta = json.loads(fetch(
            f"https://api.github.com/repos/{owner}/{repo}/commits/{branch}",
            accept="application/vnd.github+json"))
        sha = str(meta.get("sha", "")).strip()
    except Exception:
        return  # 인터넷 안 됨 / 사용량 초과 / 저장소 없음 → 기존 버전으로 실행
    if not sha:
        return

    local_sha = SHA_FILE.read_text(encoding="utf-8-sig").strip() if SHA_FILE.exists() else ""
    if sha == local_sha:
        return  # 이미 최신

    # 2) 그 커밋에 '고정된' 주소로 받음 → 캐시 영향 없이 신선하고, 파일들이 서로 어긋나지 않음
    base = f"https://raw.githubusercontent.com/{owner}/{repo}/{sha}/"

    # 교체할 파일 목록(원격 설정 우선 → 새 파일 추가도 자동 반영)
    files = cfg.get("files", [])
    try:
        rcfg = json.loads(fetch(base + "update_config.json"))
        files = rcfg.get("files", files) or files
    except Exception:
        pass
    if not files:
        return

    # 전부 임시로 받음 — 하나라도 실패하면 이번 업데이트 전체 취소(반쪽 업데이트 방지)
    staged = {}
    for name in files:
        try:
            staged[name] = fetch(base + urllib.parse.quote(name), binary=True)
        except Exception as e:
            log(f"내려받기 실패({name}): {e} → 이번 업데이트 취소")
            return

    try:
        new_ver = fetch(base + "VERSION").strip()
    except Exception:
        new_ver = ""

    # 한꺼번에 적용
    for name, data in staged.items():
        try:
            (HERE / name).write_bytes(data)
        except Exception as e:
            log(f"적용 실패({name}): {e}")

    if new_ver:
        VERSION_FILE.write_text(new_ver, encoding="utf-8")
    SHA_FILE.write_text(sha, encoding="utf-8")
    log(f"업데이트 완료: → {new_ver or '(버전미상)'} ({sha[:7]})")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log(f"예외: {e}")
