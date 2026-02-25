"""
홍익대학교 수강신청 모니터링 스크립트
=====================================

사용법:
  1. pip install -r requirements.txt
  2. python monitor.py
  3. 브라우저가 열리면 직접 로그인 & 과목 조회 페이지로 이동
  4. 콘솔에서 Enter 입력
  5. 브라우저에서 모니터링할 과목 행을 클릭하여 선택
  6. 우측 상단 패널의 '모니터링 시작' 버튼 클릭
  7. 자리 발생 시 소리 + 데스크톱 알림
"""



import ctypes
import json
import os
import sys
import time
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import Any

# ---- 콘솔 한글 출력 설정 ----
def configure_console() -> None:
    """Windows 콘솔에서 한글이 정상 출력되도록 UTF-8 설정"""
    if sys.platform.startswith("win"):
        try:
            ctypes.windll.kernel32.SetConsoleOutputCP(65001)  # type: ignore[attr-defined]
            ctypes.windll.kernel32.SetConsoleCP(65001)  # type: ignore[attr-defined]
        except Exception:
            pass
    reconfigure = getattr(sys.stdout, "reconfigure", None)
    if callable(reconfigure):
        try:
            reconfigure(encoding="utf-8")
        except Exception:
            pass


configure_console()

# ---- 의존성 임포트 ----
try:
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.common.exceptions import (
        WebDriverException,
        TimeoutException,
        UnexpectedAlertPresentException,
    )
except ImportError:
    print("[오류] selenium 패키지가 설치되지 않았습니다.")
    print("  → pip install -r requirements.txt")
    sys.exit(1)

try:
    from webdriver_manager.chrome import ChromeDriverManager
except ImportError:
    print("[오류] webdriver-manager 패키지가 설치되지 않았습니다.")
    print("  → pip install -r requirements.txt")
    sys.exit(1)

# Windows 소리 알림
import winsound

# 데스크톱 알림 (선택)
HAS_PLYER: bool = False
desktop_notification: Any = None
try:
    from plyer import notification as desktop_notification  # type: ignore[import-untyped]
    HAS_PLYER = True
except ImportError:
    pass

# ---- 경로 / 상수 ----
CONFIG_PATH = Path(__file__).resolve().parent / "config.json"
SUGANG_URL = "https://sugang.hongik.ac.kr/cn1000.jsp"


# ============================================
# 설정 로드
# ============================================

def load_config() -> dict[str, Any]:
    """config.json 로드. 없으면 기본값 생성."""
    defaults: dict[str, Any] = {
        "check_interval_seconds": 5,
        "alert": {
            "sound": True,
            "desktop_notification": True,
            "sound_repeat": 3,
        },
    }
    if not CONFIG_PATH.exists():
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(defaults, f, ensure_ascii=False, indent=2)
        print(f"[정보] 기본 설정 파일 생성: {CONFIG_PATH}")
        return defaults

    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        cfg = json.load(f)

    # 기본값 병합
    for key, val in defaults.items():
        if key not in cfg:
            cfg[key] = val
        elif isinstance(val, dict):
            for sub_key, sub_val in val.items():
                if sub_key not in cfg[key]:
                    cfg[key][sub_key] = sub_val

    # 최소 주기 보정 (0.1초 미만은 서버 부담)
    if cfg["check_interval_seconds"] < 0.1:
        print("[경고] 체크 주기가 너무 짧아 0.1초로 보정합니다.")
        cfg["check_interval_seconds"] = 0.1

    return cfg


# ============================================
# 브라우저 초기화
# ============================================

def init_driver() -> Any:
    """Chrome WebDriver 초기화 (자동 드라이버 관리)"""
    options = Options()
    options.add_argument("--lang=ko-KR")
    options.add_argument("--window-size=1500,950")
    # 자동화 감지 최소화
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    # 스크립트 종료 후에도 브라우저 유지
    options.add_experimental_option("detach", True)

    print("[정보] Chrome 드라이버 초기화 중...")
    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
    except Exception as e:
        print(f"[오류] Chrome 드라이버 시작 실패: {e}")
        print("  Chrome 브라우저가 설치되어 있는지 확인해주세요.")
        sys.exit(1)

    driver.set_page_load_timeout(30)
    return driver


# ============================================
# JavaScript 코드 (인터랙티브 과목 선택 오버레이)
# ============================================

PICKER_JS = r"""
(function() {
    // 중복 주입 방지
    if (window.__sugangPickerActive) return 'ALREADY_ACTIVE';
    window.__sugangPickerActive = true;
    window.__selectedCourses = [];
    window.__monitoringStarted = false;

    // ── 스타일 주입 ──
    const style = document.createElement('style');
    style.id = 'sugang-picker-style';
    style.textContent = `
        .sugang-hover { outline: 3px solid #42a5f5 !important; background-color: #e3f2fd !important; cursor: pointer !important; transition: all 0.15s ease; }
        .sugang-selected { outline: 3px solid #66bb6a !important; background-color: #e8f5e9 !important; }
        .sugang-selected:hover { background-color: #c8e6c9 !important; }
        #sugang-panel {
            position: fixed; top: 12px; right: 12px; z-index: 999999;
            background: #fff; border: 2px solid #1565c0; border-radius: 10px;
            padding: 18px; min-width: 340px; max-width: 440px;
            box-shadow: 0 8px 32px rgba(0,0,0,0.25);
            font-family: 'Pretendard','Malgun Gothic',sans-serif;
            max-height: 85vh; overflow-y: auto;
        }
        #sugang-panel h3 { margin:0 0 10px 0; color:#1565c0; font-size:17px; border-bottom:2px solid #1565c0; padding-bottom:8px; }
        .sg-guide { font-size:12.5px; color:#616161; margin:6px 0 12px 0; line-height:1.6; }
        .sg-item { display:flex; justify-content:space-between; align-items:center; padding:7px 10px; margin:4px 0; background:#f5f5f5; border-radius:5px; font-size:13px; }
        .sg-item-info { flex:1; }
        .sg-item-code { font-weight:bold; color:#1565c0; margin-right:6px; }
        .sg-item-enroll { color:#e53935; font-size:12px; margin-left:4px; }
        .sg-remove { color:#e53935; cursor:pointer; font-weight:bold; padding:2px 6px; font-size:18px; line-height:1; border:none; background:none; }
        .sg-remove:hover { color:#b71c1c; background:#ffebee; border-radius:3px; }
        #sg-start-btn {
            display:block; width:100%; padding:13px; margin-top:14px;
            background:#1565c0; color:#fff; border:none; border-radius:7px;
            font-size:15px; font-weight:bold; cursor:pointer;
            font-family:'Pretendard','Malgun Gothic',sans-serif;
            transition: background 0.2s;
        }
        #sg-start-btn:hover { background:#0d47a1; }
        #sg-start-btn:disabled { background:#bdbdbd; cursor:not-allowed; }
        .sg-empty { color:#9e9e9e; font-size:13px; padding:8px 0; }
        .sg-count { font-size:11px; color:#757575; margin-top:8px; }
    `;
    document.head.appendChild(style);

    // ── 과목 행 탐색 ──
    // 학수번호 패턴 (예: 101512-4, 14320-1)
    const codePattern = /^\d{5,6}-\d{1,2}$/;
    const enrollPattern = /^\d+\s*\/\s*\d+$/;
    let courseRows = [];

    // 모든 table을 순회하되, 중첩 테이블의 행은 가장 안쪽 것만 수집
    document.querySelectorAll('table tr').forEach(row => {
        const cells = row.querySelectorAll(':scope > td');  // 직계 자식 td만
        if (cells.length < 8) return;
        for (let i = 0; i < Math.min(cells.length, 7); i++) {
            const text = cells[i].textContent.trim();
            if (codePattern.test(text)) {
                courseRows.push({ row, codeIdx: i, cells });
                break;
            }
        }
    });

    if (courseRows.length === 0) {
        alert('과목 테이블을 찾을 수 없습니다.\n과목 조회 결과가 있는 페이지인지 확인해주세요.');
        window.__sugangPickerActive = false;
        return 'NO_TABLE';
    }

    // ── 컬럼 분석 (첫 번째 행 기준) ──
    function extractRowData(cells, codeIdx) {
        const courseCode = cells[codeIdx].textContent.trim();
        let courseName = '', enrollment = '', professor = '', schedule = '';

        // 과목명: 코드 다음의 긴 텍스트 (숫자만 있는 칸 건너뜀)
        for (let i = codeIdx + 1; i < cells.length; i++) {
            const t = cells[i].textContent.trim();
            if (t && !/^\d+$/.test(t) && !enrollPattern.test(t) && t.length > 1) {
                courseName = t;
                break;
            }
        }

        // 현재인원/제한인원
        for (let i = codeIdx; i < cells.length; i++) {
            const t = cells[i].textContent.trim();
            if (enrollPattern.test(t)) {
                enrollment = t;
                break;
            }
        }

        // 교수명: 한글 2~4자
        for (let i = codeIdx; i < cells.length; i++) {
            const t = cells[i].textContent.trim();
            if (/^[가-힣]{2,4}$/.test(t)) {
                professor = t;
                break;
            }
        }

        // 시간: 요일+교시 패턴
        for (let i = 0; i < cells.length; i++) {
            const t = cells[i].textContent.trim();
            if (/[월화수목금토일]\d/.test(t)) {
                schedule = t;
                break;
            }
        }

        return { course_code: courseCode, course_name: courseName, enrollment, professor, schedule };
    }

    // ── 플로팅 패널 생성 ──
    const panel = document.createElement('div');
    panel.id = 'sugang-panel';
    panel.innerHTML = `
        <h3>📋 수강 모니터링 - 과목 선택</h3>
        <div class="sg-guide">
            테이블에서 모니터링할 과목의 <b>행을 클릭</b>하세요.<br>
            선택된 행은 <span style="color:#2e7d32;font-weight:bold;">초록색</span>으로 표시됩니다.<br>
            다시 클릭하면 선택이 해제됩니다.
        </div>
        <div id="sg-list"><p class="sg-empty">선택된 과목이 없습니다.</p></div>
        <div class="sg-count" id="sg-count">총 검색된 과목: ${courseRows.length}개</div>
        <button id="sg-start-btn" disabled>모니터링 시작</button>
    `;
    document.body.appendChild(panel);

    const listEl = document.getElementById('sg-list');
    const startBtn = document.getElementById('sg-start-btn');

    // ── 패널 업데이트 ──
    function updatePanel() {
        const courses = window.__selectedCourses;
        if (courses.length === 0) {
            listEl.innerHTML = '<p class="sg-empty">선택된 과목이 없습니다.</p>';
            startBtn.disabled = true;
            startBtn.textContent = '모니터링 시작';
        } else {
            listEl.innerHTML = courses.map((c, i) =>
                `<div class="sg-item">
                    <div class="sg-item-info">
                        <span class="sg-item-code">${c.course_code}</span>
                        ${c.course_name}
                        <span class="sg-item-enroll">${c.enrollment}</span>
                        <br><small style="color:#757575">${c.professor} | ${c.schedule}</small>
                    </div>
                    <button class="sg-remove" data-idx="${i}" title="선택 해제">✕</button>
                </div>`
            ).join('');
            startBtn.disabled = false;
            startBtn.textContent = `모니터링 시작 (${courses.length}개)`;

            // 삭제 버튼 이벤트
            listEl.querySelectorAll('.sg-remove').forEach(btn => {
                btn.addEventListener('click', function(e) {
                    e.stopPropagation();
                    const idx = parseInt(this.dataset.idx);
                    const removed = courses[idx];
                    courses.splice(idx, 1);
                    // 행 하이라이트 해제
                    courseRows.forEach(cr => {
                        const data = extractRowData(Array.from(cr.cells), cr.codeIdx);
                        if (data.course_code === removed.course_code) {
                            cr.row.classList.remove('sugang-selected');
                        }
                    });
                    updatePanel();
                });
            });
        }
    }

    // ── 행 이벤트 바인딩 ──
    courseRows.forEach(({ row, codeIdx, cells }) => {
        const cellArray = Array.from(cells);

        // 호버 효과
        row.addEventListener('mouseenter', () => {
            if (!row.classList.contains('sugang-selected')) {
                row.classList.add('sugang-hover');
            }
        });
        row.addEventListener('mouseleave', () => {
            row.classList.remove('sugang-hover');
        });

        // 클릭으로 선택/해제
        row.addEventListener('click', (e) => {
            e.stopPropagation();  // 중첩 테이블 이벤트 버블링 방지
            // 기존 체크박스나 링크 클릭은 무시
            if (e.target.tagName === 'INPUT' || e.target.tagName === 'A') return;
            // 클릭된 곳이 이 행의 직계 td인지 확인 (다른 행 내부 클릭 무시)
            const clickedTd = e.target.closest('td');
            if (clickedTd && clickedTd.closest('tr') !== row) return;

            const data = extractRowData(cellArray, codeIdx);
            const isSelected = row.classList.contains('sugang-selected');

            if (isSelected) {
                row.classList.remove('sugang-selected');
                window.__selectedCourses = window.__selectedCourses.filter(
                    c => c.course_code !== data.course_code
                );
            } else {
                row.classList.add('sugang-selected');
                window.__selectedCourses.push(data);
            }
            updatePanel();
        });
    });

    // ── 시작 버튼 ──
    startBtn.addEventListener('click', () => {
        if (window.__selectedCourses.length > 0) {
            window.__monitoringStarted = true;
            panel.innerHTML = `
                <h3>🔄 모니터링 준비 중...</h3>
                <p style="font-size:13px;color:#616161;">콘솔로 전환됩니다. 브라우저를 닫지 마세요.</p>
            `;
        }
    });

    return 'OK:' + courseRows.length;
})();
"""

# ---- 페이지 DOM에서 직접 과목 데이터 추출 JS ----
# 새로고침 후 현재 페이지의 테이블에서 바로 읽는 방식
# (XHR/iframe 모두 사이트 보안에 의해 차단됨)

EXTRACT_PAGE_JS = r"""
return (function() {
    try {
    var debug = {
        url: window.location.href,
        title: document.title,
        total_rows: 0,
        rows_with_8plus_cells: 0,
        code_matches: 0,
        all_tables: document.querySelectorAll('table').length,
        all_trs: document.querySelectorAll('table tr').length,
        all_tds: document.querySelectorAll('table tr td').length,
        frames: window.frames.length,
        body_length: (document.body ? document.body.innerHTML.length : 0),
        has_login_form: !!document.querySelector('input[name="p_userid"]'),
        sample_cells: []
    };
    // \uc138\uc158 \ub9cc\ub8cc \uac10\uc9c0
    if (document.querySelector('input[name="p_userid"]')) {
        return JSON.stringify({ error: 'SESSION_EXPIRED', courses: [], debug: debug });
    }
    var results = [];
    var codePattern = /^\d{5,6}-\d{1,2}$/;
    var enrollPattern = /^\d+\s*\/\s*\d+$/;
    var rows = document.querySelectorAll('table tr');
    debug.total_rows = rows.length;
    for (var r = 0; r < rows.length; r++) {
        var row = rows[r];
        var cells = row.querySelectorAll(':scope > td');
        if (cells.length < 8) continue;
        debug.rows_with_8plus_cells++;
        if (debug.sample_cells.length < 3) {
            var sample = [];
            for (var s = 0; s < Math.min(cells.length, 7); s++) {
                sample.push(cells[s].textContent.trim().substring(0, 30));
            }
            debug.sample_cells.push(sample);
        }
        var codeIdx = -1;
        for (var i = 0; i < Math.min(cells.length, 7); i++) {
            if (codePattern.test(cells[i].textContent.trim())) {
                codeIdx = i;
                break;
            }
        }
        if (codeIdx === -1) continue;
        debug.code_matches++;
        var courseCode = cells[codeIdx].textContent.trim();
        var courseName = '';
        for (var j = codeIdx + 1; j < cells.length; j++) {
            var t = cells[j].textContent.trim();
            if (t && !/^\d+$/.test(t) && !enrollPattern.test(t) && t.length > 1) {
                courseName = t;
                break;
            }
        }
        var enrollment = '', current = 0, limit = 0, enrollIdx = -1;
        for (var j = codeIdx; j < cells.length; j++) {
            var t = cells[j].textContent.trim();
            if (enrollPattern.test(t)) {
                enrollment = t;
                enrollIdx = j;
                var parts = t.split('/');
                current = parseInt(parts[0].trim()) || 0;
                limit = parseInt(parts[1].trim()) || 0;
                break;
            }
        }
        var surplus = -1;
        if (enrollIdx >= 0 && enrollIdx + 1 < cells.length) {
            var st = cells[enrollIdx + 1].textContent.trim();
            if (/^\d+$/.test(st)) surplus = parseInt(st);
        }
        var professor = '';
        var profStart = (enrollIdx >= 0) ? enrollIdx + 3 : codeIdx + 4;
        var skipWords = {'\uc804\uccb4':1, '\uc804\uacf5':1, '\uad50\uc591':1, '\uc77c\ubc18':1, '\ubcf5\uc218':1, '\ubd80\uc804\uacf5':1};
        for (var j = profStart; j < cells.length; j++) {
            var t = cells[j].textContent.trim();
            if (/^[\uac00-\ud7a3]{2,4}$/.test(t) && !skipWords[t]) {
                professor = t;
                break;
            }
        }
        var schedule = '';
        for (var j = 0; j < cells.length; j++) {
            var t = cells[j].textContent.trim();
            if (/[\uc6d4\ud654\uc218\ubaa9\uae08\ud1a0\uc77c]\d/.test(t)) {
                schedule = t;
                break;
            }
        }
        var lastCell = cells[cells.length - 1];
        var lastText = lastCell.textContent.trim();
        var hasCheckbox = lastCell.querySelector('input[type="checkbox"]') !== null;
        var isBlocked = lastText === '\uBD88';
        results.push({
            course_code: courseCode,
            course_name: courseName,
            enrollment: enrollment,
            current: current,
            limit: limit,
            surplus: surplus,
            professor: professor,
            schedule: schedule,
            can_register: hasCheckbox,
            is_blocked: isBlocked
        });
    }
    return JSON.stringify({ error: null, courses: results, debug: debug });
    } catch(e) {
        return JSON.stringify({ error: 'JS_ERROR: ' + e.message + ' at line ' + (e.lineNumber || '?') + ' stack: ' + (e.stack || '').substring(0, 200), courses: [], debug: { url: window.location.href, body_length: (document.body ? document.body.innerHTML.length : 0), frames: window.frames.length } });
    }
})();
"""


# ============================================
# 알림 시스템
# ============================================

def alert_sound(repeat: int = 3) -> None:
    """Windows 비프음 알림 (긴급 패턴)"""
    for _ in range(repeat):
        winsound.Beep(1000, 200)
        time.sleep(0.05)
        winsound.Beep(1500, 200)
        time.sleep(0.05)
        winsound.Beep(2000, 400)
        time.sleep(0.2)


def alert_desktop(title: str, message: str) -> None:
    """데스크톱 토스트 알림"""
    if not HAS_PLYER:
        return
    try:
        desktop_notification.notify(
            title=title,
            message=message,
            app_name="수강 모니터링",
            timeout=10,
        )
    except Exception:
        pass


def fire_alerts(vacancies: list[dict[str, Any]], config: dict[str, Any]) -> None:
    """자리 발생 과목에 대해 모든 알림 발동"""
    alert_cfg = config.get("alert", {})

    for course in vacancies:
        code = course["course_code"]
        name = course["course_name"]
        enroll = course["enrollment"]
        prof = course["professor"]

        # 콘솔 강조 출력
        print()
        print("  " + "=" * 56)
        print(f"    🚨  자리 발생!  🚨")
        print(f"    {code}  {name}")
        print(f"    교수: {prof}  |  현재: {enroll}")
        print("  " + "=" * 56)
        print()

        # 데스크톱 알림
        if alert_cfg.get("desktop_notification", True):
            alert_desktop(
                f"🚨 자리 발생! {code}",
                f"{name} ({prof})\n현재: {enroll}",
            )

    # 소리는 한 번만 (여러 과목이라도)
    if vacancies and alert_cfg.get("sound", True):
        repeat = alert_cfg.get("sound_repeat", 3)
        alert_sound(repeat)


# ══════════════════════════════════════════════════════════════
# 콘솔 출력 유틸리티
# ══════════════════════════════════════════════════════════════

def display_width(text: str) -> int:
    """동아시아 전각 문자 폭을 고려한 표시 너비 계산"""
    w = 0
    for ch in text:
        eaw = unicodedata.east_asian_width(ch)
        w += 2 if eaw in ("F", "W") else 1
    return w


def pad_kr(text: str, width: int) -> str:
    """한글 폭을 고려하여 문자열을 고정 너비로 패딩"""
    current = display_width(text)
    if current > width:
        # 잘라내기
        out = []
        w = 0
        for ch in text:
            cw = 2 if unicodedata.east_asian_width(ch) in ("F", "W") else 1
            if w + cw + 1 > width:
                out.append("…")
                break
            out.append(ch)
            w += cw
        text = "".join(out)
        current = display_width(text)
    return text + " " * (width - current)


def print_banner() -> None:
    """시작 배너 출력"""
    print()
    print("  ╔══════════════════════════════════════════════════╗")
    print("  ║      홍익대학교 수강신청 모니터링 스크립트       ║")
    print("  ║      Hongik Univ. Course Availability Monitor    ║")
    print("  ╚══════════════════════════════════════════════════╝")
    print()


def print_status_table(
    targets: list[dict[str, str]],
    all_courses: list[dict[str, Any]],
    check_count: int,
    interval: int,
) -> list[dict[str, Any]]:
    """
    현재 상태를 테이블로 출력하고, 자리 있는 과목 목록을 반환.
    """
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # 화면 클리어
    os.system("cls" if os.name == "nt" else "clear")
    print_banner()
    print(f"  [체크 #{check_count}]  {now}")
    print(f"  다음 체크: {interval}초 후")
    print()

    # 타겟 코드 매핑
    target_codes = {t["course_code"] for t in targets}
    matched = [c for c in all_courses if c["course_code"] in target_codes]

    if not matched:
        print("  ⚠  선택한 과목을 페이지에서 찾을 수 없습니다.")
        print("  검색 결과 페이지가 유지되고 있는지 확인해주세요.")
        if all_courses:
            codes_found = ', '.join(c.get('course_code', '?') for c in all_courses[:5])
            codes_wanted = ', '.join(target_codes)
            extra = f' 외 {len(all_courses)-5}개' if len(all_courses) > 5 else ''
            print(f"  [참고] 서버에서 {len(all_courses)}개 과목 수신, 코드 불일치")
            print(f"    수신: {codes_found}{extra}")
            print(f"    찾는: {codes_wanted}")
        else:
            print("  [참고] 서버 응답에 과목 데이터가 없습니다 (0개).")
        return []

    # 컬럼 너비 정의 (한글 폭 기준)
    W_CODE = 13
    W_NAME = 22
    W_PROF = 8
    W_ENROLL = 10
    W_SURPLUS = 6
    W_STATUS = 16
    total_w = W_CODE + W_NAME + W_PROF + W_ENROLL + W_SURPLUS + W_STATUS + 17  # 구분선/패딩

    sep = "  " + "─" * total_w
    print(sep)
    header = (
        "  │ "
        + pad_kr("학수번호", W_CODE) + "│ "
        + pad_kr("과목명", W_NAME) + "│ "
        + pad_kr("교수", W_PROF) + "│ "
        + pad_kr("현재/제한", W_ENROLL) + "│ "
        + pad_kr("여석", W_SURPLUS) + "│ "
        + pad_kr("상태", W_STATUS) + "│"
    )
    print(header)
    print(sep)

    vacancies: list[dict[str, Any]] = []

    for c in matched:
        current = c.get("current", 0)
        limit = c.get("limit", 0)
        surplus = c.get("surplus", -1)
        can_reg = c.get("can_register", False)
        is_blocked = c.get("is_blocked", False)

        has_vacancy = (current < limit) or can_reg

        if has_vacancy:
            status = "✅ 자리있음!"
            vacancies.append(c)
        elif is_blocked:
            status = "🔴 수강불가"
        else:
            status = "❌ 만석"

        surplus_str = str(surplus) if surplus >= 0 else "-"

        row_str = (
            "  │ "
            + pad_kr(c.get("course_code", ""), W_CODE) + "│ "
            + pad_kr(c.get("course_name", ""), W_NAME) + "│ "
            + pad_kr(c.get("professor", ""), W_PROF) + "│ "
            + pad_kr(c.get("enrollment", ""), W_ENROLL) + "│ "
            + pad_kr(surplus_str, W_SURPLUS) + "│ "
            + pad_kr(status, W_STATUS) + "│"
        )
        print(row_str)

    print(sep)
    print()
    return vacancies


# ============================================
# 세션 / 페이지 상태 확인
# ============================================

def is_login_page(driver: webdriver.Chrome) -> bool:
    """현재 페이지가 로그인 페이지인지 확인"""
    try:
        url = driver.current_url
        if "cn1000.jsp" in url:
            return True
        forms = driver.find_elements(By.NAME, "login")
        fields = driver.find_elements(By.NAME, "p_userid")
        return len(forms) > 0 and len(fields) > 0
    except WebDriverException:
        return False


def has_course_table(driver: webdriver.Chrome) -> bool:
    """현재 페이지에 과목 테이블이 있는지 JS로 빠르게 확인"""
    try:
        result = driver.execute_script("""
            const tds = document.querySelectorAll('table tr td');
            for (const td of tds) {
                if (/^\\d{5,6}-\\d{1,2}$/.test(td.textContent.trim())) return true;
            }
            return false;
        """)
        return bool(result)
    except WebDriverException:
        return False


def wait_for_table(driver: webdriver.Chrome, timeout: int = 15) -> bool:
    """과목 테이블이 로드될 때까지 대기"""
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: has_course_table(d)
        )
        return True
    except TimeoutException:
        return False



def safe_refresh(driver: Any) -> bool:
    """페이지 새로고침 (POST 재전송 확인 대화상자 + alert 자동 처리)"""
    try:
        driver.refresh()
    except UnexpectedAlertPresentException:
        pass
    except WebDriverException:
        return False

    # POST 재전송 확인 alert / 기타 alert 처리
    for _ in range(3):
        time.sleep(0.3)
        try:
            driver.switch_to.alert.accept()
        except Exception:
            break
    return True

# ══════════════════════════════════════════════════════════════
# 메인 흐름
# ══════════════════════════════════════════════════════════════

def main() -> int:
    print_banner()

    # 1) 설정 로드
    config = load_config()
    # 체크 주기 입력 (config.json 값이 기본값)
    saved_interval = config.get("check_interval_seconds", 5)
    user_input = input(f"  체크 주기 (초 단위, 기본값 {saved_interval}): ").strip()
    if user_input:
        try:
            interval = float(user_input)
            if interval < 0.1:
                print("  [경고] 0.1초 미만은 서버 부담 — 0.1초로 설정")
                interval = 0.1
            config["check_interval_seconds"] = interval
        except ValueError:
            print("  [경고] 숫자 아님 — 기본값 사용")
            interval = saved_interval
    else:
        interval = saved_interval
    print(f"  체크 주기: {interval}초")
    if not HAS_PLYER:
        print("  [참고] plyer 미설치 — 데스크톱 알림 비활성화 (소리 알림은 작동)")
    print()

    # 2) 브라우저 시작
    driver = init_driver()

    # 3) 수강신청 사이트로 이동
    print(f"  수강신청 사이트로 이동합니다...")
    try:
        driver.get(SUGANG_URL)
    except WebDriverException as e:
        print(f"  [오류] 사이트 접속 실패: {e}")
        driver.quit()
        return 1

    print()
    print("  ┌──────────────────────────────────────────────────────────┐")
    print("  │                                                          │")
    print("  │   1. 브라우저에서 직접 로그인해주세요                     │")
    print("  │   2. 과목 조회 페이지로 이동해주세요                     │")
    print("  │      (과목별 수강신청 → 학과 선택 → 조회)               │")
    print("  │   3. 과목 목록이 화면에 보이면                           │")
    print("  │      이 콘솔에서 Enter를 눌러주세요                      │")
    print("  │                                                          │")
    print("  └──────────────────────────────────────────────────────────┘")
    print()

    input("  준비가 되면 Enter를 누르세요... ")

    # 4) 과목 테이블 확인
    print("\n  과목 테이블을 탐색 중...")
    if not has_course_table(driver):
        print("  ⚠  과목 테이블을 즉시 찾지 못했습니다. 잠시 대기...")
        if not wait_for_table(driver, timeout=10):
            print("  [경고] 과목 테이블을 찾을 수 없습니다.")
            print("  과목 조회 결과가 있는 페이지인지 다시 확인해주세요.")
            input("  확인 후 Enter... ")
            if not wait_for_table(driver, timeout=5):
                print("  [오류] 과목 테이블을 여전히 찾을 수 없습니다.")
                driver.quit()
                return 1

    print("  ✅ 과목 테이블 발견!")

    # 5) 모니터링 URL 저장
    monitoring_url = driver.current_url
    print(f"  모니터링 URL: {monitoring_url}")

    # 6) 인터랙티브 과목 선택 UI 주입
    print("\n  과목 선택 UI를 주입합니다...")
    try:
        result = driver.execute_script(PICKER_JS)
        if result and str(result).startswith("NO_TABLE"):
            print("  [오류] JS에서 과목 테이블을 찾지 못했습니다.")
            driver.quit()
            return 1
        if result:
            print(f"  ✅ 과목 선택 UI 활성화 ({result})")
    except WebDriverException as e:
        print(f"  [오류] UI 주입 실패: {e}")
        driver.quit()
        return 1

    print()
    print("  ┌──────────────────────────────────────────────────────────┐")
    print("  │                                                          │")
    print("  │   브라우저에서 모니터링할 과목의 행을 클릭하세요         │")
    print("  │   → 행 위에 마우스를 올리면 파란색 테두리                │")
    print("  │   → 클릭하면 초록색으로 선택됨                          │")
    print("  │   → 다시 클릭하면 선택 해제                             │")
    print("  │                                                          │")
    print("  │   선택 완료 후 우측 상단 패널의                          │")
    print("  │   [모니터링 시작] 버튼을 클릭하세요                      │")
    print("  │                                                          │")
    print("  └──────────────────────────────────────────────────────────┘")
    print()

    # 7) 사용자가 과목 선택 완료 & 시작 버튼 클릭 대기
    selected_courses: list[dict[str, str]] = []
    print("  사용자 선택 대기 중... (브라우저에서 과목 선택 후 '모니터링 시작' 클릭)")
    while True:
        time.sleep(0.5)
        try:
            started = driver.execute_script("return window.__monitoringStarted === true;")
            if started:
                raw = driver.execute_script("return JSON.stringify(window.__selectedCourses);")
                selected_courses = json.loads(raw) if raw else []
                break
        except WebDriverException:
            print("\n  [오류] 브라우저 연결이 끊어졌습니다.")
            return 1

    if not selected_courses:
        print("  [오류] 선택된 과목이 없습니다.")
        driver.quit()
        return 1

    print(f"\n  ✅ {len(selected_courses)}개 과목 선택 완료:")
    for c in selected_courses:
        print(f"     • {c['course_code']}  {c['course_name']}  ({c['professor']})")

    print(f"\n  ────────────────────────────────────")
    print(f"  모니터링 시작!  (Ctrl+C로 중지)")
    print(f"  체크 주기: {interval}초")
    print(f"  방식: 페이지 새로고침 + DOM 직접 읽기")
    print(f"  ────────────────────────────────────")
    time.sleep(1)
    # 8) 모니터링 루프 (새로고침 기반)
    check_count = 0
    consecutive_errors = 0
    MAX_CONSECUTIVE_ERRORS = 10
    try:
        while True:
            check_count += 1
            try:
                # 첫 체크는 현재 페이지 그대로, 이후는 새로고침
                if check_count > 1:
                    if not safe_refresh(driver):
                        consecutive_errors += 1
                        print(f"\n  [경고] 새로고침 실패 (#{consecutive_errors})")
                        time.sleep(interval)
                        continue
                    if not wait_for_table(driver, timeout=15):
                        consecutive_errors += 1
                        print(f"\n  [경고] 테이블 로드 타임아웃 (#{consecutive_errors})")
                        time.sleep(interval)
                        continue

                # 세션 만료 확인
                if is_login_page(driver):
                    print("\n  ⚠  세션이 만료되었습니다!")
                    print("  브라우저에서 다시 로그인 후 과목 조회 페이지로 이동해주세요.")
                    input("  준비가 되면 Enter... ")
                    consecutive_errors = 0
                    continue

                # DOM에서 과목 데이터 추출
                # 프레임셋 감지 + 올바른 프레임으로 전환
                frame_info = driver.execute_script("""
                    var frames = document.querySelectorAll('frame');
                    var iframes = document.querySelectorAll('iframe');
                    var names = [];
                    for (var i = 0; i < frames.length; i++) {
                        names.push('frame:' + (frames[i].name || frames[i].id || i));
                    }
                    for (var i = 0; i < iframes.length; i++) {
                        names.push('iframe:' + (iframes[i].name || iframes[i].id || i));
                    }
                    return JSON.stringify({
                        frames: frames.length,
                        iframes: iframes.length,
                        names: names,
                        url: window.location.href,
                        has_table: document.querySelectorAll('table tr td').length > 0
                    });
                """)
                frame_data = json.loads(frame_info) if frame_info else {}
                total_frames = frame_data.get('frames', 0) + frame_data.get('iframes', 0)

                if total_frames > 0 and not frame_data.get('has_table', False):
                    # 프레임셋 사용 중 — 올바른 프레임 찾기
                    switched = False
                    all_frame_elements = driver.find_elements(By.TAG_NAME, 'frame') + driver.find_elements(By.TAG_NAME, 'iframe')
                    for idx, frame_el in enumerate(all_frame_elements):
                        try:
                            driver.switch_to.frame(frame_el)
                            has_tbl = driver.execute_script("""
                                var tds = document.querySelectorAll('table tr td');
                                for (var i = 0; i < tds.length; i++) {
                                    if (/^\\d{5,6}-\\d{1,2}$/.test(tds[i].textContent.trim())) return true;
                                }
                                return false;
                            """)
                            if has_tbl:
                                switched = True
                                frame_name = frame_el.get_attribute('name') or frame_el.get_attribute('id') or str(idx)
                                break
                            driver.switch_to.default_content()
                        except WebDriverException:
                            driver.switch_to.default_content()
                            continue
                    if not switched:
                        driver.switch_to.default_content()
                raw_result = driver.execute_script(EXTRACT_PAGE_JS)
                result: dict[str, Any] = json.loads(raw_result) if raw_result else {}
                # 프레임 사용 시 default_content로 복귀
                if total_frames > 0:
                    driver.switch_to.default_content()
                error = result.get("error")
                if error == "SESSION_EXPIRED":
                    print("\n  ⚠  세션이 만료되었습니다!")
                    print("  브라우저에서 다시 로그인 후 과목 조회 페이지로 이동해주세요.")
                    input("  준비가 되면 Enter... ")
                    consecutive_errors = 0
                    continue
                if error:
                    consecutive_errors += 1
                    print(f"\n  [경고] 데이터 추출 오류 (#{consecutive_errors}): {error}")
                    if consecutive_errors >= MAX_CONSECUTIVE_ERRORS:
                        print("  서버 상태를 확인해주세요.")
                        input("  Enter를 눌러 재시도... ")
                        consecutive_errors = 0
                    time.sleep(interval)
                    continue

                # 정상 응답
                consecutive_errors = 0
                all_courses: list[dict[str, Any]] = result.get("courses", [])
                vacancies = print_status_table(
                    selected_courses, all_courses, check_count, interval
                )
                # 자리 발생 시 알림
                if vacancies:
                    fire_alerts(vacancies, config)
            except UnexpectedAlertPresentException:
                # 새로고침 중 alert 발생 — 무시하고 재시도
                try:
                    driver.switch_to.alert.accept()
                except Exception:
                    pass
            except WebDriverException as e:
                err_msg = str(e)
                if "no such window" in err_msg or "not connected" in err_msg:
                    print("\n  [오류] 브라우저가 닫혔습니다.")
                    break
                consecutive_errors += 1
                print(f"\n  [경고] 브라우저 오류 (#{consecutive_errors}): {err_msg[:100]}")
                if consecutive_errors >= MAX_CONSECUTIVE_ERRORS:
                    print(f"  [오류] 연속 {MAX_CONSECUTIVE_ERRORS}회 오류 발생.")
                    print("  문제 해결 후 Enter를 눌러주세요.")
                    input("  Enter... ")
                    consecutive_errors = 0
            except json.JSONDecodeError:
                print("  [경고] 데이터 파싱 오류. 재시도합니다...")
            time.sleep(interval)
    except KeyboardInterrupt:
        pass

    # 종료 요약
    print()
    print("  ┌──────────────────────────────────────────┐")
    print(f"  │  모니터링 종료                            │")
    print(f"  │  총 체크 횟수: {check_count:<25}│")
    print(f"  │  {datetime.now().strftime('%Y-%m-%d %H:%M:%S'):<39}│")
    print("  └──────────────────────────────────────────┘")
    print()

    try:
        driver.quit()
    except Exception:
        pass

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
