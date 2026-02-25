# SugangMacro

홍익대학교 수강신청 사이트에서 원하는 과목의 빈자리를 자동으로 감지하는 모니터링 도구.

## 기능
- 수강신청 페이지에서 과목 자동 모니터링
- 빈자리 발생 시 소리 + 데스크톱 알림
- 브라우저에서 직접 과목 선택 (인터랙티브 UI)
- 체크 주기 설정 가능 (0.1초 단위)

## 설치

### 요구사항
- Python 3.8+
- Chrome 브라우저
- Windows

### 패키지 설치
```bash
pip install -r requirements.txt
```

## 사용법
1. 실행
```bash
python monitor.py
# 또는
실행.bat
```
2. 브라우저에서 홍익대 수강신청 사이트 로그인
3. 과목 조회 페이지로 이동 후 콘솔에서 Enter
4. 브라우저에서 모니터링할 과목 클릭 → "모니터링 시작"
5. 빈자리 발생 시 알림

## 설정 (config.json)
```json
{
  "check_interval_seconds": 0.5,
  "alert": {
    "sound": true,
    "desktop_notification": true,
    "sound_repeat": 3
  }
}
```

## 기술 스택
- Selenium — 브라우저 자동화
- webdriver-manager — 드라이버 자동 관리
- plyer — 데스크톱 알림

## 주의사항
- 로그인은 수동 (자동 로그인 미지원)
- 자동 수강신청은 불가 (모니터링만)
- Windows 전용 (소리 알림에 winsound 사용)
