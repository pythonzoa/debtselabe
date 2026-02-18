# 패치 노트 v2 — 전체 수정 이력

---

## [수정 1] AppError → Error 상속 (`Code.gs`)

**문제** `class AppError`가 일반 객체라 `catch(e)`에서 `e instanceof AppError` 체크가 실패하고 스택 트레이스가 남지 않았습니다.

```js
// 전
class AppError {
  constructor(message, code, details) {
    this.message = message; ...
  }
}

// 후
class AppError extends Error {
  constructor(message, code, details) {
    super(message);
    this.name = 'AppError'; ...
  }
}
```

---

## [수정 2] CacheService 적용 (`Code.gs`)

**문제** `getAppConfig()`가 호출될 때마다 `__ADMIN_CONFIG__` 시트를 직접 읽어 불필요한 I/O 발생.

**변경** `CacheService.getScriptCache()`로 10분 캐싱 적용.
- 관리자 추가(`addAdmin`), 필터 추가(`addAllowedFilter`) 시 `invalidateConfigCache()`로 즉시 무효화
- 캐시 파싱 실패 시 시트에서 재조회하는 fallback 포함

> **주의** 관리자 이메일 변경 후 반드시 Apps Script 에디터에서 `invalidateConfigCache` 실행 필요

---

## [수정 3] LOG 시트 기록 (`Code.gs`)

**문제** `logActivity()`가 `Logger.log`만 사용해 이력이 휘발됨. 누가 언제 발행했는지 나중에 확인 불가.

**변경** `__ACTIVITY_LOG__` 숨겨진 시트에 행 append.

| 컬럼 | 내용 |
|------|------|
| timestamp | ISO 8601 형식 |
| user | 실행한 계정 이메일 |
| action | 수행한 작업명 |
| details | JSON 상세 정보 |

> **확인 방법** 스프레드시트 하단 시트 탭 우클릭 → 숨겨진 시트 표시 → `__ACTIVITY_LOG__`

---

## [수정 4] XSS 패치 (`Index.html`)

**문제** `FilterPresets.renderList()`에서 프리셋 이름이 `onclick` 인라인 핸들러에 직접 삽입됨.

```js
// 취약한 코드 (전)
`<span onclick="FilterPresets.load('${name}')">`
// name = '); alert('XSS') 입력 시 임의 JS 실행 가능

// 안전한 코드 (후)
const nameSpan = document.createElement('span');
nameSpan.textContent = name;  // 어떤 문자도 텍스트로만 처리
nameSpan.addEventListener('click', () => this.load(name));
```

---

## [수정 5] publishData — getValues → getDisplayValues (`Code.gs`)

**문제** `publishData()`에서 `getValues()`로 데이터를 읽으면 날짜 셀이 Date 객체로 변환됨. 이를 `setValues()`로 `__DASHBOARD_DB__`에 쓰면 한국어 로케일 기준으로 `2027. 5. 9.` 형식으로 재렌더링되어 날짜 검증 오류 발생.

```js
// 전
const sourceData = sourceSheet.getDataRange().getValues();

// 후
const sourceData = sourceSheet.getDataRange().getDisplayValues();
// 원본에 표시된 문자열(2027-05-09) 그대로 복사
```

---

## [수정 6] SYSTEM_ADMINS 권한 분리 (`Code.gs`)

**문제** `getAppConfig()`가 `SYSTEM_ADMINS` 상수와 시트 A열을 합쳐서 관리자 목록을 만들었음. 시트에서 이메일을 바꿔도 `SYSTEM_ADMINS`에 있는 이메일은 항상 관리자로 남는 문제.

```js
// 전 — 상수와 시트를 항상 합산
const uniqueAdmins = [...new Set([
  ...SYSTEM_ADMINS.map(e => e.toLowerCase()),
  ...admins  // 시트 A열
])];

// 후 — 시트 A열이 유일한 소스
//      시트 A열이 비어있는 비상 상황에서만 SYSTEM_ADMINS로 fallback
const adminsSource = admins.length > 0
  ? admins
  : SYSTEM_ADMINS.map(e => e.toLowerCase());
```

> **운영 방법** `__ADMIN_CONFIG__` 시트 A열에서만 관리자 이메일 추가/변경. A1은 헤더(`Admin ID`), A2부터 이메일 입력.

---

## [수정 7] 날짜 검증 정규식 확장 (`Code.gs`)

**문제** `validateData()`의 날짜 검증이 `YYYY-MM-DD` 형식만 허용. `__DASHBOARD_DB__` 스냅샷에 저장된 `2027. 5. 9.` 형식을 읽지 못해 전체 데이터가 만기일 누락으로 오탐.

```js
// 전 — YYYY-MM-DD 형식만 허용
if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {

// 후 — YYYY-MM-DD 와 YYYY. M. D. 형식 모두 허용
if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr) &&
    !/^\d{4}\.\s*\d{1,2}\.\s*\d{1,2}\.?$/.test(dateStr)) {
```

---

## 파일 목록

| 파일 | 수정 항목 |
|------|----------|
| `Code.gs` | 수정 1~3, 5~7 |
| `Index.html` | 수정 4 |

---

## 적용 체크리스트

- [ ] `Code.gs` 전체를 Apps Script 에디터에 붙여넣기
- [ ] `Index.html` 전체를 Apps Script 에디터에 붙여넣기
- [ ] Apps Script 에디터에서 `invalidateConfigCache` 실행
- [ ] Apps Script 에디터에서 `publishData` 실행 (스냅샷 재생성)
- [ ] 대시보드 새로고침 후 무결성 검증 확인
