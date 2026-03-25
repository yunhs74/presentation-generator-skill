# 데이터 형식 가이드

이 문서는 Presentation Generator Skill에서 사용하는 데이터 형식을 설명합니다.

## JSON 형식

### 전체 구조

```json
{
  "metadata": { ... },
  "slides": [ ... ],
  "design_preferences": { ... },
  "keep_existing_slides": false
}
```

### metadata (선택)

| 필드 | 설명 |
|------|------|
| `title` | 프레젠테이션 제목 (문서 속성) |
| `author` | 작성자 |
| `date` | 날짜 |

### slides (필수)

각 슬라이드 객체의 구조:

```json
{
  "layout": "슬라이드 유형",
  "content": { ... }
}
```

#### 지원하는 layout 유형

| Layout | 설명 | content 필드 |
|--------|------|-------------|
| `Title Slide` | 타이틀 슬라이드 | `title`, `subtitle` |
| `Content` | 텍스트/불릿 슬라이드 | `title`, `body` 또는 `bullets` |
| `Chart` | 차트 슬라이드 | `title`, `chart` |
| `Table` | 테이블 슬라이드 | `title`, `table` |
| `SmartArt` | SmartArt 스타일 | `title`, `smartart` |
| `Image` | 이미지 슬라이드 | `title`, `image_path` |

#### chart 객체

```json
{
  "type": "auto",
  "categories": ["Q1", "Q2", "Q3", "Q4"],
  "series": [
    {"name": "매출", "values": [100, 200, 150, 300]}
  ]
}
```

**type 옵션:** `auto`, `bar`, `column`, `line`, `pie`, `doughnut`, `area`, `scatter`, `radar`, `stacked_bar`

> `auto`를 사용하면 데이터 특성에 따라 최적 차트가 자동 선택됩니다.

#### table 객체

```json
{
  "headers": ["항목", "값"],
  "rows": [
    ["A", "100"],
    ["B", "200"]
  ]
}
```

#### smartart 객체

```json
{
  "type": "process",
  "items": ["단계1", "단계2", "단계3"]
}
```

**type 옵션:**
- `process` — 프로세스 흐름 (→)
- `cycle` — 순환 구조 (⟲)
- `hierarchy` — 계층 구조 (피라미드 형태)
- `comparison` — 좌우 비교 (2개 항목)

### design_preferences (선택)

| 필드 | 설명 | 기본값 |
|------|------|--------|
| `enhance` | 디자인 자동 개선 | `false` |
| `color_scheme` | 색상 테마 (`auto`, `professional`, `modern`, `minimal`) | `auto` |
| `prefer_charts_over_tables` | 가능하면 차트 우선 사용 | `false` |

---

## Excel 형식

- 각 **시트**가 하나의 **테이블 슬라이드**로 변환됩니다
- 첫 행은 **헤더**로 처리됩니다
- 시트 이름이 슬라이드 **제목**이 됩니다

## CSV 형식

- 첫 행은 **헤더**로 처리됩니다
- 파일 이름(확장자 제외)이 슬라이드 **제목**이 됩니다
- 전체 데이터가 하나의 **테이블 슬라이드**로 변환됩니다
