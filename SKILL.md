---
name: presentation-generator
description: 템플릿(.pptx)과 데이터(JSON/Excel/CSV)를 기반으로 PowerPoint 프레젠테이션을 자동 생성합니다. 템플릿 디자인을 분석하고, 더 나은 도형·구도가 있으면 개선하여 적용합니다.
---

# Presentation Generator Skill

## 개요

이 Skill은 PowerPoint 템플릿(`.pptx`)과 구조화된 데이터(JSON/Excel/CSV)를 입력으로 받아
프레젠테이션을 자동 생성합니다. 단순히 템플릿에 데이터를 삽입하는 것이 아니라,
**디자인 개선 엔진**을 통해 더 좋은 도형, 구도, 차트 유형이 있으면 자동으로 변경하여 적용합니다.

## 사전 요구사항

```bash
pip install python-pptx openpyxl
```

## 사용 방법

### 1단계: 템플릿 분석 (선택사항)

먼저 템플릿의 구조를 분석하여 어떤 슬라이드 레이아웃과 플레이스홀더가 있는지 확인합니다.

```bash
python scripts/template_analyzer.py <template.pptx> [--output analysis.json]
```

분석 결과를 보고 데이터 파일의 키를 매핑할 수 있습니다.

### 2단계: 데이터 준비

데이터는 다음 JSON 형식을 따릅니다 (예시: `examples/example_data.json` 참고):

```json
{
  "metadata": {
    "title": "프레젠테이션 제목",
    "author": "작성자",
    "date": "2026-03-25"
  },
  "slides": [
    {
      "layout": "Title Slide",
      "content": {
        "title": "메인 타이틀",
        "subtitle": "서브 타이틀"
      }
    },
    {
      "layout": "Content",
      "content": {
        "title": "슬라이드 제목",
        "body": "본문 텍스트 또는 불릿 포인트 리스트",
        "bullets": ["항목 1", "항목 2", "항목 3"]
      }
    },
    {
      "layout": "Chart",
      "content": {
        "title": "매출 추이",
        "chart": {
          "type": "auto",
          "categories": ["Q1", "Q2", "Q3", "Q4"],
          "series": [
            {"name": "2025", "values": [100, 150, 130, 200]},
            {"name": "2026", "values": [120, 180, 160, 240]}
          ]
        }
      }
    },
    {
      "layout": "Table",
      "content": {
        "title": "비교표",
        "table": {
          "headers": ["항목", "Plan A", "Plan B"],
          "rows": [
            ["가격", "100만원", "200만원"],
            ["기능", "기본", "프리미엄"]
          ]
        }
      }
    },
    {
      "layout": "SmartArt",
      "content": {
        "title": "프로세스 흐름",
        "smartart": {
          "type": "process",
          "items": ["기획", "디자인", "개발", "테스트", "배포"]
        }
      }
    }
  ],
  "design_preferences": {
    "enhance": true,
    "color_scheme": "auto",
    "prefer_charts_over_tables": true
  }
}
```

Excel/CSV 데이터도 지원합니다. 자세한 형식은 `examples/data_format_guide.md`를 참고하세요.

### 3단계: 프레젠테이션 생성

```bash
python scripts/generate_presentation.py \
  --template <template.pptx> \
  --data <data.json> \
  --output <output.pptx> \
  [--enhance]
```

**주요 옵션:**
| 옵션 | 설명 | 기본값 |
|------|------|--------|
| `--template` | 템플릿 `.pptx` 파일 경로 | (필수) |
| `--data` | 데이터 파일 경로 (JSON/Excel/CSV) | (필수) |
| `--output` | 출력 `.pptx` 파일 경로 | `output.pptx` |
| `--enhance` | 디자인 자동 개선 활성화 | 비활성 |
| `--analyze-only` | 템플릿 분석만 수행 | 비활성 |

## 디자인 개선 엔진

`--enhance` 플래그가 활성화되면 다음 개선이 자동 적용됩니다:

1. **도형 배치 최적화** — 불규칙한 정렬을 그리드 기반으로 균등 배치
2. **차트 유형 자동 선택** — 데이터 패턴(추세·비교·분포)에 따라 최적 차트 추천
3. **SmartArt 스타일 도형** — 프로세스(→), 순환(⟲), 계층(▽) 구조를 시각적 도형 그룹으로 표현
4. **텍스트 오버플로 방지** — 콘텐츠 길이에 따라 폰트 크기·텍스트 박스 자동 조정
5. **색상 일관성** — 템플릿 테마 색상 기반으로 통일된 컬러 팔레트 적용
6. **시각적 밸런스** — 슬라이드 여백·공간 분배 최적화

## Agent에서 사용 시

Agent가 이 Skill을 사용할 때는 다음 순서를 따릅니다:

1. 사용자로부터 **템플릿 파일**과 **데이터**(파일 또는 텍스트)를 받습니다.
2. 데이터가 텍스트인 경우 JSON 형식으로 변환하여 임시 파일로 저장합니다.
3. `template_analyzer.py`로 템플릿을 분석합니다.
4. 분석 결과를 참고하여 데이터의 키를 적절히 매핑합니다.
5. `generate_presentation.py --enhance`로 프레젠테이션을 생성합니다.
6. 완성된 `.pptx` 파일 경로를 사용자에게 알려줍니다.
