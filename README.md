# Presentation Generator Skill

템플릿(.pptx)과 데이터(JSON/Excel/CSV)를 기반으로 PowerPoint 프레젠테이션을 자동 생성하는 Skill.
템플릿 디자인을 분석하고, 더 나은 도형·구도가 있으면 개선하여 적용합니다.

## 주요 기능

- **템플릿 분석**: 슬라이드 레이아웃, 플레이스홀더, 도형, 색상, 폰트 분석
- **데이터 매핑**: JSON/Excel/CSV 데이터를 슬라이드에 자동 매핑
- **디자인 개선 엔진**:
  - 도형 배치 최적화 (그리드 정렬, 균등 분배)
  - 차트 유형 자동 선택 (데이터 패턴 분석)
  - SmartArt 스타일 도형 (프로세스, 순환, 계층, 비교)
  - 텍스트 오버플로 방지
  - 색상 일관성 적용
  - 시각적 밸런스 개선

## 설치

```bash
pip install python-pptx openpyxl
```

## 사용법

```bash
# 템플릿 분석
python scripts/template_analyzer.py template.pptx

# 프레젠테이션 생성 (디자인 개선 포함)
python scripts/generate_presentation.py -t template.pptx -d data.json -o output.pptx --enhance
```

자세한 내용은 [SKILL.md](SKILL.md)를 참고하세요.
