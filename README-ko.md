<p align="right">
  [한국어]
  [<a href="README.md">English</a>]
</p>

# Excel2Image

`Excel2Image`는 Excel 시트를 이미지 파일(JPG)로 변환하는 Python 스크립트입니다. 사용자는 Excel 파일의 데이터를 이미지 형식으로 생성하여 보고서, 프레젠테이션, 아카이빙 등 다양한 용도로 사용할 수 있습니다. 이 스크립트는 각 시트의 콘텐츠에 따라 출력 이미지의 크기를 동적으로 조정합니다.

## 기능:
- Excel 파일의 각 시트를 JPG 이미지로 변환합니다.
- macOS, Windows, Linux에서 한국어 폰트를 지원합니다.
- 시트 콘텐츠에 따라 출력 이미지의 크기를 동적으로 조정합니다.
- 병합된 셀 및 텍스트 정렬을 처리합니다.

## 요구 사항:
- Python 3
- `openpyxl` 라이브러리 (Excel 파일 읽기용)
- `matplotlib` 라이브러리 (이미지 생성용)
- `PIL` (Pillow) 라이브러리 (이미지 저장용)

## 설치:

필요한 라이브러리를 설치합니다: `pip3 install -r requirements.txt`

## 사용법:

명령줄에서 스크립트를 실행합니다:

- `<excel_file_path>`: Excel 파일 경로.
- `<output_directory>`: 이미지 파일을 저장할 디렉토리 경로.

예시: excel2image sample.xlsx .

## 라이선스:
이 프로젝트는 MIT 라이선스 하에 제공됩니다.