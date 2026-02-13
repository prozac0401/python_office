# python_office

`python_office` 프로젝트의 기본 사용 방법을 정리한 문서입니다.

## 1. 개발 환경 준비

### 요구 사항
- Python 3.10 이상
- Git

### 저장소 클론
```bash
git clone <REPOSITORY_URL>
cd python_office
```

### 가상환경 생성 및 활성화
```bash
python -m venv .venv
source .venv/bin/activate
```

### 의존성 설치
의존성 파일이 존재하는 경우 아래 명령으로 설치합니다.
```bash
pip install -r requirements.txt
```

## 2. 프로젝트 실행

애플리케이션 진입점 파일(예: `main.py` 또는 패키지 모듈)이 준비되어 있다면 다음과 같이 실행합니다.

```bash
python main.py
# 또는
python -m <package_name>
```

## 3. 테스트 실행

테스트 코드가 포함되어 있다면 다음 명령으로 실행합니다.

```bash
pytest -q
```

## 4. 권장 개발 워크플로

1. 기능 브랜치 생성
2. 코드 수정
3. 테스트 실행
4. 커밋 및 Pull Request 생성

## 5. 현재 상태

현재 저장소는 초기 구성 단계로 보이며, 상세 실행 방법은 실제 소스 코드/의존성 파일이 추가되면 함께 업데이트하는 것을 권장합니다.
