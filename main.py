from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware


app = FastAPI(
    title="고교학점제 교과편성 최적화 API",
    version="0.1.0",
)

# 배포 연결 시험 단계에서는 모든 출처를 허용합니다.
# GitHub Pages 주소가 확정되면 해당 주소만 허용하도록 변경합니다.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["*"],
)


@app.get("/")
def root():
    return {
        "service": "course-group-planner-api",
        "status": "ready",
        "message": "교과편성 최적화 서버가 정상 작동 중입니다.",
    }


@app.get("/health")
def health():
    return {"status": "ok"}

