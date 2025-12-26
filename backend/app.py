from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from api.routes.auth import auth_router
from api.routes.realtime import real_time_announcements_router 
from api.routes.scheduled import scheduled_announcements_router

from api.routes.account import manage_account_router
from api.routes.emergency import emergency_route
from api.routes.files import router as files_router
from fastapi.staticfiles import StaticFiles
import os

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Allow all origins for development
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/")
def read_root():
    return {"Hello": "World"}


app.include_router(auth_router)
app.include_router(real_time_announcements_router)
app.include_router(scheduled_announcements_router)
app.include_router(manage_account_router)
app.include_router(emergency_route)
app.include_router(files_router, prefix="/files", tags=["Files"])

# Mount Media Directory
if not os.path.exists("media"):
    os.makedirs("media")
app.mount("/media", StaticFiles(directory="media"), name="media")

# Graceful Shutdown
from contextlib import asynccontextmanager
from api.audio_service import audio_service

@asynccontextmanager
async def lifespan(app: FastAPI):
    # Startup
    yield
    # Shutdown
    print("[LifeSpan] Shutting down audio service...")
    try:
        audio_service.stop()
    except Exception as e:
        # Suppress CancelledError or other cleanup errors during forced shutdown
        print(f"[LifeSpan] Cleanup skipped or failed: {e}")
