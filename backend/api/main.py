from fastapi import APIRouter


main_router = APIRouter()

@main_router.get("/")
def main_router_root():
    return {"message": "Welcome to the main API route"}