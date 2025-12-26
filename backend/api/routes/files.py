from fastapi import APIRouter, UploadFile, File, HTTPException, Query
import os
import shutil
from typing import List
from ..controller import controller
import logging

router = APIRouter()
MEDIA_DIR = "media"

# Ensure media directory exists
if not os.path.exists(MEDIA_DIR):
    os.makedirs(MEDIA_DIR)

@router.get("/")
async def list_files():
    """List all audio files in the media directory."""
    try:
        files = []
        # DEBUG: correct absolute path
        abs_media_path = os.path.abspath(MEDIA_DIR)
        print(f"[Files] Scanning directory: {abs_media_path}")

        for filename in os.listdir(MEDIA_DIR):
            file_path = os.path.join(MEDIA_DIR, filename)
            if os.path.isfile(file_path):
                stats = os.stat(file_path)
                files.append({
                    "id": filename, # Use filename as ID for simplicity
                    "name": filename,
                    "size": f"{stats.st_size / 1024 / 1024:.2f} MB",
                    "date": "Local File", # Simplify date or use stats.st_mtime
                    "type": "audio/mpeg", # Generic or guess type
                    "url": f"/media/{filename}"
                })
        return files
    except Exception as e:
        logging.error(f"Error listing files: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/upload")
async def upload_file(file: UploadFile = File(...), user: str = Query("Unknown")):
    """Upload an audio file to the media directory."""
    try:
        file_location = os.path.join(MEDIA_DIR, file.filename)
        
        # Check if file exists
        if os.path.exists(file_location):
            # Allow overwrite or rename? For now, error or overwrite. 
            # AppContext logic usually checks dupes, but backend is safer.
            pass

        with open(file_location, "wb+") as file_object:
            shutil.copyfileobj(file.file, file_object)
            
        logging.info(f"User {user} uploaded file: {file.filename}")
        
        stats = os.stat(file_location)
        return {
            "id": file.filename,
            "name": file.filename,
            "size": f"{stats.st_size / 1024 / 1024:.2f} MB",
            "url": f"/media/{file.filename}"
        }
    except Exception as e:
        logging.error(f"Upload failed: {e}")
        raise HTTPException(status_code=500, detail=f"Upload failed: {str(e)}")

@router.delete("/{filename}")
async def delete_file(filename: str, user: str = Query("Unknown")):
    """Delete a file from the media directory."""
    try:
        file_path = os.path.join(MEDIA_DIR, filename)
        if os.path.exists(file_path):
            os.remove(file_path)
            logging.info(f"User {user} deleted file: {filename}")
            return {"status": "success", "message": f"File {filename} deleted"}
        else:
            raise HTTPException(status_code=404, detail="File not found")
    except Exception as e:
        logging.error(f"Delete failed: {e}")
        raise HTTPException(status_code=500, detail=str(e))
