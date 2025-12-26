from fastapi import APIRouter, HTTPException
from typing import Optional, List
from pydantic import BaseModel
from api.firebaseConfig import db, firestore_server_timestamp
from api.controller import controller, Task, TaskType, Priority

real_time_announcements_router = APIRouter(
    prefix="/realtime",
    tags=["Real Time Announcements"]
)

class BroadcastRequest(BaseModel):
    user: str
    zones: List[str]
    type: str = "voice" # 'voice' or 'text'
    content: Optional[str] = None # Text content or encoded metadata

class BroadcastAction(BaseModel):
    user: str
    type: str # 'voice' or 'text'
    action: str # 'START', 'STOP', 'MESSAGE'
    details: str
    timestamp: Optional[str] = None

@real_time_announcements_router.post("/start")
def start_broadcast(req: BroadcastRequest):
    """
    Request to start a Live Broadcast (Voice or Text) or Background Audio.
    Verified by PA Controller.
    """
    # Determine Priority and Type
    if req.type == 'background':
        task_type = TaskType.BACKGROUND
        priority = Priority.BACKGROUND
    elif req.type == 'voice':
        task_type = TaskType.VOICE
        priority = Priority.REALTIME
    else:
        task_type = TaskType.TEXT
        priority = Priority.REALTIME

    task = Task(
        type=task_type,
        priority=priority,
        data={
            "user": req.user,
            "zones": req.zones,
            "content": req.content
        }
    )
    
    success = controller.request_playback(task)
    if not success:
        raise HTTPException(status_code=409, detail="System Busy or Higher Priority Active")
    
    return {"message": "Broadcast Started", "task_id": task.id}

@real_time_announcements_router.post("/stop")
def stop_broadcast(user: str, type: str = "voice", task_id: Optional[str] = None): 
    """
    Request to stop the current broadcast.
    Type can be 'voice', 'text', 'background'
    """
    target_type = TaskType.VOICE
    if type == 'background':
        target_type = TaskType.BACKGROUND
    elif type == 'text':
        target_type = TaskType.TEXT
    
    controller.stop_task(task_id, task_type=target_type, user=user)
    return {"message": "Broadcast Stopped"}

class CompleteRequest(BaseModel):
    task_id: str

@real_time_announcements_router.post("/complete")
def complete_task(req: CompleteRequest):
    """
    Signal that a task (e.g. Schedule playback) has finished.
    """
    controller.stop_task(req.task_id, user="System")
    return {"message": "Task Completed"}

@real_time_announcements_router.post("/log")
def log_broadcast(action: BroadcastAction):
    # Log history only
    try:
        log_entry = action.dict()
        log_entry["timestamp"] = firestore_server_timestamp()
        update_time, doc_ref = db.collection("logs").add(log_entry)
        return {"message": "Logged successfully", "id": doc_ref.id}
    except Exception as e:
        print(f"Logging failed: {e}")
        return {"message": "Logged (fallback)", "id": None}

@real_time_announcements_router.get("/logs")
def get_logs():
    try:
        from firebase_admin import firestore
        docs = db.collection("logs").order_by("timestamp", direction=firestore.Query.DESCENDING).limit(50).stream()
        logs = []
        for doc in docs:
            data = doc.to_dict()
            data["id"] = doc.id
            if "timestamp" in data and data["timestamp"]:
                ts = data["timestamp"]
                if hasattr(ts, 'isoformat'):
                    data["timestamp"] = ts.isoformat()
                else:
                    data["timestamp"] = str(ts)
            logs.append(data)
        return logs
    except Exception as e:
        print(f"Fetch logs failed: {e}")
        return []

class LogUpdate(BaseModel):
    action: str = None
    details: str = None

@real_time_announcements_router.put("/log/{log_id}")
def update_log(log_id: str, update: LogUpdate):
    try:
        doc_ref = db.collection("logs").document(log_id)
        if not doc_ref.get().exists:
            raise HTTPException(status_code=404, detail="Log not found")
        fields_to_update = {k: v for k, v in update.dict().items() if v is not None}
        if fields_to_update:
            doc_ref.update(fields_to_update)
        return {"message": "Log updated successfully"}
    except Exception as e:
         raise HTTPException(status_code=500, detail=str(e))
         
@real_time_announcements_router.delete("/log/{log_id}")
def delete_log(log_id: str):
    try:
        db.collection("logs").document(log_id).delete()
        return {"message": "Log deleted successfully"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
