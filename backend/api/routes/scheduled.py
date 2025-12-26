from fastapi import APIRouter, HTTPException
from api.firebaseConfig import db, firestore_server_timestamp
from pydantic import BaseModel
from typing import Optional
from api.controller import controller, Task, TaskType, Priority
from datetime import datetime

scheduled_announcements_router = APIRouter(prefix="/scheduled", tags=["scheduled"])

class ScheduleItem(BaseModel):
    message: str
    date: str
    time: str
    repeat: str
    zones: str # comma separated
    type: str = 'text' # or 'voice'
    audio: Optional[str] = None # Base64

@scheduled_announcements_router.get("/")
def get_schedules():
    try:
        docs = db.collection("schedules").stream()
        schedules = []
        for doc in docs:
            data = doc.to_dict()
            data["id"] = doc.id
            schedules.append(data)
        # Sort by date/time (optional, can be done in frontend)
        return schedules
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to fetch schedules: {str(e)}")

@scheduled_announcements_router.post("/")
def create_schedule(schedule: dict):
    try:
        # 1. Validation
        required = ["message", "date", "time", "repeat", "zones"]
        for f in required:
            if f not in schedule or not schedule[f]:
                 raise HTTPException(status_code=400, detail=f"Missing field: {f}")

        # 2. Conflict Check (Prevent duplicate times)
        existing = db.collection("schedules")\
            .where("date", "==", schedule["date"])\
            .where("time", "==", schedule["time"])\
            .where("status", "==", "Pending")\
            .get()
            
        if len(existing) > 0:
            raise HTTPException(status_code=409, detail="Time slot already occupied. Please choose another time.")

        # 3. Persistence (Firestore)
        if "id" in schedule: del schedule["id"]
        schedule['status'] = 'Pending' # Default
        
        _, doc_ref = db.collection("schedules").add(schedule)
        doc_id = doc_ref.id
        
        # 3. Sync to Controller Queue
        # Parse datetime for sorting
        # 3. Queueing
        try:
            dt_str = f"{schedule['date']} {schedule['time']}"
            scheduled_time = datetime.strptime(dt_str, "%Y-%m-%d %H:%M")
        except ValueError:
             raise HTTPException(status_code=400, detail="Invalid date/time format. Use YYYY-MM-DD and HH:MM")

        task = Task(
            id=doc_id, # Use Firestore ID for consistency
            type=TaskType.SCHEDULE,
            priority=Priority.SCHEDULE, # 20
            data=schedule,
            scheduled_time=scheduled_time
        )
        controller.request_playback(task)

        # 4. Log
        db.collection("logs").add({
            "user": schedule.get("user", "Admin"),
            "action": "Schedule Created",
            "type": "Schedule",
            "details": f"Scheduled: {schedule.get('message')}",
            "timestamp": firestore_server_timestamp()
        })

        return {"id": doc_id, "message": "Schedule created and queued"}
    except HTTPException as he:
        raise he
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to create schedule: {str(e)}")

@scheduled_announcements_router.put("/{id}")
def update_schedule(id: str, schedule: dict):
    try:
        # 1. Persistence
        db.collection("schedules").document(id).set(schedule, merge=True)
        
        # 2. Sync Controller (Remove old, Add new)
        controller.remove_from_queue(id)
        
        # Re-add
        try:
            dt_str = f"{schedule.get('date')} {schedule.get('time')}"
            scheduled_time = datetime.strptime(dt_str, "%Y-%m-%d %H:%M")
        except ValueError:
             raise HTTPException(status_code=400, detail="Invalid date/time format")

        task = Task(
            id=id,
            type=TaskType.SCHEDULE,
            priority=Priority.SCHEDULE,
            data=schedule,
            scheduled_time=scheduled_time
        )
        controller.request_playback(task)
        
        # 3. Log
        db.collection("logs").add({
            "user": schedule.get("user", "Admin"),
            "action": "Schedule Updated",
            "type": "Schedule",
            "details": f"Updated schedule ID: {id}",
            "timestamp": firestore_server_timestamp()
        })

        return {"message": "Schedule updated and re-queued"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to update schedule: {str(e)}")

@scheduled_announcements_router.delete("/{id}")
def delete_schedule(id: str, user: str = "Admin"):
    try:
        # 1. Persistence
        db.collection("schedules").document(id).delete()
        
        # 2. Sync Controller
        controller.remove_from_queue(id)
        # Also stop if it's currently playing?
        # controller.stop_task(id) # Optional, strictly speaking deleting a schedule usually stops it if it's playing.
        controller.stop_task(id)

        # 3. Log
        db.collection("logs").add({
            "user": user,
            "action": "Schedule Deleted",
            "type": "Schedule",
            "details": f"Deleted schedule ID: {id}",
            "timestamp": firestore_server_timestamp()
        })

        return {"message": "Schedule deleted and unqueued"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to delete schedule: {str(e)}")
