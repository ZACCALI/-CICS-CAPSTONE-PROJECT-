from fastapi import APIRouter, HTTPException
from api.firebaseConfig import db
from pydantic import BaseModel
import datetime
from firebase_admin import firestore
from api.controller import controller, Task, TaskType, Priority

emergency_route = APIRouter(prefix="/emergency", tags=["emergency"])

class EmergencyAction(BaseModel):
    user: str
    action: str # ACTIVATED / DEACTIVATED

@emergency_route.get("/")
def get_emergency_status():
    try:
        doc = db.collection("emergency").document("status").get()
        if doc.exists:
            return doc.to_dict()
        return {"active": False, "history": []}
    except Exception as e:
         return {"active": False, "history": [], "error": str(e)}

def log_to_file(msg):
    try:
        with open("debug_log.txt", "a") as f:
            f.write(f"{datetime.datetime.now()}: {msg}\n")
    except:
        pass

@emergency_route.post("/toggle")
def toggle_emergency(action: EmergencyAction):
    try:
        should_activate = action.action == "ACTIVATED"
        
        # 1. Controller State Management (The Source of Truth)
        if should_activate:
            task = Task(
                type=TaskType.EMERGENCY,
                priority=Priority.EMERGENCY,
                data={"user": action.user}
            )
            success = controller.request_playback(task)
        else:
            # Check Permission: Only the activator can deactivate
            active_user = controller.get_active_emergency_user()
            
            if active_user and active_user != action.user:
                 raise HTTPException(status_code=403, detail=f"Only the user who activated the emergency ({active_user}) can stop it.")
            
            # FIX: Must pass Task ID to bypass "Refresh Protection" logic in Controller
            task_id_to_stop = None
            if controller.current_task and controller.current_task.type == TaskType.EMERGENCY:
                task_id_to_stop = controller.current_task.id
            
            controller.stop_task(task_id_to_stop, TaskType.EMERGENCY, user=action.user)

        # 2. History & Logging (Preserved for Frontend compatibility)
        ref = db.collection("emergency").document("status")
        doc = ref.get()
        data = doc.to_dict() if doc.exists else {"active": False, "history": []}
        
        if should_activate:
            # ACTIVATED: Prepend new history entry
            history_entry = {
                "id": datetime.datetime.now().isoformat(),
                "action": "ACTIVATED",
                "time": datetime.datetime.now().strftime("%Y-%m-%d %I:%M %p"),
                "user": action.user
            }
            history = [history_entry] + data.get("history", [])

            # Log to Unified History
            _, doc_ref = db.collection("logs").add({
                "user": action.user,
                "action": "ACTIVATED Emergency",
                "type": "Emergency",
                "details": "Emergency Signal Broadcasting...", 
                "timestamp": firestore.SERVER_TIMESTAMP
            })
            
            ref.set({
                "active": True,
                "history": history,
                "current_log_id": doc_ref.id
            })
        else:
             # DEACTIVATED Logic (Closing the session)
             current_log_id = data.get("current_log_id")
             current_history = data.get("history", [])
             
             if current_history and current_history[0].get('action') == 'ACTIVATED':
                 end_time = datetime.datetime.now().strftime("%I:%M %p")
                 current_history[0]['action'] = "Emergency Session"
                 current_history[0]['time'] = f"{current_history[0]['time']} - {end_time}"
                 history = current_history
             else:
                 history_entry = {
                    "id": datetime.datetime.now().isoformat(),
                    "action": "DEACTIVATED",
                    "time": datetime.datetime.now().strftime("%Y-%m-%d %I:%M %p"),
                    "user": action.user
                }
                 history = [history_entry] + current_history

             ref.set({
                "active": False,
                "history": history,
                "current_log_id": None
             })
             
             # Update the unified log
             if current_log_id:
                 try:
                     db.collection("logs").document(current_log_id).update({
                         "action": "Emergency Session",
                         "details": f"Emergency Session Ended (Deactivated by {action.user})"
                     })
                 except:
                     pass

        return {"active": should_activate, "history": history}
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to toggle emergency: {str(e)}")

@emergency_route.delete("/history")
def clear_emergency_history(user: str = None):
    try:
        ref = db.collection("emergency").document("status")
        doc = ref.get()
        if doc.exists:
            current_history = doc.to_dict().get("history", [])
            if user:
                new_history = [h for h in current_history if h.get("user") != user]
                ref.update({"history": new_history})
            else:
                ref.update({"history": []})
        
        return {"message": "Emergency history cleared"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to clear history: {str(e)}")
