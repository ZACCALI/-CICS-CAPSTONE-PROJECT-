from api.controller import controller, TaskType
from api.firebaseConfig import db
import firebase_admin
from firebase_admin import firestore

# 1. Force Controller Stop
controller.stop_task(None, TaskType.EMERGENCY)
print("Controller Emergency Stopped.")

# 2. Force Firestore Reset
db.collection("system").document("state").set({
    "mode": "IDLE",
    "priority": 0,
    "active_task": None
})
db.collection("emergency").document("status").set({
    "active": False,
    "history": []
})
print("Firestore State Reset.")
