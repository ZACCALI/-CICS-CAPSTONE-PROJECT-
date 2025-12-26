
import sys
import os
import time
from datetime import datetime, timedelta

# Ensure we can import backend modules
sys.path.append(os.path.join(os.getcwd()))

from api.controller import controller, Task, TaskType, Priority
from api.firebaseConfig import db

def test_recurrence():
    print("--- TESTING RECURRENCE (DAILY) ---")
    
    # 1. Create a "Past" Daily Task (so it runs immediately)
    now = datetime.now()
    past_time = now - timedelta(minutes=1)
    
    data = {
        'message': 'TEST_RECURRENCE_MSG',
        'date': past_time.strftime("%Y-%m-%d"),
        'time': past_time.strftime("%H:%M"),
        'repeat': 'daily',
        'zones': 'Test Zone',
        'status': 'Pending'
    }
    
    print(f"1. Injecting Test Schedule (Scheduled for: {data['time']})...")
    # We simulate what routes/scheduled.py does
    _, ref = db.collection('schedules').add(data)
    doc_id = ref.id
    
    task = Task(
        id=doc_id,
        type=TaskType.SCHEDULE,
        priority=Priority.SCHEDULE,
        data=data,
        scheduled_time=past_time
    )
    
    # 2. Add to Controller Queue
    controller.request_playback(task)
    
    # 3. Wait for Scheduler to pick it up (it sleeps 1s)
    print("2. Waiting for Scheduler to pick it up (5s)...")
    time.sleep(5)
    
    # 4. Check if it ran (Status should be Completed for original)
    doc = db.collection('schedules').document(doc_id).get()
    print(f"3. Original Task Status: {doc.to_dict().get('status')}")
    
    # 5. Check if New Task Created (Tomorrow)
    print("4. Checking for New Recurred Task...")
    pending_docs = db.collection('schedules')\
        .where('message', '==', 'TEST_RECURRENCE_MSG')\
        .where('status', '==', 'Pending')\
        .stream()
        
    found_recurrence = False
    for d in pending_docs:
        d_data = d.to_dict()
        if d_data['date'] != data['date']: # Should be different date
            print(f"   [SUCCESS] Found Recurring Task! Date: {d_data['date']} (ID: {d.id})")
            # Cleanup
            db.collection('schedules').document(d.id).delete()
            found_recurrence = True
            
    # Cleanup Original
    db.collection('schedules').document(doc_id).delete()
    
    if found_recurrence:
        print("✅ RECURRENCE TEST PASSED")
    else:
        print("❌ RECURRENCE TEST FAILED")

def test_persistence():
    print("\n--- TESTING PERSISTENCE (RESTART) ---")
    
    # 1. Create a Dummy Pending Task (Future)
    future_time = datetime.now() + timedelta(hours=1)
    data = {
        'message': 'TEST_PERSISTENCE_MSG',
        'date': future_time.strftime("%Y-%m-%d"),
        'time': future_time.strftime("%H:%M"),
        'repeat': 'once',
        'status': 'Pending'
    }
    _, ref = db.collection('schedules').add(data)
    doc_id = ref.id
    print(f"1. Created Pending Task in DB: {doc_id}")
    
    # 2. Simulate Server Restart (Manually call load function)
    # Clear queue first to prove it loads
    controller.queue = [] 
    print("2. Cleared In-Memory Queue (Simulating Restart)")
    
    print("3. Triggering _load_pending_schedules()...")
    controller._load_pending_schedules()
    
    # 3. Verify
    found = False
    for t in controller.queue:
        if t.id == doc_id:
            found = True
            print(f"   [SUCCESS] Task {t.id} loaded back into queue!")
            break
            
    # Cleanup
    db.collection('schedules').document(doc_id).delete()
    
    if found:
        print("✅ PERSISTENCE TEST PASSED")
    else:
        print("❌ PERSISTENCE TEST FAILED")

if __name__ == "__main__":
    test_recurrence()
    test_persistence()
    print("\nTests Completed.")
