import requests
import time
from datetime import datetime

BASE_URL = "http://localhost:8000"

def test_controller_logic():
    print("Testing PA Controller Logic (With Time Shift)...")

    # 1. Reset (Clear Queue)
    # Since we don't have a 'clear queue' endpoint for testing, we rely on previous tests being clean
    # or we can clear schedules via DELETE.
    # Assuming clean state for now or low likelihood of interference.

    # 2. Submit a Schedule (Due NOW)
    print("\n1. Submitting Schedule (Due Immediately)...")
    now_str = datetime.now()
    sch_data = {
        "message": "Shift Test Message",
        "date": now_str.strftime("%Y-%m-%d"),
        "time": now_str.strftime("%H:%M"),
        "repeat": "once",
        "zones": "Zone Test"
    }
    res = requests.post(f"{BASE_URL}/scheduled/", json=sch_data)
    sch_id = res.json().get("id")
    print(f"   -> Schedule Submitted: {sch_id}")

    # 3. IMMEDIATELY Start Realtime (to preempt it before it finishes playing or right as it starts)
    # Wait a split second to let scheduler maybe pick it up, or preempt it from queue.
    # If scheduler picks it up, it becomes PLAYING. Then Realtime PREEMPTS it.
    time.sleep(1.1) 
    
    print("\n2. Interrupting with Realtime (Blocking for 3 seconds)...")
    rt_data = {"user": "Interrupter", "zones": ["Z1"], "type": "voice"}
    requests.post(f"{BASE_URL}/realtime/start", json=rt_data)
    
    # 4. Wait to simulate duration
    time.sleep(3)
    
    # 5. Stop Realtime
    print("\n3. Stopping Realtime (Should Trigger Shift)...")
    requests.post(f"{BASE_URL}/realtime/stop?user=Interrupter")

    # 6. Verify System Recovers
    time.sleep(1) # Wait for tick
    
    # We can't easily read internal queue state via API to verify exact timestamp 
    # unless we add a debug endpoint or check logs.
    # But we can verify it eventually plays.
    # For verifiable proof of shift, we'd need to inspect the backend logs.
    print("\n4. Verification: Check Backend Logs for '[Controller] Applying Time Shift'")
    
    # Cleanup
    requests.delete(f"{BASE_URL}/scheduled/{sch_id}")

if __name__ == "__main__":
    test_controller_logic()
