
import requests
import time

BASE_URL = "http://127.0.0.1:8000"

def test_logs():
    print("Testing Unified Logging...")

    # 1. Trigger Emergency Logs
    print("1. Triggering Emergency Toggle...")
    try:
        res = requests.post(f"{BASE_URL}/emergency/toggle", json={"user": "Admin", "action": "ACTIVATED"})
        if res.status_code == 200:
            print("   -> Emergency Activated Successfully")
        else:
            print(f"   -> Failed: {res.text}")
            
        time.sleep(1) # Wait for Firestore
        
        # Deactivate
        res = requests.post(f"{BASE_URL}/emergency/toggle", json={"user": "Admin", "action": "DEACTIVATED"})
        
        # Verify consolidation: History should have 1 entry (merged), not 2
        history = res.json().get('history', [])
        # We expect the top entry to be "Emergency Session" for Admin
        # And if we just cleared history before this test, it should be 1 total? 
        # But we might have other history. Let's check the top one.
        if history and history[0].get('action') == 'Emergency Session':
             print("   -> [PASS] Logs Consolidated to 'Emergency Session'")
        else:
             print(f"   -> [FAIL] Log consolidation failed. Top action: {history[0].get('action') if history else 'None'}")
    except Exception as e:
        print(f"   -> Error: {e}")

    # 2. Trigger Schedule Logs
    print("2. Creating Schedule...")
    try:
        schedule_data = {
            "message": "Test Announcement",
            "date": "2025-12-25",
            "time": "10:00",
            "repeat": "None",
            "zones": "All Zones",
            "status": "Scheduled",
            "type": "Text"
        }
        res = requests.post(f"{BASE_URL}/scheduled/", json=schedule_data)
        if res.status_code == 200:
            sch_id = res.json().get("id")
            print(f"   -> Schedule Created (ID: {sch_id})")
            
            # Delete it to clean up
            if sch_id:
                requests.delete(f"{BASE_URL}/scheduled/{sch_id}")
                print("   -> Schedule Deleted")
        else:
            print(f"   -> Failed: {res.text}")
    except Exception as e:
        print(f"   -> Error: {e}")

    # 3. Trigger Log Deletion (to verify audit)
    print("3. Deleting a Log to Verify Audit...")
    try:
        # Create a dummy log first
        dummy = {"user": "Admin", "type": "System", "action": "Dummy Log", "details": "To be deleted"}
        res = requests.post(f"{BASE_URL}/realtime/log", json=dummy)
        
        # We need the ID to delete it. Fetched logs are needed.
        # So we fetch, find our dummy, and delete it.
        time.sleep(1)
        logs_res = requests.get(f"{BASE_URL}/realtime/logs")
        if logs_res.status_code == 200:
            all_logs = logs_res.json()
            # Find closest Dummy Log
            target = next((l for l in all_logs if l.get('action') == 'Dummy Log'), None)
            if target:
                del_res = requests.delete(f"{BASE_URL}/realtime/log/{target['id']}?user=TestAudit")
                if del_res.status_code == 200:
                    print("   -> Log Deleted Successfully via Realtime API")
                else:
                    print(f"   -> Delete Failed: {del_res.text}")
            else:
                 print("   -> Could not find dummy log to delete")
        
    except Exception as e:
        print(f"   -> Error: {e}")

    # 4. Trigger Silent Deletion (delete the 'Log Deleted' log)
    print("4. Verify Silent Deletion (Delete a log and ensure NO audit log is created)...")
    try:
        # Create a log to delete
        temp_log = {"user": "Admin", "type": "System", "action": "Temp Log", "details": "To be silenty deleted"}
        res = requests.post(f"{BASE_URL}/realtime/log", json=temp_log)
        temp_id = res.json().get("id")
        time.sleep(1)

        if temp_id:
            # Delete it
            del_res = requests.delete(f"{BASE_URL}/realtime/log/{temp_id}?user=TestAudit")
            if del_res.status_code == 200:
                 print("   -> Log deleted successfully.")
                 time.sleep(1)
                 
                 # Verify NO "Log Deleted" type log appeared recently
                 logs_res = requests.get(f"{BASE_URL}/realtime/logs")
                 all_logs = logs_res.json()
                 
                 # Check top logs (most recent)
                 recent_audit = next((l for l in all_logs[:5] if l.get('action') == 'Log Deleted'), None)
                 
                 if recent_audit:
                      print(f"   -> [FAIL] Found 'Log Deleted' audit log: {recent_audit['id']} (Should be silent)")
                 else:
                      print("   -> [PASS] No 'Log Deleted' audit entry found (Silent Deletion Confirmed)")
            else:
                 print(f"   -> Delete Failed: {del_res.text}")
        else:
             print("   -> Failed to create temp log for silent test")

    except Exception as e:
        print(f"   -> Error: {e}")

    # 5. Check Logs
    print("3. Fetching Logs to Verify Persistence...")
    try:
        res = requests.get(f"{BASE_URL}/realtime/logs")
        if res.status_code == 200:
            logs = res.json()
            print(f"   -> Fetched {len(logs)} logs.")
            
            # Verify recent logs exist
            found_emergency = False
            found_schedule = False
            found_deletion = False
            
            for log in logs[:15]: # Check last 15
                print(f"      Log: [{log.get('type')}] {log.get('action')} - {log.get('details')}")
                if log.get('type') == 'Emergency' and 'ACTIVATED' in log.get('action', ''):
                    found_emergency = True
                if log.get('type') == 'Schedule' and 'Schedule Created' in log.get('action', ''):
                    found_schedule = True
                if log.get('type') == 'System' and 'Log Deleted' in log.get('action', ''):
                    found_deletion = True
            
            if found_emergency: print("   [PASS] Emergency Log Found")
            else: print("   [FAIL] Emergency Log NOT Found")
            
            if found_schedule: print("   [PASS] Schedule Log Found")
            else: print("   [FAIL] Schedule Log NOT Found")

            if found_deletion: print("   [PASS] Deletion Audit Log Found")
            else: print("   [FAIL] Deletion Audit Log NOT Found")
            
        else:
            print(f"   -> Failed to fetch logs: {res.text}")
    except Exception as e:
        print(f"   -> Error: {e}")

    # 5. Test Clear Emergency History (User Specific)
    print("5. Testing Clear Emergency History (User Scoped)...")
    try:
        # Clear it for Admin
        clear_res = requests.delete(f"{BASE_URL}/emergency/history?user=Admin")
        if clear_res.status_code == 200:
             print("   -> History Cleared Successfully for Admin")
        else:
             print(f"   -> Failed to clear history: {clear_res.text}")

        # specific check
        status_res = requests.get(f"{BASE_URL}/emergency/")
        data = status_res.json()
        
        # Check if Admin entries are gone
        remaining = [h for h in data.get('history', []) if h.get('user') == 'Admin']
        if len(remaining) == 0:
             print("   -> [PASS] Admin history is empty")
        else:
             print(f"   -> [FAIL] Admin history remaining: {len(remaining)} items")

    except Exception as e:
        print(f"   -> Error: {e}")

    # 6. Test Music Logs
    print("6. Testing Music Logs...")
    try:
         # Manually create a Music log (mocking what frontend does)
         music_log = {
             "user": "Admin",
             "type": "Music",
             "action": "Started Audio Playback",
             "details": "Playing file: VerificationTrack.mp3"
         }
         res = requests.post(f"{BASE_URL}/realtime/log", json=music_log)
         if res.status_code == 200:
             print("   -> Music Log Created Successfully")
         else:
             print(f"   -> Failed to create Music log: {res.text}")
             
    except Exception as e:
        print(f"   -> Error: {e}")

    print("7. Testing Unified Session Logging (Start -> Update)...")
    try:
         # 1. Start Session
         start_data = {
             "user": "Admin",
             "type": "Music",
             "action": "Active Session",
             "details": "Playing..."
         }
         res1 = requests.post(f"{BASE_URL}/realtime/log", json=start_data)
         log_id = res1.json().get("id")
         
         if log_id:
             print(f"   -> Session Started (ID: {log_id})")
             time.sleep(1)
             
             # Update it
             update_payload = {
                 "action": "Music Session",
                 "details": "Music Session (Start: 10:00 AM - End: 10:01 AM)"
             }
             res2 = requests.put(f"{BASE_URL}/realtime/log/{log_id}", json=update_payload)
             
             if res2.status_code == 200:
                 print("   -> Session Updated Successfully")
             else:
                 print(f"   -> Failed to update session: {res2.text}")
         else:
             print("   -> Failed to start session (No ID returned)")

    except Exception as e:
        print(f"   -> Error: {e}")

    # 8. Check Logs (Updated)
    print("8. Fetching Logs to Verify Persistence...")
    try:
        time.sleep(1)
        res = requests.get(f"{BASE_URL}/realtime/logs")
        if res.status_code == 200:
            logs = res.json()
            print(f"   -> Fetched {len(logs)} logs.")
            
            # Verify recent logs exist
            found_music_session = False
            
            for log in logs[:20]: # Check last 20
                print(f"      Log: [{log.get('type')}] {log.get('action')} - {log.get('details')}")
                # We look for the UPDATED state
                if log.get('type') == 'Music' and log.get('action') == 'Music Session':
                    details = log.get('details', '')
                    if "(Start:" in details and "- End:" in details:
                         found_music_session = True
            
            if found_music_session: print("   [PASS] Unified Music Session Log Found (With Details)")
            else: print("   [FAIL] Unified Music Session Log NOT Found (Or missing details)")
            
        else:
            print(f"   -> Failed to fetch logs: {res.text}")
    except Exception as e:
        print(f"   -> Error: {e}")

if __name__ == "__main__":
    test_logs()
