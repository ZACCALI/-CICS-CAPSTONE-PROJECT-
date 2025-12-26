import requests
import time

BASE_URL = "http://localhost:8000"

def verify_emergency_permission():
    print("=== VERIFYING: Emergency Owner Policy ===")

    # 1. Activate as ADMIN
    print("\n1. Admin activates Emergency...")
    res = requests.post(f"{BASE_URL}/emergency/toggle", json={"user": "Admin", "action": "ACTIVATED"})
    if res.status_code == 200:
        print("   -> Emergency ACTIVATED by Admin [OK]")
    else:
        print(f"   -> Failed to activate: {res.text}")
        return

    time.sleep(1)

    # 2. Intruder tries to Deactivate
    print("\n2. 'UserB' tries to Deactivate...")
    res = requests.post(f"{BASE_URL}/emergency/toggle", json={"user": "UserB", "action": "DEACTIVATED"})
    
    if res.status_code == 403:
        print(f"   -> BLOCKED as expected: {res.json()['detail']} [PASS]")
    else:
        print(f"   -> FAILED! UserB was allowed (Status: {res.status_code}) [FAIL]")
        # Cleanup anyway
        requests.post(f"{BASE_URL}/emergency/toggle", json={"user": "Admin", "action": "DEACTIVATED"})
        return

    # 3. Admin Deactivates
    print("\n3. Admin deactivates...")
    res = requests.post(f"{BASE_URL}/emergency/toggle", json={"user": "Admin", "action": "DEACTIVATED"})
    
    if res.status_code == 200:
        print("   -> Deactivated by Owner [PASS]")
    else:
         print(f"   -> Owner failed to deactivate: {res.text} [FAIL]")

if __name__ == "__main__":
    verify_emergency_permission()
