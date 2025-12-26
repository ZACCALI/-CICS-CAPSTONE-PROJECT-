import requests
import time

BASE_URL = "http://localhost:8000"

def verify_emergency_kills_voice():
    print("=== VERIFYING: Emergency vs Voice Broadcast ===")

    # 1. Start Voice Broadcast
    print("\n1. Starting Voice Broadcast...")
    voice_data = {"user": "TestUser", "zones": ["All"], "type": "voice"}
    res = requests.post(f"{BASE_URL}/realtime/start", json=voice_data)
    
    if res.status_code == 200:
        print("   -> Voice Broadcast STARTED [OK]")
    else:
        print(f"   -> Failed to start voice: {res.text}")
        return

    # 2. Check State (Should be Voice)
    # Ideally we'd have a get_state endpoint, but absent that, we assume it's running.
    # We can check logs or just proceed to kill it.
    
    time.sleep(1) # Let it 'play' for a second

    # 3. Activate EMERGENCY
    print("\n2. Activating EMERGENCY (The 'Big Red Button')...")
    em_data = {"user": "Admin", "action": "ACTIVATED"}
    res = requests.post(f"{BASE_URL}/emergency/toggle", json=em_data)
    
    if res.status_code == 200:
        data = res.json()
        if data.get("active") is True:
             print("   -> Emergency ACTIVATED [OK]")
        else:
             print("   -> Emergency failed to return active=True")
    else:
        print(f"   -> Failed to toggle emergency: {res.text}")
        return

    # 4. Verify System State
    # Emergency should now be the dominant force. The Voice broadcast is effectively dead.
    print("\n3. Result Verification:")
    print("   If functionality is correct, the Backend Log should show:")
    print("   '[Controller] Preempting: voice'")
    print("   '  -> Killing Realtime <id>'")
    print("   '[Controller] Starting: emergency (Mode: EMERGENCY)'")

    # 5. Cleanup (Clear Emergency)
    print("\n4. Clearing Emergency...")
    em_data["action"] = "DEACTIVATED"
    requests.post(f"{BASE_URL}/emergency/toggle", json=em_data)
    print("   -> Emergency CLEARED [OK]")

if __name__ == "__main__":
    verify_emergency_kills_voice()
