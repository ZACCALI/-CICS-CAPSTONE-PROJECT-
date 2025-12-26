
import sys
import os
import time

# Ensure we can import backend modules
sys.path.append(os.path.join(os.getcwd()))

from api.audio_service import audio_service

def test_sound():
    print("Testing Audio Service...")
    print("1. Testing TTS (You should hear 'System Check')")
    audio_service.play_text("System Check. One, Two, Three.")
    
    time.sleep(3)
    
    print("2. Testing again (You should hear 'Test Complete')")
    audio_service.play_text("Test Complete.")

if __name__ == "__main__":
    test_sound()
