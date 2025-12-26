import os
import firebase_admin
from firebase_admin import credentials, firestore, auth

# Path to service account key
cred_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "serviceAccountKey.json")

# Initialize Firebase App
if not firebase_admin._apps:
    try:
        cred = credentials.Certificate(cred_path)
        firebase_admin.initialize_app(cred)
        print("Firebase initialized successfully.")
    except Exception as e:
        print(f"Error initializing Firebase: {e}")
        raise e

# Export Firestore and Auth clients
db = firestore.client()
# auth is the module itself, exported for convenience if needed, 
# but usually accessed via firebase_admin.auth directly in other files.

def firestore_server_timestamp():
    return firestore.SERVER_TIMESTAMP

