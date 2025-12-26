
import firebase_admin
from firebase_admin import credentials, firestore
import os

try:
    # Setup path
    cred_path = "serviceAccountKey.json"
    
    # Initialize
    if not firebase_admin._apps:
        cred = credentials.Certificate(cred_path)
        firebase_admin.initialize_app(cred)
    
    db = firestore.client()
    
    # Try to write and delete a test document
    doc_ref = db.collection('system_checks').document('connection_test')
    doc_ref.set({
        'status': 'connected',
        'timestamp': firestore.SERVER_TIMESTAMP
    })
    
    # Read it back
    doc = doc_ref.get()
    if doc.exists:
        print("SUCCESS: Successfully wrote to and read from Firestore.")
        print(f"Data: {doc.to_dict()}")
        
        # Cleanup
        doc_ref.delete()
        print("SUCCESS: Cleanup successful.")
    else:
        print("FAILURE: Document was written but could not be found.")

except Exception as e:
    print(f"FAILURE: Connection failed. Error: {e}")
