from api.firebaseConfig import db
from firebase_admin import firestore

def cleanup_logs():
    print("Starting Log Cleanup...")
    users_to_delete = ['SessionTester', 'DJ Test', 'TestAdmin']
    
    try:
        logs_ref = db.collection("logs")
        
        # Firestore 'in' query supports up to 10 values
        query = logs_ref.where("user", "in", users_to_delete)
        docs = query.stream()
        
        count = 0
        batch = db.batch()
        
        for doc in docs:
            batch.delete(doc.reference)
            count += 1
            
            # Commit batch every 400 items (limit is 500)
            if count % 400 == 0:
                batch.commit()
                batch = db.batch()
                print(f"   -> Deleted batch of 400...")
        
        if count > 0:
            batch.commit() # Commit remaining
            print(f"   -> Cleanup Complete. Deleted {count} logs.")
        else:
            print("   -> No logs found for target users.")
            
    except Exception as e:
        print(f"Error during cleanup: {e}")

if __name__ == "__main__":
    cleanup_logs()
