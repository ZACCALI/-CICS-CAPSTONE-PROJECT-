import firebase_admin
from firebase_admin import credentials, firestore, auth
import os

# Initialize Firebase
if not firebase_admin._apps:
    cred = credentials.Certificate("serviceAccountKey.json")
    firebase_admin.initialize_app(cred)

db = firestore.client()

def enforce_single_admin():
    print("--- Enforcing Single Admin Policy ---")
    
    # 1. Fetch all users
    users_ref = db.collection("users")
    docs = users_ref.stream()
    
    admins = []
    all_users = []
    
    for doc in docs:
        data = doc.to_dict()
        uid = doc.id
        email = data.get('email', '')
        role = data.get('role', 'user')
        
        user_info = {'uid': uid, 'email': email, 'role': role, 'data': data}
        all_users.append(user_info)
        
        if role == 'admin':
            admins.append(user_info)
            
    print(f"Total Users: {len(all_users)}")
    print(f"Total Admins Found: {len(admins)}")
    
    target_admin = None
    
    # 2. Logic to handle admin count
    if len(admins) == 1:
        target_admin = admins[0]
        print(f"Good: Exact number of admins (1). Admin is: {target_admin['email']}")
        
    elif len(admins) > 1:
        print(f"WARNING: Multiple admins found ({len(admins)}). Cleaning up...")
        # Strategy: Keep 'admin@gmail.com' if exists, otherwise keep the one with lowest UID/Email (deterministic) or just first.
        # Let's prioritize 'admin@gmail.com'
        
        target_admin = next((a for a in admins if 'admin@gmail.com' in a['email']), admins[0])
        
        for admin in admins:
            if admin['uid'] != target_admin['uid']:
                print(f" -> Demoting {admin['email']} ({admin['uid']}) to 'user'")
                db.collection("users").document(admin['uid']).update({'role': 'user'})
        
        print(f"Cleanup complete. Key Admin is: {target_admin['email']}")

    else:
        # 0 Admins - Create one
        print("No admins found. Creating default admin...")
        email = "admin@gmail.com"
        password = "admin123" # Default
        
        try:
            # Check if user exists in Auth but not Firestore or just needs role update
            try:
                user = auth.get_user_by_email(email)
                uid = user.uid
                print(f" -> Auth user exists ({uid}). Promoting to admin...")
            except auth.UserNotFoundError:
                user = auth.create_user(email=email, password=password, display_name="System Administrator")
                uid = user.uid
                print(f" -> Created new Auth user ({uid}).")
            
            # Set Firestore Doc
            target_admin = {
                'uid': uid,
                'email': email,
                'role': 'admin',
                'status': 'approved',
                'name': 'System Administrator'
            }
            db.collection("users").document(uid).set(target_admin, merge=True)
            print(" -> Admin created/promoted successfully.")
            
        except Exception as e:
            print(f"CRITICAL ERROR creating admin: {e}")
            return

    # 3. Ensure the single admin is APPROVED and has correct role
    if target_admin:
        uid = target_admin['uid']
        print(f"Verifying final admin status for {uid}...")
        users_ref.document(uid).update({
            'role': 'admin',
            'status': 'approved'
        })
        print(" -> Verified: Role=Admin, Status=Approved")

def cleanup_non_admins():
    print("\n--- Cleaning up Non-Admin Users ---")
    users = db.collection("users").stream()
    count = 0
    for doc in users:
        data = doc.to_dict()
        uid = doc.id
        role = data.get('role', 'user')
        email = data.get('email', 'Unknown')
        
        if role != 'admin':
            print(f"Deleting user: {email} ({uid})")
            # 1. Delete Firestore
            db.collection("users").document(uid).delete()
            # 2. Delete Auth
            try:
                auth.delete_user(uid)
                print(f" -> Deleted from Auth.")
            except auth.UserNotFoundError:
                print(f" -> Not found in Auth (already deleted).")
            except Exception as e:
                print(f" -> Error deleting auth: {e}")
            count += 1
        else:
            print(f"SKIPPING Admin: {email}")

    print(f"Cleanup Complete. Deleted {count} users.")

def reset_emergency_state():
    print("--- Resetting Emergency State ---")
    try:
        # Old path cleanup (optional but good)
        db.collection("system_state").document("emergency").delete()
        
        # New path reset
        db.collection("emergency").document("status").set({
            "active": False,
            "history": [],
            "current_log_id": None
        })
        print(" -> Emergency state reset to inactive/empty.")
    except Exception as e:
        print(f" -> Error resetting emergency: {e}")

if __name__ == "__main__":
    enforce_single_admin()
    cleanup_non_admins()
    reset_emergency_state()
