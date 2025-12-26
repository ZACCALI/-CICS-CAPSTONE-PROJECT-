from fastapi import APIRouter, Depends, HTTPException, status
from api.firebaseConfig import db, auth
from api.routes.auth import verify_admin
from pydantic import BaseModel
import datetime

manage_account_router = APIRouter(prefix="/account", tags=["account"])

from firebase_admin import firestore

@manage_account_router.get("/")
def get_users(admin_user: dict = Depends(verify_admin)):
    """
    List all users from Firestore.
    Protected: Admin only.
    """
    try:
        users = []
        # Query Firestore users collection
        docs = db.collection("users").stream()
        for doc in docs:
            user_data = doc.to_dict()
            user_data["uid"] = doc.id
            users.append(user_data)
        return users
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to fetch users: {str(e)}")

class CreateUserRequest(BaseModel):
    email: str
    password: str
    name: str
    role: str = "user"

@manage_account_router.post("/create")
def create_user(user: CreateUserRequest, admin_user: dict = Depends(verify_admin)):
    """
    Create a new user (Auth + Firestore).
    Protected: Admin only.
    """
    try:
        # 1. Create in Firebase Auth
        auth_user = auth.create_user(
            email=user.email,
            password=user.password,
            display_name=user.name
        )
        
        # 2. Create in Firestore
        new_user_doc = {
            "uid": auth_user.uid,
            "name": user.name,
            "email": user.email,
            "role": user.role,
            "status": "approved", # Admin created = auto approved
            "createdAt": datetime.datetime.now().isoformat(),
            "lastLogin": None
        }
        db.collection("users").document(auth_user.uid).set(new_user_doc)
        
        # Log
        db.collection("logs").add({
            "user": "Admin", # Admin action
            "action": "User Created",
            "type": "Account",
            "details": f"Created user: {user.email} ({user.role})",
            "timestamp": firestore.SERVER_TIMESTAMP
        })
        
        return {"message": f"User {user.email} created successfully", "uid": auth_user.uid}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to create user: {str(e)}")

@manage_account_router.put("/approve/{uid}")
def approve_user(uid: str, admin_user: dict = Depends(verify_admin)):
    """
    Approve a user account by setting status to 'approved'.
    Protected: Admin only.
    """
    try:
        user_ref = db.collection("users").document(uid)
        user_doc = user_ref.get()
        if not user_doc.exists:
            raise HTTPException(status_code=404, detail="User not found")
        
        # Get user details for log
        user_data = user_doc.to_dict()
        user_email = user_data.get('email', 'Unknown')
        
        user_ref.update({
            "status": "approved",
            "isOnline": True, # Auto-online for immediate badge visibility
            "lastLogin": datetime.datetime.now().isoformat() # Set lastLogin to now
        })
        
        # Log
        db.collection("logs").add({
            "user": "Admin",
            "action": "User Approved",
            "type": "Account",
            "details": f"Approved user: {user_email}",
            "timestamp": firestore.SERVER_TIMESTAMP
        })
        
        return {"message": f"User {uid} approved successfully"}
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to approve user: {str(e)}")

@manage_account_router.post("/reset/{uid}")
def reset_user(uid: str, admin_user: dict = Depends(verify_admin)):
    """
    Reset a user account password to '12345678' and status to 'pending' (optional).
    Protected: Admin only.
    """
    try:
        # Get user details for log
        user_doc = db.collection("users").document(uid).get()
        user_email = "Unknown"
        if user_doc.exists:
             user_email = user_doc.to_dict().get('email', 'Unknown')

        # 1. Reset Password in Firebase Auth
        auth.update_user(uid, password="12345678")
        
        # Log
        db.collection("logs").add({
            "user": "Admin",
            "action": "Password Reset",
            "type": "Account",
            "details": f"Reset password for: {user_email}",
            "timestamp": firestore.SERVER_TIMESTAMP
        })
        
        return {"message": f"User {uid} password reset to 12345678 successfully"}
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to reset user: {str(e)}")

@manage_account_router.delete("/{uid}")
def delete_user(uid: str, admin_user: dict = Depends(verify_admin)):
    """
    Delete a user from both Firebase Authentication and Firestore.
    Protected: Admin only.
    """
    try:
        # Get user details for log before deletion
        user_doc = db.collection("users").document(uid).get()
        user_email = "Unknown"
        if user_doc.exists:
             user_email = user_doc.to_dict().get('email', 'Unknown')

        # 1. Delete from Firestore
        db.collection("users").document(uid).delete()
        
        # 2. Delete from Firebase Authentication
        try:
            auth.delete_user(uid)
        except auth.UserNotFoundError:
            # If user not in Auth but in Firestore, we proceed
            pass
        
        # Log
        db.collection("logs").add({
            "user": "Admin",
            "action": "User Deleted",
            "type": "Account",
            "details": f"Deleted user: {user_email}",
            "timestamp": firestore.SERVER_TIMESTAMP
        })
            
        return {"message": f"User {uid} deleted successfully"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to delete user: {str(e)}")

class UpdateAdminProfileRequest(BaseModel):
    name: str | None = None
    email: str | None = None
    password: str | None = None
    avatar: str | None = None

@manage_account_router.put("/profile")
def update_admin_profile(data: UpdateAdminProfileRequest, admin_user: dict = Depends(verify_admin)):
    """
    Update Admin Profile (Name, Email, Password) directly via Admin SDK.
    Bypasses 'Recent Login' requirement of client SDKs.
    Protected: Admin only.
    """
    uid = admin_user['uid']
    try:
        updates = {}
        # 1. Update Authentication (Email, Password, Name)
        auth_updates = {}
        if data.email: auth_updates['email'] = data.email
        if data.password: auth_updates['password'] = data.password
        if data.name: auth_updates['display_name'] = data.name
        
        if auth_updates:
            auth.update_user(uid, **auth_updates)

        # 2. Update Firestore
        firestore_updates = {}
        if data.name: firestore_updates['name'] = data.name
        if data.email: firestore_updates['email'] = data.email
        if data.avatar: firestore_updates['avatar'] = data.avatar
        
        if firestore_updates:
            db.collection("users").document(uid).update(firestore_updates)
            
        # Log
        db.collection("logs").add({
            "user": "Admin",
            "action": "Profile Updated",
            "type": "System",
            "details": f"System Administrator updated their profile/credentials.",
            "timestamp": firestore.SERVER_TIMESTAMP
        })

        return {"message": "Admin profile updated successfully"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to update admin profile: {str(e)}")
