from fastapi import APIRouter, Depends, HTTPException, status, Header
from firebase_admin import auth
from api.firebaseConfig import db

auth_router = APIRouter(prefix="/auth", tags=["auth"])

async def verify_token(authorization: str = Header(...)):
    """
    Verifies the Firebase ID token passed in the Authorization header.
    Returns the decoded token if valid.
    """
    if not authorization.startswith("Bearer "):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Invalid authentication credentials",
        )
    
    token = authorization.split("Bearer ")[1]
    
    try:
        # Allow 60 seconds of clock skew to prevent "Token used too early" errors
        decoded_token = auth.verify_id_token(token, clock_skew_seconds=60)
        return decoded_token
    except Exception as e:
        print(f"Error verifying token: {e}") # Debug logging
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail=f"Invalid token: {str(e)}",
        )

async def verify_admin(decoded_token: dict = Depends(verify_token)):
    """
    Verifies if the user associated with the token has admin privileges.
    Checks Firestore 'users' collection for the 'role' field.
    """
    uid = decoded_token.get("uid")
    if not uid:
        raise HTTPException(status_code=401, detail="Invalid token payload")

    user_doc = db.collection("users").document(uid).get()
    if not user_doc.exists:
        raise HTTPException(status_code=403, detail="User not found")

    user_data = user_doc.to_dict()
    if user_data.get("role") != "admin":
        raise HTTPException(status_code=403, detail="Admins only")
    
    return decoded_token

@auth_router.get("/")
def auth_check():
    return {"message": "Auth module loaded"}
