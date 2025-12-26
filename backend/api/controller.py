import threading
import time
import os
import uuid
from enum import IntEnum
from datetime import datetime, timedelta
from typing import Optional, List, Dict
from firebase_admin import firestore
from firebase_admin import firestore
from api.firebaseConfig import db
from api.audio_service import audio_service # <--- NEW IMPORT

# --- 1. Constants & Enums ---
class Priority(IntEnum):
    IDLE = 0
    BACKGROUND = 10
    SCHEDULE = 20
    REALTIME = 30
    EMERGENCY = 100

class State(IntEnum):
    PENDING = 1
    PLAYING = 2
    INTERRUPTED = 3
    COMPLETED = 4

class TaskType:
    VOICE = 'voice'
    TEXT = 'text'
    EMERGENCY = 'emergency'
    SCHEDULE = 'schedule'
    BACKGROUND = 'background'

# --- 2. Data Structures ---
class Task:
    def __init__(self, 
                 type: str, 
                 priority: int, 
                 data: dict, 
                 id: str = None, 
                 status: State = State.PENDING, 
                 created_at: datetime = None,
                 scheduled_time: datetime = None):
        self.id = id if id else str(uuid.uuid4())
        self.type = type
        self.priority = priority
        self.data = data
        self.status = status
        self.created_at = created_at if created_at else datetime.now()
        self.scheduled_time = scheduled_time if scheduled_time else datetime.now()

    def to_dict(self):
        return {
            'id': self.id,
            'type': self.type,
            'priority': int(self.priority),
            'data': self.data,
            'status': int(self.status),
            'created_at': self.created_at.isoformat(),
            'scheduled_time': self.scheduled_time.isoformat()
        }

# --- 3. The Controller ---
class PAController:
    _instance = None

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(PAController, cls).__new__(cls)
            cls._instance._initialized = False
        return cls._instance

    def __init__(self):
        if self._initialized:
            return
        
        self._lock = threading.Lock()  # THE MUTEX
        self.current_task: Optional[Task] = None
        self.queue: List[Task] = []    # Priority Queue for Schedules
        self.emergency_mode = False
        self._running = True
        
        # Track interruption duration to shift queue
        self.pause_start_time: Optional[datetime] = None
        
        # Track Suspended Task (for Resume)
        self.suspended_task: Optional[Task] = None

        # Reset Logic on init to ensure clean state
        self._reset_state()
        
        # Cleanup State
        self.last_cleanup = datetime.now()

        # Start Scheduler Thread
        self.scheduler_thread = threading.Thread(target=self._scheduler_loop, daemon=True)
        self.scheduler_thread.start()
        
        self._initialized = True
        print("PA Controller Initialized")

    def _reset_state(self):
        """Resets Firestore state to Idle on startup"""
        try:
            db.collection('system').document('state').set({
                'active_task': None,
                'priority': 0,
                'mode': 'IDLE',
                'timestamp': firestore.SERVER_TIMESTAMP
            })
            self._load_pending_schedules() # <--- NEW: Load schedules on reset/startup
        except Exception as e:
            print(f"Failed to reset state: {e}")

    def _load_pending_schedules(self):
        """Resilience: Loads 'Pending' schedules from Firestore into Queue on startup"""
        print("[Controller] Loading pending schedules from database...")
        try:
            # Query all Pending schedules
            # FIX: Use keyword arguments to silence UserWarning
            docs = db.collection('schedules').where(filter=firestore.FieldFilter('status', '==', 'Pending')).stream()
            count = 0
            for doc in docs:
                data = doc.to_dict()
                try:
                    # Parse Date/Time
                    dt_str = f"{data.get('date')} {data.get('time')}"
                    scheduled_time = datetime.strptime(dt_str, "%Y-%m-%d %H:%M")
                    
                    # Create Task
                    task = Task(
                        id=doc.id,
                        type=TaskType.SCHEDULE,
                        priority=Priority.SCHEDULE,
                        data=data,
                        scheduled_time=scheduled_time
                    )
                    
                    # Add to internal queue (avoiding re-triggering logic)
                    self.queue.append(task)
                    count += 1
                except ValueError:
                    print(f"  -> Skipping invalid date format in {doc.id}")
                    continue
            
            # Sort queue
            self.queue.sort(key=lambda x: x.scheduled_time)
            print(f"[Controller] Resilience: Loaded {count} pending tasks.")
            
        except Exception as e:
            print(f"[Controller] Failed to load pending schedules: {e}")

    # --- MAIN ENTRY POINT ---
    def request_playback(self, new_task: Task) -> bool:
        with self._lock:  # Critical Section
            print(f"[Controller] Request: {new_task.type} (Pri: {new_task.priority})")

            # 1. Emergency Check (Invincible)
            if self.emergency_mode and new_task.priority < Priority.EMERGENCY:
                print(f"[Controller] Denied: Emergency Active")
                return False

            # 2. Schedule Check (Always Queue first)
            if new_task.type == TaskType.SCHEDULE:
                 # Standard Schedule submission (queued)
                 print(f"[Controller] Queued Schedule: {new_task.id}")
                 self._add_to_queue(new_task)
                 return True 

            # 3. Priority Check
            current_pri = self.current_task.priority if self.current_task else Priority.IDLE
            
            # CHECK OWNERSHIP: Allow user to interrupt THEMSELVES (e.g. Refresh page)
            is_same_user = False
            if self.current_task and self.current_task.data.get('user') == new_task.data.get('user'):
                is_same_user = True

            # Logic: Higher Priority WINS OR (Equal Priority AND Same User WINS)
            if new_task.priority > current_pri or (new_task.priority == current_pri and is_same_user):
                # PREEMPTION
                self._preempt_current_task(new_task.priority)
                self._start_task(new_task)
                return True
            
            else:
                # Lower/Equal priority (different user) -> Busy
                print(f"[Controller] Denied: Busy (Current: {current_pri}, New: {new_task.priority})")
                return False

    def stop_task(self, task_id: str, task_type: str = None, user: str = None):
        """Called to manually stop a task (e.g., Stop Broadcast, Clear Emergency)"""
        with self._lock:
            if not self.current_task:
                return

            # If requesting to stop specific task, check ID
            if task_id and self.current_task.id != task_id:
                print(f"[Controller] Denied Stop: ID Mismatch ({task_id} vs {self.current_task.id})")
                return

            # Anti-Zombie: Require ID for Realtime Tasks to prevent race conditions
            # (e.g. Frontend "Stop" from previous session killing new session)
            # Anti-Zombie: Require ID for Realtime Tasks to prevent race conditions
            # (e.g. Frontend "Stop" from previous session killing new session)
            if not task_id:
                 # --- NEW PROTECTION START ---
                 # If no ID is provided (e.g. generic 'handleUnload' stop), 
                 # STRICTLY PROTECT persistent tasks (Schedule, Emergency).
                 if self.current_task.type in [TaskType.SCHEDULE, TaskType.EMERGENCY]:
                     print(f"[Controller] Denied Generic Stop: Cannot kill {self.current_task.type} without Task ID.")
                     return
                 # --- NEW PROTECTION END ---

                 # RECOVERY: If ID is missing, but the User requesting stop matches the Task Owner, ALLOW IT.
                 current_owner = self.current_task.data.get('user')
                 if user and current_owner and user == current_owner:
                     print(f"[Controller] ALLOWING STOP (Recovery): User '{user}' matches Task Owner")
                 elif user in ['System', 'System Admin', 'Admin']: # Admin Override
                     print(f"[Controller] ALLOWING STOP (Admin Override): User '{user}'")
                 else:
                     print(f"[Controller] Denied Stop: Missing Task ID for Realtime Task (Requester: {user}, Owner: {current_owner})")
                     return

            print(f"[Controller] Stopping Task: {self.current_task.id}")
            
            if self.current_task.priority == Priority.EMERGENCY:
                self.emergency_mode = False
            
            self.current_task = None
            self._update_firestore_state(None, Priority.IDLE, 'IDLE')
            
            # Stop Audio
            audio_service.stop()
            
            # Application of Time Shift (System became IDLE)
            self._apply_queue_shift()

            # RESUME SUSPENDED TASK
            if self.suspended_task:
                 print(f"[Controller] [RESUME] Found Suspended Task: {self.suspended_task.type} (ID: {self.suspended_task.id})")
                 # Small delay for smooth transition
                 time.sleep(1)
                 # Force status reset just in case
                 self.suspended_task.status = State.PENDING
                 self._start_task(self.suspended_task)
                 # IMPORTANT: Do NOT clear suspended_task here if _start_task fails (it doesn't fail).
                 # We must ensure we don't lose it.
                 self.suspended_task = None
                 print(f"[Controller] [RESUME] Task Resumed. Suspended cleared.")
            else:
                 print(f"[Controller] No Suspended Task to resume.")

    def get_queue(self):
        return self.queue
        
    def remove_from_queue(self, schedule_id: str):
        with self._lock:
             self.queue = [t for t in self.queue if t.id != schedule_id]

    def get_active_emergency_user(self) -> Optional[str]:
        with self._lock:
            if self.current_task and self.current_task.priority == Priority.EMERGENCY:
                 return self.current_task.data.get('user')
            return None

    # --- INTERNAL LOGIC ---
    def _add_to_queue(self, task: Task):
        self.queue.append(task)
        # Sort by scheduled_time
        self.queue.sort(key=lambda x: x.scheduled_time)

    def _preempt_current_task(self, new_priority):
        if not self.current_task:
            return

        print(f"[Controller] Preempting: {self.current_task.type}")

        # Specific Logic per Type
        if self.current_task.type == TaskType.SCHEDULE:
            # Soft Stop: Re-queue at HEAD
            print(f"  -> Re-queueing Schedule {self.current_task.id}")
            self.current_task.status = State.INTERRUPTED
            # Push to front of queue
            self.queue.insert(0, self.current_task) 
        
        elif self.current_task.type == TaskType.VOICE or self.current_task.type == TaskType.TEXT:
            # Hard Stop: Kill completely
            print(f"  -> Killing Realtime {self.current_task.id}")
            self.current_task.status = State.COMPLETED
            
        elif self.current_task.type == TaskType.BACKGROUND:
            # Soft Stop: Suspend
            print(f"  -> [SUSPEND] Suspending Background Task {self.current_task.id} for {new_priority}")
            self.suspended_task = self.current_task
            self.current_task = None
            # Do NOT mark COMPLETED. State remains valid in object.
        
        else:
             # Default Fallback
             self.current_task = None

        self.current_task = None
        
        # Stop Audio logic
        audio_service.stop()

    def _start_task(self, task: Task):
        self.current_task = task
        self.current_task.status = State.PLAYING
        
        # Start Time Shift Tracking if High Priority
        if task.priority >= Priority.REALTIME:
            if self.pause_start_time is None:
                self.pause_start_time = datetime.now()
                print(f"[Controller] Time Shift Started at {self.pause_start_time}")

        if task.priority == Priority.EMERGENCY:
            self.emergency_mode = True
        
        mode = 'BROADCAST'
        if task.type == TaskType.EMERGENCY: mode = 'EMERGENCY'
        elif task.type == TaskType.SCHEDULE: mode = 'SCHEDULE'
        elif task.type == TaskType.BACKGROUND: mode = 'BACKGROUND'

        print(f"[Controller] Starting: {task.type} (Mode: {mode})")
        self._update_firestore_state(task, task.priority, mode)
        
        # --- AUDIO OUTPUT START ---
        if task.type == TaskType.SCHEDULE:
             # Check if it's Text or Audio
             msg = task.data.get('message', '')
             if not msg: msg = "Scheduled Announcement."
             
             # UPDATED: Use chained playback (Intro -> Text) Non-Blocking
             intro_path = os.path.join("system_sounds", "intro.mp3")
             abs_intro = os.path.abspath(intro_path)
             
             audio_service.play_announcement(abs_intro, msg)

        elif task.type == TaskType.VOICE:
            # Voice is Realtime (Microphone). 
            # UPDATED: Use Async Intro so we don't block Start Command
            intro_path = os.path.join("system_sounds", "intro.mp3")
            abs_intro = os.path.abspath(intro_path)
            audio_service.play_intro_async(abs_intro)
            
        elif task.type == TaskType.TEXT:
            # FIX: Route sends 'content', not 'message'
            msg = task.data.get('content') or task.data.get('message', '')
            if msg:
                print(f"[Controller] Speaking Text: {msg}")
                # UPDATED: Use chained playback
                intro_path = os.path.join("system_sounds", "intro.mp3")
                abs_intro = os.path.abspath(intro_path)
                audio_service.play_announcement(abs_intro, msg)
            else:
                print("[Controller] Error: Text task has no content/message to speak.")
            
        elif task.type == TaskType.EMERGENCY:
             # UPDATED EMERGENCY SCRIPT
             script = ("Attention. This is an emergency alert. Please remain calm and follow the instructions carefully. "
                       "The situation is urgent. Stay tuned for further information.")
             # Emergency doesn't use Intro Chime
             audio_service.play_text(script)
        # --- AUDIO OUTPUT END ---

    def _apply_queue_shift(self):
        """Shifts all queued items by the duration of the High Priority Interruption"""
        if self.pause_start_time:
            now = datetime.now()
            duration = now - self.pause_start_time
            print(f"[Controller] Applying Time Shift: +{duration}")
            
            batch = db.batch()
            updated_count = 0

            for task in self.queue:
                task.scheduled_time += duration
                
                # Update Firestore so UI reflects new time
                try:
                     ref = db.collection('schedules').document(task.id)
                     new_date = task.scheduled_time.strftime("%Y-%m-%d")
                     new_time = task.scheduled_time.strftime("%H:%M")
                     batch.update(ref, {'date': new_date, 'time': new_time})
                     updated_count += 1
                except ValueError:
                    pass # Skip if invalid ID/Doc

            # Sort again just in case (though relative order shouldn't change)
            self.queue.sort(key=lambda x: x.scheduled_time)
            
            if updated_count > 0:
                try:
                    batch.commit()
                    print(f"[Controller] Persisted shift for {updated_count} schedules")
                except Exception as e:
                    print(f"[Controller] Batch update failed: {e}")

            self.pause_start_time = None

    def _update_firestore_state(self, task, priority, mode):
        try:
            data = {
                'active_task': task.to_dict() if task else None,
                'priority': int(priority),
                'mode': mode,
                'timestamp': firestore.SERVER_TIMESTAMP
            }
            db.collection('system').document('state').set(data)
        except Exception as e:
            print(f"[Controller] DB Error: {e}")

    # --- SCHEDULER LOOP ---
    def _scheduler_loop(self):
        while self._running:
            time.sleep(1) # Tiick every 1s
            
            # --- OPTIMIZATION: PERIODIC CLEANUP (Every 24 Hours) ---
            if (datetime.now() - self.last_cleanup).total_seconds() > 86400:
                self._cleanup_old_data()
                self.last_cleanup = datetime.now()

            with self._lock:
                # 1. Check for Due Tasks
                now = datetime.now()
                candidates = [t for t in self.queue if t.scheduled_time <= now]
                
                if not candidates:
                    continue

                next_task = candidates[0]
                
                # 2. Priority Check
                # If system is busy with Higher or Equal priority, wait.
                if self.current_task and self.current_task.priority >= next_task.priority:
                    continue
                    
                # 3. Promote & Execute
                self.queue.remove(next_task)
                next_task.priority = Priority.SCHEDULE # Ensure it has correct priority
                
                print(f"[Scheduler] Promoting Schedule {next_task.id}")
                
                # Mark as Completed in DB (for the specific instance)
                try:
                    db.collection('schedules').document(next_task.id).update({'status': 'Completed'})
                except Exception as e:
                    print(f"[Scheduler] Failed to mark completed: {e}")

                # Preempt lower priority if needed
                if self.current_task:
                    self._preempt_current_task(next_task.priority)

                self._start_task(next_task)
                
                # NEW: RECURRENCE LOGIC (Daily/Weekly)
                self._handle_recurrence(next_task)

    def _handle_recurrence(self, task: Task):
        """Checks if task needs to repeat and schedules the next instance"""
        repeat = task.data.get('repeat', 'once').lower()
        if repeat == 'once':
            return
            
        print(f"[Scheduler] Processing Recurrence: {repeat}")
        
        # FIX: DRIFT PREVENTION
        # Instead of adding 24h to the 'current executing time' (which might be delayed),
        # we calculate the next date but force the Original Time.
        
        try:
            # 1. Get Original Time Metadata
            original_time_str = task.data.get('time') # e.g. "12:00"
            if not original_time_str:
                 # Fallback if missing
                 original_time_str = task.scheduled_time.strftime("%H:%M")

            # 2. Calculate Next Date
            current_date = task.scheduled_time.date()
            next_date = None
            
            if repeat == 'daily':
                next_date = current_date + timedelta(days=1)
            elif repeat == 'weekly':
                next_date = current_date + timedelta(weeks=1)

            if next_date:
                # 3. Combine Next Date + Original Time
                next_dt_str = f"{next_date.strftime('%Y-%m-%d')} {original_time_str}"
                next_time_obj = datetime.strptime(next_dt_str, "%Y-%m-%d %H:%M")
                
                # 4. Create New Task
                new_data = task.data.copy()
                new_data['date'] = next_date.strftime("%Y-%m-%d")
                new_data['time'] = original_time_str
                new_data['status'] = 'Pending'
                if 'id' in new_data: del new_data['id'] 
                
                # 5. Save & Queue
                _, new_ref = db.collection('schedules').add(new_data)
                new_id = new_ref.id
                print(f"[Scheduler] Created recurring instance: {new_id} for {new_data['date']} at {new_data['time']}")
                
                new_task = Task(
                    id=new_id,
                    type=TaskType.SCHEDULE,
                    priority=Priority.SCHEDULE,
                    data=new_data,
                    scheduled_time=next_time_obj
                )
                self._add_to_queue(new_task)
                
        except Exception as e:
            print(f"[Scheduler] Recurrence Failed: {e}")

    def _cleanup_old_data(self):
        """Optimization: Garbage Collect old data to keep DB lean"""
        print("[Controller] Running Daily Cleanup...")
        try:
             # 1. Delete Old Logs (> 7 Days)
             cutoff = datetime.now() - timedelta(days=7)
             logs = db.collection('logs').where('timestamp', '<', cutoff).limit(100).stream()
             batch = db.batch()
             count = 0
             for doc in logs:
                 batch.delete(doc.reference)
                 count += 1
             
             # 2. Delete Completed Schedules (> 7 Days)
             # Note: 'schedules' store strings for date/time usually, so verify format or query carefully.
             # Ideally we rely on status='Completed'. 
             # For simpler logic, we'll just check status for now, or skip if complex date parsing needed query-side.
             # Let's just do Logs for now to be safe from deleting future recurrents by accident.
             
             if count > 0:
                 batch.commit()
                 print(f"[Controller] Cleanup: Deleted {count} old log entries.")
             else:
                 print("[Controller] Cleanup: No old data to delete.")
                 
        except Exception as e:
            print(f"[Controller] Cleanup Failed: {e}")

# Global Instance
controller = PAController()

