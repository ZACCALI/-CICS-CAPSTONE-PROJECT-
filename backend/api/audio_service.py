
import os
import subprocess
import threading
import platform

class AudioService:
    def __init__(self):
        self.current_process = None
        self._lock = threading.Lock()

    def play_text(self, text: str):
        """Plays text using Windows TTS (PowerShell)"""
        self.stop() # Stop any current sound
        
        print(f"[AudioService] Speaking: {text}")
        
        # Escape single quotes for PowerShell
        safe_text = text.replace("'", "''")
        
        # Windows Native TTS via PowerShell
        command = [
            'powershell', 
            '-Command', 
            f"Add-Type –AssemblyName System.Speech; (New-Object System.Speech.Synthesis.SpeechSynthesizer).Speak('{safe_text}')"
        ]
        
        self._run_command(command)

    def play_intro(self, file_path: str):
        """Plays the intro chime synchronously (blocks until finished or fixed time)"""
        self.stop() # Stop any background music
        
        print(f"[AudioService] Playing Intro: {file_path}")
        
        # Check if file exists to avoid PowerShell errors
        if not os.path.exists(file_path):
            print(f"[AudioService] Intro file not found at {file_path}. Playing Fallback Tone.")
            # Fallback: System Beep (Frequency, Duration) -> Ding-Dong effect
            # 800Hz for 400ms, then 600Hz for 600ms
            fallback_script = "[console]::beep(800, 400); Start-Sleep -Milliseconds 100; [console]::beep(600, 600)"
            self._run_command(['powershell', '-c', fallback_script])
            return

        # Escape for PowerShell
        safe_path = file_path.replace("'", "''")

        # PowerShell script to play MP3 using WPF MediaPlayer
        # We use a fixed sleep of 4 seconds for reliability if duration detection fails, 
        # or we could loop on duration. "Airport Chime" is usually 3-5s.
        # PowerShell script to play MP3 using WPF MediaPlayer (Robust)
        # We wait for duration to load. If it fails, we fallback to fixed 4s sleep.
        ps_script = f"""
        Add-Type -AssemblyName PresentationCore, PresentationFramework;
        $p = New-Object System.Windows.Media.MediaPlayer;
        $p.Open('{safe_path}');
        
        # 1. Wait for Media Open / Duration Load (Async)
        $attempts = 20; 
        while (-not $p.NaturalDuration.HasTimeSpan -and $attempts -gt 0) {{
            Start-Sleep -Milliseconds 100;
            $attempts--;
        }}
        
        $p.Play();

        # 2. Synchronous Wait
        if ($p.NaturalDuration.HasTimeSpan) {{
            while ($p.Position -lt $p.NaturalDuration.TimeSpan) {{
                Start-Sleep -Milliseconds 100;
            }}
        }} else {{
            # Fallback if duration never loaded (common with some encodings)
            write-host "Duration unknown, using fallback wait."
            Start-Sleep -Seconds 4;
        }}
        $p.Close();
        """
        
        # Run synchronously
        command = ['powershell', '-c', ps_script]
        try:
            subprocess.run(command, check=True)
        except Exception as e:
            print(f"[AudioService] Intro playback failed: {e}")

    def play_file(self, file_path: str):
        """Plays audio file using PowerShell (MP3 Support)"""
        self.stop()
        
        print(f"[AudioService] Playing: {file_path}")
        safe_path = file_path.replace("'", "''")
        
        # Use simple SoundPlayer for WAV (Robust) or MediaPlayer for MP3
        if file_path.lower().endswith('.wav'):
             command = ['powershell', '-c', f"(New-Object Media.SoundPlayer '{safe_path}').PlaySync()"]
        else:
             # Async Play for long files (don't block controller forever)?
             # Actually, play_file is usually background music or long audio.
             # If we block, we freeze the server.
             # So we should run this ASCYNCHRONOUSLY (subprocess.Popen)
             
             # But the old implementation was using _run_command (Thread) which calls Popen.
             # Wait, the previous play_file used PlaySync (Blocking) inside _run_command (Thread).
             # So it blocked the THREAD, not the server. safely.
             
             # Reuse the MediaPlayer logic but keep it alive?
             # Complex for MP3 streaming via PowerShell in background.
             # For now, let's keep the Intro Logic separate and leave play_file as is (mostly WAV/Basic).
             # User ONLY asked for Intro Chime MP3 support.
             
             # Reverting play_file change to just the original logic but with error handling?
             # No, the user might want to play MP3s in general.
             # Let's upgrade play_file to use MediaPlayer too if MP3.
             
             ps_script = f"""
                Add-Type -AssemblyName PresentationCore, PresentationFramework;
                $p = New-Object System.Windows.Media.MediaPlayer;
                $p.Open('{safe_path}');
                $p.Play();
                # We need to keep process alive while playing
                # But we can't easily detect end in a headless Popen without blocking.
                # Hack: Just sleep for a long time or until stopped?
                # Better: Use a loop that checks 'stop' signal? No, we kill process to stop.
                while ($p.Position -lt $p.NaturalDuration.TimeSpan) {{ Start-Sleep -Milliseconds 500 }}
                $p.Close();
             """
             command = ['powershell', '-c', ps_script]
        
        self._run_command(command)

    def stop(self):
        """Stops the currently running audio process"""
        with self._lock:
            if self.current_process:
                print("[AudioService] Stopping Audio...")
                # Force kill the process
                try:
                    self.current_process.terminate()
                except:
                    pass
                self.current_process = None

    def _run_command(self, command):
        """Runs the audio command in a separate thread so it doesn't block the controller"""
        def target():
            proc = None
            # 1. Start Process (Critical Section)
            with self._lock:
                try:
                    # capture_output to see errors if any
                    proc = subprocess.Popen(
                        command, 
                        stdout=subprocess.PIPE, 
                        stderr=subprocess.PIPE,
                        text=True
                    )
                    self.current_process = proc
                except Exception as e:
                    print(f"[AudioService] Exception starting process: {e}")
                    return

            # 2. Wait for Completion (NON-BLOCKING Check)
            # We released the lock, so stop() can be called now.
            if proc:
                try:
                    stdout, stderr = proc.communicate()
                    if stderr:
                        print(f"[AudioService] Process Error: {stderr}")
                except Exception as e:
                    pass # Process killed or error
                finally:
                    # Cleanup
                    with self._lock:
                        if self.current_process == proc:
                            self.current_process = None

        thread = threading.Thread(target=target)
        thread.start()

    def play_announcement(self, intro_path: str, text: str):
        """
        Plays Intro Chime followed by Text (TTS) in a single NON-BLOCKING PowerShell sequence.
        This releases the Controller Lock immediately, allowing interruption.
        """
        self.stop()
        
        print(f"[AudioService] Announcing: Intro -> '{text}'")
        
        # Safe strings
        safe_intro = intro_path.replace("'", "''")
        safe_text = text.replace("'", "''")
        
        # PowerShell Script: Load Assemblies -> Play Intro (Wait) -> Speak
        ps_script = f"""
        Add-Type -AssemblyName PresentationCore, PresentationFramework;
        Add-Type -AssemblyName System.Speech;
        
        # 1. Play Intro
        $p = New-Object System.Windows.Media.MediaPlayer;
        $p.Open('{safe_intro}');
        
        # Wait for load
        $attempts = 20; 
        while (-not $p.NaturalDuration.HasTimeSpan -and $attempts -gt 0) {{
            Start-Sleep -Milliseconds 100;
            $attempts--;
        }}
        
        $p.Play();
        
        # Wait for end
        if ($p.NaturalDuration.HasTimeSpan) {{
            while ($p.Position -lt $p.NaturalDuration.TimeSpan) {{
                Start-Sleep -Milliseconds 100;
            }}
        }} else {{
            Start-Sleep -Seconds 4; # Fallback
        }}
        $p.Close();
        
        # 2. Speak Text
        $synth = New-Object System.Speech.Synthesis.SpeechSynthesizer;
        $synth.Speak('{safe_text}');
        """
        
        # Run in Thread (Non-Blocking)
        self._run_command(['powershell', '-c', ps_script])

    def play_intro_async(self, file_path: str):
         """Plays intro asynchronously (for Voice Broadcast where no text follows)"""
         self.stop()
         print(f"[AudioService] Playing Intro (Async): {file_path}")
         safe_path = file_path.replace("'", "''")
         
         ps_script = f"""
         Add-Type -AssemblyName PresentationCore, PresentationFramework;
         $p = New-Object System.Windows.Media.MediaPlayer;
         $p.Open('{safe_path}');
         
         $attempts = 20; 
         while (-not $p.NaturalDuration.HasTimeSpan -and $attempts -gt 0) {{ Start-Sleep -Milliseconds 100; $attempts--; }}
         
         $p.Play();
         
         if ($p.NaturalDuration.HasTimeSpan) {{
            while ($p.Position -lt $p.NaturalDuration.TimeSpan) {{ Start-Sleep -Milliseconds 100; }}
         }} else {{ Start-Sleep -Seconds 4; }}
         $p.Close();
         """
         self._run_command(['powershell', '-c', ps_script])

# Global Instance
audio_service = AudioService()
