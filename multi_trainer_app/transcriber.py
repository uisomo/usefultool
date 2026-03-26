"""
Real-time Speech-to-Text Transcriber

Uses speech_recognition for microphone capture and either:
  - OpenAI Whisper (local) for transcription
  - Google Speech Recognition (fallback, free)

Transcribed text is instantly available via a callback.
"""

import io
import queue
import threading
import wave
import numpy as np

try:
    import speech_recognition as sr
except ImportError:
    sr = None

try:
    import whisper
except ImportError:
    whisper = None


class Transcriber:
    """
    Captures microphone audio and transcribes in real-time.

    Usage:
        def on_text(text, audio_bytes):
            print(f"User said: {text}")

        t = Transcriber(callback=on_text)
        t.start()
        ...
        t.stop()
    """

    def __init__(
        self,
        callback=None,
        whisper_model: str = "base",
        use_whisper: bool = True,
        energy_threshold: int = 300,
        pause_threshold: float = 1.5,
    ):
        """
        Args:
            callback: function(text: str, audio_wav_bytes: bytes) called on each utterance
            whisper_model: Whisper model size (tiny/base/small/medium/large)
            use_whisper: If True, use local Whisper; else use Google Speech API
            energy_threshold: Mic sensitivity
            pause_threshold: Seconds of silence to consider end of utterance
        """
        self.callback = callback
        self.use_whisper = use_whisper and whisper is not None
        self._whisper_model = None
        self._whisper_model_name = whisper_model
        self._running = False
        self._thread = None
        self._audio_queue = queue.Queue()

        if sr is None:
            raise ImportError("speech_recognition is required: pip install SpeechRecognition")

        self.recognizer = sr.Recognizer()
        self.recognizer.energy_threshold = energy_threshold
        self.recognizer.pause_threshold = pause_threshold
        self.recognizer.dynamic_energy_threshold = True

    def _load_whisper(self):
        if self.use_whisper and self._whisper_model is None:
            print(f"[Transcriber] Loading Whisper model '{self._whisper_model_name}'...")
            self._whisper_model = whisper.load_model(self._whisper_model_name)
            print("[Transcriber] Whisper model loaded.")

    def _transcribe_audio(self, audio: "sr.AudioData") -> str:
        """Transcribe an AudioData object to text."""
        if self.use_whisper:
            self._load_whisper()
            # Convert to WAV bytes for Whisper
            wav_bytes = audio.get_wav_data()
            # Write to temp buffer and transcribe
            import tempfile, os
            with tempfile.NamedTemporaryFile(suffix=".wav", delete=False) as f:
                f.write(wav_bytes)
                tmp_path = f.name
            try:
                result = self._whisper_model.transcribe(tmp_path)
                return result.get("text", "").strip()
            finally:
                os.unlink(tmp_path)
        else:
            # Fallback to Google Speech Recognition
            try:
                return self.recognizer.recognize_google(audio)
            except sr.UnknownValueError:
                return ""
            except sr.RequestError as e:
                print(f"[Transcriber] Google API error: {e}")
                return ""

    def _get_wav_bytes(self, audio: "sr.AudioData") -> bytes:
        """Extract raw WAV bytes from AudioData for recording."""
        return audio.get_wav_data()

    def _listen_loop(self):
        """Background thread: listen → transcribe → callback."""
        mic = sr.Microphone()
        with mic as source:
            print("[Transcriber] Adjusting for ambient noise...")
            self.recognizer.adjust_for_ambient_noise(source, duration=1)
            print("[Transcriber] Listening... (speak now)")

        stop_listening = self.recognizer.listen_in_background(
            mic, self._audio_callback, phrase_time_limit=30
        )
        self._stop_listening = stop_listening

        # Keep thread alive while running
        while self._running:
            try:
                audio = self._audio_queue.get(timeout=0.5)
                text = self._transcribe_audio(audio)
                if text and self.callback:
                    wav_bytes = self._get_wav_bytes(audio)
                    self.callback(text, wav_bytes)
            except queue.Empty:
                continue

    def _audio_callback(self, recognizer, audio):
        """Called by background listener when audio is captured."""
        if self._running:
            self._audio_queue.put(audio)

    def start(self):
        """Start listening in background."""
        if self._running:
            return
        self._running = True
        self._thread = threading.Thread(target=self._listen_loop, daemon=True)
        self._thread.start()
        print("[Transcriber] Started.")

    def stop(self):
        """Stop listening."""
        self._running = False
        if hasattr(self, "_stop_listening"):
            self._stop_listening(wait_for_stop=False)
        if self._thread:
            self._thread.join(timeout=3)
        print("[Transcriber] Stopped.")

    def transcribe_file(self, audio_path: str) -> str:
        """Transcribe an audio file (for testing / offline mode)."""
        with sr.AudioFile(audio_path) as source:
            audio = self.recognizer.record(source)
        return self._transcribe_audio(audio)
