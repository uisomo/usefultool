"""
Text-to-Speech Engine

Uses edge-tts (Microsoft Edge's free TTS) for high-quality, multi-voice
speech synthesis. Each trainer gets a distinct voice.

Falls back to gTTS or pyttsx3 if edge-tts is unavailable.
"""

import asyncio
import os
import tempfile
from pathlib import Path

# Try edge-tts first (best quality, free, many voices)
try:
    import edge_tts
    HAS_EDGE_TTS = True
except ImportError:
    HAS_EDGE_TTS = False

# Fallback: gTTS
try:
    from gtts import gTTS
    HAS_GTTS = True
except ImportError:
    HAS_GTTS = False

# Fallback: pyttsx3 (offline)
try:
    import pyttsx3
    HAS_PYTTSX3 = True
except ImportError:
    HAS_PYTTSX3 = False


class TTSEngine:
    """Generate speech audio files from text with per-trainer voices."""

    def __init__(self, output_dir: str = "output/tts_cache"):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self._counter = 0

        if HAS_EDGE_TTS:
            self.backend = "edge-tts"
        elif HAS_GTTS:
            self.backend = "gtts"
        elif HAS_PYTTSX3:
            self.backend = "pyttsx3"
        else:
            raise ImportError(
                "No TTS backend available. Install one of: edge-tts, gTTS, pyttsx3"
            )
        print(f"[TTS] Using backend: {self.backend}")

    def synthesize(self, text: str, voice: str = "en-US-GuyNeural") -> str:
        """
        Convert text to speech audio file.

        Args:
            text: Text to speak
            voice: Voice identifier (edge-tts voice name)

        Returns:
            Path to the generated audio file (.mp3)
        """
        self._counter += 1
        output_path = str(self.output_dir / f"tts_{self._counter:04d}.mp3")

        if self.backend == "edge-tts":
            return self._edge_tts(text, voice, output_path)
        elif self.backend == "gtts":
            return self._gtts(text, output_path)
        else:
            return self._pyttsx3(text, output_path)

    def _edge_tts(self, text: str, voice: str, output_path: str) -> str:
        """Synthesize using edge-tts."""
        async def _synth():
            communicate = edge_tts.Communicate(text, voice)
            await communicate.save(output_path)

        # Run async in sync context
        try:
            loop = asyncio.get_event_loop()
            if loop.is_running():
                import concurrent.futures
                with concurrent.futures.ThreadPoolExecutor() as pool:
                    pool.submit(lambda: asyncio.run(_synth())).result()
            else:
                loop.run_until_complete(_synth())
        except RuntimeError:
            asyncio.run(_synth())

        return output_path

    def _gtts(self, text: str, output_path: str) -> str:
        """Synthesize using gTTS (Google)."""
        tts = gTTS(text=text, lang="en", slow=False)
        tts.save(output_path)
        return output_path

    def _pyttsx3(self, text: str, output_path: str) -> str:
        """Synthesize using pyttsx3 (offline)."""
        engine = pyttsx3.init()
        # pyttsx3 can save to file
        engine.save_to_file(text, output_path)
        engine.runAndWait()
        return output_path

    def cleanup(self):
        """Remove cached TTS files."""
        for f in self.output_dir.glob("tts_*.mp3"):
            f.unlink()
