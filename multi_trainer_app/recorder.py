"""
Session Recorder

Records the entire training session as a horizontal 16:9 YouTube video.
Captures:
  - Video frames from SceneComposer (puppet animations + UI)
  - Audio from user mic and trainer TTS responses
  - Syncs audio with video frames

Uses MoviePy for final video assembly.
"""

import os
import time
import tempfile
from pathlib import Path
from PIL import Image
import numpy as np

from moviepy import (
    ImageClip,
    AudioFileClip,
    CompositeAudioClip,
    concatenate_videoclips,
    concatenate_audioclips,
)


class SessionSegment:
    """One segment of the session (a single speaker's turn)."""

    def __init__(
        self,
        speaker: str,
        text: str,
        frame_idle: Image.Image,
        frame_speaking: Image.Image,
        audio_path: str = None,
        duration: float = 2.0,
    ):
        self.speaker = speaker
        self.text = text
        self.frame_idle = frame_idle
        self.frame_speaking = frame_speaking
        self.audio_path = audio_path
        self.duration = duration  # Will be updated based on audio length


class SessionRecorder:
    """Records and exports the full training session as MP4."""

    def __init__(self, output_dir: str = "output", fps: int = 30):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.fps = fps
        self.segments: list[SessionSegment] = []

    def add_segment(self, segment: SessionSegment):
        """Add a recorded segment to the session."""
        # If there's audio, use its duration
        if segment.audio_path and os.path.exists(segment.audio_path):
            try:
                audio = AudioFileClip(segment.audio_path)
                segment.duration = audio.duration + 0.5  # Small padding
                audio.close()
            except Exception:
                pass  # Keep default duration
        self.segments.append(segment)

    def export(self, filename: str = "training_session.mp4") -> str:
        """
        Export the full session as an MP4 video.

        Returns:
            Path to the exported video file.
        """
        if not self.segments:
            print("[Recorder] No segments to export.")
            return ""

        output_path = str(self.output_dir / filename)
        print(f"[Recorder] Exporting {len(self.segments)} segments to {output_path}...")

        video_clips = []
        audio_clips = []
        current_time = 0.0

        for i, seg in enumerate(self.segments):
            print(f"  Processing segment {i + 1}/{len(self.segments)}: {seg.speaker}")

            # Create video clip from the speaking frame
            # Use speaking frame for first 80% of duration, idle for last 20%
            speaking_duration = seg.duration * 0.85
            idle_duration = seg.duration * 0.15

            # Speaking part
            speaking_array = np.array(seg.frame_speaking.convert("RGB"))
            speaking_clip = ImageClip(speaking_array, duration=speaking_duration)

            # Idle part (mouth closed at end)
            idle_array = np.array(seg.frame_idle.convert("RGB"))
            idle_clip = ImageClip(idle_array, duration=idle_duration)

            # Combine
            segment_clip = concatenate_videoclips(
                [speaking_clip, idle_clip], method="compose"
            )
            video_clips.append(segment_clip)

            # Audio
            if seg.audio_path and os.path.exists(seg.audio_path):
                try:
                    audio = AudioFileClip(seg.audio_path)
                    audio = audio.with_start(current_time)
                    audio_clips.append(audio)
                except Exception as e:
                    print(f"  Warning: Could not load audio for segment {i}: {e}")

            current_time += seg.duration

        # Concatenate all video
        final_video = concatenate_videoclips(video_clips, method="compose")

        # Composite all audio
        if audio_clips:
            final_audio = CompositeAudioClip(audio_clips)
            final_video = final_video.with_audio(final_audio)

        # Write final video
        final_video.write_videofile(
            output_path,
            fps=self.fps,
            codec="libx264",
            audio_codec="aac",
            threads=4,
            preset="medium",
            logger="bar",
        )

        # Cleanup
        final_video.close()
        for clip in video_clips:
            clip.close()

        print(f"[Recorder] Exported: {output_path}")
        return output_path

    def clear(self):
        """Clear all recorded segments."""
        self.segments.clear()
