#!/usr/bin/env python3
"""
Multi-Trainer App - Main Orchestrator

Runs an interactive training session where you speak to a team of AI trainers.
Your speech is transcribed in real-time, routed to the appropriate trainer(s),
and they respond via audio with animated puppet avatars.

The entire session is recorded as a horizontal 16:9 YouTube-ready video.

Usage:
    # Interactive mode (live mic):
    python main.py

    # With a specific topic:
    python main.py --topic "How to build good habits"

    # Offline / demo mode (text input instead of mic):
    python main.py --demo

    # Custom number of trainers:
    python main.py --trainers 5
"""

import argparse
import os
import sys
import threading
import time
from pathlib import Path

# Add parent dir to path for imports
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config import AppConfig, TrainerProfile, DEFAULT_TRAINERS
from transcriber import Transcriber
from router import Router
from trainer import TrainerTeam
from tts_engine import TTSEngine
from puppet import PuppetFactory
from composer import SceneComposer
from recorder import SessionRecorder, SessionSegment


# Extra trainer templates for when user wants more than 3
EXTRA_TRAINERS = [
    TrainerProfile(
        name="Captain Blaze",
        personality=(
            "You are Captain Blaze, a bold leadership and teamwork trainer. "
            "You use military-style motivation, emphasize discipline and "
            "camaraderie. Keep responses concise (2-3 sentences max)."
        ),
        voice="en-US-DavisNeural",
        color=(220, 150, 30),
        accent_color=(180, 120, 20),
        emoji="🟡",
    ),
    TrainerProfile(
        name="Prof. Aria",
        personality=(
            "You are Prof. Aria, an academic and research-oriented trainer. "
            "You cite studies, use data-driven reasoning, and encourage "
            "critical thinking. Keep responses concise (2-3 sentences max)."
        ),
        voice="en-US-SaraNeural",
        color=(180, 80, 220),
        accent_color=(140, 50, 180),
        emoji="🟣",
    ),
    TrainerProfile(
        name="Zen Master Kai",
        personality=(
            "You are Zen Master Kai, a philosophical trainer who draws from "
            "Eastern and Western wisdom traditions. You use koans, paradoxes, "
            "and deep questions. Keep responses concise (2-3 sentences max)."
        ),
        voice="en-US-TonyNeural",
        color=(60, 180, 180),
        accent_color=(40, 140, 140),
        emoji="🩵",
    ),
    TrainerProfile(
        name="Nova Storm",
        personality=(
            "You are Nova Storm, a creative and unconventional trainer. "
            "You think outside the box, use humor, and challenge assumptions. "
            "Keep responses concise (2-3 sentences max)."
        ),
        voice="en-US-JaneNeural",
        color=(220, 100, 180),
        accent_color=(180, 60, 140),
        emoji="🩷",
    ),
]


def build_trainer_list(count: int) -> list[TrainerProfile]:
    """Build a list of trainer profiles for the requested count."""
    all_trainers = list(DEFAULT_TRAINERS) + EXTRA_TRAINERS
    if count <= 0:
        count = 3
    if count > len(all_trainers):
        # Repeat trainers with modified names if we need more
        extended = list(all_trainers)
        i = 0
        while len(extended) < count:
            base = all_trainers[i % len(all_trainers)]
            suffix = len(extended) - len(all_trainers) + 1
            extended.append(
                TrainerProfile(
                    name=f"{base.name} #{suffix + 1}",
                    personality=base.personality,
                    voice=base.voice,
                    color=tuple(min(255, c + suffix * 20) for c in base.color),
                    accent_color=base.accent_color,
                    emoji=base.emoji,
                )
            )
            i += 1
        return extended[:count]
    return all_trainers[:count]


class MultiTrainerSession:
    """Orchestrates the full interactive training session."""

    def __init__(self, config: AppConfig, topic: str = ""):
        self.config = config
        self.topic = topic
        self.running = False

        # Core components
        self.router = Router(config.trainers)
        self.team = TrainerTeam(config)
        self.tts = TTSEngine(output_dir=os.path.join(config.output_dir, "tts_cache"))
        self.puppet_factory = PuppetFactory()
        self.composer = SceneComposer(config, self.puppet_factory)
        self.recorder = SessionRecorder(output_dir=config.output_dir, fps=config.fps)

        # Session state
        self._turn_count = 0
        self._lock = threading.Lock()

    def _process_user_input(self, text: str, audio_wav_bytes: bytes = None):
        """Process a single user utterance through the full pipeline."""
        with self._lock:
            self._turn_count += 1
            turn = self._turn_count

        print(f"\n{'='*60}")
        print(f"[Turn {turn}] You: {text}")
        print(f"{'='*60}")

        # Log user message
        self.team.log_message(self.config.user_name, text)

        # Record user segment (no TTS audio - user spoke live)
        user_frame_idle = self.composer.compose_frame(
            topic=self.topic,
            subtitle_speaker=self.config.user_name,
            subtitle_text=text,
            speaking_name="",
        )
        user_frame_speaking = self.composer.compose_frame(
            topic=self.topic,
            subtitle_speaker=self.config.user_name,
            subtitle_text=text,
            speaking_name=self.config.user_name,
        )

        # Save user audio if available
        user_audio_path = None
        if audio_wav_bytes:
            user_audio_path = os.path.join(
                self.config.output_dir, "tts_cache", f"user_{turn:04d}.wav"
            )
            with open(user_audio_path, "wb") as f:
                f.write(audio_wav_bytes)

        self.recorder.add_segment(SessionSegment(
            speaker=self.config.user_name,
            text=text,
            frame_idle=user_frame_idle,
            frame_speaking=user_frame_speaking,
            audio_path=user_audio_path,
            duration=max(2.0, len(text.split()) * 0.4),  # Estimate if no audio
        ))

        # Route to trainer(s)
        targets = self.router.route(text)
        group_context = self.team.get_group_context()

        print(f"[Router] Sending to: {', '.join(t.name for t in targets)}")

        # Each targeted trainer responds
        for trainer_profile in targets:
            trainer_ai = self.team.get_trainer(trainer_profile.name)

            # Get LLM response
            response = trainer_ai.respond(text, group_context)
            self.team.log_message(trainer_profile.name, response)

            print(f"\n{trainer_profile.emoji} {trainer_profile.name}: {response}")

            # Generate TTS audio
            audio_path = self.tts.synthesize(response, voice=trainer_profile.voice)

            # Compose video frames
            frame_idle = self.composer.compose_frame(
                topic=self.topic,
                subtitle_speaker=trainer_profile.name,
                subtitle_text=response,
                speaking_name="",
            )
            frame_speaking = self.composer.compose_frame(
                topic=self.topic,
                subtitle_speaker=trainer_profile.name,
                subtitle_text=response,
                speaking_name=trainer_profile.name,
            )

            # Record segment
            self.recorder.add_segment(SessionSegment(
                speaker=trainer_profile.name,
                text=response,
                frame_idle=frame_idle,
                frame_speaking=frame_speaking,
                audio_path=audio_path,
            ))

        # Check turn limit
        if self._turn_count >= self.config.max_turns:
            print(f"\n[Session] Reached max turns ({self.config.max_turns}). Ending.")
            self.running = False

    def run_interactive(self):
        """Run with live microphone input."""
        print("\n" + "=" * 60)
        print("  MULTI-TRAINER SESSION")
        print(f"  Topic: {self.topic or '(open discussion)'}")
        print(f"  Trainers: {', '.join(t.name for t in self.config.trainers)}")
        print("  Speak to interact. Say 'stop' or 'exit' to end.")
        print("=" * 60 + "\n")

        self.running = True

        def on_transcription(text: str, audio_bytes: bytes):
            if not self.running:
                return
            lower = text.lower().strip()
            if lower in ("stop", "exit", "quit", "end session", "goodbye"):
                print("\n[Session] Ending session...")
                self.running = False
                return
            if len(text.strip()) < 2:
                return
            self._process_user_input(text, audio_bytes)

        transcriber = Transcriber(callback=on_transcription)
        transcriber.start()

        try:
            while self.running:
                time.sleep(0.5)
        except KeyboardInterrupt:
            print("\n[Session] Interrupted by user.")
        finally:
            transcriber.stop()
            self._export_session()

    def run_demo(self):
        """Run in demo mode with text input (no mic required)."""
        print("\n" + "=" * 60)
        print("  MULTI-TRAINER SESSION (Demo Mode)")
        print(f"  Topic: {self.topic or '(open discussion)'}")
        print(f"  Trainers: {', '.join(t.name for t in self.config.trainers)}")
        print("  Type your messages. Type 'exit' to end.")
        print("  Tip: mention a trainer by name, or say 'all' for everyone.")
        print("=" * 60 + "\n")

        self.running = True

        try:
            while self.running:
                try:
                    user_input = input(f"\n[{self.config.user_name}] > ").strip()
                except EOFError:
                    break

                if not user_input:
                    continue
                if user_input.lower() in ("exit", "quit", "stop", "bye"):
                    break

                self._process_user_input(user_input)
        except KeyboardInterrupt:
            print("\n[Session] Interrupted.")
        finally:
            self.running = False
            self._export_session()

    def _export_session(self):
        """Export the recorded session to video."""
        if not self.recorder.segments:
            print("[Session] No segments recorded. Nothing to export.")
            return

        print(f"\n[Session] Exporting {len(self.recorder.segments)} segments...")
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        filename = f"training_session_{timestamp}.mp4"
        path = self.recorder.export(filename)
        if path:
            print(f"\n{'='*60}")
            print(f"  VIDEO SAVED: {path}")
            print(f"  Resolution: {self.config.video_width}x{self.config.video_height}")
            print(f"  Segments: {len(self.recorder.segments)}")
            print(f"{'='*60}\n")

        # Cleanup TTS cache
        self.tts.cleanup()


def main():
    parser = argparse.ArgumentParser(
        description="Multi-Trainer App - Interactive AI training with puppet avatars"
    )
    parser.add_argument(
        "--topic", "-t",
        type=str,
        default="",
        help="Discussion topic for the session",
    )
    parser.add_argument(
        "--trainers", "-n",
        type=int,
        default=3,
        help="Number of trainers (default: 3)",
    )
    parser.add_argument(
        "--demo", "-d",
        action="store_true",
        help="Run in demo mode (text input instead of mic)",
    )
    parser.add_argument(
        "--model", "-m",
        type=str,
        default="",
        help="LLM model to use (default: deepseek/deepseek-r1-distill-llama-70b)",
    )
    parser.add_argument(
        "--output", "-o",
        type=str,
        default="output",
        help="Output directory for videos",
    )
    args = parser.parse_args()

    # Build config
    trainers = build_trainer_list(args.trainers)
    config = AppConfig(
        trainers=trainers,
        output_dir=args.output,
    )
    if args.model:
        config.llm_model = args.model

    # Validate API key
    if not config.llm_api_key:
        print("ERROR: No API key found.")
        print("Set OPENROUTER_API_KEY environment variable or create a .env file.")
        sys.exit(1)

    # Run session
    session = MultiTrainerSession(config=config, topic=args.topic)

    if args.demo:
        session.run_demo()
    else:
        session.run_interactive()


if __name__ == "__main__":
    main()
