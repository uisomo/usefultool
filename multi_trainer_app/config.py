"""
Configuration for trainers, app settings, and video output.

Trainers are defined as a list of dicts - add/remove entries to change
the number of trainers. Everything is driven by this config.
"""

import os
from dataclasses import dataclass, field
from dotenv import load_dotenv

load_dotenv()


@dataclass
class TrainerProfile:
    name: str
    personality: str  # System prompt personality description
    voice: str  # edge-tts voice name
    color: tuple  # RGB puppet body color
    accent_color: tuple  # RGB accent (hair, accessories)
    emoji: str  # Identifier emoji for logs


# Default trainers - modify this list to change trainer count
DEFAULT_TRAINERS = [
    TrainerProfile(
        name="Coach Rex",
        personality=(
            "You are Coach Rex, an energetic and motivational fitness trainer. "
            "You speak with high energy, use action words, and push people to "
            "give their best. Keep responses concise (2-3 sentences max)."
        ),
        voice="en-US-GuyNeural",
        color=(220, 60, 60),       # Red
        accent_color=(180, 40, 40),
        emoji="🔴",
    ),
    TrainerProfile(
        name="Sage Luna",
        personality=(
            "You are Sage Luna, a calm and wise mindfulness trainer. "
            "You speak softly, use metaphors, and guide people toward "
            "inner balance. Keep responses concise (2-3 sentences max)."
        ),
        voice="en-US-JennyNeural",
        color=(60, 160, 220),      # Blue
        accent_color=(40, 120, 180),
        emoji="🔵",
    ),
    TrainerProfile(
        name="Dr. Vex",
        personality=(
            "You are Dr. Vex, a sharp and analytical strategy trainer. "
            "You speak precisely, cite logic, and break problems into steps. "
            "Keep responses concise (2-3 sentences max)."
        ),
        voice="en-US-AriaNeural",
        color=(100, 200, 100),     # Green
        accent_color=(60, 160, 60),
        emoji="🟢",
    ),
]


@dataclass
class AppConfig:
    # --- LLM ---
    llm_base_url: str = "https://openrouter.ai/api/v1"
    llm_api_key: str = ""
    llm_model: str = "deepseek/deepseek-r1-distill-llama-70b"

    # --- Video output ---
    video_width: int = 1920
    video_height: int = 1080
    fps: int = 30
    output_dir: str = "output"

    # --- Audio ---
    sample_rate: int = 24000

    # --- Session ---
    max_turns: int = 50  # Safety cap on conversation turns
    silence_timeout: float = 30.0  # Seconds of silence before auto-ending

    # --- Trainers ---
    trainers: list = field(default_factory=lambda: list(DEFAULT_TRAINERS))

    # --- User puppet ---
    user_name: str = "You"
    user_color: tuple = (240, 200, 60)       # Yellow
    user_accent_color: tuple = (200, 160, 40)

    def __post_init__(self):
        if not self.llm_api_key:
            self.llm_api_key = os.getenv("OPENROUTER_API_KEY", "")
