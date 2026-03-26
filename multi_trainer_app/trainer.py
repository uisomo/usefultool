"""
AI Trainer

Each trainer is backed by an LLM (via OpenRouter/OpenAI-compatible API).
Maintains per-trainer conversation history for context continuity.
"""

from openai import OpenAI
from config import TrainerProfile, AppConfig


class TrainerAI:
    """Manages LLM interaction for a single trainer."""

    def __init__(self, profile: TrainerProfile, config: AppConfig):
        self.profile = profile
        self.config = config
        self.client = OpenAI(
            base_url=config.llm_base_url,
            api_key=config.llm_api_key,
        )
        self.history: list[dict] = [
            {"role": "system", "content": profile.personality}
        ]

    def respond(self, user_text: str, group_context: str = "") -> str:
        """
        Generate a response from this trainer.

        Args:
            user_text: What the user said
            group_context: Recent messages from other trainers (for group awareness)

        Returns:
            The trainer's text response
        """
        # Build the user message with optional group context
        message = user_text
        if group_context:
            message = (
                f"[Group conversation context:\n{group_context}]\n\n"
                f"User says to you: {user_text}"
            )

        self.history.append({"role": "user", "content": message})

        try:
            completion = self.client.chat.completions.create(
                extra_headers={
                    "HTTP-Referer": "http://localhost",
                    "X-Title": "Multi-Trainer App",
                },
                model=self.config.llm_model,
                messages=self.history,
                max_tokens=200,
                temperature=0.8,
            )
            response = completion.choices[0].message.content.strip()
        except Exception as e:
            response = f"[{self.profile.name} is having trouble responding: {e}]"

        self.history.append({"role": "assistant", "content": response})

        # Keep history manageable (last 20 exchanges)
        if len(self.history) > 41:  # system + 20 pairs
            self.history = [self.history[0]] + self.history[-40:]

        return response

    def reset(self):
        """Clear conversation history."""
        self.history = [self.history[0]]


class TrainerTeam:
    """Manages the full team of trainers."""

    def __init__(self, config: AppConfig):
        self.config = config
        self.trainers: dict[str, TrainerAI] = {}
        for profile in config.trainers:
            self.trainers[profile.name] = TrainerAI(profile, config)

        # Shared conversation log for group context
        self.conversation_log: list[dict] = []

    def get_trainer(self, name: str) -> TrainerAI:
        return self.trainers[name]

    def get_group_context(self, limit: int = 5) -> str:
        """Get recent conversation as context string."""
        if not self.conversation_log:
            return ""
        recent = self.conversation_log[-limit:]
        lines = [f"{entry['speaker']}: {entry['text']}" for entry in recent]
        return "\n".join(lines)

    def log_message(self, speaker: str, text: str):
        """Log a message to the shared conversation."""
        self.conversation_log.append({"speaker": speaker, "text": text})

    def reset_all(self):
        """Reset all trainers."""
        for t in self.trainers.values():
            t.reset()
        self.conversation_log.clear()
