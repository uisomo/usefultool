"""
Message Router

Determines which trainer(s) should respond to user input based on:
  1. Explicit name mentions ("Coach Rex, what do you think?")
  2. Keywords like "all", "everyone", "team"
  3. Default: picks the most relevant trainer or round-robins

The router is intentionally simple - it parses the user's text and
returns a list of trainer names that should respond.
"""

from config import TrainerProfile


class Router:
    """Routes user messages to the appropriate trainer(s)."""

    # Keywords that trigger all trainers
    ALL_KEYWORDS = {"all", "everyone", "team", "together", "group", "all of you"}

    def __init__(self, trainers: list[TrainerProfile]):
        self.trainers = trainers
        self._turn_index = 0  # For round-robin fallback

    def route(self, user_text: str) -> list[TrainerProfile]:
        """
        Determine which trainer(s) should respond.

        Returns:
            List of TrainerProfile objects that should respond.
        """
        text_lower = user_text.lower().strip()

        # 1. Check for "all" keywords
        for kw in self.ALL_KEYWORDS:
            if kw in text_lower:
                return list(self.trainers)

        # 2. Check for name mentions
        mentioned = []
        for trainer in self.trainers:
            # Check full name and first name
            full_name = trainer.name.lower()
            first_name = full_name.split()[0] if " " in full_name else full_name
            # Also check last name/word
            last_name = full_name.split()[-1] if " " in full_name else ""

            if full_name in text_lower or first_name in text_lower:
                mentioned.append(trainer)
            elif last_name and last_name in text_lower:
                mentioned.append(trainer)

        if mentioned:
            return mentioned

        # 3. Default: round-robin single trainer
        trainer = self.trainers[self._turn_index % len(self.trainers)]
        self._turn_index += 1
        return [trainer]

    def reset(self):
        """Reset round-robin counter."""
        self._turn_index = 0
