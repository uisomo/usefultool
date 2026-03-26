"""
Scene Composer

Arranges puppet avatars and UI elements into a 16:9 horizontal frame.

Layout (1920x1080):
┌─────────────────────────────────────────────────┐
│  TOPIC BAR (discussion topic text)              │
├─────────┬───────────────────────────────────────┤
│         │                                       │
│  USER   │   TRAINER 1  │  TRAINER 2  │ TRAINER 3│
│ PUPPET  │   PUPPET     │  PUPPET     │ PUPPET   │
│         │                                       │
├─────────┴───────────────────────────────────────┤
│  SUBTITLE BAR (current speaker's text)          │
└─────────────────────────────────────────────────┘

Puppets scale automatically based on trainer count.
"""

from PIL import Image, ImageDraw, ImageFont
from puppet import Puppet, PuppetFactory, _get_font
from config import AppConfig


class SceneComposer:
    """Composes a full video frame with puppets and UI."""

    def __init__(self, config: AppConfig, puppet_factory: PuppetFactory):
        self.config = config
        self.puppet_factory = puppet_factory
        self.width = config.video_width
        self.height = config.video_height

        # Layout zones
        self.topic_bar_height = 80
        self.subtitle_bar_height = 120
        self.puppet_zone_top = self.topic_bar_height
        self.puppet_zone_height = self.height - self.topic_bar_height - self.subtitle_bar_height

        # Pre-create user puppet
        self.user_puppet = puppet_factory.get_puppet(
            config.user_name, config.user_color, config.user_accent_color
        )

        # Pre-create trainer puppets
        self.trainer_puppets: dict[str, Puppet] = {}
        for t in config.trainers:
            self.trainer_puppets[t.name] = puppet_factory.get_puppet(
                t.name, t.color, t.accent_color
            )

    def compose_frame(
        self,
        topic: str = "",
        subtitle_speaker: str = "",
        subtitle_text: str = "",
        speaking_name: str = "",
    ) -> Image.Image:
        """
        Compose a single video frame.

        Args:
            topic: Current discussion topic (shown in top bar)
            subtitle_speaker: Name of current speaker
            subtitle_text: Current speech text (shown in bottom bar)
            speaking_name: Name of the entity currently speaking (for mouth animation)

        Returns:
            PIL Image (1920x1080 RGBA)
        """
        frame = Image.new("RGB", (self.width, self.height), (30, 30, 40))
        draw = ImageDraw.Draw(frame)

        # --- Topic bar ---
        self._draw_topic_bar(draw, topic)

        # --- Puppet zone ---
        self._draw_puppets(frame, speaking_name)

        # --- Subtitle bar ---
        self._draw_subtitle_bar(draw, subtitle_speaker, subtitle_text)

        return frame

    def _draw_topic_bar(self, draw: ImageDraw.Draw, topic: str):
        """Draw the topic bar at the top."""
        # Background
        draw.rectangle(
            [0, 0, self.width, self.topic_bar_height],
            fill=(20, 20, 30),
        )
        # Accent line
        draw.rectangle(
            [0, self.topic_bar_height - 3, self.width, self.topic_bar_height],
            fill=(100, 100, 200),
        )
        # Topic text
        if topic:
            font = _get_font(28)
            # Truncate if too long
            max_chars = 80
            display_text = topic[:max_chars] + "..." if len(topic) > max_chars else topic
            text_bbox = draw.textbbox((0, 0), display_text, font=font)
            text_w = text_bbox[2] - text_bbox[0]
            x = (self.width - text_w) // 2
            y = (self.topic_bar_height - (text_bbox[3] - text_bbox[1])) // 2
            draw.text((x, y), display_text, fill=(200, 200, 255), font=font)

    def _draw_puppets(self, frame: Image.Image, speaking_name: str):
        """Draw all puppets in the puppet zone."""
        num_trainers = len(self.trainer_puppets)
        total_slots = 1 + num_trainers  # user + trainers

        # Calculate puppet size to fit
        available_width = self.width
        slot_width = available_width // total_slots
        puppet_scale = min(1.0, slot_width / 320)  # 320 = puppet width + padding

        puppet_h = int(400 * puppet_scale)
        puppet_w = int(300 * puppet_scale)

        # Vertical center in puppet zone
        y_offset = self.puppet_zone_top + (self.puppet_zone_height - puppet_h) // 2

        # Draw user puppet (leftmost)
        user_frame = self.user_puppet.get_frame(
            speaking=speaking_name == self.config.user_name
        )
        if puppet_scale < 1.0:
            user_frame = user_frame.resize((puppet_w, puppet_h), Image.LANCZOS)
        x = slot_width // 2 - puppet_w // 2
        frame.paste(user_frame, (x, y_offset), user_frame)

        # Draw a subtle divider between user and trainers
        div_x = slot_width
        draw = ImageDraw.Draw(frame)
        draw.line(
            [(div_x, self.puppet_zone_top + 20),
             (div_x, self.puppet_zone_top + self.puppet_zone_height - 20)],
            fill=(80, 80, 100),
            width=2,
        )

        # Draw trainer puppets
        for i, (name, puppet) in enumerate(self.trainer_puppets.items()):
            is_speaking = (speaking_name == name)
            p_frame = puppet.get_frame(speaking=is_speaking)
            if puppet_scale < 1.0:
                p_frame = p_frame.resize((puppet_w, puppet_h), Image.LANCZOS)

            x = (i + 1) * slot_width + slot_width // 2 - puppet_w // 2
            frame.paste(p_frame, (x, y_offset), p_frame)

            # Glow effect behind speaking puppet
            if is_speaking:
                glow_draw = ImageDraw.Draw(frame)
                glow_draw.rounded_rectangle(
                    [x - 10, y_offset - 10, x + puppet_w + 10, y_offset + puppet_h + 10],
                    radius=15,
                    outline=(255, 255, 100),
                    width=3,
                )

    def _draw_subtitle_bar(
        self, draw: ImageDraw.Draw, speaker: str, text: str
    ):
        """Draw the subtitle bar at the bottom."""
        bar_top = self.height - self.subtitle_bar_height

        # Background (semi-dark)
        draw.rectangle(
            [0, bar_top, self.width, self.height],
            fill=(15, 15, 25),
        )
        # Accent line
        draw.rectangle(
            [0, bar_top, self.width, bar_top + 3],
            fill=(100, 100, 200),
        )

        if not text:
            return

        # Speaker name
        name_font = _get_font(24)
        text_font = _get_font(22)

        if speaker:
            name_text = f"{speaker}:"
            draw.text((40, bar_top + 15), name_text, fill=(255, 220, 100), font=name_font)

        # Subtitle text (word-wrapped)
        max_width = self.width - 80
        words = text.split()
        lines = []
        current_line = ""
        for word in words:
            test_line = f"{current_line} {word}".strip()
            bbox = draw.textbbox((0, 0), test_line, font=text_font)
            if bbox[2] - bbox[0] > max_width:
                if current_line:
                    lines.append(current_line)
                current_line = word
            else:
                current_line = test_line
        if current_line:
            lines.append(current_line)

        # Draw lines (max 3)
        y = bar_top + 50
        for line in lines[:3]:
            draw.text((40, y), line, fill=(220, 220, 240), font=text_font)
            y += 28
