"""
Puppet Avatar Generator

Generates cartoon-style puppet characters programmatically using Pillow.
Each puppet has:
  - A distinct body color
  - Simple face with eyes and animated mouth
  - Name label
  - Two states: mouth open (speaking) and mouth closed (idle)

Puppets are rendered as PIL Images and cached for performance.
"""

from PIL import Image, ImageDraw, ImageFont
import math


def _get_font(size: int) -> ImageFont.FreeTypeFont:
    """Try to load a good font, fall back to default."""
    font_paths = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
        "/usr/share/fonts/truetype/freefont/FreeSansBold.ttf",
    ]
    for path in font_paths:
        try:
            return ImageFont.truetype(path, size)
        except (OSError, IOError):
            continue
    return ImageFont.load_default()


class Puppet:
    """A single puppet avatar with idle and speaking frames."""

    def __init__(
        self,
        name: str,
        body_color: tuple,
        accent_color: tuple,
        width: int = 300,
        height: int = 400,
    ):
        self.name = name
        self.body_color = body_color
        self.accent_color = accent_color
        self.width = width
        self.height = height

        # Pre-render both states
        self._frame_idle = self._render(mouth_open=False)
        self._frame_speaking = self._render(mouth_open=True)

    def _render(self, mouth_open: bool) -> Image.Image:
        """Render a single puppet frame."""
        img = Image.new("RGBA", (self.width, self.height), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)

        cx = self.width // 2
        head_radius = 70
        head_cy = 120

        # --- Hair / hat (accent color arc on top) ---
        hair_bbox = [
            cx - head_radius - 10,
            head_cy - head_radius - 20,
            cx + head_radius + 10,
            head_cy + 10,
        ]
        draw.ellipse(hair_bbox, fill=self.accent_color)

        # --- Head (circle) ---
        head_bbox = [
            cx - head_radius,
            head_cy - head_radius,
            cx + head_radius,
            head_cy + head_radius,
        ]
        # Skin tone
        skin = (255, 220, 185)
        draw.ellipse(head_bbox, fill=skin, outline=(180, 150, 120), width=2)

        # --- Eyes ---
        eye_y = head_cy - 10
        eye_offset = 25
        eye_radius = 10
        for ex in [cx - eye_offset, cx + eye_offset]:
            # White
            draw.ellipse(
                [ex - eye_radius, eye_y - eye_radius, ex + eye_radius, eye_y + eye_radius],
                fill="white",
                outline=(100, 100, 100),
                width=1,
            )
            # Pupil
            pr = 5
            draw.ellipse(
                [ex - pr, eye_y - pr, ex + pr, eye_y + pr],
                fill=(40, 40, 40),
            )

        # --- Mouth ---
        mouth_y = head_cy + 30
        if mouth_open:
            # Open mouth - ellipse
            draw.ellipse(
                [cx - 20, mouth_y - 12, cx + 20, mouth_y + 12],
                fill=(180, 60, 60),
                outline=(120, 40, 40),
                width=2,
            )
        else:
            # Closed mouth - slight smile arc
            draw.arc(
                [cx - 20, mouth_y - 10, cx + 20, mouth_y + 10],
                start=0,
                end=180,
                fill=(180, 80, 80),
                width=3,
            )

        # --- Body (rounded rectangle below head) ---
        body_top = head_cy + head_radius + 5
        body_left = cx - 65
        body_right = cx + 65
        body_bottom = self.height - 50
        draw.rounded_rectangle(
            [body_left, body_top, body_right, body_bottom],
            radius=20,
            fill=self.body_color,
            outline=self.accent_color,
            width=3,
        )

        # --- Arms (small arcs on sides) ---
        arm_y = body_top + 40
        # Left arm
        draw.arc(
            [body_left - 30, arm_y - 15, body_left + 10, arm_y + 50],
            start=180,
            end=360,
            fill=self.body_color,
            width=8,
        )
        # Right arm
        draw.arc(
            [body_right - 10, arm_y - 15, body_right + 30, arm_y + 50],
            start=180,
            end=360,
            fill=self.body_color,
            width=8,
        )

        # --- Name label ---
        font = _get_font(20)
        text_bbox = draw.textbbox((0, 0), self.name, font=font)
        text_w = text_bbox[2] - text_bbox[0]
        name_x = cx - text_w // 2
        name_y = self.height - 40
        # Background pill
        pill_pad = 8
        draw.rounded_rectangle(
            [name_x - pill_pad, name_y - pill_pad,
             name_x + text_w + pill_pad, name_y + (text_bbox[3] - text_bbox[1]) + pill_pad],
            radius=10,
            fill=self.accent_color,
        )
        draw.text((name_x, name_y), self.name, fill="white", font=font)

        return img

    def get_frame(self, speaking: bool) -> Image.Image:
        """Return the appropriate frame based on speaking state."""
        return self._frame_speaking if speaking else self._frame_idle


class PuppetFactory:
    """Creates and caches puppets from trainer profiles."""

    def __init__(self, puppet_width: int = 300, puppet_height: int = 400):
        self.puppet_width = puppet_width
        self.puppet_height = puppet_height
        self._cache: dict[str, Puppet] = {}

    def get_puppet(self, name: str, body_color: tuple, accent_color: tuple) -> Puppet:
        if name not in self._cache:
            self._cache[name] = Puppet(
                name=name,
                body_color=body_color,
                accent_color=accent_color,
                width=self.puppet_width,
                height=self.puppet_height,
            )
        return self._cache[name]
