from __future__ import annotations

import math
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, r"C:\Users\tyler\GitHub\wild-eyez-packing-list\.codex_pptx_build\pydeps")

import imageio_ffmpeg  # type: ignore
from PIL import Image, ImageDraw, ImageFilter


REPO = Path(r"C:\Users\tyler\GitHub\wild-eyez-packing-list")
PAGES_DIR = REPO / "assets" / "pages"
AUDIO = REPO / "assets" / "audio" / "AUDIO_ADOBE_UNDER_16MB_128K.mp3"
OUT = REPO / "WILD EYEZ PACKING GUIDE_READ_ALONG_VIDEO_PAGE1_VISIBLE.mp4"

WIDTH = 1920
HEIGHT = 1080
FPS = 15
PAGE_WIDTH = 1700
PAGE_HEIGHT = 2200
PAGE_X = (WIDTH - PAGE_WIDTH) // 2
MAX_SCROLL = PAGE_HEIGHT - HEIGHT

# Page durations estimated from the current 12:06 audio and page-by-page script.
PAGE_DURATIONS = [43.68, 153.47, 106.76, 93.43, 77.37, 38.83, 137.98, 103.72, 71.17]

# Page 1 has branding/header content above the actual read-along copy. Start lower
# so the quote and How to Use This Guide text are visible as soon as narration begins.
PAGE_START_SCROLL = {
    1: 500,
}


def make_background() -> Image.Image:
    bg_path = REPO / "assets" / "backgrounds" / "dark-brown-topo.jpg"
    if bg_path.exists():
        bg = Image.open(bg_path).convert("RGB")
        scale = max(WIDTH / bg.width, HEIGHT / bg.height)
        resized = bg.resize((math.ceil(bg.width * scale), math.ceil(bg.height * scale)), Image.Resampling.LANCZOS)
        left = (resized.width - WIDTH) // 2
        top = (resized.height - HEIGHT) // 2
        bg = resized.crop((left, top, left + WIDTH, top + HEIGHT))
    else:
        bg = Image.new("RGB", (WIDTH, HEIGHT), (34, 22, 13))

    veil = Image.new("RGBA", (WIDTH, HEIGHT), (20, 12, 7, 80))
    bg = Image.alpha_composite(bg.convert("RGBA"), veil)
    return bg.convert("RGB")


def make_page_shadow() -> Image.Image:
    shadow = Image.new("RGBA", (PAGE_WIDTH + 44, PAGE_HEIGHT + 44), (0, 0, 0, 0))
    draw = ImageDraw.Draw(shadow)
    draw.rectangle((22, 22, PAGE_WIDTH + 22, PAGE_HEIGHT + 22), fill=(0, 0, 0, 135))
    shadow = shadow.filter(ImageFilter.GaussianBlur(14))
    return shadow


def load_pages() -> list[Image.Image]:
    pages: list[Image.Image] = []
    for i in range(1, 10):
        p = PAGES_DIR / f"WILD EYEZ PACKING AND PREPARATION GUIDE6_Page_{i}.jpg"
        if not p.exists():
            raise FileNotFoundError(p)
        page = Image.open(p).convert("RGB")
        if page.size != (PAGE_WIDTH, PAGE_HEIGHT):
            page = page.resize((PAGE_WIDTH, PAGE_HEIGHT), Image.Resampling.LANCZOS)
        pages.append(page)
    return pages


def write_video() -> None:
    if not AUDIO.exists():
        raise FileNotFoundError(AUDIO)

    ffmpeg = imageio_ffmpeg.get_ffmpeg_exe()
    tmp_video = REPO / ".codex_video_build" / "read_along_silent.mp4"
    tmp_video.parent.mkdir(parents=True, exist_ok=True)
    if tmp_video.exists():
        tmp_video.unlink()
    if OUT.exists():
        OUT.unlink()

    bg = make_background()
    shadow = make_page_shadow()
    pages = load_pages()

    cmd = [
        ffmpeg,
        "-y",
        "-f",
        "rawvideo",
        "-vcodec",
        "rawvideo",
        "-pix_fmt",
        "rgb24",
        "-s",
        f"{WIDTH}x{HEIGHT}",
        "-r",
        str(FPS),
        "-i",
        "-",
        "-an",
        "-c:v",
        "libx264",
        "-preset",
        "veryfast",
        "-crf",
        "25",
        "-pix_fmt",
        "yuv420p",
        str(tmp_video),
    ]

    proc = subprocess.Popen(cmd, stdin=subprocess.PIPE)
    assert proc.stdin is not None

    try:
        for page_index, (page, duration) in enumerate(zip(pages, PAGE_DURATIONS), start=1):
            frames = max(1, round(duration * FPS))
            for frame_i in range(frames):
                if frames == 1:
                    progress = 0.0
                else:
                    progress = frame_i / (frames - 1)

                # Constant teleprompter-like movement inside each page.
                start_scroll = PAGE_START_SCROLL.get(page_index, 0)
                scroll_y = int(round(start_scroll + progress * (MAX_SCROLL - start_scroll)))
                frame = bg.copy().convert("RGBA")

                page_y = -scroll_y
                frame.alpha_composite(shadow, (PAGE_X - 22, page_y - 22))
                frame.paste(page, (PAGE_X, page_y))

                # Thin page frame to keep the document feeling physical.
                draw = ImageDraw.Draw(frame)
                draw.rectangle((PAGE_X, page_y, PAGE_X + PAGE_WIDTH - 1, page_y + PAGE_HEIGHT - 1), outline=(236, 133, 32, 255), width=5)

                proc.stdin.write(frame.convert("RGB").tobytes())
    finally:
        proc.stdin.close()
        code = proc.wait()
        if code != 0:
            raise RuntimeError(f"ffmpeg video encode failed with code {code}")

    mux_cmd = [
        ffmpeg,
        "-y",
        "-i",
        str(tmp_video),
        "-i",
        str(AUDIO),
        "-c:v",
        "copy",
        "-c:a",
        "aac",
        "-b:a",
        "128k",
        "-shortest",
        str(OUT),
    ]
    subprocess.run(mux_cmd, check=True)


if __name__ == "__main__":
    write_video()
    print(OUT)
