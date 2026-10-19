"""ShortsMaker - offline long video -> vertical short clips with automatic captions.

Everything runs on your own computer:
  * speech to text with faster-whisper (model downloaded once, then offline)
  * best moments picked from the transcript (keyword/energy scoring, or a
    local Ollama model if you have one running)
  * 9:16 crop that follows the speaker's face (OpenCV)
  * word-by-word highlighted captions burned into the video (ffmpeg + ASS)

Usage:
    python shorts_maker.py "my video.mp4"
    python shorts_maker.py "my video.mp4" --clips 5 --min 20 --max 45 --lang hi
    python shorts_maker.py "my video.mp4" --ollama qwen2.5:7b
Run `python shorts_maker.py -h` for every option.

Use it on videos you own or have permission to use.
"""
import argparse
import json
import math
import os
import re
import subprocess
import sys
import urllib.request
import wave

import numpy as np

# ---------------------------------------------------------------- helpers


def ffmpeg_exe():
    try:
        import imageio_ffmpeg
        return imageio_ffmpeg.get_ffmpeg_exe()
    except Exception:
        return "ffmpeg"


FFMPEG = ffmpeg_exe()


def run(cmd):
    p = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True,
                       encoding="utf-8", errors="replace")
    if p.returncode != 0:
        sys.exit("ffmpeg failed:\n" + p.stderr[-2000:])
    return p


def log(msg):
    print(msg, flush=True)


def video_info(path):
    """Return (width, height, duration_seconds) using ffmpeg's banner."""
    p = subprocess.run([FFMPEG, "-hide_banner", "-i", path], stderr=subprocess.PIPE,
                       text=True, encoding="utf-8", errors="replace")
    err = p.stderr
    m = re.search(r"Duration: (\d+):(\d+):(\d+\.?\d*)", err)
    dur = int(m.group(1)) * 3600 + int(m.group(2)) * 60 + float(m.group(3)) if m else 0
    m = re.search(r"Video:.*?(\d{2,5})x(\d{2,5})", err)
    if not m:
        sys.exit("Could not read the video size. Is this a video file?")
    w, h = int(m.group(1)), int(m.group(2))
    r = re.search(r"rotate\s*:\s*(-?\d+)", err) or re.search(r"rotation of (-?\d+)", err)
    if r and abs(int(float(r.group(1)))) in (90, 270):
        w, h = h, w
    return w, h, dur


# ---------------------------------------------------------------- 1. audio + transcript


def extract_audio(video, wav):
    run([FFMPEG, "-y", "-i", video, "-vn", "-ac", "1", "-ar", "16000", "-c:a", "pcm_s16le", wav])


def transcribe(wav, model_size, lang, device):
    from faster_whisper import WhisperModel
    compute = "float16" if device == "cuda" else "int8"
    log(f"  loading Whisper '{model_size}' on {device} (first run downloads it once)...")
    model = WhisperModel(model_size, device=device, compute_type=compute)
    segments, info = model.transcribe(wav, language=lang, word_timestamps=True, vad_filter=True,
                                      beam_size=5)
    log(f"  language: {info.language}")
    words, segs = [], []
    for s in segments:
        sw = [{"w": w.word.strip(), "s": round(w.start, 3), "e": round(w.end, 3)}
              for w in (s.words or []) if w.word.strip()]
        if not sw:
            continue
        words += sw
        segs.append({"s": sw[0]["s"], "e": sw[-1]["e"], "text": s.text.strip()})
        log(f"  [{fmt_t(s.start)}] {s.text.strip()[:80]}")
    return {"language": info.language, "segments": segs, "words": words}


def fmt_t(t):
    return f"{int(t // 60):02d}:{t % 60:04.1f}"


# ---------------------------------------------------------------- 2. pick moments

HOOK_WORDS = [
    # English
    "secret", "mistake", "never", "always", "truth", "how to", "why", "best", "worst", "money",
    "free", "biggest", "nobody", "everyone", "shocking", "important", "must", "warning", "stop",
    "simple", "trick", "tip", "rich", "fail", "success", "problem", "solution", "story",
    # Hinglish / Hindi (Roman)
    "galti", "sach", "raaz", "kabhi", "sabse", "zaroor", "jaruri", "paisa", "kaise", "kyun",
    "dhyan", "khatarnak", "asli", "seekh", "mehnat", "sapna",
    # Hindi (Devanagari)
    "गलती", "सच", "राज़", "राज", "कभी", "सबसे", "ज़रूर", "जरूरी", "पैसा", "पैसे", "कैसे", "क्यों",
    "ध्यान", "खतरनाक", "असली", "सीख", "मेहनत", "सपना", "सफलता",
]


def loudness(wav):
    """RMS loudness per second (0..1) from the 16 kHz mono wav."""
    with wave.open(wav) as f:
        a = np.frombuffer(f.readframes(f.getnframes()), np.int16).astype(np.float32)
    n = len(a) // 16000
    if n == 0:
        return np.zeros(1)
    rms = np.sqrt((a[: n * 16000].reshape(n, 16000) ** 2).mean(1))
    return rms / (rms.max() or 1)


def candidate_windows(segs, mn, mx):
    """Windows that start and end on sentence boundaries, mn..mx seconds long."""
    out = []
    for i in range(len(segs)):
        for j in range(i, len(segs)):
            d = segs[j]["e"] - segs[i]["s"]
            if d > mx:
                break
            if d >= mn:
                out.append((i, j))
    return out


def heuristic_pick(tr, wav, n, mn, mx):
    segs, words = tr["segments"], tr["words"]
    loud = loudness(wav)
    scored = []
    for i, j in candidate_windows(segs, mn, mx):
        s, e = segs[i]["s"], segs[j]["e"]
        text = " ".join(x["text"] for x in segs[i:j + 1])
        low = text.lower()
        nwords = sum(1 for w in words if s <= w["s"] < e)
        pace = nwords / max(e - s, 1)                       # words per second
        hooks = sum(low.count(k) for k in HOOK_WORDS)
        first = segs[i]["text"].lower()
        hook_start = 2.0 if (any(k in first for k in HOOK_WORDS) or "?" in first) else 0
        energy = float(loud[int(s):int(math.ceil(e))].mean()) if int(s) < len(loud) else 0
        questions = text.count("?")
        score = pace * 1.5 + hooks * 1.2 + hook_start + energy * 3 + questions * 0.8
        scored.append((score, s, e, text))
    return non_overlapping(scored, n)


def non_overlapping(scored, n):
    scored.sort(reverse=True)
    picked = []
    for sc, s, e, text in scored:
        if all(e <= ps or s >= pe for _, ps, pe, _ in picked):
            picked.append((sc, s, e, text))
        if len(picked) == n:
            break
    return sorted(picked, key=lambda x: x[1])


def ollama_pick(tr, n, mn, mx, model, host):
    """Ask a local Ollama model for the best moments. Returns None if unavailable."""
    lines = "\n".join(f"[{x['s']:.1f}-{x['e']:.1f}] {x['text']}" for x in tr["segments"])
    prompt = (
        f"You are a short-video editor. From this timestamped transcript pick the {n} most engaging, "
        f"self-contained moments for YouTube Shorts / Instagram Reels. Each must be {mn}-{mx} seconds, "
        "start with a strong hook and end on a complete thought. Do not overlap them.\n"
        'Reply ONLY with JSON: {"clips":[{"start":seconds,"end":seconds,"title":"short catchy title"}]}\n\n'
        + lines[:60000])
    body = json.dumps({"model": model, "prompt": prompt, "stream": False, "format": "json"}).encode()
    try:
        req = urllib.request.Request(host.rstrip("/") + "/api/generate", body,
                                     {"Content-Type": "application/json"})
        with urllib.request.urlopen(req, timeout=900) as r:
            resp = json.loads(json.loads(r.read())["response"])
    except Exception as ex:
        log(f"  Ollama not usable ({ex}); using the built-in scoring instead.")
        return None
    segs = tr["segments"]
    out = []
    for c in resp.get("clips", []):
        try:
            s, e = float(c["start"]), float(c["end"])
        except Exception:
            continue
        # snap to sentence boundaries
        s = min((x["s"] for x in segs), key=lambda v: abs(v - s))
        e = min((x["e"] for x in segs), key=lambda v: abs(v - e))
        if e - s < 5:
            continue
        text = " ".join(x["text"] for x in segs if x["s"] >= s and x["e"] <= e)
        out.append((0, s, e, text, c.get("title", "")))
    return out or None


# ---------------------------------------------------------------- 3. framing


def face_center_x(video, s, e, vw, vh):
    """Median horizontal face position in the clip, as 0..1 (0.5 if no face)."""
    import cv2
    if not hasattr(cv2, "CascadeClassifier"):      # OpenCV 5 moved it out; install opencv-python<5
        log("  face tracking unavailable (needs opencv-python 4.x); centring the crop")
        return 0.5
    cascade = cv2.CascadeClassifier(os.path.join(cv2.data.haarcascades,
                                                 "haarcascade_frontalface_default.xml"))
    cap = cv2.VideoCapture(video)
    xs = []
    t = s
    step = max((e - s) / 24, 0.5)
    while t < e:
        cap.set(cv2.CAP_PROP_POS_MSEC, t * 1000)
        ok, frame = cap.read()
        t += step
        if not ok:
            continue
        g = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        k = 480 / g.shape[0] if g.shape[0] > 480 else 1
        if k != 1:
            g = cv2.resize(g, None, fx=k, fy=k)
        faces = cascade.detectMultiScale(g, 1.1, 5, minSize=(30, 30))
        if len(faces):
            x, y, w, h = max(faces, key=lambda f: f[2] * f[3])
            xs.append((x + w / 2) / g.shape[1])
    cap.release()
    return float(np.median(xs)) if xs else 0.5


# ---------------------------------------------------------------- 4. captions (ASS)


def ass_time(t):
    t = max(t, 0)
    h, r = divmod(t, 3600)
    m, s = divmod(r, 60)
    return f"{int(h)}:{int(m):02d}:{s:05.2f}"


def make_ass(words, s, e, path, font, size, color, highlight, upper, per_line):
    def hexc(c):            # "#RRGGBB" -> ASS &H00BBGGRR
        c = c.lstrip("#")
        return f"&H00{c[4:6]}{c[2:4]}{c[0:2]}".upper()
    head = f"""[Script Info]
ScriptType: v4.00+
PlayResX: 1080
PlayResY: 1920
WrapStyle: 0
ScaledBorderAndShadow: yes

[V4+ Styles]
Format: Name, Fontname, Fontsize, PrimaryColour, SecondaryColour, OutlineColour, BackColour, Bold, Italic, Underline, StrikeOut, ScaleX, ScaleY, Spacing, Angle, BorderStyle, Outline, Shadow, Alignment, MarginL, MarginR, MarginV, Encoding
Style: Cap,{font},{size},{hexc(color)},{hexc(color)},&H00000000,&H80000000,-1,0,0,0,100,100,0,0,1,6,2,2,80,80,520,1

[Events]
Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text
"""
    ws = [w for w in words if w["s"] >= s - 0.05 and w["e"] <= e + 0.05]
    lines = []
    for k in range(0, len(ws), per_line):
        chunk = ws[k:k + per_line]
        for idx, w in enumerate(chunk):
            st = w["s"] - s
            en = (chunk[idx + 1]["s"] if idx + 1 < len(chunk) else w["e"]) - s
            parts = []
            for jdx, x in enumerate(chunk):
                txt = x["w"].upper() if upper else x["w"]
                txt = txt.replace("{", "(").replace("}", ")")
                if jdx == idx:
                    parts.append("{\\c" + hexc(highlight) + "\\fscx110\\fscy110}" + txt + "{\\r}")
                else:
                    parts.append(txt)
            lines.append(f"Dialogue: 0,{ass_time(st)},{ass_time(max(en, st + 0.05))},Cap,,0,0,0,,"
                         + " ".join(parts))
    with open(path, "w", encoding="utf-8") as f:
        f.write(head + "\n".join(lines) + "\n")


# ---------------------------------------------------------------- 5. render


def render(video, s, e, cx, vw, vh, ass_name, out_dir, out_name, fit, captions):
    tw, th = 1080, 1920
    if fit == "blur":
        vf = (f"split[a][b];[a]scale={tw}:{th}:force_original_aspect_ratio=increase,crop={tw}:{th},"
              f"boxblur=20:2[bg];[b]scale={tw}:-2[fg];[bg][fg]overlay=(W-w)/2:(H-h)/2,setsar=1")
    else:
        cw = min(vw, int(vh * 9 / 16)) // 2 * 2
        x = int(min(max(cx * vw - cw / 2, 0), vw - cw))
        vf = f"crop={cw}:{vh // 2 * 2}:{x}:0,scale={tw}:{th},setsar=1"
    if captions:
        vf += f",subtitles={ass_name}"
    cmd = [FFMPEG, "-y", "-ss", f"{s:.3f}", "-t", f"{e - s:.3f}", "-i", os.path.abspath(video),
           "-vf", vf, "-c:v", "libx264", "-preset", "veryfast", "-crf", "20",
           "-c:a", "aac", "-b:a", "160k", "-movflags", "+faststart", out_name]
    p = subprocess.run(cmd, cwd=out_dir, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True,
                       encoding="utf-8", errors="replace")
    if p.returncode != 0:
        sys.exit("ffmpeg failed:\n" + p.stderr[-2000:])


# ---------------------------------------------------------------- main


def main():
    ap = argparse.ArgumentParser(description="Offline long video -> viral vertical shorts with captions")
    ap.add_argument("video")
    ap.add_argument("--out", help="output folder (default: <video name>_shorts next to the video)")
    ap.add_argument("--clips", type=int, default=5, help="number of clips (default 5)")
    ap.add_argument("--min", type=float, default=25, help="minimum clip length in seconds")
    ap.add_argument("--max", type=float, default=55, help="maximum clip length in seconds")
    ap.add_argument("--model", default="small",
                    help="Whisper size: tiny, base, small, medium, large-v3 (bigger = better, slower)")
    ap.add_argument("--lang", default=None, help="speech language, e.g. hi or en (default: auto)")
    ap.add_argument("--device", default="auto", choices=["auto", "cpu", "cuda"])
    ap.add_argument("--ollama", metavar="MODEL", help="use a local Ollama model to pick moments")
    ap.add_argument("--ollama-host", default="http://localhost:11434")
    ap.add_argument("--fit", default="crop", choices=["crop", "blur"],
                    help="crop = follow the face, blur = whole frame on a blurred background")
    ap.add_argument("--no-captions", action="store_true")
    ap.add_argument("--font", default=None, help="caption font (default: Nirmala UI for Hindi, else Arial Black)")
    ap.add_argument("--font-size", type=int, default=78)
    ap.add_argument("--color", default="#FFFFFF")
    ap.add_argument("--highlight", default="#FFD400")
    ap.add_argument("--words-per-line", type=int, default=3)
    ap.add_argument("--no-upper", action="store_true", help="do not CAPITALISE captions")
    a = ap.parse_args()

    if not os.path.isfile(a.video):
        sys.exit(f"Video not found: {a.video}")
    base = os.path.splitext(os.path.basename(a.video))[0]
    out = a.out or os.path.join(os.path.dirname(os.path.abspath(a.video)), base + "_shorts")
    os.makedirs(out, exist_ok=True)
    vw, vh, dur = video_info(a.video)
    log(f"Video: {vw}x{vh}, {fmt_t(dur)}  ->  {out}")

    wav = os.path.join(out, "audio.wav")
    tjson = os.path.join(out, "transcript.json")
    log("1/4 Extracting audio...")
    extract_audio(a.video, wav)
    if os.path.isfile(tjson):
        log("2/4 Using saved transcript.json (delete it to transcribe again)")
        tr = json.load(open(tjson, encoding="utf-8"))
    else:
        log("2/4 Transcribing speech...")
        device = a.device
        if device == "auto":
            try:
                import ctranslate2
                device = "cuda" if ctranslate2.get_cuda_device_count() > 0 else "cpu"
            except Exception:
                device = "cpu"
        tr = transcribe(wav, a.model, a.lang, device)
        json.dump(tr, open(tjson, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    if not tr["segments"]:
        sys.exit("No speech found in the video.")

    log("3/4 Finding the best moments...")
    mx = min(a.max, dur)
    mn = min(a.min, mx)
    picks = ollama_pick(tr, a.clips, mn, mx, a.ollama, a.ollama_host) if a.ollama else None
    if picks is None:
        picks = [p + ("",) for p in heuristic_pick(tr, wav, a.clips, mn, mx)]
    if not picks:
        sys.exit("Video too short for the chosen clip length. Try a smaller --min.")

    hindi = bool(re.search(r"[ऀ-ॿ]", " ".join(w["w"] for w in tr["words"])))
    font = a.font or ("Nirmala UI" if hindi else "Arial Black")

    log(f"4/4 Rendering {len(picks)} clips...")
    summary = []
    for k, (_, s, e, text, title) in enumerate(picks, 1):
        name = f"short_{k:02d}.mp4"
        ass = f"short_{k:02d}.ass"
        cx = 0.5 if a.fit == "blur" else face_center_x(a.video, s, e, vw, vh)
        if not a.no_captions:
            make_ass(tr["words"], s, e, os.path.join(out, ass), font, a.font_size, a.color,
                     a.highlight, not a.no_upper and not hindi, a.words_per_line)
        render(a.video, s, e, cx, vw, vh, ass, out, name, a.fit, not a.no_captions)
        title = title or re.split(r"(?<=[.?!।])\s", text)[0][:80]
        summary.append(f"{name}  [{fmt_t(s)} - {fmt_t(e)}]  ({e - s:.0f}s)\n  Title idea: {title}\n  {text}\n")
        log(f"  done: {name}  {fmt_t(s)}-{fmt_t(e)}")
    with open(os.path.join(out, "clips.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(summary))
    os.remove(wav)
    log(f"\nFinished. {len(picks)} shorts are in: {out}")


if __name__ == "__main__":
    main()
