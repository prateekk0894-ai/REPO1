# ShortsMaker – Offline Long Video se Shorts

Apni lambi video (podcast, lecture, vlog) se **9:16 vertical short clips** banata hai,
**automatic word-by-word captions** ke saath. Sab kuch aapke PC par chalta hai: free, bina
API key ke, aur setup ke baad bina internet ke.

## Kya karta hai

1. Video se audio nikalta hai.
2. **Whisper** (faster-whisper) se bolne ko text banata hai, har shabd ke time ke saath. Hindi aur English dono.
3. **Best moments chunta hai:** hook words ("secret", "galti", "sabse", "कैसे"...), bolne ki speed,
   awaaz ki energy aur sawaal (?) dekh kar. Har clip poore sentence par shuru aur khatam hoti hai.
   Agar aapke PC par [Ollama](https://ollama.com) chal raha ho, to local AI model se bhi chun sakta hai.
4. **9:16 crop** jo bolne wale ke chehre par rehta hai, ya `--fit blur` se poora frame blurred background par.
5. **Captions:** 3 shabd ek saath, bola ja raha shabd peela aur bada. Video par jale hue (burned-in).
6. Har clip `short_01.mp4`, `short_02.mp4`... aur `clips.txt` mein timing, title idea aur text.

## Setup (sirf ek baar)

1. **Python 3.10 ya naya** install karo: <https://www.python.org/downloads/>
   (installer mein **"Add python.exe to PATH"** zaroor tick karo).
2. Is `shorts_maker` folder mein **`setup.bat`** par double-click karo.
   Ye zaroori cheezein aur Whisper model (~500 MB) download karega. Isme 5–15 minute lag sakte hain.
   ffmpeg alag se install karne ki zaroorat nahi, wo saath mein aa jaata hai.

## Chalana

- Apni video file ko **`run.bat` par drag-and-drop** karo.
- Ya `run.bat` kholo aur video ka path daalo.
- Puchhega: kitni clips (Enter = 5) aur bhasha (`hi`, `en`, ya Enter = auto).
- Clips video ke paas ek naye folder **`<video ka naam>_shorts`** mein aa jaayengi.

10 minute ki video, bina graphics card ke, lagbhag 5–15 minute leti hai. NVIDIA card ho to kaafi tez chalega.

## Zyada options (Command Prompt se)

```bat
venv\Scripts\python.exe shorts_maker.py "D:\videos\podcast.mp4" --clips 8 --min 20 --max 45 --lang hi
```

| Option | Matlab | Default |
|---|---|---|
| `--clips 8` | kitni clips | 5 |
| `--min 20 --max 45` | clip ki lambai (second) | 25 – 55 |
| `--lang hi` / `--lang en` | bhasha (khaali = auto) | auto |
| `--model medium` | Whisper size: `tiny`, `base`, `small`, `medium`, `large-v3`. Bada = zyada sahi, par slow | small |
| `--fit blur` | crop ki jagah poora frame, blurred background par | crop |
| `--ollama qwen2.5:7b` | local Ollama AI se moments chunna | band |
| `--font "Arial Black"` | caption font (Hindi ke liye default `Nirmala UI`) | auto |
| `--font-size 90` | caption size | 78 |
| `--highlight "#00FF88"` | bola ja raha shabd kis rang mein | peela |
| `--words-per-line 2` | ek baar mein kitne shabd | 3 |
| `--no-captions` | captions ke bina | |
| `--no-upper` | English captions CAPITAL na karo | |

**Bhasha:** Hindi audio par captions Devanagari (हिंदी) mein aate hain. `--lang en` dene par Whisper Hindi ko aksar English mein translate kar deta hai.
**Galat shabd theek karna ho?** Output folder mein `transcript.json` edit karo aur dobara chalao.
Ye transcript dobara use hota hai, isliye dusri baar sirf rendering hoti hai.

## Ollama (optional, smarter selection)

1. <https://ollama.com> se Ollama install karo.
2. Command Prompt mein: `ollama pull qwen2.5:7b` (8 GB+ RAM chahiye).
3. Command Prompt mein is folder se chalao:
   `venv\Scripts\python.exe shorts_maker.py "video.mp4" --ollama qwen2.5:7b`

Ollama na mile to tool apne aap built-in scoring use karta hai.

## Copyright

Ye tool **apni khud ki videos** ya jinki permission aapke paas ho, unke liye hai. Ye kisi aur ke
content ka copyright nahi hatata, aur Content ID se bachne ke liye video ko chhupata bhi nahi.
Kisi aur ki video ke clips upload karne par claim ya strike aa sakti hai.
