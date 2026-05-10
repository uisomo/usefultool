# Video transcript analysis

Pipeline for pulling a YouTube transcript and getting a sentence-by-sentence
rhetorical breakdown.

## 1. Fetch the transcript (run locally, not in this sandbox)

This sandbox's egress IP is blocked by YouTube (HTTP 403 on every endpoint),
so run on your own machine:

```bash
pip install yt-dlp
./fetch_transcript.sh "https://youtu.be/5qKJj59nalc"
```

Produces:

- `video.en.srt` — captions with timestamps
- `video.txt` — cleaned plain text, ready for analysis

## 2. Get the analysis

Paste the contents of `video.txt` (or `video.en.srt`) back into the Claude
session. The analysis output is a Markdown report with:

- **High-level structure** — intro → thesis → evidence → counterpoint → close
- **Sentence-by-sentence table** with columns:
  `#`, `timestamp`, `sentence`, `type` (claim / example / question /
  transition / definition), `technique` (analogy, contrast, anecdote,
  statistic, hypothetical, repetition, …), `role-in-arc`, `notes`
- **Patterns observed** — recurring techniques, pacing, where examples land
  relative to claims, question density, etc.

## Notes

- Auto-captions have no punctuation. The cleaner joins lines into one blob;
  the analysis step re-punctuates and segments into sentences.
- For higher-quality transcripts (or videos without captions), swap step 1 for
  Whisper: `yt-dlp -x --audio-format mp3 <url>` then `whisper video.mp3
  --model small --language en`.
