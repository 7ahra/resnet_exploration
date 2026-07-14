# BCCE 2026 — What chemistry faculty believe about doctoral education

Slide deck for the BCCE 2026 talk, generated from the paper draft
*"What chemistry faculty believe about doctoral education: a phenomenographic
and thematic characterization"* (Zoe Rahimi Pirkoohi & Jordan Harshman,
University of Iowa).

## Files
- **`BCCE2026_faculty_beliefs.pptx`** — the editable deck (16 slides, 16:9).
  Open in PowerPoint / Keynote / Google Slides. Fonts: Georgia (display/quotes)
  + Calibri (labels/body), both standard with Office.
- **`BCCE2026_preview.pdf`** — a quick-look preview. *Note:* the preview is
  rendered with substitute fonts (Liberation Serif/Sans), so exact letterforms
  differ slightly from the real deck; use the `.pptx` for the final look.
- **`build_deck.py`** — regenerates the `.pptx` from scratch (`python-pptx`).
- **`pptx_render.py`** — renders the `.pptx` to PNGs for QA.
- **`assets/`** — figures reused from the paper (the speech-bubble opener and
  the Figure 3 theme/boundary map).

## Deck outline (~15 min, visual-forward)
1. Title
2. The enterprise & its strains
3. The gap — teaching a class vs. advising a PhD
4. Research question
5. Why beliefs matter — Teacher-Centered Systemic Reform
6. What counts as a belief — the 4-part decision rule
7. Methods — 20 faculty, interviews + card sort, 990 → 15 → 4
8. The sample (N = 20)
9. **The finding — agree on propositions, divide on boundaries**
10–13. The four themes (with a representative quote + the axis of division)
14. The full outcome space (Figure 3)
15. What this adds
16. Implications + thank you

## Regenerate
```bash
pip install python-pptx Pillow
python build_deck.py      # -> BCCE2026_faculty_beliefs.pptx
```

Palette (from the paper's figures): teal `#456A73`, gold `#A37730`,
green `#5E7B4F`, mauve `#9D6473`.

Supported by NSF CAREER #2142873 & #2602955 (Jordan Harshman, PI).
