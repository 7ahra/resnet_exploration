# BCCE 2026 — What chemistry faculty believe about doctoral education

Slide deck for the BCCE 2026 talk, generated from the paper draft
*"What chemistry faculty believe about doctoral education: a phenomenographic
and thematic characterization"* (Zoe Rahimi Pirkoohi & Jordan Harshman,
University of Iowa).

## Design (v4 — assertion-evidence STORYTELLING, current)
Matches the research group's deck not just in look but in *rhetoric*: it is an
assertion-evidence story, so every slide title is a full-sentence **claim** and
the body is only the evidence for it. Read the titles in order and they tell the
whole argument — stakes → why reform stalls → the gap → the question → (what a
belief is) → how we looked → **the reveal** (agree, but divide) → the four
beliefs, each with its divide → the full map → why it matters → one-line close.

- **Storytelling moves borrowed from the reference:** claim-titles that chain;
  concepts built *before* they are used (belief defined, TCSR mechanism shown);
  the finding dramatized with **opposing quotes** (P13 vs P7; creativity as the
  faculty's vs the student's to build) instead of stated abstractly.
- **Look (unchanged from v3):** sage-green #7D9263 rounded boxes, #2C2C2C black
  sharp boxes for the opposing side, terrazzo background, Century Gothic, Times
  New Roman italic quotes, bottom-left citations, slide numbers.
- 15 slides.

## Files
- **`BCCE2026_faculty_beliefs.pptx`** — the editable deck (16 slides, 16:9).
- **`BCCE2026_preview.pdf`** — quick-look preview. *Note:* rendered with
  substitute fonts (Liberation for Century Gothic / Times), so exact letterforms
  differ slightly from the real deck; use the `.pptx` for the final look.
- **`build_v4.py`** — regenerates the current (storytelling house-style) `.pptx`.
- **`make_map_v3.py`** — rebuilds the boundary map in the green house style.
- **`pptx_render.py`** — renders the `.pptx` to PNGs for QA.
- **Earlier versions (kept for reference):** `build_v3.py` (house-style, section
  order), `build_v2.py` (restrained flamingo/Iowa), `build_deck.py` (v1),
  `make_assets.py`.
- **`assets/`** — figures. v3 uses `bg_texture.png` (terrazzo) and
  `boundary_map_v3.png`.

## Deck outline (~15 min, minimal-text)
1. Title
2. The enterprise & its strains
3. The gap — teaching a class vs. advising a PhD
4. Research question
5. Why beliefs matter — Teacher-Centered Systemic Reform
6. What counts as a belief — the 4-part decision rule
7. Methods — 20 faculty, interviews + card sort, 990 → 15 → 4
8. The sample (N = 20)
9. **The finding — agree on propositions, divide on boundaries** (black slide)
10–13. The four themes (representative quote + the axis of division)
14. The full outcome space (recreated boundary map)
15. What this adds
16. Implications + thank you

## Regenerate
```bash
pip install python-pptx Pillow
python make_map_v3.py     # -> assets/boundary_map_v3.png
python build_v4.py        # -> BCCE2026_faculty_beliefs.pptx
```

Supported by NSF CAREER #2142873 & #2602955 (Jordan Harshman, PI).
