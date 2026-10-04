---
name: presentation-workshop
description: Collaborate with a presenter to plan, organise, and refine presentation content interactively, then build an editable PowerPoint from this project's Teal Editorial template. Use for presentation planning, slide storyboarding, turning notes into a talk, or building a template-based deck.
compatibility: This Presentations repository, Python 3 with python-pptx and lxml; LibreOffice and a PDF renderer for visual validation.
---

# Presentation Workshop

Help the presenter find the story before fitting it into slides. The presenter owns the message and voice; you help clarify, organise, edit, and produce it. Work in conversation, not by dropping an unsolicited finished deck.

## Local resources

Resolve all relative paths from this directory, not the shell's working directory:

- `../../../Teal_Editorial_Template.potx`: the approved, native-layout template.
- `../../../Teal_Editorial_Starter.pptx`: visual examples only.
- `../../../Teal_Editorial_Preview.pdf`: layout preview.
- `../../../README.md`: layout choices and content-length guidance.
- `../../../build_template.py`: shared text-formatting helper; **do not run it to build a talk**.
- `scripts/build_presentation.py`: builds approved content directly from the POTX.
- `references/example-content.json`: a small, valid content manifest; not a suggested talk.

Use the project's file/shell tools and instructions. Do not upload presenter materials to external services. Check the actual files before asking for facts you can discover.

## 1. Discover the brief

Read supplied notes, decks, source documents, and existing workshop state first. Ask **one focused question per turn**, starting with the missing answer that most affects the presentation. If nothing is known, start: **“Who will you be presenting to?”**

Discover only what matters: audience and prior knowledge; desired understanding, action, or decision; topic and central message; duration including questions/activities; setting and tone; required material and reliable sources; presenter's preferred voice and boundaries. Do not repeat answered questions or insist on a fixed questionnaire. Usually no more than seven initial questions are needed; a complete brief needs none.

Offer a useful proposal with each subsequent question when possible. Example: “For these clinicians, I suggest starting with a case rather than definitions. Does that fit how you like to teach?” Allow rambling, revisions, and uncertainty; turn them into a concise brief without inventing facts.

Create `presentations/<descriptive-slug>/CONTENT.md` under the repository. Use one working document containing:

- **Brief:** audience, desired outcome, central message, duration, voice, requirements, exclusions.
- **Working state:** current phase, decisions the presenter has accepted, unresolved assumptions, next question, and any requested scope changes.
- **Storyboard:** slide number, audience-facing title, single takeaway, purpose, layout, and estimated minutes.
- **Slide content:** approved on-slide wording, separate speaker notes, transitions, activity prompts, and sources.

Confirm the brief when material uncertainty remains. Otherwise present a concise summary and move to the story. Do not stall on administrative approval of already supplied decisions.

## 2. Organise the story together

Propose a compact narrative suited to this audience: opening hook/problem, necessary context, two to four major ideas, application/example, and a clear closing takeaway or action. This is a starting point, not a compulsory structure. Reserve time for interaction and questions; make the sum of estimates fit the agreed session length rather than treating every slide as one minute.

Show a storyboard, not full prose for every slide. Use message-led titles where appropriate. Explain only the important sequencing choices. Ask one useful question about the proposed story; offer at most two alternatives if the presenter is undecided. Update the storyboard when they reorder, add, or remove material.

Choose layouts by the message, not by a desire to use all ten. Inspect actual placeholder names with:

```bash
python3 <absolute-skill-directory>/scripts/build_presentation.py --layouts
```

Use the layout guidance in `../../../README.md`. Do not fill eight blocks just because an eight-block layout exists. A discussion or closing slide can use an existing bullet/takeaway layout; do not redesign the template without a request.

## 3. Craft content interactively

Draft **one to three slides at a time**, adapting the batch size to the presenter's preference. For each, show the title, short on-slide text, and the main speaker-note beats. End with one focused editorial question. Avoid presenting the whole deck for approval unless the presenter requests a single-pass draft.

- Preserve the presenter's terminology and tone. Distinguish their original words from suggested rewrites when the distinction matters.
- Give each slide a clear job and one main takeaway. Keep explanations, nuance, and transitions in notes; put essential caveats on the slide when leaving them out would mislead.
- Treat examples and exercises as teaching choices. Label invented scenarios as hypothetical/composite; de-identify real cases and ask before including sensitive details.
- Do not invent data, quotations, references, or clinical claims. Record verified sources; mark unsupported claims for resolution, and do not build them as established facts.
- Split dense content or choose a more suitable layout before shrinking the text. Keep approved fonts, colours, teal square bullets, and white bullets on dark panels.
- Offer optional questions, pauses, and discussion prompts when useful to the session's goal. Do not add activities purely to make the talk longer.

Track accepted changes in `CONTENT.md`. If the goal or audience changes substantially, revisit the brief and sequence. On resumption, read this file and ask only what is still unresolved.

## 4. Approve, then build

Summarise the final sequence, estimated timing, remaining uncertainties, and output filename. Obtain the presenter's approval of the content/sequence **before generating the PPTX**. An explicit “build it” applying to the current draft is approval; do not ask again. Planning permission alone is not build permission. If the presenter explicitly requests a fast single-pass deck, state the assumptions and respect that authorisation rather than enforcing extra interview gates.

Translate accepted wording into `presentations/<slug>/content.json` using the example manifest and the actual layout catalogue:

- Top level: `title`, `presenter`, `footer`, `slides`.
- Each slide: exact `layout` name, `fields` with all and only that layout's named text placeholders, optional `notes`, optional `sources` (list of reference strings), and optional `duration_minutes`.
- Field values are plain strings; use `\n` between bullets or paragraphs. Do **not** type bullet characters into the text: the template supplies them.
- An intentionally unused placeholder has an empty string. Do not leave template example wording or “Click to add…” prompts in the final content.
- Sources and timing are written to speaker notes; timing does not create automatic slide transitions.

Build using absolute paths:

```bash
python3 <absolute-skill-directory>/scripts/build_presentation.py \
  <absolute-repository>/presentations/<slug>/content.json \
  <absolute-repository>/presentations/<slug>/<presentation-name>-v1.pptx
```

The builder retains native masters/layouts, creates editable text placeholders, replaces the shared footer, adds speaker notes, and removes example slides. It never overwrites an existing output. For revisions use `-v2`, `-v3`, etc.; preserve the previous version and regenerate from the accepted content rather than hand-patching ZIP files.

Do not overwrite source decks, the template, or other people's files. Do not commit, push, install globally, or change global Pi settings unless asked.

## 5. Validate and review with the presenter

Run the builder's checks, then export the candidate to PDF with an isolated LibreOffice profile if available. Render **every** slide and inspect for clipped text, overlaps, low contrast, crowded layouts, broken bullet styling, and unexpected font substitution. Check that speaker notes and sources are present and timing still fits. The builder is not an overflow detector; a successful save is not visual acceptance.

If needed, refine wording or layout choice, update the content manifest, and create a new candidate. Material content changes return to presenter approval; cosmetic fixes within the accepted design do not require another interview. If rendering is unavailable, report the limitation and call the deck an unverified draft, not visually checked. Do not claim direct PowerPoint UI testing unless it happened.

Deliver the local PPTX and, when generated, a PDF preview. Give a short summary of slide count, duration estimate, verification, and any remaining source issues. Invite one targeted review question rather than a generic list of tasks for the presenter.

## Done means

The presenter has an accepted story and content, an editable deck built from the approved template, matching speaker notes/sources where supplied, and an honest validation report. If they wanted planning only, stop at the accepted content plan.
