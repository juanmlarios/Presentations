# Teal Editorial PowerPoint Template

A reusable 16:9 PowerPoint design based on `Understanding_Neurodivergence.pptx`, retaining its Georgia typography, teal palette, white space, and editable geometric shapes. The original presentation is unchanged.

## Files

- **[Teal_Editorial_Template.potx](Teal_Editorial_Template.potx)** — PowerPoint template with 10 native, named slide layouts and editable placeholders. Includes examples that you can replace or delete.
- **[Teal_Editorial_Starter.pptx](Teal_Editorial_Starter.pptx)** — an ordinary presentation containing the same layouts and one example slide per layout. Save a copy before editing.
- **[Teal_Editorial_Preview.pdf](Teal_Editorial_Preview.pdf)** — visual preview of the 10 example slides; not an editable template.

## Start a presentation

1. Open the `.potx` in Microsoft PowerPoint to create a new presentation, then save it as `.pptx`. Alternatively, open the starter deck and use **Save As** to make your own copy.
2. Select **Home → New Slide** and choose one of the named layouts. Use **Home → Layout** to change an existing slide's layout.
3. Click a placeholder to enter your text. Paste text without source formatting to retain the template styling.
4. Delete unwanted example slides. The layouts remain available even after their example slides are deleted.
5. Use **Home → Reset** to restore placeholder positions. Reset does not remove your text; it may also reset placeholder formatting.

Decorative panels, lines, and cards live on the layouts, not on the editable slide layer. They remain editable through **View → Slide Master**.

## Layouts

| Layout | Best for | Suggested content limit |
|---|---|---|
| 01 Cover | Presentation title, subtitle, presenter | Title: 1–2 lines; subtitle: 1 line |
| 02 Introduction + Bullets | Lead statement plus key points | Lead: 1–2 lines; 3–5 concise bullets |
| 03 Four Cards Across | Four ideas or categories | Headings: 1–2 lines; descriptions: about 20–30 words |
| 04 Contrasting Panels | Two perspectives or alternatives | 3–5 short bullets per panel |
| 05 Eight Detail Blocks | Eight concepts with explanations | Headings: 1 line; descriptions: 1–2 lines |
| 06 Four Cards Grid | Four observations or recommendations | 1–2 short sentences per card |
| 07 Two Columns + Takeaway | Two complementary perspectives | About 3 bullets per column; takeaway: 1 line |
| 08 Six Detail Blocks + Takeaway | A six-part framework | Headings: 1 line; descriptions: 1–2 lines |
| 09 Bullet List + Takeaway | Key points and a conclusion | 3–5 concise bullets; takeaway: 1 line |
| 10 Case Study + Discussion | Scenario and discussion questions | Narrative: about 80–120 words; 4–6 short questions |

These limits are guidelines, not enforced restrictions. Split dense content across slides rather than shrinking it until it is hard to read.

## Customise the branding

- **Footer:** open **View → Slide Master**, select the top master, and edit the text box named **EDIT HERE: presentation footer**. This updates all content slides. The cover intentionally has no footer.
- **Colours and fonts:** use PowerPoint's theme colour/font controls, or the controls on the **Slide Master** tab. The template uses Georgia for both heading and body theme fonts. Keep Georgia installed for the intended appearance.
- **Bullets:** filled square bullets match the approved reference: teal on light backgrounds, white on dark teal panels. The symbol uses Arial at 70% of the text size; the surrounding text retains Georgia.
- **Slide numbers:** the bottom-right number is a dynamic field on the master, not typed text. It updates as slides are inserted, removed, or reordered. The cover is counted but does not display a number. Do not replace the field with plain text.
- **Layout artwork:** select the relevant child layout in Slide Master view to edit its cards, panels, and other decorations.

Layout changes between substantially different designs can require moving or re-entering content: a four-card layout and a single bullet list do not have a one-to-one content mapping.

## Rebuild and validation

Run `python3 build_template.py` with `python-pptx` and `lxml` installed. This recreates both PowerPoint deliverables from the original source and runs assertions for ZIP/XML integrity, internal relationships, layout/placeholder counts, shape boundaries, and inserting/saving/reopening new slides from every layout. It does not overwrite the source deck or rebuild the PDF preview.

Both deliverables were exported successfully with LibreOffice. All 10 example layouts and all 10 newly inserted, populated layouts were visually inspected. Direct Microsoft PowerPoint UI validation is still recommended before broad distribution; LibreOffice rendering is not a substitute for testing PowerPoint's layout picker and Reset command.

## Interactive presentation workshop (local Pi skill)

The project-local skill is [`.pi/skills/presentation-workshop/SKILL.md`](.pi/skills/presentation-workshop/SKILL.md). It is not installed globally.

In Pi, run `/reload`, then:

```text
/skill:presentation-workshop
```

You can add a topic or a path to your notes after the command. The skill interviews the presenter one question at a time, agrees a story, refines slides in small batches, and builds only after the content is approved. It keeps the working plan and accepted content under `presentations/<slug>/` and uses the actual POTX rather than recreating its design.

The bundled builder adds editable text, speaker notes, references, and a custom footer. It retains all ten layouts and the approved square bullets. Existing output files are never overwritten; use versioned filenames for revisions. Visual review remains a separate required step.

```bash
# Inspect actual template fields.
python3 .pi/skills/presentation-workshop/scripts/build_presentation.py --layouts

# Build accepted content (replace these paths with your workshop files).
python3 .pi/skills/presentation-workshop/scripts/build_presentation.py \
  presentations/my-talk/content.json presentations/my-talk/my-talk-v1.pptx

# Run the builder's regression checks.
python3 .pi/skills/presentation-workshop/scripts/test_build_presentation.py
```

A valid small manifest is provided at [`.pi/skills/presentation-workshop/references/example-content.json`](.pi/skills/presentation-workshop/references/example-content.json).
