# DOCX Implementation Tasks

## Phase 1: Text Placeholder System
- [x] Create `_DOCX_TEXT_PLACEHOLDER_STASH` dict and `docx_stash_user_text()` helper
- [x] Create `docx_restore_all_stashed_text()` that replaces all placeholders with XML-escaped text
- [x] Create `docx_reset_text_placeholder_stash()` to clear state between conversions
- [x] Test: stash and restore a string with `&`, `<`, `>`, `"` characters — passed

## Phase 2: Update Inline Rules to Use Placeholders
- [ ] bold/italic/strikethrough — stash the captured text, emit XML structure with placeholder
- [ ] inline code — stash the code content
- [ ] subscript/superscript/underline — stash the text
- [ ] Test: run inline rules in docx mode, verify output has placeholders not raw text

## Phase 3: Links and Mentions
- [ ] md_link — stash link display text, emit hyperlink XML with placeholder
- [ ] bare_url — stash URL text
- [ ] mention — stash @username text
- [ ] issue_ref — stash #number text
- [ ] footnote_ref — stash footnote ID text
- [ ] Test: verify hyperlink relationship IDs are still collected correctly

## Phase 4: Block Rules
- [ ] paragraph — collect inline output, wrap in `<w:p>`, call restore at block level
- [ ] heading — emit heading style XML, stash heading text
- [ ] blockquote/alerts — stash quote text content
- [ ] list items — stash item text content
- [ ] table cells — stash cell content
- [ ] code blocks — stash each line of code
- [ ] footnote section — stash footnote text
- [ ] horizontal rule — no text, just XML structure
- [ ] Test: generate DOCX, validate XML with validator script

## Phase 5: Document Assembly
- [ ] Wire `_docx_restore_text()` into final XML before zipping
- [ ] Build document.xml.rels with hyperlink relationships
- [ ] Build styles.xml with heading styles
- [ ] Zip into .docx
- [ ] Test: open in Word/LibreOffice, verify content renders

## Phase 6: Verify RTF Unchanged
- [ ] Diff RTF output against known-good to confirm no regressions
