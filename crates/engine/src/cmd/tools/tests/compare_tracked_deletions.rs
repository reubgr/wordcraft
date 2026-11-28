//! Regression tests for #5920: pre-existing tracked deletions in either compare input must never
//! be resurrected as live text or spurious insertions/deletions, and Accept All / Reject All over
//! the comparison output must stay coherent with the inputs' accept-resolved semantics across all
//! stories and multi-paragraph deletions.

use super::*;
use std::sync::atomic::{AtomicUsize, Ordering};

static TMP_SEQ: AtomicUsize = AtomicUsize::new(0);

fn save_temp_docx(doc: &Document, tag: &str) -> std::path::PathBuf {
    let n = TMP_SEQ.fetch_add(1, Ordering::Relaxed);
    let path = std::env::temp_dir().join(format!("wc-5920-{tag}-{}-{n}.docx", std::process::id()));
    crate::io::save_path(&path, doc).unwrap();
    path
}

fn compare_docs(orig: &Document, rev: &Document, tag: &str) -> (Session, Value) {
    let rev_path = save_temp_docx(rev, tag);
    let mut s = Session::new(orig.clone());
    let res = s.run("review.compare", &json!({"path": rev_path.to_string_lossy(), "author": "Compare Study", "showIn": "new"})).unwrap();
    let _ = std::fs::remove_file(&rev_path);
    (s, res)
}

fn accept_all_text(doc: &Document, story: StoryRef) -> String {
    let mut s = Session::new(doc.clone());
    s.run("review.acceptAll", &json!({})).unwrap();
    s.doc.plain_text(story)
}

fn reject_all_text(doc: &Document, story: StoryRef) -> String {
    let mut s = Session::new(doc.clone());
    s.run("review.rejectAll", &json!({})).unwrap();
    s.doc.plain_text(story)
}

/// Builds a document from `text` and marks the first occurrence of `needle` as a tracked deletion
/// authored by `"Prior Author"`.
fn doc_with_deleted_text(text: &str, needle: &str) -> Document {
    let mut s = Session::new(Document::from_text(text));
    s.author = "Prior Author".into();
    s.run("review.trackChanges", &json!({"value": true})).unwrap();
    s.run("select.text", &json!({"text": needle})).unwrap();
    s.run("text.delete", &json!({})).unwrap();
    s.doc
}

/// Builds a document from `text`, selects `anchor` (inserting `inserted` before it) as a tracked
/// insertion authored by `"Prior Author"`.
fn doc_with_inserted_text(prefix: &str, inserted: &str, suffix: &str) -> Document {
    let mut s = Session::new(Document::from_text(&format!("{prefix}{suffix}")));
    s.author = "Prior Author".into();
    s.run("review.trackChanges", &json!({"value": true})).unwrap();
    s.run("caret.set", &json!({"pos": Pos::body(0, prefix.len())})).unwrap();
    s.run("text.insert", &json!({"text": inserted})).unwrap();
    s.doc
}

#[test]
fn compare_does_not_resurrect_preexisting_tracked_deletion_in_base() {
    let base = doc_with_deleted_text("The draft audit timetable was agreed.", "draft ");
    let rev = Document::from_text("The audit timetable was agreed.");

    let (compared, res) = compare_docs(&base, &rev, "base-del");
    assert_eq!(res["changes"], 0, "accept-resolved inputs are identical: {res}");
    assert!(!compared.doc.plain_text(StoryRef::Body).contains("draft"), "deleted base text must not appear in compare output");
    assert_eq!(accept_all_text(&compared.doc, StoryRef::Body), "The audit timetable was agreed.");
    assert_eq!(reject_all_text(&compared.doc, StoryRef::Body), "The audit timetable was agreed.");

    // Base has a pre-existing deletion plus a real text change against revised.
    let base_edit = doc_with_deleted_text("The draft old timetable was agreed.", "draft ");
    let rev_edit = Document::from_text("The new timetable was agreed.");
    let (compared_edit, res_edit) = compare_docs(&base_edit, &rev_edit, "base-del-edit");
    assert!(res_edit["changes"].as_u64().unwrap() >= 1);
    assert!(!compared_edit.doc.plain_text(StoryRef::Body).contains("draft"), "pre-existing base deletion must stay absent");
    assert_eq!(accept_all_text(&compared_edit.doc, StoryRef::Body), "The new timetable was agreed.");
    assert_eq!(reject_all_text(&compared_edit.doc, StoryRef::Body), "The old timetable was agreed.");
}

#[test]
fn compare_does_not_resurrect_preexisting_tracked_deletion_in_revised() {
    let base = Document::from_text("The audit timetable was agreed.");
    let rev = doc_with_deleted_text("The draft audit timetable was agreed.", "draft ");

    let (compared, res) = compare_docs(&base, &rev, "rev-del");
    assert_eq!(res["changes"], 0, "accept-resolved inputs are identical: {res}");
    assert!(!compared.doc.plain_text(StoryRef::Body).contains("draft"), "deleted revised text must not appear as live text or insertion");
    assert_eq!(accept_all_text(&compared.doc, StoryRef::Body), "The audit timetable was agreed.");
    assert_eq!(reject_all_text(&compared.doc, StoryRef::Body), "The audit timetable was agreed.");

    // Revised has a pre-existing deletion plus a real edit relative to base.
    let base_edit = Document::from_text("The old timetable was agreed.");
    let rev_edit = doc_with_deleted_text("The draft new timetable was agreed.", "draft ");
    let (compared_edit, res_edit) = compare_docs(&base_edit, &rev_edit, "rev-del-edit");
    assert!(res_edit["changes"].as_u64().unwrap() >= 1);
    assert!(!compared_edit.doc.plain_text(StoryRef::Body).contains("draft"), "pre-existing revised deletion must stay absent");
    assert_eq!(accept_all_text(&compared_edit.doc, StoryRef::Body), "The new timetable was agreed.");
    assert_eq!(reject_all_text(&compared_edit.doc, StoryRef::Body), "The old timetable was agreed.");
}

#[test]
fn compare_handles_preexisting_tracked_deletions_in_both_inputs_and_mixed_same_paragraph() {
    // Different pre-existing deletions on each side, identical accept-resolved text.
    let base = doc_with_deleted_text("The draft audit timetable was agreed.", "draft ");
    let rev = doc_with_deleted_text("The interim audit timetable was agreed.", "interim ");
    let (compared, res) = compare_docs(&base, &rev, "both-del-diff");
    assert_eq!(res["changes"], 0);
    let body = compared.doc.plain_text(StoryRef::Body);
    assert!(!body.contains("draft") && !body.contains("interim"), "neither input's deletion may appear: {body:?}");
    assert_eq!(accept_all_text(&compared.doc, StoryRef::Body), "The audit timetable was agreed.");
    assert_eq!(reject_all_text(&compared.doc, StoryRef::Body), "The audit timetable was agreed.");

    // Same pre-existing deleted word on both sides: must not survive as live text or aliased del.
    let rev_same = doc_with_deleted_text("The draft audit timetable was agreed.", "draft ");
    let (compared_same, res_same) = compare_docs(&base, &rev_same, "both-del-same");
    assert_eq!(res_same["changes"], 0);
    assert!(!compared_same.doc.plain_text(StoryRef::Body).contains("draft"));
    assert_eq!(accept_all_text(&compared_same.doc, StoryRef::Body), "The audit timetable was agreed.");
    assert_eq!(reject_all_text(&compared_same.doc, StoryRef::Body), "The audit timetable was agreed.");

    // Same paragraph with base tracked insertion ("final ") and revised tracked deletion ("old ").
    let base_ins = doc_with_inserted_text("The ", "final ", "audit timetable was agreed.");
    let rev_del = doc_with_deleted_text("The old audit timetable was agreed.", "old ");
    let (compared_mix, res_mix) = compare_docs(&base_ins, &rev_del, "both-samepara");
    assert!(res_mix["changes"].as_u64().unwrap() >= 1);
    assert!(!compared_mix.doc.plain_text(StoryRef::Body).contains("old"), "revised deletion must not resurrect as insertion");
    assert_eq!(accept_all_text(&compared_mix.doc, StoryRef::Body), "The audit timetable was agreed.");
    assert_eq!(reject_all_text(&compared_mix.doc, StoryRef::Body), "The final audit timetable was agreed.");
}

#[test]
fn compare_handles_multi_paragraph_tracked_deletions_on_either_side() {
    // Delete across paragraph boundaries: tail of para 0, all of para 1, head of para 2.
    let make_multi_del = || {
        let mut s = Session::new(Document::from_text("Alpha drop0.\nWhole deleted middle.\ndrop2 Omega."));
        s.author = "Prior Author".into();
        s.run("review.trackChanges", &json!({"value": true})).unwrap();
        let a = Pos::body(0, "Alpha ".len());
        let b = Pos::body(2, "drop2 ".len());
        s.sel = Selection { anchor: a, focus: b };
        s.run("text.delete", &json!({})).unwrap();
        s.doc
    };
    let expected_accepted = "Alpha Omega.";
    let plain = Document::from_text(expected_accepted);

    // Multi-paragraph deletion in base vs accepted plain revised.
    let base_multi = make_multi_del();
    assert_eq!(accept_all_text(&base_multi, StoryRef::Body), expected_accepted);
    let (cmp_base, res_base) = compare_docs(&base_multi, &plain, "multi-base");
    assert_eq!(res_base["changes"], 0, "multi-paragraph base deletion matches accepted revised: {res_base}");
    assert_eq!(cmp_base.doc.body.len(), 1);
    assert_eq!(cmp_base.doc.plain_text(StoryRef::Body), expected_accepted);
    assert_eq!(accept_all_text(&cmp_base.doc, StoryRef::Body), expected_accepted);
    assert_eq!(reject_all_text(&cmp_base.doc, StoryRef::Body), expected_accepted);

    // Multi-paragraph deletion in revised vs accepted plain base.
    let rev_multi = make_multi_del();
    let (cmp_rev, res_rev) = compare_docs(&plain, &rev_multi, "multi-rev");
    assert_eq!(res_rev["changes"], 0, "multi-paragraph revised deletion matches accepted base: {res_rev}");
    assert_eq!(cmp_rev.doc.body.len(), 1);
    assert_eq!(cmp_rev.doc.plain_text(StoryRef::Body), expected_accepted);
    assert_eq!(accept_all_text(&cmp_rev.doc, StoryRef::Body), expected_accepted);
    assert_eq!(reject_all_text(&cmp_rev.doc, StoryRef::Body), expected_accepted);

    // Whole middle paragraph deleted (including its paragraph mark) in revised vs different surviving text in base.
    let mut s_whole = Session::new(Document::from_text("First.\nGone paragraph.\nLast."));
    s_whole.author = "Prior Author".into();
    s_whole.run("review.trackChanges", &json!({"value": true})).unwrap();
    s_whole.sel = Selection { anchor: Pos::body(1, 0), focus: Pos::body(2, 0) };
    s_whole.run("text.delete", &json!({})).unwrap();
    let base_two = Document::from_text("First.\nUpdated last.");
    let (cmp_whole, _) = compare_docs(&base_two, &s_whole.doc, "multi-whole-rev");
    assert!(!cmp_whole.doc.plain_text(StoryRef::Body).contains("Gone paragraph"));
    assert_eq!(accept_all_text(&cmp_whole.doc, StoryRef::Body), "First.\nLast.");
    assert_eq!(reject_all_text(&cmp_whole.doc, StoryRef::Body), "First.\nUpdated last.");
}

#[test]
fn compare_handles_tracked_deletions_in_tables_footnotes_endnotes_and_text_boxes() {
    // 1. Table cell tracked deletion on revised side (and base side).
    let make_table_doc = |cell_text: &str, del_word: Option<&str>| {
        let mut s = Session::new(Document::new());
        s.run("insert.table", &json!({"rows": 1, "cols": 2})).unwrap();
        s.run("text.insert", &json!({"text": cell_text})).unwrap();
        if let Some(word) = del_word {
            s.author = "Prior Author".into();
            s.run("review.trackChanges", &json!({"value": true})).unwrap();
            s.run("select.text", &json!({"text": word})).unwrap();
            s.run("text.delete", &json!({})).unwrap();
        }
        s.doc
    };
    let tbl_plain = make_table_doc("Cell value", None);
    let tbl_del = make_table_doc("Cell draft value", Some("draft "));
    let (cmp_tbl_rev, res_tbl_rev) = compare_docs(&tbl_plain, &tbl_del, "table-rev-del");
    assert_eq!(res_tbl_rev["changes"], 0);
    assert!(!cmp_tbl_rev.doc.plain_text(StoryRef::Body).contains("draft"));
    let (cmp_tbl_base, res_tbl_base) = compare_docs(&tbl_del, &tbl_plain, "table-base-del");
    assert_eq!(res_tbl_base["changes"], 0);
    assert!(!cmp_tbl_base.doc.plain_text(StoryRef::Body).contains("draft"));

    // 2. Footnote & endnote tracked deletions (both paired and newly inserted notes).
    let make_note_doc = |note_text: &str, del_word: Option<&str>, endnote: bool| {
        let mut s = session_with_note("Body paragraph.", note_text, endnote);
        if let Some(word) = del_word {
            let kind = if endnote { wordcraft_doc::PartKind::Endnote } else { wordcraft_doc::PartKind::Footnote };
            let note_id = s.doc.parts.iter().find(|(_, p)| p.kind == kind).map(|(&id, _)| id).unwrap();
            let rev_id = s.doc.revisions.len() as u32;
            s.doc.revisions.push(wordcraft_doc::Revision::new(RevisionKind::Delete, "Prior Author", "2026-10-15T09:30:00Z"));
            let p = s.doc.para_mut(StoryRef::Part(note_id), &wordcraft_doc::Path::top(0)).unwrap();
            let start = p.text.find(word).unwrap();
            p.format(start, start + word.len(), &|c| c.del = Some(rev_id)).unwrap();
        }
        s.doc
    };
    for endnote in [false, true] {
        let kind = if endnote { wordcraft_doc::PartKind::Endnote } else { wordcraft_doc::PartKind::Footnote };
        let plain_note = make_note_doc("Clean note.", None, endnote);
        let del_note = make_note_doc("Clean draft note.", Some("draft "), endnote);

        for (tag, a, b) in [("note-base-del", &del_note, &plain_note), ("note-rev-del", &plain_note, &del_note)] {
            let (cmp, res) = compare_docs(a, b, tag);
            assert_eq!(res["changes"], 0, "{tag} endnote={endnote}: {res}");
            let note_id = cmp.doc.parts.iter().find(|(_, p)| p.kind == kind).map(|(&id, _)| id).unwrap();
            assert!(!cmp.doc.plain_text(StoryRef::Part(note_id)).contains("draft"), "{tag}: deleted note word must not resurrect");
            assert_eq!(accept_all_text(&cmp.doc, StoryRef::Part(note_id)), " Clean note.");
            assert_eq!(reject_all_text(&cmp.doc, StoryRef::Part(note_id)), " Clean note.");
        }
    }

    // 3. Revised footnote whose entire text was tracked-deleted must not resurrect as unmarked live
    //    text in a paired note or as inserted text in a newly added note.
    let mut orig_fn = Session::new(Document::from_text("Para one.\nPara two."));
    orig_fn.run("caret.set", &json!({"pos": Pos::body(0, "Para one.".len())})).unwrap();
    orig_fn.run("references.footnote", &json!({"text": "Kept note."})).unwrap();
    let orig_fn_id = orig_fn.doc.parts.iter().find(|(_, p)| p.kind == wordcraft_doc::PartKind::Footnote).map(|(&id, _)| id).unwrap();

    let mut rev_fn = Session::new(Document::from_text("Para one.\nPara two."));
    rev_fn.run("caret.set", &json!({"pos": Pos::body(0, "Para one.".len())})).unwrap();
    rev_fn.run("references.footnote", &json!({"text": "Kept note."})).unwrap();
    rev_fn.run("caret.set", &json!({"pos": Pos::body(1, "Para two.".len())})).unwrap();
    rev_fn.run("references.footnote", &json!({"text": "Gone inserted note."})).unwrap();
    rev_fn.doc.revisions.push(wordcraft_doc::Revision::new(RevisionKind::Delete, "Prior Author", "2026-10-20T14:05:00Z"));
    for part in rev_fn.doc.parts.values_mut() {
        if part.kind == wordcraft_doc::PartKind::Footnote
            && let Some(Block::Para(p)) = part.blocks.first_mut().map(Arc::make_mut)
            && let Some(r) = p.runs.last_mut()
        {
            r.props.del = Some(0);
        }
    }
    let (cmp_fn, _) = compare_docs(&orig_fn.doc, &rev_fn.doc, "rev-fn-full-del");
    // Paired note: revised deleted " Kept note.", so Compare marks it deleted (Accept drops it, Reject restores it).
    assert_eq!(accept_all_text(&cmp_fn.doc, StoryRef::Part(orig_fn_id)), "");
    assert_eq!(reject_all_text(&cmp_fn.doc, StoryRef::Part(orig_fn_id)), " Kept note.");
    // Newly added note in revised had its text deleted: must not resurrect on Accept All.
    let added_fn_id =
        cmp_fn.doc.parts.iter().find(|&(&id, p)| id != orig_fn_id && p.kind == wordcraft_doc::PartKind::Footnote).map(|(&id, _)| id).unwrap();
    assert!(!cmp_fn.doc.plain_text(StoryRef::Part(added_fn_id)).contains("Gone inserted note"));
    assert_eq!(accept_all_text(&cmp_fn.doc, StoryRef::Part(added_fn_id)), "");

    // 4. Tracked deletion of a footnote reference in body drops the orphaned footnote part before compare.
    let mut rev_del_ref = session_with_note("Body paragraph.", "Orphaned note.", false);
    let ref_off = rev_del_ref.doc.body[0].as_para().unwrap().object_offsets()[0];
    let rev_id = rev_del_ref.doc.revisions.len() as u32;
    rev_del_ref.doc.revisions.push(wordcraft_doc::Revision::new(RevisionKind::Delete, "Prior Author", "2026-10-20T14:05:00Z"));
    rev_del_ref
        .doc
        .para_mut(StoryRef::Body, &wordcraft_doc::Path::top(0))
        .unwrap()
        .format(ref_off, ref_off + wordcraft_doc::para::OBJ.len_utf8(), &|c| c.del = Some(rev_id))
        .unwrap();
    let plain_body = Document::from_text("Body paragraph.");
    let (cmp_orphan, res_orphan) = compare_docs(&plain_body, &rev_del_ref.doc, "rev-del-noteref");
    assert_eq!(res_orphan["changes"], 0);
    assert!(
        !cmp_orphan.doc.parts.values().any(|p| p.kind == wordcraft_doc::PartKind::Footnote),
        "deleted footnote reference must not adopt an orphaned footnote part"
    );

    // 5. Text box tracked deletion on both sides.
    let make_tb_doc = |tb_text: &str, del_word: Option<&str>| {
        let mut s = Session::new(Document::from_text("Main body."));
        s.run("insert.textBox", &json!({"text": tb_text})).unwrap();
        if let Some(word) = del_word {
            let tb_id = s.doc.parts.iter().find(|(_, p)| p.kind == wordcraft_doc::PartKind::TextBox).map(|(&id, _)| id).unwrap();
            let rev_id = s.doc.revisions.len() as u32;
            s.doc.revisions.push(wordcraft_doc::Revision::new(RevisionKind::Delete, "Prior Author", "2026-10-15T09:30:00Z"));
            let p = s.doc.para_mut(StoryRef::Part(tb_id), &wordcraft_doc::Path::top(0)).unwrap();
            let start = p.text.find(word).unwrap();
            p.format(start, start + word.len(), &|c| c.del = Some(rev_id)).unwrap();
        }
        s.doc
    };
    let tb_plain = make_tb_doc("Box content.", None);
    let tb_del = make_tb_doc("Box draft content.", Some("draft "));
    for (tag, a, b) in [("tb-base-del", &tb_del, &tb_plain), ("tb-rev-del", &tb_plain, &tb_del)] {
        let (cmp_tb, res_tb) = compare_docs(a, b, tag);
        assert_eq!(res_tb["changes"], 0, "{tag}: {res_tb}");
        for (&id, part) in &cmp_tb.doc.parts {
            if part.kind == wordcraft_doc::PartKind::TextBox {
                assert!(!cmp_tb.doc.plain_text(StoryRef::Part(id)).contains("draft"), "{tag}: deleted text box word must not resurrect");
            }
        }
    }
}

#[test]
fn compare_preserves_preexisting_insertions_and_format_revisions_behaviour() {
    // Pre-existing insertion in revised still surfaces as a tracked insertion; Accept/Reject match.
    let base = Document::from_text("The audit timetable was agreed.");
    let rev_ins = doc_with_inserted_text("The ", "final ", "audit timetable was agreed.");
    let (cmp_ins, res_ins) = compare_docs(&base, &rev_ins, "rev-ins");
    assert!(res_ins["changes"].as_u64().unwrap() >= 1);
    assert_eq!(accept_all_text(&cmp_ins.doc, StoryRef::Body), "The final audit timetable was agreed.");
    assert_eq!(reject_all_text(&cmp_ins.doc, StoryRef::Body), "The audit timetable was agreed.");

    // Pre-existing formatting revision in base survives untouched when visible formatting matches revised.
    let mut s_base = Session::new(Document::from_text("The audit timetable."));
    s_base.author = "Prior Author".into();
    s_base.run("review.trackChanges", &json!({"value": true})).unwrap();
    s_base.run("select.text", &json!({"text": "audit"})).unwrap();
    s_base.run("format.bold", &json!({})).unwrap();
    let mut s_rev = Session::new(Document::from_text("The audit timetable."));
    s_rev.run("select.text", &json!({"text": "audit"})).unwrap();
    s_rev.run("format.bold", &json!({})).unwrap();
    let (cmp_fmt, _) = compare_docs(&s_base.doc, &s_rev.doc, "base-fmt");
    let p = cmp_fmt.doc.body[0].as_para().unwrap();
    let fmt_rev = p.run_ranges().find_map(|(_, c)| c.fmt_change.as_ref().map(|ch| ch.rev));
    assert_eq!(
        fmt_rev.and_then(|id| cmp_fmt.doc.revisions.get(id as usize)).map(|r| r.author.as_str()),
        Some("Prior Author"),
        "pre-existing base formatting revision keeps its author when formatting matches"
    );
}
