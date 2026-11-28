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

/// Builds a paragraph containing `discard` deleted by author `"D"` immediately followed by `keep`
/// inserted by author `"I"` (plus optional bold formatting on `keep`).
fn para_adjacent_foreign_del_and_ins(doc: &mut Document, discard: &str, keep: &str, keep_bold: bool) -> Paragraph {
    let del_id = doc.revisions.len() as u32;
    doc.revisions.push(wordcraft_doc::Revision::new(RevisionKind::Delete, "Author D", "2026-10-15T09:30:00Z"));
    let ins_id = doc.revisions.len() as u32;
    doc.revisions.push(wordcraft_doc::Revision::new(RevisionKind::Insert, "Author I", "2026-10-16T10:00:00Z"));
    let mut p = Paragraph::new();
    p.insert_text(0, discard, &CharProps { del: Some(del_id), ..Default::default() }).unwrap();
    let at = p.len();
    p.insert_text(at, keep, &CharProps { ins: Some(ins_id), bold: keep_bold.then_some(true), ..Default::default() }).unwrap();
    p
}

#[test]
fn compare_equal_path_strips_foreign_insertions_across_adjacent_authors_tables_and_text_boxes() {
    // 1. Body paragraph: base plain "keep"; revised DOCX has Author D deleting "discard" right
    //    before Author I inserting "keep". Both accept-resolved inputs read "keep".
    let base = Document::from_text("keep");
    let mut rev = Document::new();
    rev.body[0] = para_block(para_adjacent_foreign_del_and_ins(&mut rev, "discard", "keep", false));

    let (compared, res) = compare_docs(&base, &rev, "adj-authors");
    assert_eq!(res["changes"], 0, "accept-resolved texts are identical: {res}");
    assert_eq!(compared.doc.plain_text(StoryRef::Body), "keep");
    assert_eq!(accept_all_text(&compared.doc, StoryRef::Body), "keep");
    assert_eq!(reject_all_text(&compared.doc, StoryRef::Body), "keep", "Reject All must not erase unchanged text");

    // Same probe with formatting: false.
    let rev_path = save_temp_docx(&rev, "adj-authors-nofmt");
    let mut s_nf = Session::new(base.clone());
    let res_nf = s_nf
        .run("review.compare", &json!({"path": rev_path.to_string_lossy(), "author": "Compare Study", "formatting": false, "showIn": "new"}))
        .unwrap();
    let _ = std::fs::remove_file(&rev_path);
    assert_eq!(res_nf["changes"], 0);
    assert_eq!(reject_all_text(&s_nf.doc, StoryRef::Body), "keep");

    // Same probe where Author I also made "keep" bold: formatting change must be tracked, and
    // Reject All must restore unbold "keep" rather than deleting it.
    let mut rev_bold = Document::new();
    rev_bold.body[0] = para_block(para_adjacent_foreign_del_and_ins(&mut rev_bold, "discard", "keep", true));
    let (cmp_bold, res_bold) = compare_docs(&base, &rev_bold, "adj-authors-bold");
    assert!(res_bold["changes"].as_u64().unwrap() >= 1);
    let mut s_acc = Session::new(cmp_bold.doc.clone());
    s_acc.run("review.acceptAll", &json!({})).unwrap();
    assert_eq!(s_acc.doc.plain_text(StoryRef::Body), "keep");
    assert_eq!(s_acc.doc.body[0].as_para().unwrap().props_of_char(0).bold, Some(true));
    let mut s_rej = Session::new(cmp_bold.doc.clone());
    s_rej.run("review.rejectAll", &json!({})).unwrap();
    assert_eq!(s_rej.doc.plain_text(StoryRef::Body), "keep");
    assert_ne!(s_rej.doc.body[0].as_para().unwrap().props_of_char(0).bold, Some(true));

    // 2. Matched table cell with adjacent foreign deletion + insertion on revised side.
    let make_single_cell_table = |p: Paragraph| {
        let mut cell = wordcraft_doc::table::Cell::empty();
        cell.blocks = vec![para_block(p)];
        let tbl = wordcraft_doc::table::Table {
            grid: vec![200.0],
            rows: vec![wordcraft_doc::table::Row { cells: vec![cell], ..Default::default() }],
            ..Default::default()
        };
        Arc::new(Block::Table(tbl))
    };
    let mut tbl_base = Document::new();
    tbl_base.body = vec![make_single_cell_table(Paragraph::with_text("keep", CharProps::default())), para_block(Paragraph::new())];
    let mut tbl_rev = Document::new();
    let rev_cell_para = para_adjacent_foreign_del_and_ins(&mut tbl_rev, "discard", "keep", false);
    tbl_rev.body = vec![make_single_cell_table(rev_cell_para), para_block(Paragraph::new())];
    let (cmp_tbl, res_tbl) = compare_docs(&tbl_base, &tbl_rev, "adj-authors-table");
    assert_eq!(res_tbl["changes"], 0, "matched table must report zero changes: {res_tbl}");
    assert_eq!(accept_all_text(&cmp_tbl.doc, StoryRef::Body).trim(), "keep");
    assert_eq!(reject_all_text(&cmp_tbl.doc, StoryRef::Body).trim(), "keep");

    // 3. Adopted text box with adjacent foreign deletion + insertion on revised side.
    let mut tb_base = Session::new(Document::from_text("Main."));
    tb_base.run("insert.textBox", &json!({"text": "keep"})).unwrap();
    let mut tb_rev = Session::new(Document::from_text("Main."));
    tb_rev.run("insert.textBox", &json!({"text": "placeholder"})).unwrap();
    let rev_tb_id = tb_rev.doc.parts.iter().find(|(_, p)| p.kind == wordcraft_doc::PartKind::TextBox).map(|(&id, _)| id).unwrap();
    let tb_para = para_adjacent_foreign_del_and_ins(&mut tb_rev.doc, "discard", "keep", false);
    tb_rev.doc.parts.get_mut(&rev_tb_id).unwrap().blocks = vec![para_block(tb_para)];
    let (cmp_tb, res_tb) = compare_docs(&tb_base.doc, &tb_rev.doc, "adj-authors-tb");
    assert_eq!(res_tb["changes"], 0, "equal text box must report zero changes: {res_tb}");
    let active_tb_id = cmp_tb.doc.body[0].as_para().unwrap().objects.iter().find_map(|o| o.text_box()).unwrap();
    assert_eq!(accept_all_text(&cmp_tb.doc, StoryRef::Part(active_tb_id)), "keep");
    assert_eq!(reject_all_text(&cmp_tb.doc, StoryRef::Part(active_tb_id)), "keep");

    // 4. Foreign paragraph mark insertion on equal multi-paragraph revised document.
    let base_multi = Document::from_text("Alpha\nBeta");
    let mut rev_multi = Document::from_text("Alpha\nBeta");
    rev_multi.revisions.push(wordcraft_doc::Revision::new(RevisionKind::Insert, "Author I", "2026-10-16T10:00:00Z"));
    rev_multi.para_mut(StoryRef::Body, &wordcraft_doc::Path::top(0)).unwrap().mark.ins = Some(0);
    let (cmp_mark, res_mark) = compare_docs(&base_multi, &rev_multi, "adj-mark-ins");
    assert_eq!(res_mark["changes"], 0);
    assert_eq!(accept_all_text(&cmp_mark.doc, StoryRef::Body), "Alpha\nBeta");
    assert_eq!(reject_all_text(&cmp_mark.doc, StoryRef::Body), "Alpha\nBeta");

    // 5. Whole-paragraph pre-existing insertion in base removed in revised: Reject All restores it.
    let mut base_whole_ins = Document::from_text("Extra\nKeep");
    base_whole_ins.revisions.push(wordcraft_doc::Revision::new(RevisionKind::Insert, "Author I", "2026-10-15T09:30:00Z"));
    let p0 = base_whole_ins.para_mut(StoryRef::Body, &wordcraft_doc::Path::top(0)).unwrap();
    let len0 = p0.len();
    p0.format(0, len0, &|c| c.ins = Some(0)).unwrap();
    p0.mark.ins = Some(0);
    let rev_only_keep = Document::from_text("Keep");
    let (cmp_whole, res_whole) = compare_docs(&base_whole_ins, &rev_only_keep, "base-whole-ins");
    assert!(res_whole["changes"].as_u64().unwrap() >= 1);
    assert_eq!(accept_all_text(&cmp_whole.doc, StoryRef::Body), "Keep");
    assert_eq!(reject_all_text(&cmp_whole.doc, StoryRef::Body), "Extra\nKeep");
}

/// Builds a document containing a chain of `depth` nested 1x1 tables whose innermost cell holds
/// `leaf`, returning `(doc, leaf_path)`. Enclosing cells end with a trailing empty paragraph as
/// OOXML (`w:tc`) requires, so DOCX round-trips keep the exact same cell block structure.
fn doc_with_nested_table_leaf(depth: usize, leaf: Paragraph) -> (Document, wordcraft_doc::Path) {
    let mut current: Arc<Block> = para_block(leaf);
    for level in 0..depth {
        let mut cell = wordcraft_doc::table::Cell::empty();
        cell.blocks = if level == 0 { vec![current] } else { vec![current, para_block(Paragraph::new())] };
        let tbl = wordcraft_doc::table::Table {
            grid: vec![120.0],
            rows: vec![wordcraft_doc::table::Row { cells: vec![cell], ..Default::default() }],
            ..Default::default()
        };
        current = Arc::new(Block::Table(tbl));
    }
    let mut doc = Document::new();
    doc.body = vec![current, para_block(Paragraph::new())];
    let mut raw = Vec::with_capacity(depth * 3 + 1);
    for _ in 0..depth {
        raw.extend_from_slice(&[0, 0, 0]);
    }
    raw.push(0);
    (doc, wordcraft_doc::Path(raw))
}

/// Round-trips `doc` through DOCX so in-memory base fixtures and DOCX-loaded revised fixtures
/// share identical serialization defaults (table properties, note reference formatting, etc.).
fn roundtrip_docx(doc: &Document, tag: &str) -> Document {
    let path = save_temp_docx(doc, tag);
    let reopened = crate::io::open_path(&path).unwrap();
    let _ = std::fs::remove_file(&path);
    reopened
}

#[test]
fn compare_resolves_tracked_deletions_past_depth_16_and_preserves_deep_note_refs() {
    // 1. Tracked deletion at depth 17 and depth 18 (within MAX_TABLE_DEPTH = 18, past PARA_WALK_DEPTH = 16).
    for depth in [17usize, 18] {
        let (plain_doc, leaf_path) = doc_with_nested_table_leaf(depth, Paragraph::new());
        let mut del_leaf = Paragraph::new();
        del_leaf.insert_text(0, "draft", &CharProps { del: Some(0), ..Default::default() }).unwrap();
        let (mut del_doc, del_path) = doc_with_nested_table_leaf(depth, del_leaf);
        del_doc.revisions.push(wordcraft_doc::Revision::new(RevisionKind::Delete, "Prior Author", "2026-10-15T09:30:00Z"));
        assert_eq!(leaf_path, del_path);
        assert_eq!(del_doc.para(StoryRef::Body, &leaf_path).unwrap().plain_text(), "draft");

        for (tag, a, b) in [(format!("deep-{depth}-base"), &del_doc, &plain_doc), (format!("deep-{depth}-rev"), &plain_doc, &del_doc)] {
            let (cmp, res) = compare_docs(a, b, &tag);
            assert_eq!(res["changes"], 0, "{tag}: accepted depth-{depth} cells match: {res}");
            let leaf = cmp.doc.para(StoryRef::Body, &leaf_path).expect("explicit deep leaf path must exist");
            assert_eq!(leaf.plain_text(), "", "{tag}: depth-{depth} tracked deletion must be removed");
            assert!(leaf.runs.iter().all(|r| r.props.del.is_none()), "{tag}: no leftover del run at depth {depth}");
        }
    }

    // 2. Footnote whose shallow body reference is tracked-deleted while a depth-17 cell holds a
    //    surviving reference: the full-depth liveness scan must preserve the footnote part.
    let mut deep_ref_leaf = Paragraph::with_text("deep ", CharProps::default());
    let (mut doc_shared, leaf_path) = doc_with_nested_table_leaf(17, Paragraph::new());
    let fn_id = doc_shared
        .add_part(wordcraft_doc::PartKind::Footnote, vec![para_block(Paragraph::with_text("Deeply referenced note.", CharProps::default()))]);
    let note_ref = InlineObject::NoteRef { kind: wordcraft_doc::para::NoteKind::Footnote, id: fn_id, custom: String::new() };
    let at_leaf = deep_ref_leaf.len();
    deep_ref_leaf.insert_object(at_leaf, note_ref, &CharProps::default()).unwrap();
    *doc_shared.para_mut(StoryRef::Body, &leaf_path).unwrap() = deep_ref_leaf;

    let surviving_doc = roundtrip_docx(&doc_shared, "deep-note-ref-rt");
    let surviving_fn_id = surviving_doc.parts.iter().find(|(_, p)| p.kind == wordcraft_doc::PartKind::Footnote).map(|(&id, _)| id).unwrap();
    let mut doc_with_shallow_del = surviving_doc.clone();
    let del_rev = doc_with_shallow_del.revisions.len() as u32;
    doc_with_shallow_del.revisions.push(wordcraft_doc::Revision::new(RevisionKind::Delete, "Prior Author", "2026-10-15T09:30:00Z"));
    let shallow = doc_with_shallow_del.para_mut(StoryRef::Body, &wordcraft_doc::Path::top(1)).unwrap();
    shallow
        .insert_object(
            0,
            InlineObject::NoteRef { kind: wordcraft_doc::para::NoteKind::Footnote, id: surviving_fn_id, custom: String::new() },
            &CharProps { del: Some(del_rev), ..Default::default() },
        )
        .unwrap();

    let (cmp_deep_ref, res_deep_ref) = compare_docs(&doc_with_shallow_del, &surviving_doc, "deep-note-ref");
    assert_eq!(res_deep_ref["changes"], 0, "surviving deep footnote ref must match: {res_deep_ref}");
    let kept_fn = cmp_deep_ref
        .doc
        .parts
        .values()
        .find(|p| p.kind == wordcraft_doc::PartKind::Footnote)
        .expect("footnote referenced at depth 17 must not be pruned");
    assert!(kept_fn.blocks[0].as_para().unwrap().plain_text().contains("Deeply referenced note."));
}

fn text_box_shape(story_id: u32) -> InlineObject {
    InlineObject::Shape {
        kind: wordcraft_doc::para::ShapeKind::TextBox,
        w: 120.0,
        h: 40.0,
        fill: None,
        stroke: None,
        stroke_width: 1.0,
        float: Default::default(),
        story: Some(story_id),
        freeform: None,
        effects: Default::default(),
        extra: Default::default(),
    }
}

/// Builds a document whose body paragraph holds `"Body."` followed by a tracked-deleted text-box
/// anchor; inside the text box sits a live `NoteRef` to a footnote with `note_text`.
fn doc_with_deleted_text_box_holding_footnote(note_text: &str) -> (Document, u32, u32) {
    let mut doc = Document::from_text("Body.");
    let mut fn_para = Paragraph::new();
    let fn_id = doc.add_part(wordcraft_doc::PartKind::Footnote, Vec::new());
    fn_para
        .insert_object(
            0,
            InlineObject::NoteRef { kind: wordcraft_doc::para::NoteKind::Footnote, id: fn_id, custom: String::new() },
            &CharProps::default(),
        )
        .unwrap();
    let at = fn_para.len();
    fn_para.insert_text(at, &format!(" {note_text}"), &CharProps::default()).unwrap();
    doc.parts.get_mut(&fn_id).unwrap().blocks = vec![para_block(fn_para)];

    let mut tb_para = Paragraph::with_text("Box", CharProps::default());
    let at_tb = tb_para.len();
    tb_para
        .insert_object(
            at_tb,
            InlineObject::NoteRef { kind: wordcraft_doc::para::NoteKind::Footnote, id: fn_id, custom: String::new() },
            &CharProps::default(),
        )
        .unwrap();
    let tb_id = doc.add_part(wordcraft_doc::PartKind::TextBox, vec![para_block(tb_para)]);

    let del_id = doc.revisions.len() as u32;
    doc.revisions.push(wordcraft_doc::Revision::new(RevisionKind::Delete, "Prior Author", "2026-10-15T09:30:00Z"));
    let body_p = doc.para_mut(StoryRef::Body, &wordcraft_doc::Path::top(0)).unwrap();
    let at_body = body_p.len();
    body_p.insert_object(at_body, text_box_shape(tb_id), &CharProps { del: Some(del_id), ..Default::default() }).unwrap();
    (doc, tb_id, fn_id)
}

#[test]
fn compare_prunes_notes_orphaned_by_deleted_text_box_anchors_and_preserves_shared_and_unreferenced_parts() {
    let plain = Document::from_text("Body.");

    // 1. Deleted text-box anchor with internal footnote in the in-memory base.
    let (base_del_tb, _, _) = doc_with_deleted_text_box_holding_footnote("Orphan box note.");
    let (cmp_base, res_base) = compare_docs(&base_del_tb, &plain, "del-tb-fn-base");
    assert_eq!(res_base["changes"], 0, "base deleted box + footnote matches plain revised: {res_base}");
    assert!(
        !cmp_base.doc.parts.values().any(|p| p.kind == wordcraft_doc::PartKind::Footnote),
        "base deleted text-box anchor must not leave orphan footnote in compare output"
    );
    assert!(
        !cmp_base.doc.parts.values().any(|p| p.kind == wordcraft_doc::PartKind::TextBox),
        "base deleted text-box part should be pruned once unreachable"
    );

    // 2. Deleted text-box anchor with internal footnote in the revised DOCX.
    let (rev_del_tb, _, _) = doc_with_deleted_text_box_holding_footnote("Orphan box note.");
    let rev_path = save_temp_docx(&rev_del_tb, "del-tb-fn-rev-check");
    let reopened = crate::io::open_path(&rev_path).unwrap();
    let _ = std::fs::remove_file(&rev_path);
    assert!(
        reopened.parts.values().any(|p| p.kind == wordcraft_doc::PartKind::Footnote),
        "DOCX round-trip must preserve the footnote inside the deleted text box"
    );
    let (cmp_rev, res_rev) = compare_docs(&plain, &rev_del_tb, "del-tb-fn-rev");
    assert_eq!(res_rev["changes"], 0, "revised deleted box + footnote matches plain base: {res_rev}");
    assert!(
        !cmp_rev.doc.parts.values().any(|p| p.kind == wordcraft_doc::PartKind::Footnote),
        "revised deleted text-box anchor must not adopt orphan footnote"
    );

    // 3. Nested text box inside deleted text box holding a footnote (transitive reachability).
    let mut nested_doc = Document::from_text("Body.");
    let nested_fn =
        nested_doc.add_part(wordcraft_doc::PartKind::Footnote, vec![para_block(Paragraph::with_text("Nested box note.", CharProps::default()))]);
    let mut inner_p = Paragraph::with_text("Inner", CharProps::default());
    let at_inner = inner_p.len();
    inner_p
        .insert_object(
            at_inner,
            InlineObject::NoteRef { kind: wordcraft_doc::para::NoteKind::Footnote, id: nested_fn, custom: String::new() },
            &CharProps::default(),
        )
        .unwrap();
    let inner_tb = nested_doc.add_part(wordcraft_doc::PartKind::TextBox, vec![para_block(inner_p)]);
    let mut outer_p = Paragraph::with_text("Outer", CharProps::default());
    let at_outer = outer_p.len();
    outer_p.insert_object(at_outer, text_box_shape(inner_tb), &CharProps::default()).unwrap();
    let outer_tb = nested_doc.add_part(wordcraft_doc::PartKind::TextBox, vec![para_block(outer_p)]);
    nested_doc.revisions.push(wordcraft_doc::Revision::new(RevisionKind::Delete, "Prior Author", "2026-10-15T09:30:00Z"));
    let bp = nested_doc.para_mut(StoryRef::Body, &wordcraft_doc::Path::top(0)).unwrap();
    let at_bp = bp.len();
    bp.insert_object(at_bp, text_box_shape(outer_tb), &CharProps { del: Some(0), ..Default::default() }).unwrap();
    let (cmp_nested, res_nested) = compare_docs(&nested_doc, &plain, "del-nested-tb-fn");
    assert_eq!(res_nested["changes"], 0);
    assert!(!cmp_nested.doc.parts.values().any(|p| p.kind == wordcraft_doc::PartKind::Footnote));

    // 4. Shared footnote: referenced inside deleted text box AND by a surviving body NoteRef.
    let clean_shared = session_with_note("Body.", "Shared note.", false).doc;
    let shared_fn_id = clean_shared.parts.iter().find(|(_, p)| p.kind == wordcraft_doc::PartKind::Footnote).map(|(&id, _)| id).unwrap();
    let mut shared_doc = clean_shared.clone();
    let mut tb_para = Paragraph::with_text("Box", CharProps::default());
    let at_tb = tb_para.len();
    tb_para
        .insert_object(
            at_tb,
            InlineObject::NoteRef { kind: wordcraft_doc::para::NoteKind::Footnote, id: shared_fn_id, custom: String::new() },
            &CharProps::default(),
        )
        .unwrap();
    let shared_tb_id = shared_doc.add_part(wordcraft_doc::PartKind::TextBox, vec![para_block(tb_para)]);
    let del_id = shared_doc.revisions.len() as u32;
    shared_doc.revisions.push(wordcraft_doc::Revision::new(RevisionKind::Delete, "Prior Author", "2026-10-15T09:30:00Z"));
    let bp_shared = shared_doc.para_mut(StoryRef::Body, &wordcraft_doc::Path::top(0)).unwrap();
    let at_bp = bp_shared.len();
    bp_shared.insert_object(at_bp, text_box_shape(shared_tb_id), &CharProps { del: Some(del_id), ..Default::default() }).unwrap();
    let (cmp_shared, res_shared) = compare_docs(&shared_doc, &clean_shared, "del-tb-shared-fn");
    assert_eq!(res_shared["changes"], 0, "shared footnote must survive deleted box anchor: {res_shared}");
    assert!(
        cmp_shared
            .doc
            .parts
            .values()
            .any(|p| p.kind == wordcraft_doc::PartKind::Footnote && p.blocks[0].as_para().unwrap().plain_text().contains("Shared note.")),
        "genuinely shared footnote must be preserved"
    );

    // 5. Unrelated previously unreferenced footnote part is preserved when another deleted text
    //    box's footnote is pruned.
    let (mut unrel_base, _, _) = doc_with_deleted_text_box_holding_footnote("Orphan box note.");
    unrel_base
        .add_part(wordcraft_doc::PartKind::Footnote, vec![para_block(Paragraph::with_text("Unrelated pre-existing note.", CharProps::default()))]);
    let mut unrel_rev = Document::from_text("Body.");
    unrel_rev
        .add_part(wordcraft_doc::PartKind::Footnote, vec![para_block(Paragraph::with_text("Unrelated pre-existing note.", CharProps::default()))]);
    let (cmp_unrel, res_unrel) = compare_docs(&unrel_base, &unrel_rev, "del-tb-unrel-fn");
    assert_eq!(res_unrel["changes"], 0);
    let surviving_notes: Vec<String> = cmp_unrel
        .doc
        .parts
        .values()
        .filter(|p| p.kind == wordcraft_doc::PartKind::Footnote)
        .map(|p| p.blocks[0].as_para().unwrap().plain_text())
        .collect();
    assert_eq!(surviving_notes, vec!["Unrelated pre-existing note."]);

    // 6. Shared text box: two anchors point to the same box, only one is deleted -> box and note survive.
    let (mut shared_tb, tb_id, _) = doc_with_deleted_text_box_holding_footnote("Shared box note.");
    shared_tb.para_mut(StoryRef::Body, &wordcraft_doc::Path::top(0)).unwrap().insert_object(0, text_box_shape(tb_id), &CharProps::default()).unwrap();
    let shared_tb = roundtrip_docx(&shared_tb, "shared-tb-rt");
    let mut accepted_shared_tb = Session::new(shared_tb.clone());
    accepted_shared_tb.run("review.acceptAll", &json!({})).unwrap();
    let (cmp_shared_tb, res_shared_tb) = compare_docs(&shared_tb, &accepted_shared_tb.doc, "shared-tb-anchor");
    assert_eq!(res_shared_tb["changes"], 0, "shared text box must match accepted counterpart: {res_shared_tb}");
    assert!(cmp_shared_tb.doc.parts.values().any(|p| p.kind == wordcraft_doc::PartKind::Footnote));
}

#[test]
fn compare_deletion_prepass_matches_accept_all_for_sections_headers_footers_hyperlinks_and_fields() {
    let mut doc = Document::new();
    let del_rev = doc.revisions.len() as u32;
    doc.revisions.push(wordcraft_doc::Revision::new(RevisionKind::Delete, "Prior Author", "2026-10-15T09:30:00Z"));

    let mut hdr_para = Paragraph::new();
    hdr_para.insert_text(0, "DropHdr ", &CharProps { del: Some(del_rev), ..Default::default() }).unwrap();
    let at = hdr_para.len();
    hdr_para.insert_text(at, "Link", &CharProps { link: Some("https://example.org/docs".into()), ..Default::default() }).unwrap();
    let hdr_id = doc.add_part(wordcraft_doc::PartKind::Header, vec![para_block(hdr_para)]);

    let mut ftr_para = Paragraph::new();
    ftr_para.insert_object(0, InlineObject::FieldStart { instr: "PAGE".into(), locked: false, dirty: false }, &CharProps::default()).unwrap();
    let at = ftr_para.len();
    ftr_para.insert_text(at, "old ", &CharProps { del: Some(del_rev), ..Default::default() }).unwrap();
    let at = ftr_para.len();
    ftr_para.insert_text(at, "42", &CharProps::default()).unwrap();
    let at = ftr_para.len();
    ftr_para.insert_object(at, InlineObject::FieldEnd, &CharProps::default()).unwrap();
    let at = ftr_para.len();
    ftr_para
        .insert_object(
            at,
            InlineObject::Field { instr: "NUMPAGES".into(), result: "9".into(), locked: false, dirty: false, code: false, form: String::new() },
            &CharProps::default(),
        )
        .unwrap();
    let ftr_id = doc.add_part(wordcraft_doc::PartKind::Footer, vec![para_block(ftr_para)]);

    // Paragraph 0: empty predecessor whose mark is tracked-deleted and carries a section break.
    let mut s0 = wordcraft_doc::section::SectionProps::default();
    s0.headers.default = Some(hdr_id);
    let mut p0 = Paragraph::new();
    p0.insert_text(0, "gone0", &CharProps { del: Some(del_rev), ..Default::default() }).unwrap();
    p0.mark.del = Some(del_rev);
    p0.section = Some(Box::new(s0));

    // Paragraph 1: nonempty predecessor whose mark is tracked-deleted and carries a section break.
    let mut s1 = wordcraft_doc::section::SectionProps::default();
    s1.footers.default = Some(ftr_id);
    let mut p1 = Paragraph::with_text("Keep1 ", CharProps::default());
    let at1 = p1.len();
    p1.insert_text(at1, "drop1", &CharProps { del: Some(del_rev), ..Default::default() }).unwrap();
    p1.mark.del = Some(del_rev);
    p1.section = Some(Box::new(s1));

    // Paragraph 2: trailing paragraph joined into p1.
    let p2 = Paragraph::with_text("Keep2", CharProps::default());

    doc.body = vec![para_block(p0), para_block(p1), para_block(p2)];

    let mut accepted = Session::new(doc.clone());
    accepted.run("review.acceptAll", &json!({})).unwrap();

    let (cmp, res) = compare_docs(&doc, &accepted.doc, "sections-hf-link-field");
    assert_eq!(res["changes"], 0, "normalized body must match engine Accept All: {res}");
    assert_eq!(cmp.doc.plain_text(StoryRef::Body), accepted.doc.plain_text(StoryRef::Body));
    assert_eq!(cmp.doc.last_section, accepted.doc.last_section);

    let cmp_hdr = cmp.doc.parts.get(&hdr_id).unwrap().blocks[0].as_para().unwrap();
    let acc_hdr = accepted.doc.parts.get(&hdr_id).unwrap().blocks[0].as_para().unwrap();
    assert_eq!(cmp_hdr.plain_text(), acc_hdr.plain_text());
    assert_eq!(cmp_hdr.props_of_char(0).link.as_deref(), Some("https://example.org/docs"));

    let cmp_ftr = cmp.doc.parts.get(&ftr_id).unwrap().blocks[0].as_para().unwrap();
    let acc_ftr = accepted.doc.parts.get(&ftr_id).unwrap().blocks[0].as_para().unwrap();
    assert_eq!(cmp_ftr.plain_text(), acc_ftr.plain_text());
    assert_eq!(cmp_ftr.objects, acc_ftr.objects, "rich-field start/end markers and simple field must match Accept All");
}

#[test]
fn compare_compacts_many_alternating_deleted_runs_and_note_separators() {
    let mut doc = session_with_note("", "Kept footnote.", false).doc;
    let kept_fn = doc.parts.iter().find(|(_, p)| p.kind == wordcraft_doc::PartKind::Footnote).map(|(&id, _)| id).unwrap();
    let kept_ref_props = doc.body[0].as_para().unwrap().props_of_char(0).clone();
    let kept_ref_obj = doc.body[0].as_para().unwrap().objects[0].clone();

    let del_rev = doc.revisions.len() as u32;
    doc.revisions.push(wordcraft_doc::Revision::new(RevisionKind::Delete, "Prior Author", "2026-10-15T09:30:00Z"));
    let dropped_fn =
        doc.add_part(wordcraft_doc::PartKind::Footnote, vec![para_block(Paragraph::with_text("Dropped footnote.", CharProps::default()))]);

    // Paragraph 0: 250 alternating differently-formatted tracked-deleted runs interleaved with
    // surviving runs and note references.
    let mut p0 = Paragraph::new();
    let mut expected_p0 = Paragraph::new();
    for i in 0..250usize {
        let at = p0.len();
        let del_props = CharProps {
            del: Some(del_rev),
            bold: Some(i % 2 == 0),
            italic: Some(i % 3 == 0),
            size: Some(10.0 + (i % 5) as f32),
            ..Default::default()
        };
        p0.insert_text(at, &format!("d{i}"), &del_props).unwrap();
        if i == 100 {
            let at_obj = p0.len();
            p0.insert_object(
                at_obj,
                InlineObject::NoteRef { kind: wordcraft_doc::para::NoteKind::Footnote, id: dropped_fn, custom: String::new() },
                &del_props,
            )
            .unwrap();
        }
        if i % 50 == 0 {
            let live = format!("k{i} ");
            let at_live = p0.len();
            p0.insert_text(at_live, &live, &CharProps::default()).unwrap();
            let at_exp = expected_p0.len();
            expected_p0.insert_text(at_exp, &live, &CharProps::default()).unwrap();
        }
        if i == 150 {
            let at_obj = p0.len();
            p0.insert_object(at_obj, kept_ref_obj.clone(), &kept_ref_props).unwrap();
            let at_exp = expected_p0.len();
            expected_p0.insert_object(at_exp, kept_ref_obj.clone(), &kept_ref_props).unwrap();
        }
    }

    // Paragraph 1: all runs deleted across distinct formats -> emptying paragraph keeps first run's
    // formatting on `mark` (matching `Paragraph::delete`).
    let mut p1 = Paragraph::new();
    p1.insert_text(0, "goneA", &CharProps { del: Some(del_rev), bold: Some(true), ..Default::default() }).unwrap();
    let at1 = p1.len();
    p1.insert_text(at1, "goneB", &CharProps { del: Some(del_rev), italic: Some(true), ..Default::default() }).unwrap();

    doc.body = vec![para_block(p0), para_block(p1)];

    // Footnote separator with multiple alternating deleted runs and a deleted paragraph mark.
    let mut sep_p = Paragraph::new();
    sep_p.insert_text(0, "del1", &CharProps { del: Some(del_rev), bold: Some(true), ..Default::default() }).unwrap();
    let s_at = sep_p.len();
    sep_p.insert_text(s_at, "del2", &CharProps { del: Some(del_rev), italic: Some(true), ..Default::default() }).unwrap();
    let s_at = sep_p.len();
    sep_p.insert_text(s_at, "---", &CharProps::default()).unwrap();
    sep_p.mark.del = Some(del_rev);
    doc.footnote_separators.separator = Some(vec![para_block(sep_p)]);

    let mut expected_doc = session_with_note("", "Kept footnote.", false).doc;
    let exp_fn = expected_doc.parts.iter().find(|(_, p)| p.kind == wordcraft_doc::PartKind::Footnote).map(|(&id, _)| id).unwrap();
    for o in &mut expected_p0.objects {
        if let InlineObject::NoteRef { id, .. } = o {
            *id = exp_fn;
        }
    }
    let mut exp_p1 = Paragraph::new();
    exp_p1.mark.bold = Some(true);
    expected_doc.body = vec![para_block(expected_p0), para_block(exp_p1)];

    let (cmp, res) = compare_docs(&doc, &expected_doc, "compact-many-runs");
    assert_eq!(res["changes"], 0, "compacted document must match expected: {res}");
    let out_p0 = cmp.doc.body[0].as_para().unwrap();
    assert_eq!(out_p0.plain_text(), "k0 k50 k100 k150 k200 ");
    assert_eq!(out_p0.objects.len(), 1, "only the surviving NoteRef remains");
    assert!(cmp.doc.parts.contains_key(&kept_fn), "surviving footnote must stay");
    assert!(!cmp.doc.parts.contains_key(&dropped_fn), "deleted footnote must be pruned");
    assert_eq!(cmp.doc.body[1].as_para().unwrap().mark.bold, Some(true), "emptied paragraph must keep first deleted run's formatting on mark");
    let sep_out = cmp.doc.footnote_separators.separator.as_ref().unwrap()[0].as_para().unwrap();
    assert_eq!(sep_out.plain_text(), "---");
    assert!(sep_out.mark.del.is_none());
}
