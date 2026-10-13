# Spreadsheet formats

The app reads two Excel files, both picked on the Configuration page. For each file it searches the **first 15 rows** for the header row, so a title or blank rows above the headers are fine. Header matching ignores case, extra spaces and punctuation, and treats `#`, `No` and `Number` as the same word.

If a required column can't be found, the step stops with a message that names the file, the tab, the missing columns and the headers it did see. Logic: `_read_required_columns` in `Merge.py`.

## Cover Sheet Spreadsheet

The app reads the first tab.

| Required column | Also accepted |
|---|---|
| Account Number | Account # |
| Preacher Number | Preacher #, Preacher ID |
| Preacher Name | |
| Envelope Number | Envelope #, Donor #, Donor Number |
| Envelope Name | Donor Name |

These columns are optional and are used when present: `Primary Street`, `Primary City`, `Primary State`, `Primary ZIP Code`, `Primary Email Address` and `Send Quarterly Report Via`. When the last one is `E-Mail`, the output file gets the `e` suffix.

The app drops rows with no Account Number. Each row is one donor–preacher pair, so a donor with several preachers has several rows.

## Extra Gift Report Spreadsheet

The app reads the tab named **Spons Xtra**. Matching ignores case and spacing. If no tab has that name, it falls back to the first tab, but only when that tab has the required columns.

| Required column | Also accepted |
|---|---|
| ID (preacher number) | Preacher ID, Preacher # |
| Donor # | Donor Number, Donor ID |

The cover-sheet names `Preacher Number` and `Envelope Number` are deliberately **not** accepted here. If the cover sheet is picked in this box by mistake, the step fails instead of quietly flagging every donor for a thank-you note.

Rows are skipped when they:
- have a blank ID or Donor #,
- have note text in the Donor # column,
- or list a donor # on the hard-coded `sponsorBlacklist`.

Donor #s that aren't in the cover sheet are listed in the app's log.

## Troubleshooting

**"…is missing required column(s): 'ID', 'Donor #'…"** (before v1.5.8 this showed up as `None of [Index(['eNum', 'pNum'])] are in the [columns]`)
- Check that the **Spons Xtra** tab has `ID` and `Donor #` headers.
- Check that the Extra Gift Report box points at the extra gift export, not another report.

**"…no tab named 'Spons Xtra'…"**
- The wrong file is selected, or the export tab was renamed and its first tab doesn't have the required columns.

**Donors not getting TY notes they should**
- Look in the log for "donor #(s) … aren't in the cover sheet". It means the extra gift report and the cover sheet come from different periods or different exports.
