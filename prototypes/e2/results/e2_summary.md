# E2 results

Instances: target X pid 30676, decoy Y pid 28816

## Correctness (first VBA handoff per backend; later runs compared to it)

### pywin32: 27/27 checks passed; all 5 runs identical: True
- PASS — Correct instance and file (work/x): got 'x', expected 'x'
- PASS — Correct workbook among several open (Inputs!A1 = 10, not decoy 77 or FixtureB 5555): got 10.0, expected 10
- PASS — Unsaved edit: Calc!B2 value: got 999.0, expected 999
- PASS — Unsaved edit: Calc!B3 formula: got '=B2*3', expected '=B2*3'
- PASS — Unsaved edit: Calc!B3 recalculated value: got 2997.0, expected 2997
- PASS — Workbook reported as unsaved: got False, expected False
- PASS — Multi-area selection: sheet: got 'Calc', expected 'Calc'
- PASS — Multi-area selection: areas: got ['$B$2:$B$5', '$D$7:$E$9'], expected ['$B$2:$B$5', '$D$7:$E$9']
- PASS — Cross-sheet formula: Calc!A1: got '=Inputs!A1*2', expected '=Inputs!A1*2'
- PASS — Cross-sheet value: Calc!A1: got 20.0, expected 20
- PASS — Reference into hidden sheet: Calc!C1: got 43.0, expected 43
- PASS — Error types preserved (Errors!A1:A7): got [{'error': '#DIV/0!'}, {'error': '#N/A'}, {'error': '#REF!'}, {'error': '#VALUE!'}, {'error': '#NAME?'}, {'error': '#NUM!'}, {'error': '#SPILL!'}], expected [{'error': '#DIV/0!'}, {'error': '#N/A'}, {'error': '#REF!'}, {'error': '#VALUE!'}, {'error': '#NAME?'}, {'error': '#NUM!'}, {'error': '#SPILL!'}]
- PASS — Text that looks like an error stays text (Errors!B1): got '#N/A', expected '#N/A'
- PASS — Dynamic array formula: Dyn!A1: got '=SEQUENCE(5)', expected '=SEQUENCE(5)'
- PASS — Spilled cell has no formula: Dyn!A2: got '', expected ''
- PASS — Spilled values: Dyn!A1:A5: got [1.0, 2.0, 3.0, 4.0, 5.0], expected [1, 2, 3, 4, 5]
- PASS — Spill reference: Dyn!E1 formula: got '=SUM(A1#)', expected '=SUM(A1#)'
- PASS — Spill reference: Dyn!E1 value: got 15.0, expected 15
- PASS — Spill range detected: got {'Dyn!A1_has_spill': True, 'Dyn!A1_spill': '$A$1:$A$5'}, expected {'Dyn!A1_has_spill': True, 'Dyn!A1_spill': '$A$1:$A$5'}
- PASS — Hidden sheet visibility (0): got 0, expected 0
- PASS — Very hidden sheet visibility (2): got 2, expected 2
- PASS — Very hidden sheet formula readable: got '=Inputs!A1', expected '=Inputs!A1'
- PASS — Protected sheet detected: got True, expected True
- PASS — Protected sheet value readable: got 110.0, expected 110
- PASS — Protected sheet, hidden formula: what COM returns (informational): got '<formula not readable>'
- PASS — Bulk 1k: count, first and last cell: 1000 cells, first ['=Inputs!$A$1*ROW()+COLUMN()', 11.0], last ['=Inputs!$A$1*ROW()+COLUMN()', 1010.0]
- PASS — Bulk 10k: count, first and last cell: 10000 cells, first ['=Inputs!$A$1*ROW()+COLUMN()', 11.0], last ['=Inputs!$A$1*ROW()+COLUMN()', 10010.0]
- API fallbacks needed: none

### xlwings: 25/27 checks passed; all 5 runs identical: True
- PASS — Correct instance and file (work/x): got 'x', expected 'x'
- PASS — Correct workbook among several open (Inputs!A1 = 10, not decoy 77 or FixtureB 5555): got 10.0, expected 10
- PASS — Unsaved edit: Calc!B2 value: got 999.0, expected 999
- PASS — Unsaved edit: Calc!B3 formula: got '=B2*3', expected '=B2*3'
- PASS — Unsaved edit: Calc!B3 recalculated value: got 2997.0, expected 2997
- PASS — Workbook reported as unsaved: got False, expected False
- PASS — Multi-area selection: sheet: got 'Calc', expected 'Calc'
- PASS — Multi-area selection: areas: got ['$B$2:$B$5', '$D$7:$E$9'], expected ['$B$2:$B$5', '$D$7:$E$9']
- PASS — Cross-sheet formula: Calc!A1: got '=Inputs!A1*2', expected '=Inputs!A1*2'
- PASS — Cross-sheet value: Calc!A1: got 20.0, expected 20
- PASS — Reference into hidden sheet: Calc!C1: got 43.0, expected 43
- FAIL — Error types preserved (Errors!A1:A7): got [{'error': '#DIV/0!'}, {'error': '#N/A'}, {'error': '#REF!'}, {'error': '#VALUE!'}, {'error': '#NAME?'}, {'error': '#NUM!'}, -2146826243], expected [{'error': '#DIV/0!'}, {'error': '#N/A'}, {'error': '#REF!'}, {'error': '#VALUE!'}, {'error': '#NAME?'}, {'error': '#NUM!'}, {'error': '#SPILL!'}]
- FAIL — Text that looks like an error stays text (Errors!B1): got {'error': '#N/A'}, expected '#N/A'
- PASS — Dynamic array formula: Dyn!A1: got '=SEQUENCE(5)', expected '=SEQUENCE(5)'
- PASS — Spilled cell has no formula: Dyn!A2: got '', expected ''
- PASS — Spilled values: Dyn!A1:A5: got [1.0, 2.0, 3.0, 4.0, 5.0], expected [1, 2, 3, 4, 5]
- PASS — Spill reference: Dyn!E1 formula: got '=SUM(A1#)', expected '=SUM(A1#)'
- PASS — Spill reference: Dyn!E1 value: got 15.0, expected 15
- PASS — Spill range detected: got {'Dyn!A1_has_spill': True, 'Dyn!A1_spill': '$A$1:$A$5'}, expected {'Dyn!A1_has_spill': True, 'Dyn!A1_spill': '$A$1:$A$5'}
- PASS — Hidden sheet visibility (0): got 0, expected 0
- PASS — Very hidden sheet visibility (2): got 2, expected 2
- PASS — Very hidden sheet formula readable: got '=Inputs!A1', expected '=Inputs!A1'
- PASS — Protected sheet detected: got True, expected True
- PASS — Protected sheet value readable: got 110.0, expected 110
- PASS — Protected sheet, hidden formula: what COM returns (informational): got '<formula not readable>'
- PASS — Bulk 1k: count, first and last cell: 1000 cells, first ['=Inputs!$A$1*ROW()+COLUMN()', 11.0], last ['=Inputs!$A$1*ROW()+COLUMN()', 1010.0]
- PASS — Bulk 10k: count, first and last cell: 10000 cells, first ['=Inputs!$A$1*ROW()+COLUMN()', 11.0], last ['=Inputs!$A$1*ROW()+COLUMN()', 10010.0]
- API fallbacks needed: ['Workbook.Saved', 'Range.Areas (multi-area selection)', 'Worksheet.Visible (xlwings .visible is bool; cannot tell hidden from very hidden)', 'Worksheet.ProtectContents', 'Range.HasSpill / SpillingToRange']

## Explicit handoff to the other instance (decoy, same file name)

- pywin32: read y/E2_FixtureA.xlsx, Inputs!A1 = 77.0 (expected y, 77)
- xlwings: read y/E2_FixtureA.xlsx, Inputs!A1 = 77.0 (expected y, 77)

## Naive GetActiveObject (comparison only)

- Attached to pid 11752 (target X is 30676); target workbook present: False; workbooks open there: 2

## Timings (ms), warm runs 2–5 median; run 1 shown separately

| Backend | Run 1 wall | Warm wall | Warm process | Warm attach | Warm read small | Warm bulk 1k | Warm bulk 10k |
|---|---|---|---|---|---|---|---|
| pywin32 | 854 | 893 | 783 | 629.1 | 88.6 | 4.1 | 19.6 |
| xlwings | 1177 | 1134 | 1034 | 824.4 | 115.3 | 8.0 | 27.2 |

## Failure handling

| Case | ok | Exit code | Wall ms | Error |
|---|---|---|---|---|
| pywin32: workbook not open | False | 2 | 720 | workbook 'Missing.xlsx' is not open in the target Excel instance |
| pywin32: window/pid mismatch | False | 2 | 161 | window 396202 belongs to process 30676, not 28816 |
| xlwings: workbook not open | False | 2 | 1275 | workbook 'Missing.xlsx' is not open in the target Excel instance |
| xlwings: window/pid mismatch | False | 2 | 1302 | window 396202 belongs to process 30676, not 28816 |
| pywin32: target workbook closed | False | 2 | 723 | workbook 'E2_FixtureA.xlsx' is not open in the target Excel instance |
| xlwings: target workbook closed | False | 2 | 1275 | workbook 'E2_FixtureA.xlsx' is not open in the target Excel instance |
| pywin32: Excel instance closed | False | 2 | 171 | window 527408 does not exist |
| xlwings: Excel instance closed | False | 2 | 1242 | window 527408 does not exist |

Decoy Excel process exited before the 'instance closed' test: True

## Read-only verification

- target: unchanged
- decoy: unchanged
