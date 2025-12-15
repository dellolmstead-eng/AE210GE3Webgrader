# AE210 GE3 Web Grader (v2.0)

Browser grader for AE210 GE3 Jet11 submissions (iteration 1 and cutout). Aligns with the MATLAB `GE3_autograde_Olmstead_Fall_2025_v2.m` logic.

## Usage

1. Open `docs/index.html` (or serve the `docs/` folder).
2. Drop a GE3 Jet11 workbook that matches the GE3 template.
3. Review the computed score and any deduction notes; processing stays client-side.

## Parity testing

Open `docs/test_runner.html` and compare results to MATLAB logs for the same workbook names.

## Notes

- GE3 keeps its own README/docs set; shared workbook templates (baseline/v3/YF-22) are symlinked from `CommonAssets/`.
- If GE3 should consume the shared docs, let me know and I will link them.
