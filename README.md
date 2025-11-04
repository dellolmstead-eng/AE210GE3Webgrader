# AE210GE3Webgrader
Autograder for GE3 that is web hosted so cadets can submit an excel file, and it will read it and give feedback on whether it meets the requirements or not.
Goal is to have a website that cadets upload a JET11 .xlsm file to and it grades it and gives feedback that matches the autograder results. The truth source for requirements is the design project RFP, and the aircraft created must meet the threshold requirements and be a viable aircraft.

# Graduate to V1.0 on 2025/11/4 15:14
Bugs on cost (due to shift in cost-year from 2009 to 2022) and gear (The cells had shifted down one) cleared, appears to be giving satisfactory results. Calling this complete



## Parity testing

Use the GitHub Pages test runner to compare browser output with MATLAB baselines:

1. Visit `https://dellolmstead.github.io/AE210GE3Webgrader/test_runner.html`.
2. Upload one of the sample spreadsheets (the file name must match an entry in `docs/testdata/matlab_expected.json`).
3. The page will run the web grader locally and highlight any line-by-line differences from the MATLAB log captured in `r1/textout_2025-11-01_15-58-24.txt`.

