STREAMLIT PERFORMANCE FIX

Replace the original project with this folder.

Changed files:
1. dashboard.py
   - Removed input() which can hang in cloud environments.
   - Added flush=True progress messages.

2. excel_writer.py
   - Removed an expensive full-workbook formatting pass.
   - Removed four expensive full-workbook post-processing/validation scans from save().
   - Dashboard generation, sheets, formulas, and normal formatting creation remain unchanged.

IMPORTANT:
Keep your currently working streamlit_app.py unchanged.
Do not use the previous Popen/live-log modification.

After uploading the files to GitHub, test the app again.
