DAILY SAP IMPORT FOLDER

1. Download the new SAP sales report as .xlsx or .xlsm.
2. Put the file directly in this Daily_Input folder (not in Processed).
3. Double-click RunDashboard.bat.

The program will:
- detect the current-year master sales file in Input,
- append only new SAP rows,
- skip rows already imported,
- create a backup before changing the master,
- move the processed daily file into Daily_Input\Processed,
- regenerate Output\Result.xlsx.

Do not put target/objective files in this folder.
