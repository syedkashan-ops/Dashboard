OBJECTIVE FILES - AUTOMATIC DETECTION

Keep exactly TWO objective workbooks in the Input folder.
Their filenames can be anything; the dashboard identifies them from their structure.

1) OUTLET OBJECTIVE WORKBOOK
   Required product sheets: PMG, HSD, R95/HOBC, LUBE/Lubes
   Required columns in each sheet: Area, Name, Cost Center, Jul ... Jun
   Used ONLY for individual outlet / Cost Center objectives in Control sheets and outlet-level reports.

2) AREA OBJECTIVE WORKBOOK
   Required product sheets: PMG, HSD, HOBC/R95, Lubes/LUBE
   Required layout: Sales Area rows (CoRos/A01...A08) with Jul ... Jun monthly values
   Used for Main Dashboard, DDM Dashboard, Area Dashboard, Analytics and all aggregate objectives.

HOW TO REPLACE FILES
- Delete or overwrite the old outlet-objective workbook and paste the new one with ANY filename.
- Delete or overwrite the old area-objective workbook and paste the new one with ANY filename.
- Do not keep old copies in Input, otherwise the program will report that more than one matching workbook was found.
- Do not change the required sheet/column structure.
- Run RunDashboard.bat after replacement.

IMPORTANT
Outlet objectives are not added together to create area objectives. This prevents duplication and ensures that official area targets always come from the area-objective workbook.
