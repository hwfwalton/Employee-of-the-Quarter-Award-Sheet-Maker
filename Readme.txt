CRC of the Quarter Plaque Sheet Generator

Authorship and Design
=======================
Designed and written by Henry Walton (itlm073) during Spring and Summer quarter
of 2014 with extensive help from the internet.
The program uses a VBScript to fill in a Word Template with information drawn 
from an Excel book. The GUI is an HTA that uses a combination of more VBScript
and some Javascript to provide tools to add new employees to the archives. The
Word template and final generated sheets are based on the formerly used Adobe
InDesign sheets which can still be found buried in the depths of CLMDoc.
Pictures of program:
http://imgur.com/a/iWj5E

Included Files
=======================
EXECUTABLES:
  AddCRC.hta          - Program used to add a new CRC of the Quarter to the
                        archives. Can be opened directly or from Sheetmaker.
  EditLabs.hta        - Program used to change the name of or delete existing an
                        existing lab or add a new lab.
  SheetMaker.hta      - The executable used to generate the plaque pages.

RESOURCES:
  /pictures/*         - The CRC photos used in the plaque sheets.
  /res/Archives.xlsx  - The Excel book that stores the information on CRCs.
  /res/ArchivesBACKUP - A copy of the archives excel file automatically made
                        before any changes are made to the archives. It is
                        replaced each time another change is made.
  /res/Template.dotm  - The Word template the plaque sheets modify.
  
  /res/AddCRC.vbs     - Contains scripts used for adding new CRC entries and
                        handling photos.
  /res/WordGen.vbs    - A single script that formats Template.docm with the CRC
                        info drawn from archive.xlsx.
  /res/EditLabs.vbs   - Contains scripts used by Editlabs.HTA for editing,
                        adding, and deleting labs.

REFERENCE:
  Readme.txt          - What you're currently looking at.
  ToDos.txt           - A list of additional features and bug fixes that haven't
                        yet been implemented.
  /assets/*           - Original assets used in development. Not used by the
                        program, but included here for reference and future
                        development.

Requirements
=======================
To use this program you will need:
  - A PC running Windows 7 or later.
  - Word and Excel 2007 or later.
  - A color printer, preferably with photo paper.
  - The BerkeleyUCDavis and FuturaUCDavis fonts. These should already be 
    installed on both HWS PCs, but if they are not or are you are working on a
    different computer, you can find them in /assets/fonts/.

Usage Walkthroughs
=======================
PLEASE NOTE
  If you copy this program to your computer rather than just running it from
  the server, DO NOT COPY THE ASSETS FOLDER. It is nearly three times the size
  of the rest of the program put together as it has all the original assets.

GENERATING AND PRINTING A PLAQUE SHEET
1. Open PlaqueSheetGen.hta by double clicking it.
2. Choose the lab to generate a sheet for from the dropdown menu.
  - Please note that though Hutchison and Wellman are in the same lab group, 
     they have different plaques as their histories differ (The CRCs from Spring
     2012 and earlier on the Wellman sheet are from the now closed MU lab).
3. Click Generate Sheet. After a 5-10 seconds (the button will remain depressed,
    do not click it again) the formatted document should appear.
4. Verify that the formatting is correct.
5. In Word, click File in the top right corner and then click Print.
6. Select your color printer from the dropdown menu under "Printer". If you're
    in the office (307 Surge IV), this will be the HP Color LaserJet in Admin.
7. Set the printer to use photo paper. For the HP Color LaserJet in the office,
    click Printer Properties and in the window that appears, select "Tray 3"
    from the dropdown menu under "Source is:". Click OK.
8. Print your beautifully formatted document!
9. If you're done printing, you can close Word. There is no need to save the
    document. If you're going to print another sheet, it's faster to leave Word
    open so it doesn't have to launch fresh for each sheet.


ADDING A NEW CRC OF THE QUARTER ENTRY
1. Open AddCRC.hta by double clicking it or clicking "Add New Employee" from
    the plaque sheet generator.
2. Enter the employee's first and last names into the respective fields.
3. Check for an existing photo of the employee by clicking "Check for Existing"
    which scans the CRC of the Quarter archives for their name. If the employee
    has not preveiously received a CRC of the Quarter award, no photo will be
    found. In this case click "New Photo" and navigate to the file to use.
  - If there is no an existing photo of the CRC you will likely need to find
    the new award recipient's CRC photo in CLMDOC. You should find it somewhere
    in Z:\CLMDOC\Pictures\Students\. If you can't find it there, talk to Joe.
  - If a file with the same name is already in the pictures folder, a random
    number will be appended to the file name.
4. Select the lab to which the employee belongs from the dropdown list.
5. Either leave the checkbox for "Use current year and quarter" selected or 
    uncheck it and manually set the date.
6. Click "Add Employee" to add the employee to the archives!
  - If there is already an entry for the selected date, you will be prompted if
    you woud like to replace the old entry with your new one.
  - When adding an entry to "Hutch & Wellman" you will be initially prompted if
    you would like to add your entry to Hutch and then immediately asked if you
    would like to add your entry to Wellman. You should choose yes for both
    unless you've been specifically told otherwise.


Known Issues and Bugs
=======================
- If the script crashes partway through generating a sheet, Excel may still be
  silently running in the background. Before doing anything else, OPEN TASK
  MANAGER AND CLOSE ANY RUNNING PROCESSES OF EXCEL.EXE. Failing to do so could
  potentially wreck the Excel book or put Excel in read-only mode, causing
  further issues.

- Clicking the "Generate Sheet" button in PlaqueSheetGen multiple times can
  launch multiple instances of Word and generate the same sheet repeatedly.

