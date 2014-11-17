' This file has all the functions used by AddCRC.hta to add new entries to the Excel book and for managing the CRC photos

' Utility function that formats a given date from 5 digit year&quarterIndex to written out
' (4 digits of year Year and one for the quarter)YYYYQ -> "Quarter Year"
' e.g. 20113 -> Fall 2011
Function dateText(awardDate)
    ' Extract the year
    awardYear = mid(awardDate, 1, len(awardDate)-1)
    
    ' Pull the last character to find the quarter    
    Select Case mid(awardDate, len(awardDate))
    Case "1"
        quarter = "Winter"
    Case "2"
        quarter = "Spring"
    Case "3"
        quarter = "Fall"
    End Select
    
    ' Extract and concatenate the quarter and year strings together
    dateText = quarter & Space(1) & awardYear
End Function

' Prompts the user to choose an image file which is then added to the pictures folder
Function ChooseFile
    ' This function modified from one found at
    ' http://stackoverflow.com/questions/21559775/vbscript-to-open-a-dialog-to-select-a-filepath
    
    ' Create a shell object and use an ActiveX object to create the file open dialogue 
    Set wShell=CreateObject("WScript.Shell")
    Set oExec=wShell.Exec("mshta.exe ""about:<input type=file id=FILE><script>FILE.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""")
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Pull the path of the file and parse out the file name and extension
    fullPath = oExec.StdOut.ReadLine
    
    If Not len(fullPath) = "0" Then
    
        pwd = CreateObject("WScript.shell").CurrentDirectory
        fullFile = Mid(fullPath, InStrRev(fullPath, "\") + 1)
        fileName = Mid(fullFile, 1, InStr(fullFile, ".") - 1)
        extension = Mid(fullFile, InStrRev(fullFile, "."))
        
        ' Make sure we have a valid image format
        If extension = ".jpg" or extension = ".bmp" or extension = ".png" Then
            
            ' Check if the file already exists
            If (Not fso.FileExists(pwd & "\pictures\" & fullFile)) Then
                ' If the file does not already exist, just copy in the new one.
                fso.CopyFile fullPath, pwd & "\pictures\"
                ' Set the value ChooseFile which is returned to the HTA
                ChooseFile = fullFile
            Else
                ' Append the number of seconds elapsed since midnight to differentiate the file
                thisTime = Timer
                newFile = fileName & "-" & mid(thisTime, 1, Len(thisTime)-3) & extension
                fso.CopyFile fullPath, pwd & "\pictures\" & newFile
                ChooseFile = newFile
            End If
            
        Else
            If len(fullFile) > 0 Then
                ' Through an error if the file isn't a recognized image type
                MsgBox "|" & len(fullFile) & "| is not a recognized image file. Please use a jpg, bmp, or png.", 48
            End If
        End If
    
    End If
End Function

' Function to check through the excel archives and find if a given person is already present. Returns the file name
' of the person's portrait file if they do already exist.
Function FindPhoto(name)
    ' Create a File System Object and use it to extract the present working directory.
    pwd = CreateObject("WScript.shell").CurrentDirectory

    ' Use the pwd to concatenate the absolute path to the appropriate Excel sheet.
    ePath = pwd & "\res\Archives.xlsx"
    
    ' Create Excel COM object
    Set eObj = CreateObject("Excel.Application")
    
    ' Open the Excel sheet, but leave visibility set to false.  Select the sheet of the lab
    eObj.Workbooks.Open(ePath)
    
    ' Index we use to go through all the sheets
    FindPhoto = "NONE"
    photoDate = 0
    
    ' Find the number of worksheets in the file
    sheetCnt = eObj.Worksheets.Count
        
    ' Loop through all the sheets looking for the given name. If the name is found, pull the corresponding file name
    For i=1 To sheetCnt
        eObj.Sheets(i).Activate
        
        ' Find the number of rows in the current sheet
        eObj.ActiveSheet.UsedRange.Select
        rc = eObj.Selection.Rows.Count
        
        ' It probably would have been better to use one of the built in Excel find functions, but they proved surprisingly
        ' unwieldy. This was so much easier.
        For j=1 To rc
            ' If we find a name match and the awardDate is greater than any previous matches we've found (or 0 if we haven't found any other entries)
            ' then record the file name of the photo
            If eObj.Cells(j,3) = name And eObj.Cells(j,1) > photoDate Then
                photoDate = eObj.Cells(j,1)
                FindPhoto = eObj.Cells(j,2)
            End If
        Next
    Next
    
    ' Make sure to quit the Excel object, otherwise, it just continues to invisibly run in the background
    eObj.Quit
End Function

Sub BackUpArchives
    ' Create a File System Object and use it to make a copy of the current Archives file
    Set fso = CreateObject("Scripting.FileSystemObject")
    pwd = CreateObject("WScript.shell").CurrentDirectory
    fso.CopyFile pwd & "\res\Archives.xlsx", pwd & "\res\ArchivesBACKUP.xlsx"
    
    ' Free the memory
    Set fso = Nothing
End Sub

' Adds a new entry to the excel sheet with the given info
Sub AddCRC(name, photo, lab, awardDate)
    ' Use the given lab name to determine the proper value for empTitle
    Select Case lab
    Case "HWS"
        empTitle = "employee"
    Case "Admin"
        empTitle = "Support employee"
    Case Else
        empTitle = "CRC"
    End Select
    
    fullTitle = dateText(awardDate)+" "+lab+" "+empTitle+" of the quarter"

    ' Before opening Excel or doing anything else, confirm the info is correct
    response = MsgBox ("Would you like to add "+name+" as the "+fullTitle+"?",36,"Add CRC Entry?")

    ' Check user's response and either add the entry or exit
    If response = 6 Then    ' vbYes = 6
        ' Before adding the new entry, create a new backup
        BackUpArchives
        
        ' Create a File System Object and use it to extract the present working directory
        pwd = CreateObject("WScript.shell").CurrentDirectory

        ' Use the pwd to concatenate the absolute path to the appropriate Excel sheet.
        ePath = pwd & "\res\Archives.xlsx"
        
        ' Create Excel COM object
        Set eObj = CreateObject("Excel.Application")
        
        ' Open the Excel sheet, but leave visibility set to false.  Select the sheet of the lab
        eObj.Workbooks.Open(ePath)
        eObj.Sheets(lab).Activate

        ' Counts the number of used rows so that we can get the index of the most recent entry
        eObj.ActiveSheet.UsedRange.Select
        rc = eObj.Selection.Rows.Count
        
        ' The excel sheet is laid out where columns 1, 2, and 3 are year+quarter, photoFileName, and CRCname respectively
        ' If a repeated entry is found, then we only need to replace the photoFileName and CRCname cells. If we add a new entry, fill in all three 
        If CInt(awardDate) <= eObj.Cells(rc, 1).Value Then
            ' If the given awardDate is less than or equal to the given date, then there already exists an entry with that date, so we scan for it.
            ' I had it start at the top of the list and decrement as it seemed more likely that it's a recent entry that would be replaced
            For i = rc To 1 Step -1
                ' Check if the entry has the same awardDate value. If so, prompt the user whether they want to overwrite.
                If eObj.Cells(i, 1).Value = CInt(awardDate) Then
                    existingName = eObj.Cells(i, 3).Value
                    overwrite = MsgBox (existingName+" is already listed as the "+fullTitle+". Would you like to replace this entry?",52,"Overwrite Existing CRC Entry?")
                    If overwrite = 6 Then
                        eObj.Cells(i, 2).Value = photo
                        eObj.Cells(i, 3).Value = name
                        MsgBox name+" has been added as the "+fullTitle+"!",64,"Successfully added CRC"
                    Else
                        MsgBox "No changes made. "+existingName+" is recorded as the "+fullTitle+".",64,"No Changes Made"
                    End If
                    ' Terminate early if we find an existing entry. This might be better done with a Do...While loop but I hate defining my own indexes
                    Exit For
                End If
            Next
        Else
            eObj.Cells(rc+1, 1).Value = awardDate
            eObj.Cells(rc+1, 2).Value = photo
            eObj.Cells(rc+1, 3).Value = name
            MsgBox name+" has been added as the "+fullTitle+"!",64,"Successfully added CRC"
        End If
        
        ' Save changes, close the file, and quit Excel
        eObj.ActiveWorkbook.Save
        eObj.ActiveWorkbook.Close
        eObj.Quit
        

    End If
End Sub