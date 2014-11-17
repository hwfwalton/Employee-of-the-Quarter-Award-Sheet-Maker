' This file has the functions used by EditLabs.HTA to change lab names, delete labs, and add new labs

Function openExcel
    pwd = CreateObject("WScript.shell").CurrentDirectory
    ePath = pwd & "\res\Archives.xlsx"
    Set eObj = CreateObject("Excel.Application")
    eObj.Workbooks.Open(ePath)
    
    set openExcel = eObj
End Function

Function closeExcel(eObj)
    eObj.ActiveWorkbook.Save
    eObj.ActiveWorkbook.Close
    eObj.quit
End Function

Function entryCount
    cnt = 3*(Year(Date)-2010)
    
    Select Case Month(Date) 
        ' Winter Quarter
        Case 0, 1, 2
            cnt = cnt + 1
            
        ' Spring/Summer Quarter
        Case 3, 4, 5, 6, 7
            cnt = cnt + 2
        
        ' Fall Quarter
        Case 8, 9, 10, 11
            cnt = cnt + 3
    End Select    
    
    entryCount = cnt
End Function

Function labExists(byVal lab)
    set eObj = openExcel
    nameCollision = False
    
    For Each sh In eObj.Sheets
        If sh.Name = lab Then
            nameCollision = True
        End If
    Next
    
    labExists = nameCollision
    
    closeExcel(eObj)
End Function

' Called after adding a lab or changing a lab's name. Alphabetizes the Excel sheets
Sub AlphabetizeSheets
    Set eObj = openExcel
    sheetCnt = eObj.Sheets.Count
    
    For i = 1 To sheetCnt
        For j = 1 To sheetCnt - 1
            sheet1 = eObj.Sheets(j).Name
            sheet2 = eObj.Sheets(j+1).Name
            If UCase(sheet1) > UCase(sheet2) Then
                eObj.Sheets(j).Move ,eObj.Sheets(j+1)
            End If
        Next
    Next
    
    closeExcel(eObj)
End Sub

' Backs up
Sub BackUpArchives
    ' Create a File System Object and use it to make a copy of the current Archives file
    Set fso = CreateObject("Scripting.FileSystemObject")
    pwd = CreateObject("WScript.shell").CurrentDirectory
    fso.CopyFile pwd & "\res\Archives.xlsx", pwd & "\res\ArchivesBACKUP.xlsx"
    
    ' Free the memory
    Set fso = Nothing
End Sub

' 
Sub ChangeLabName(byVal lab)
    ' Prompt the user to choose a new name
    Input = InputBox("Enter new name for the "+lab+" lab below.","Change Lab Name",lab)
    
    ' Hitting Cancel returns a zero length string so check if that is the case and take no action if so
    ' I suppose this could malfunction in the unusual use case that a user wanted to change a lab's name to 
    ' and empty string, but I assume this won't happen.
    If len(Input) <> 0 Then
        response = MsgBox("Are you sure you would like to change the name of the " + lab + " lab to " + Input + "?",52, "Change Lab Name?")
        BackUpArchives
        If response = 6 Then
            ' Open the Excel file.
            Set eObj = openExcel
            
            ' Find the sheet using the existing lab name and change it to the new name.
            eObj.Sheets(lab).Name = Input
            
            ' Save and close the Excel file.
            closeExcel(eObj)
            AlphabetizeSheets
        Else
            MsgBox "No changes made. The lab name remains " + lab + ".",64,"No Changes Made"
        End If
    End If
    
End Sub

Sub DeleteLab(byVal lab)
    ' Confirm the user's actions
    response = MsgBox ("Are you sure you want to delete the entry and records for " + lab + "? This cannot be undone.",52,"Delete Lab Records?")
    
    If response = 6 Then
        BackUpArchives
        ' Open the Excel file 
        Set eObj = openExcel
        
        ' TODO: Automatically backup deleted labs to a separate "deletedLabs.xlsx" file in \res\.
        
        ' Find the sheet using the lab name and delete it. Alerts must be disabled first as Excel attempts to do its
        ' own confirmation but the user is unable to respond as its visibility is False.
        eObj.DisplayAlerts = False
        eObj.Sheets(lab).Delete
        eObj.DisplayAlerts = True
        
        ' Save and close the Excel file.
        closeExcel(eObj)
    End If
End Sub

Sub AddLab(byVal labName, copyLab, existingLab)
    BackUpArchives
    Set eObj = openExcel
    
    ' Create a new sheet and move it to the end of the list.
    Set newSheet = eObj.Sheets.Add
    newSheet.name = labName
    newSheet.Move ,eObj.Sheets(eObj.Sheets.Count)
    
    If copyLab Then    
        ' Copy the entries from the existing lab.
        eObj.Sheets(existingLab).Activate
        eObj.ActiveSheet.UsedRange.Select
        rc = eObj.Selection.Rows.Count
        eObj.ActiveSheet.Range("A1:C"&rc).Copy
        
        ' Paste the entries into the new lab.
        newSheet.Range("A1").PasteSpecial

        
        MsgBox "Copied entries from " & existingLab & " into new lab named " & labName & ".",64,"Successfully Added New Lab"
    Else
        ' Check the number of rows to create based on the current date.
        rowCnt = entryCount
        
        ' Fill in the date column and then flatten the values.
        newSheet.Range("A1:A"&rowCnt).Formula = "=(2010+FLOOR.MATH((ROW()-1)/3))&(MOD(ROW()-1,3)+1)"
        newSheet.Range("A1:A"&rowCnt).Value = newSheet.Range("A1:A"&rowCnt).Value
        
        ' Fill in the CRC photo file column.
        newSheet.Range("B1:B"&rowCnt).Value = "itlm000.jpg"
        
        ' Fill in the CRC name column and then flatten the values.
        newSheet.Range("C1:C"&rowCnt).Formula = "=(""" & labName & """&ROW())"
        newSheet.Range("C1:C"&rowCnt).Value = newSheet.Range("C1:C"&rowCnt).Value
    
        MsgBox "Created new lab called " & labName & " with template CRC entries.",64,"Successfully Added New Lab"
    End If
    
    AlphabetizeSheets
    closeExcel(eObj)
End Sub

















