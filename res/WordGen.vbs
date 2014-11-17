' This file has the function used by PlaqueSheetGen.hta to generate the formatted Word doc for the CRC of the Quarter 
' Plaques as well as related utility functions.

' Utility function that formats a given date from 5 digit year&quarterIndex to written out
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

' Inches to points
Function itp(inches)
    itp = 72*inches
End Function

Sub WordGen(ByVal lab)
    ' Use the given lab name to determine the proper values for empTitle (employee title)
    Select Case lab
    Case "HWS"
        empTitle = "Hardware Support Employee"
    Case "Admin"
        empTitle = "Admin Support Employee"
    Case Else
        empTitle = "Computer Room Consultant"
    End Select

    ' *****************
    ' TODO: When Hutch and Wellman fully integrate we can merge the lab entries and this If should be removed.
    ' The separate excel sheets should be replaced with a single Hutchison-Wellman entry in the archives. 
    ' The else will remain as excel sheets can't handle slashes in their names (see its comments for more details)
    
    If lab = "Hutchison" or lab = "Wellman" Then
        labName = "Hutchison/Wellman"
    Else
        ' Change the dashes to slashes in multi-building labgroups on the sheet as that was how it was original formatted.
        ' Probably unnecessary to do this. If you want to just use dashes, remove this line and replace all instances
        ' of labName with lab.
        labName = Replace(lab,"-","/")
    End If


    ' Create a Shell object and use it to extract the present working directory
    pwd = CreateObject("WScript.shell").CurrentDirectory

    ' Use the pwd to concatenate the absolute paths to the Word template and appropriate Excel sheet 
    ePath = pwd & "\res\Archives.xlsx"
    tPath = pwd & "\res\Template.dotm"

    ' Creates Excel and Word COM objects
    Set eObj = CreateObject("Excel.Application")
    Set wObj = CreateObject("Word.Application")
    
    ' Set to false while it's working so user doesn't see the craziness.
    ' Set to True if making changes so you can see what's happening.
    wObj.Visible = False

    
    ' Creates a new document using the template specified by tPath and defines our document object
    set docObj = wObj.Documents.Add(tPath)

    ' Open the Excel sheet, but leave visibility set to false. Select the sheet of the with the same name as the given 
    ' lab. There should never be an issue finding a sheet with the given value as the options given to the user are 
    ' determined by the sheet names.
    eObj.Workbooks.Open(ePath)
    eObj.Sheets(lab).Activate
    
    ' Counts the number of used rows so that we can get the index of the most recent entry
    eObj.ActiveSheet.UsedRange.Select
    rc = eObj.Selection.Rows.Count

    ' Now that we have the document open and ready, we can begin formatting it. There are two passes over the 
    ' document, the first inserts all the text, the second inserts the pictures. They could probably be combined,
    ' but it was simpler to separate them as image insertion does strange things
    
    '=====================================================================================================
    ' TEXT PASS
    
    ' Make sure we're at the beginning of the document
    wObj.Selection.HomeKey 6, 0
 
    ' Inserts the appropriate employee title
    With wObj.Selection.Find
        .Text = "title"
        .MatchWholeWord = True
        .Replacement.Text = empTitle
        .Execute ,,,,,,,,,,2
    End With

    ' Inserts the appropriate Lab name
    With wObj.Selection.Find
        .Text = "crclab"
        .MatchWholeWord = True
        .Replacement.Text = labName
        .Execute ,,,,,,,,,,2
    End With
    
	' Loops through the most recent 9 CRCs and finds and replaces the relevant name and date tags in the template
    For i = 0 to 9
        ' Inserts CRC's name
        With wObj.Selection.Find
            .Text = "name" & CStr(i)
            .MatchWholeWord = True
            .Replacement.Text = eObj.Cells(rc-i,3).Value      ' Pulls the name from the Excel sheet
            .Execute ,,,,,,,,,,2
        End With
		
        ' Inserts award receipt date
        With wObj.Selection.Find
            .Text = "date" & CStr(i)
            .MatchWholeWord = True
            .Replacement.Text = dateText(eObj.Cells(rc-i,1).Value)
            .Execute ,,,,,,,,,,2
        End With
    Next
    
    wObj.Selection.HomeKey 6, 0

    '=====================================================================================================
    ' IMAGE PASS
 
    ' I'm sorry for the following code. I'm not proud of what I've created.
    firstImg = True
   
    ' Loops through 10 times (once for the newest winner and once for each of the past 9
    For i = 0 to 9
        Dim img                             ' Using a Dim helps keeps things organized here.
        ' Pull the image file name from the excel sheet to pull the correct image
        Set img = docObj.InlineShapes.AddPicture(pwd & "\pictures\" & eObj.Cells(rc-i,2).Value)
        
        If firstImg = True then
            img.Height = itp(3.14)
            img.Width = itp(2.42)
            ' img.ScaleHeight = 90
            ' img.ScaleWidth = 90
                                            ' For reasons beyond me, the selection methods are in the document object
            docObj.InlineShapes(1).Select   ' but the selection and cut methods are tied to the Word object.
            wObj.Selection.Cut              ' Cut the image after scaling it so we can move it.
            firstImg = False
        Else
            img.Height = itp(0.405)         ' Those numbers are the exact inch dimensions we want the small images
            img.Width = itp(0.3116)         ' to be. I got these numbers from the inDesign files.
            ' img.ScaleHeight = 10          ' I tried using scaling rather than absolute dimensions to create the mini images
            ' img.ScaleWidth = 10           ' as the scaled ones come out better, but the variable sizes of the source
            docObj.InlineShapes(1).Select   ' images made this impractical.
            wObj.Selection.Cut
        End If
        
        With wObj.Selection.Find            ' For each picture, find the correct pic# tag in the template and pastes
            .Text = "pic" & CStr(i)         ' the picture there
            .Forward = True
            .MatchWholeWord = True
            .Replacement.Text = ""
            .Execute ,,,,,,,,,,1            ' We use 1 this time because 1 leaves the text selected after finding it
        End With
        wObj.Selection.Paste
        wObj.Selection.HomeKey 6, 0
    Next
  
    ' Reset selection back to the beginning of the document again.
    wObj.Selection.HomeKey 6, 0

    ' Find Quarter, copy it, and then deselect it. The reason for this is to clear the clipboard from giant images
    ' because Word otherwise prompts if you would like to save the contents of the clipboard before exiting and
    ' I couldn't find a cleaner way to clear the clipboard.
	With wObj.Selection.Find
        .Text = "Quarter"
        .Replacement.Text = "Quarter"
        .Execute ,,,,,,,,,,1
    End With
    wObj.Selection.Copy
    wObj.Selection.HomeKey 6, 0

    ' Now that the formatting is done, set the word instance Visibility to true and bring it to the front
    wObj.Visible = True
    wObj.WindowState = 2
    wObj.WindowState = 1
    
    ' Make sure to quit the Excel object otherwise it just continues to invisibly run in the background.
    eObj.Quit

End Sub




