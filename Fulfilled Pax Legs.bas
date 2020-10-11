Attribute VB_Name = "Weekly_HTG_Fullfilled_PL_LW"
Sub Weekly_HTGPaxLegsLWMTD()

Application.DisplayAlerts = False

Dim Source As Workbook
Dim Template As Workbook
Dim olapp As Object
Dim olmail As Object
Dim olsubject As String

    Dim iWeekday As Integer, LastSundayDate As Date

    iWeekday = Weekday(Now(), vbSunday)

    LastSundayDate = Format(Now - (iWeekday - 1), "dd-mmm-yy")

' NEWPATH is THIS year's path
NEWPATH = "***\Weekly Reports " & Format(DateAdd("d", 0, Now), "yyyy")
' OLDPATH is the previous year's path,will be useful when in January you want the Previous month(December) report to be saved in previous year's path, and not in this year's path
OLDPATH = "***\Weekly Reports " & Format(DateAdd("d", 0, Now), "yyyy")
' If the folder doesn't exist, then create it, useful again in February when running the January Report, it will create then new year's Folder and save it in there
If Dir(NEWPATH, vbDirectory) = "" _
Then MkDir NEWPATH

' SAVE report in the Right PATH
If Format(DateAdd("m", 0, Now), "yyyy") = Format(DateAdd("w", -1, Now), "yyyy") _
Then RPATH = NEWPATH & "\HTG Fullfilled Pax Legs vs LY Report WE " & Format(Now - (iWeekday - 1), "DD.MM.YY") & ".xlsx"
If Not Format(DateAdd("m", 0, Now), "yyyy") = Format(DateAdd("w", -1, Now), "yyyy") _
Then RPATH = OLDPATH & "\HTG Fullfilled Pax Legs vs LY Report WE " & Format(Now - (iWeekday - 1), "DD.MM.YY") & ".xlsx"


 '-----------------------------------------------------------Outlook Macro Start-------------------------------------------------------------------------------------------
    Dim gappOutlook As Object
Dim ns As Namespace
Dim inbox As MAPIFolder
Dim subfolder As MAPIFolder
Dim item As Object
Dim atmt As Attachment
Dim filename As String
Dim I As Integer
Dim varResponse As VbMsgBoxResult
Set ns = GetNamespace("MAPI")
Set gappOutlook = CreateObject("Outlook.Application")
Set inbox = gappOutlook.Session.Folders("Commercial Commercial")
Set subfolder = inbox.Folders("Inbox")
I = 0
' Check subfolder for messages and exit of none found
If subfolder.Items.Count = 0 Then
    MsgBox "There are no messages in the Subm from Arch folder.", vbInformation, _
           "Nothing Found"
    Exit Sub
End If
' Check each message for attachments

For Each item In subfolder.Items
If TypeName(item) = "MailItem" Then 'Change the below string to      WorksheetFunction.WeekNum((item.ReceivedTime)) = WorksheetFunction.WeekNum((Now))-1 _     if for any reason you are running the macro the week after(-1 means one week after, -2 would mean two weeks after, anyway it would be very unlikely,unless you need an old report)
If WorksheetFunction.WeekNum((item.ReceivedTime)) = WorksheetFunction.WeekNum((Now)) _
Then
    For Each atmt In item.Attachments
' Check filename of each attachment and save if it has "xlsx" extension
        If atmt.filename = "HTG Fullfilled Pax Legs vs LY Report.xlsx" _
        Then
        With item
.UnRead = False
End With
        ' This path must exist! Change folder name as necessary.
            filename = "***\HTG Fullfilled Pax Legs vs LY Report\" & _
                atmt.filename
            atmt.SaveAsFile filename
            I = I + 1
        End If
    Next atmt
    End If
    End If
Next item

' Clear memory
OlAtmt1stMonth_exit:
Set atmt = Nothing
Set item = Nothing
Set ns = Nothing

'-----------------------------------------------------------Outlook Macro End-------------------------------------------------------------------------------------------

Application.ScreenUpdating = False
Set Source = Workbooks.Open(filename:="***\HTG Fullfilled Pax Legs vs LY Report\HTG Fullfilled Pax Legs vs LY Report.xlsx")
Set Template = Workbooks.Open(filename:="***\HTG Fullfilled Pax Legs vs LY Report\HTG Fullfilled Pax Legs vs LY Template.xlsx")
Source.Sheets("Page1_1").Range("A1:G10").UnMerge
Source.Sheets("Page1_1").Range("A1:G10").Copy
Template.Sheets("Back Data").Visible = True
Template.Sheets("Back Data").Range("A1:G10").PasteSpecial xlPasteValues
Template.Sheets("Back Data").Visible = False
Source.Close
Kill ("***\HTG Fullfilled Pax Legs vs LY Report\HTG Fullfilled Pax Legs vs LY Report.xlsx")
Template.Save
Template.SaveAs filename:=(RPATH)


    Set olapp = CreateObject("Outlook.Application")
    Set olmail = olapp.createitem(olmailitem)

    olsubject = "HTG Fullfilled Pax Legs vs LY Weekly Report for w/e " & Format(Now - (iWeekday - 1), "DD.MM.YY")

    With olmail
        .display
    End With

    With olmail
        .To = "***@***.com"
        .CC = "***@***.com"
        .BCC = ""
        .Subject = olsubject
        .HTMLBody = "Good Morning, Please find attached HTG Fullfilled Pax Legs vs LY Report WE " & Format(Now - (iWeekday - 1), "DD.MM.YY") & .HTMLBody
        .Attachments.Add (ActiveWorkbook.FullName)
        '.Attachments.Add ("C:\test.txt") ' add other file
'        .Send   'or use .Display
        .display
    End With

    Set olmail = Nothing
    Set olapp = Nothing

Application.ScreenUpdating = True
ActiveWorkbook.Close savechanges:=True
Application.DisplayAlerts = True
End Sub
