Const sWarning As String = "Table is required to send email. " & vbNewLine & _
                      
Private Const msTimesNewRoman       As String = "Times New Roman"
Private Const msHTMLstyle           As String = "table, th, td {border: 1px solid gray; border-collapse:collapse;}"
Private Const msAttachPath          As String = "\\PATH\TO\ATTACHMENTS\"
Private Const msNumberFormat        As String = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
Private moOutlook                   As Object
Private moMailItem                  As Object
Private mcTotalOut                  As Currency
Private bCriteriaWarning            As Boolean
Private miNumberOfAreas             As Integer
Private miArea                      As Integer
Private miNumberofRows              As Integer
Private miRow                       As Integer

Public Sub GenerateEmail(ByVal eRLevel As enuLetterLevel)

If ActiveSheet.ListObjects.Count < 1 Then
    MsgBox sWarning
Else
   On Error GoTo err_handler:
    CreateMessage eRLevel
End If
err_handler:
Logger "Generating Level " & eRLevel & " Email"
End Sub



Public Sub CreateMessage(ByVal eRLevel As enuLetterLevel, Optional bAttachment As Boolean = False)

Dim sBody           As String: sBody = ""
Dim sSubject        As String: sSubject = ""
Dim sData           As String: sData = ""
Dim cAttachments    As New Collection
Dim cCustomerCol    As New Collection
Dim eLineCount      As enuEmailTableCount
    mcTotalOut = 0
    bCriteriaWarning = False
If Dir(msAttachPath) <> "" Then bAttachment = True
    If Selection.Count > 1 Then
    Selection.SpecialCells(xlCellTypeVisible).Select
    End If

miNumberOfAreas = Selection.Areas.Count

Dim iArea As Integer
Dim iRow As Integer
Dim iRowSize As Integer
Dim iRowStart As Integer
Dim iRowEnd As Integer

'TODO: Add Selected items to collections item=row number value = customer number'''''''

For iArea = 1 To miNumberOfAreas
      iRowSize = Selection.Areas(iArea).Rows.Count
     iRowStart = Selection.Areas(iArea).Row
       iRowEnd = iRowStart + iRowSize - 1
For iRow = iRowStart To iRowEnd
    mcTotalOut = mcTotalOut + AddToTotalOutstanding(iRow)
    
    sData = sData & HTMLTableRowWrapper(HTMLRowDataBuilder(iRow))
    If bAttachment Then cAttachments.Add AttachmentPaths(iRow, sAttachments)
    Next iRow
Next iArea

eLineCount = IIf(miNumberOfAreas = 1 And iRowSize = 1, enuEmailTableCount.SINGLE_LINE, enuEmailTableCount.MULTI_LINE)

sSubject = ReminderSubect(eRLevel, eLineCount)
sBody = HTMLHeadWrapper(HTMLStyleWrapper(msHTMLstyle))
sBody = sBody & HTMLTableWrapper(sData)
sBody = IIf(miNumberOfAreas = 1 And iRowSize = 1, "", sBody)
sBody = HTMLReminderLetterWrapper(eRLevel, sBody)
sBody = HTMLfontfaceWrapper(SETTINGS.Cells(2, 4), sBody)

SendEmail sBody, sSubject, cAttachments
Set cAttachments = Nothing

End Sub
Private Function AttachmentPaths(ByVal iRow As Integer, _
                        Optional ByVal sAttachment As String = "", _
                        Optional ByVal sFileType As String = ".pdf") As String

Dim iAttachmentCol As Integer
iAttachmentCol = ActiveSheet.ListObjects(FIRST_TABLE_OBJECT).ListColumns("ATTACHMENT_COLUMN").Index

Dim sPath As String

sPath = msAttachPath & ActiveSheet.Cells(iRow, iAttachmentCol) & sFileType
If Dir(sPath) <> "" Then
        
        AttachmentPaths = sPath

    Else
    AttachmentPaths = ""
End If

End Function
Private Function AddToTotalOutstanding(ByVal lRow As Long) As Currency
With ActiveSheet

AddToTotalOutstanding = _
.Cells(lRow, .ListObjects(FIRST_TABLE_OBJECT).ListColumns("Column_that_adds_something").Index).Value

End With
End Function
Private Function HTMLRowDataBuilder(ByVal lRow As Long) As String
Dim sRowData As String
Dim cData As New Collection
    
If ActiveSheet.ListObjects.Count > 0 Then

With ActiveSheet
     cData.Add .Cells(lRow, .ListObjects(FIRST_TABLE_OBJECT).ListColumns("COLUMN_NAME1").Index)
     cData.Add .Cells(lRow, .ListObjects(FIRST_TABLE_OBJECT).ListColumns("COLUMN_NAME2").Index)
     cData.Add .Cells(lRow, .ListObjects(FIRST_TABLE_OBJECT).ListColumns("COLUMN_NAME3").Index)
     cData.Add .Cells(lRow, .ListObjects(FIRST_TABLE_OBJECT).ListColumns("COLUMN_NAME4").Index)
     cData.Add .Cells(lRow, .ListObjects(FIRST_TABLE_OBJECT).ListColumns("COLUMN_NAME5").Index)
     cData.Add .Cells(lRow, .ListObjects(FIRST_TABLE_OBJECT).ListColumns("COLUMN_NAME6").Index)
     cData.Add .Cells(lRow, .ListObjects(FIRST_TABLE_OBJECT).ListColumns("COLUMN_NAME7").Index)
     cData.Add .Cells(lRow, .ListObjects(FIRST_TABLE_OBJECT).ListColumns("COLUMN_NAME8").Index)
     cData.Add Format( _
               .Cells(lRow, .ListObjects(FIRST_TABLE_OBJECT).ListColumns("COLUMN_NAME9").Index), _
               "#,##0.00")

On Error GoTo eval_handler:
If bCriteriaWarning = False And .Cells(lRow, .ListObjects(FIRST_TABLE_OBJECT).ListColumns("COLUMN_NAME").Index) <> "SOME_CRITERIA" Then
    MsgBox "Warning:: You are sending 'SOME_CRITERIA'. Please be aware."
    bCriteriaWarning = True
End If
eval_handler:
End With

For Each Item In cData
sRowData = sRowData & HTMLTableDataWrapper(Item)
Next Item
'DoEvents
HTMLRowDataBuilder = sRowData
    Set cData = Nothing
Else
    HTMLRowDataBuilder = ""
    Set cData = Nothing
End If

    
End Function

Public Sub SendEmail(ByVal sHTMLBody As String, ByVal sSubject As String, Optional ByVal cAttachments As Collection = Nothing)

Set moOutlook = CreateObject("Outlook.Application")
Set moMailItem = moOutlook.CreateItem(olMailItem)

 With moMailItem
 
  .Subject = sSubject
  .htmlbody = sHTMLBody & OutlookSignature
 
 If Not cAttachments Is Nothing Then
 For Each filePath In cAttachments
 If filePath <> "" Then
  .attachments.Add filePath
'  DoEvents
  End If
 Next filePath
 End If
 
 If SETTINGS.Cells(5, 4) Then
 .SentOnBehalfOfName = CStr(SETTINGS.Cells(4, 4))
 End If
'.SendUsingAccount = moAccount

'.Save
   .display
 End With

Set moOutlook = Nothing
Set moMailItem = Nothing

End Sub



Private Function HTMLTableDataWrapper(ByVal sData As String) As String
Const sTdOpen       As String = "<td align=""center"">"
Const sTdClose      As String = "</td>"

HTMLTableDataWrapper = sTdOpen & sData & sTdClose

End Function

Private Function HTMLTableRowWrapper(ByVal sRow As String) As String

Const sTrOpen       As String = "<tr>"
Const sTrClose      As String = "</tr>"

HTMLTableRowWrapper = sTrOpen & sRow & sTrClose

End Function

Private Function HTMLHeadWrapper(ByVal sHTMLBlock As String) As String
Const sHeadOpen     As String = "<head>"
Const sHeadClose    As String = "</head>"

HTMLHeadWrapper = sHeadOpen & sHTMLBlock & sHeadClose
End Function

Private Function HTMLStyleWrapper(ByVal sUserStyle As String)
Const sStyleOpen    As String = "<Style>"
Const sStyleClose   As String = "</style>"
HTMLStyleWrapper = sStyleOpen & sUserStyle & sStyleClose
End Function

Private Function HTMLTableWrapper(ByVal sTablecells As String) As String

Const sHEXCOLOR As String = "CCFFFF"
Dim sHEXCOLORUSER As String
    sHEXCOLORUSER = SETTINGS.Cells(3, 4)
Dim msHTMLtableStyleOpen As String
    msHTMLtableStyleOpen = _
        "<table style=""table-layout: fixed;""><tr>" & _
        "<th bgcolor=" & sHEXCOLORUSER & ">Column Name</th>" & _
        "<th bgcolor=" & sHEXCOLORUSER & ">Column Name</th>" & _
        "<th bgcolor=" & sHEXCOLORUSER & ">Column Name</th>" & _
        "<th bgcolor=" & sHEXCOLORUSER & ">Column Name</th>" & _
        "<th bgcolor=" & sHEXCOLORUSER & ">Column Name</th>" & _
        "<th bgcolor=" & sHEXCOLORUSER & ">Column Name</th>" & _
        "<th bgcolor=" & sHEXCOLORUSER & ">Column Name</th>" & _
        "<th bgcolor=" & sHEXCOLORUSER & ">Column Name</th>" & _
        "<th bgcolor=" & sHEXCOLORUSER & ">Column Name</th></tr>"
        
        
Const msHTMLTableTagClose As String = "</table>"

HTMLTableWrapper = msHTMLtableStyleOpen & sTablecells & msHTMLTableTagClose

End Function

Private Function HTMLBodyWrapper(ByVal sContent As String) As String

Const msHTMLBodyTagOpen         As String = "<body>"
Const msHTMLBodyTagClose        As String = "</body>"

HTMLBodyWrapper = msHTMLBodyTagOpen & sContent & msHTMLBodyTagClose
End Function

Private Function ReminderSubect(ByVal eRLevel As enuLetterLevel, ByVal eLineCount As enuEmailTableCount) As String
Const smdStamp As String = " - COMPANY NAME"
Const smdPastDue As String = "EMAIL BLURB"
Dim sLineDetails As String

If eLineCount = SINGLE_LINE Then

With ActiveSheet
    sLineDetails = _
    " DETAIL1" & .Cells(Selection.Row, .ListObjects(FIRST_TABLE_OBJECT).ListColumns("COLUMN_NAME").Index) & _
    " DETAIL2" & .Cells(Selection.Row, .ListObjects(FIRST_TABLE_OBJECT).ListColumns("COLUMN_NAME").Index) & _
    " DETAIL3" & .Cells(Selection.Row, .ListObjects(FIRST_TABLE_OBJECT).ListColumns("COLUMN_NAME").Index) & _
    smdStamp
End With

Else

    sLineDetails = IIf(eRLevel = FriendlyReminder, "TOPIC BLURB", smdPastDue) & smdStamp

End If
    
Select Case eRLevel
Case enuLetterLevel.Other
ReminderSubect = "TOPIC BLURB 1 - " & sLineDetails
Case enuLetterLevel.FirstLetter
ReminderSubect = "TOPIC BLURB 2 - " & sLineDetails
Case enuLetterLevel.SecondLetter
ReminderSubect = "TOPIC BLURB 3 - " & sLineDetails
Case enuLetterLevel.ThirdLetter
ReminderSubect = "TOPIC BLURB 4 - " & sLineDetails
Case enuLetterLevel.FinalLetter
ReminderSubect = "TOPIC BLURB 5 - " & sLineDetails

Case enuLetterLevel.FinalReminder
ReminderSubect = sIntroFirst & sTable & sGenericBodyMessageC

Case Else
ReminderSubect = "SOMETHING WENT WRONG in ReminderSubect"
End Select

End Function
Private Function HTMLReminderLetterWrapper(ByVal eRLevel As enuLetterLevel, ByVal sTable As String) As String

Const sBreak                    As String = "<br/>"
Const sIntroFriendly            As String = _
    sBreak & "Hello," & _
    sBreak & sBreak & "Nunc viverra imperdiet enim. Fusce est. Vivamus a tellus. "

    
Const sIntroFirst               As String = _
    sBreak & "Dear person, " _
    & sBreak & sBreak & "Nunc viverra imperdiet enim. Fusce est. Vivamus a tellus."

Const sTotalOut As String = " The total amount of something "

Dim sTotalDetails As String
    sTotalDetails = sTotalOut & _
    IIf(sTable <> "", "", "for criteria " & _
    ActiveSheet.Cells(Selection.Row, ActiveSheet.ListObjects(FIRST_TABLE_OBJECT).ListColumns("COLUMN_NAME").Index)) _
    & " is " & Format(mcTotalOut, "$#,##0.00") & "." _
    & sBreak & sBreak

Dim sFriendlyLineDetails As String
    sFriendlyLineDetails = IIf(sTable = "" And eRLevel = FriendlyReminder, _
    "Document# " & ActiveSheet.Cells(Selection.Row, ActiveSheet.ListObjects(FIRST_TABLE_OBJECT).ListColumns("COLUMN_NAME").Index) & _
    " viverra imperdiet enim. Fusce est. Vivamus a tellus " & Format(mcTotalOut, "$#,##0,00") & " Nunc viverra " & _
    ActiveSheet.Cells(Selection.Row, ActiveSheet.ListObjects(FIRST_TABLE_OBJECT).ListColumns("COLUMN_NAME").Index) & ".", _
    "consectetuer adipiscing. " & sTotalDetails)

Const sGenericBodyMessage As String = _
sBreak & "Lorem ipsum dolor sit amet, consectetuer adipiscing elit."


Const sGenericBodyMessageB As String = _
sBreak & "Lorem ipsum dolor sit amet, consectetuer adipiscing elit." 


Const sGenericBodyMessageC As String = _
sBreak & sBreak & "Lorem ipsum dolor sit amet, consectetuer adipiscing elit." 

Const sGenericBodyMessageD As String = _
sBreak & sBreak & "Lorem ipsum dolor sit amet, consectetuer adipiscing elit." 

Select Case eRLevel

Case enuLetterLevel.FirstLetter
HTMLReminderLetterWrapper = sIntroFirst & sTotalDetails & sTable & sGenericBodyMessage

Case enuLetterLevel.SecondLetter
HTMLReminderLetterWrapper = sIntroFirst & sTotalDetails & sTable & sGenericBodyMessage

Case enuLetterLevel.ThirdLetter
HTMLReminderLetterWrapper = sIntroFirst & sTotalDetails & sTable & sGenericBodyMessageB

Case enuLetterLevel.FinalLetter
HTMLReminderLetterWrapper = sIntroFirst & sTotalDetails & sTable & sGenericBodyMessageC

Case enuLetterLevel.Other
HTMLReminderLetterWrapper = sIntroFriendly & sFriendlyLineDetails & sTable & sGenericBodyMessageD

Case Else
HTMLReminderLetterWrapper = "SOMETHING WENT WRONG in HTMLReminderLetterWrapper"
End Select


End Function



Private Function HTMLfontfaceWrapper(ByVal sFontStyle As String, _
                                     ByVal sHTMLBlock As String) As String
Const sQuote As String = """"
Dim sHTML As String
sHTML = _
"<font face= " & sQuote & sFontStyle & sQuote & ">" & sHTMLBlock & "</font>"

HTMLfontfaceWrapper = sHTML

End Function

Private Function OutlookSignature() As String

Dim sSignature As String
sSignature = Environ("appdata") & "\Microsoft\Signatures\"

If Dir(sSignature, vbDirectory) <> vbNullString Then
            sSignature = sSignature & Dir$(sSignature & "*.htm")
            
            sSignature = CreateObject("Scripting.FileSystemObject").GetFile(sSignature).OpenAsTextStream(1, -2).ReadAll
        Else
            sSignature = ""
    End If
OutlookSignature = sSignature
End Function
