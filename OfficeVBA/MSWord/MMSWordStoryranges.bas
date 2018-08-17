' ==========================================================================
' Module      : MMSWordStoryranges
' Type        : Module
' Description : Procedures for StoryRanges Object functions
' --------------------------------------------------------------------------
' Procedures  : SentencesToExcelRow
' --------------------------------------------------------------------------
' Functions   : RegExReplace            Returns String
' --------------------------------------------------------------------------
' References  : VBScript.RegExp
'               Excel.Application
' ==========================================================================

' -----------------------------------
' Option statements
' -----------------------------------
Option Explicit
Option Private Module

Public Sub SentencesToExcelRow()
' ==========================================================================
' Description : Reads ActiveDocument and builds a list to a new Excel file.
'
' Parameters  : None
'
' Comments    : Intended for short lists, not recommended for long documents
'
' ==========================================================================

    Dim oExcelApp As Object
    Dim oExcelWB  As Object
    Dim oExcelWS  As Object
    Set oExcelApp = CreateObject("Excel.Application")
    Set oExcelWB = oExcelApp.Workbooks.Add
    Set oExcelWS = oExcelWB.Sheets.Add
    
    Const sRegExPattern As String = "[^a-zA-Z\d\s:]"
    Const sBlank        As String = ""
    Dim sSentence       As Variant
    Dim sNewContent     As String
        sNewContent = ""
    Dim lRowCounter     As Long
        lRowCounter = 2
    Dim lChar           As Long
    Dim sMainContent    As String
    '----------------------------------------------------
    'Removes invisible non-ascii characters
    '----------------------------------------------------
        sMainContent = RegExReplace(CStr(ActiveDocument.StoryRanges.Item(wdMainTextStory)), _
                                    sRegExPattern, sBlank)
    
        For lChar = 1 To Len(sMainContent)
    '------------------------------------------------------
    'Build new words/sentences usnig char(13) as terminator
    '------------------------------------------------------
            If Mid(sMainContent, lChar, 1) = vbCr Then
                If Len(sNewContent) < 2 Then GoTo next_item:
    '------------------------------------------------------
    'Copy word/sentence to new row
    '------------------------------------------------------
                oExcelWS.Cells(lRowCounter, 1).Value = sNewContent
                lRowCounter = lRowCounter + 1
                sNewContent = ""
            Else
                sNewContent = sNewContent & Mid(sMainContent, lChar, 1)
            End If
next_item:
        Next lChar
    
    '------------------------------------------------------
    'Show Excel
    '------------------------------------------------------
    oExcelApp.Visible = True
    
    '------------------------------------------------------
    'Clean up memory
    '------------------------------------------------------
    Set oExcelWS = Nothing
    Set oExcelWB = Nothing
    Set oExcelApp = Nothing
    
End Sub


Private Function RegExReplace(ByRef sWord As String, ByRef sPattern As String, ByRef sReplaceStr As String) As String
' ==========================================================================
' Description : Replace contents of a string passed in with regEx patterns.
'
' Parameters  : sWord           Original String value
'               sPattern        RegExp Pattern
'               sReplaceStr     New string against regex replace with pattern.
'
' Returns     : String
'
' ==========================================================================
    Dim oRegEx As Object
    Set oRegEx = CreateObject("VBScript.RegExp")
        
        With oRegEx
            .Pattern = sPattern
            .Global = True
        End With

            RegExReplace = oRegEx.Replace(sWord, sReplaceStr)
End Function


