Public Sub CompleteAllFlagsItems()

Dim olItems As Object
Set olItems = Outlook.Application.Session.GetDefaultFolder(olFolderInbox).Items
    
    For Each i In olItems
        On Error Resume Next
        i.FlagStatus = olFlagComplete
    Next i

End Sub