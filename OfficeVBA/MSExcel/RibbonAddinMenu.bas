Private Sub MainMenuSetup()
    'Alt dim as object
    Dim objWksMenuBar As CommandBar
    Set objWksMenuBar = Application.CommandBars("WorkSheet Menu Bar")

    Dim objMainMenu As CommandBarPopup
    Set objMainMenu = objWksMenuBar.Controls.Add(Type:=msoControlPopup, before:=8, Temporary:=True)
        objMainMenu.Caption = strMainMenuTitle

    Dim objMainMenu_Popup As CommandBarPopup
    Set objMainMenu_Popup = objMainMenu.Controls.Add(Type:=msoControlPopup, Temporary:=True)
        objMainMenu_Popup.Caption = "Popup Control"
    
    Dim objPopup_SubControlBtn As CommandBarButton
    Set objPopup_SubControlBtn = objMainMenu_Popup.Controls.Add(Type:=msoControlButton)
        objPopup_SubControlBtn.Caption = "Sub Control Button"
        objPopup_SubControlBtn.OnAction = ThisWorkbook.FullName &  "!mysub"
        objPopup_SubControlBtn.FaceId = 9341
End Sub