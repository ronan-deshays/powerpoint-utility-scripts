Attribute VB_Name = "UserInterface"
' not supported by ppt vba

Sub addButtonsToCommandBarApp()
    'clearCommandBarApp
    Dim objBtn As CommandBarButton

    Set objBtn = Application.CommandBars("Worksheet Menu Bar").Controls.Add(msoControlButton)
    With objBtn
        .Caption = "WriteRGBToShape"
        .OnAction = "ColorInfos.WriteRGBToShape"
        .Style = msoButtonCaption
    End With
    
End Sub

Sub clearCommandBarApp()
    
    With Application.CommandBars("Worksheet Menu Bar")
        On Error Resume Next
        .Controls("WriteRGBToShape").Delete
        On Error GoTo 0
    End With
    
End Sub
