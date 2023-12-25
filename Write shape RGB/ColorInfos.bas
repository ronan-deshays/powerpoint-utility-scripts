Attribute VB_Name = "ColorInfos"
Function Red(rgb As Long) As Integer
    Red = rgb Mod 256
End Function

Function Green(rgb As Long) As Integer
    Green = rgb \ 256 Mod 256
End Function

Function Blue(rgb As Long) As Integer
    Blue = rgb \ 65536 Mod 256
End Function

Sub WriteRGBToShape()
    ' Check if a shape is selected
    If ActiveWindow.Selection.ShapeRange.Count < 1 Then
        MsgBox "Please select a shape first."
        Exit Sub
    End If
    
    ' Get the selected shape
    Dim oShape As Shape
    Set oShape = ActiveWindow.Selection.ShapeRange(1)
    
    ' Get the fill color of the shape
    Dim oColor As Long
    oColor = oShape.Fill.ForeColor.rgb
    
    ' Convert the color to an RGB string
    Dim colorCode As String
    colorCode = CStr(Red(oColor)) & "," & CStr(Green(oColor)) & "," & CStr(Blue(oColor))
    
    ' Write the RGB string into the shape
    oShape.TextFrame.TextRange.Text = colorCode
End Sub
