Public Sub AddTextStyle

Dim objTextStyle As AcadTextStyle

Set objTextStyle = ThisDrawing.TextStyles.Add("Bold Greek Symbols") 
objTextStyle.SetFont "Symbol", True, False, 0, 0 
End Sub




Public Sub GetTextSettings() 
Dim objTextStyle As AcadTextStyle 
Dim strTextStyleName As String 
Dim strTextStyles As String 
Dim strTypeFace As String 
Dim blnBold As Boolean 
Dim blnItalic As Boolean 
Dim lngCharacterSet As Long 
Dim lngPitchandFamily As Long 
Dim strText As String

' Get the name of each text style in the drawing 
For Each objTextStyle In ThisDrawing.TextStyles 

strTextStyles = strTextStyles & vbCr & objTextStyle.Name

Next


' Ask the user to select the Text Style to look at 
strTextStyleName = InputBox("Please enter the name of the TextStyle " & _ "whose setting you would like to see" & vbCr & _ strTextStyles,"TextStyles", ThisDrawing.ActiveTextStyle.Name) 
' Exit the program if the user input was cancelled or empty 
If strTextStyleName = "" Then Exit Sub

On Error Resume Next

Set objTextStyle = ThisDrawing.TextStyles(strTextStyleName) ' Check for existence the text style If objTextStyle Is Nothing Then

MsgBox "This text style does not exist" Exit Sub End If

' Get the Font properties objTextStyle.GetFont strTypeFace, blnBold, blnItalic, lngCharacterSet, _

lngPitchandFamily ' Check for Type face If strTypeFace = "" Then ' No True type

MsgBox "Text Style: " & objTextStyle.Name & vbCr & _ "Using file font: " & objTextStyle.fontFile, _ vbInformation, "Text Style: " & objTextStyle.Name

Else

' True Type font info strText = "The text style: " & strTextStyleName & " has " & vbCrLf & _

"a " & strTypeFace & " type face" 

If blnBold Then 
strText = strText & vbCrLf & " and is bold" 
If blnItalic Then 
strText = strText & vbCrLf & " and is italicized" MsgBox strText & vbCr & "Using file font: " & objTextStyle.fontFile, _ vbInformation, "Text Style: " & objTextStyle.Name
End If 

End Sub



Public Sub SetFontFile

Dim objTextStyle As AcadTextStyle

Set objTextStyle = ThisDrawing.TextStyles.Add("Roman") objTextStyle.fontFile = "romand.shx" 

End Sub
