'option explicit


Public Sub PickColor(color As String)
    SendKeys("{Home}")

    Select Case color
        Case "Black"
            SendKeys("{Right}~")
            Exit Sub
        Case "Grey"
            SendKeys("{Right 2}~")
            Exit Sub
        Case "White"
            SendKeys("{Enter}")
            Exit Sub
        Case "None"
            SendKeys("{Down 8}~")
            Exit Sub
    End Select

    SendKeys("{Home}{Down 6}")

    Select Case color
        Case "Green"
            SendKeys("{Right 5}{Enter}")
        Case "Light Green"
            SendKeys("{Right 4}{Enter}")
        Case "Blue", "Light Blue", "Teal"
            SendKeys("{Right 6}{Enter}")
        Case "Dark Blue"
            SendKeys("{Right 7}{Enter}")
        Case "Navy Blue"
            SendKeys("{Right 8}{Enter}")
        Case "Maroon"
            SendKeys("{Enter}")
        Case "Yellow"
            SendKeys("{Right 3}{Enter}")
        Case "Purple"
            SendKeys("{Right 9}{Enter}")
        Case "Red"
            SendKeys("{Right}{Enter}")
        Case "Orange"
            SendKeys("{Right 2}{Enter}")
    End Select
End Sub