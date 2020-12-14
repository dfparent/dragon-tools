'#Uses "utilities.bas"
'#Uses "keyboard.bas"
'#Uses "KeyboardConstants.bas"
'#Language "WWB.NET"
Option Explicit On

Private Sub SelectTable(tablename As String)
    SendKeys("{F11 2}")
    SendKeys("^f")
    SendKeys formatPascalCase(tablename)
End Sub

Private Sub ShowTableContextMenu(tablename As String)
    SelectTable(tablename)
    SendKeys("{Tab 3}")
    ShowContextMenu
End Sub

Public Sub NavigationFind(findText As String)
    SelectTable(findText)
End Sub

Public Sub CancelNavigationFind()
    SendKeys("{F11 2}")
    SendKeys("^f")
    SendKeys("{Escape}")
    Wait 0.1
    RepeatKeyStrokes("{tab}", 3, 1)
End Sub

Public Sub OpenTable(tablename As String)
    SelectTable(tablename)
    SendKeys("{Escape}{Enter}")
End Sub

Public Sub RenameTable(tablename As String)
    SelectTable(tablename)
    SendKeys("{F2}")
End Sub

Public Sub DesignTable(tablename As String)
    ShowTableContextMenu(tablename)
    SendKeys("d")
    Wait 0.1
    CancelNavigationFind()
End Sub