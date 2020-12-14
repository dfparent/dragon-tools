'#Uses "Keyboard.bas"
'#Uses "KeyboardConstants.bas"
'option explicit


Public Sub ShowContextMenu()
    KeyPress(VK_CONTEXT_MENU)
    Wait 0.1
End Sub

