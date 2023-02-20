'#Uses "utilities.bas"
'#Uses "cache.bas"
'#Language "WWB.NET"
Option Explicit On

' Clicks the menu on the current process
Public sub ClickMenu(menuName as string)
    Dim menus As Object
    menus = GetMenus(GetForegroundProcessName())

    If menus.ContainsKey(menuName) Then
        SendKeys("%{" & menus(menuName)(0) & "}")
    Else
        ' Guess at the menu mnemonic by using the first character
        SendKeys("%" + menuName.Substring(0, 1))
    End if 
end sub

' Clicks a submenu on the current process
public Sub ClickSubMenu(menuName as string)
    Dim menus As Object
    menus = GetMenus(GetForegroundProcessName())

    If menus.ContainsKey(menuName) Then
        SendKeys("{" & menus(menuName)(0) & "}")
    Else
        ' Guess at the menu mnemonic by using the first character
        SendKeys(menuName.Substring(0, 1))
    End If
End Sub

