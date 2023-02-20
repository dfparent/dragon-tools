Public Class frmHelp
    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmHelp_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        rtxtHelp.Text = "To perform any below function, first bring the focus to the tool by pressing the Global Hot Key combination:  ctrl + shift + alt + " & GetBringToForegroundHotKeyString() & " and then do the following:

	To move the pointer to an element:  type in the callout number and then <tab> 
	To click on an element: type in the callout number and then <enter>.
	To shift+click on an element: type in the callout number and then shift+<enter>.
	To ctrl+click on an element: type in the callout number and then ctrl+<enter>.
	To double click on an element: type in the callout number and then alt+<enter>.
	To right click on an element: type in the callout number and then ctrl+shift+<enter>.
	To pick up an element (mouse down): type in the callout number and then <down arrow>.
	To drop an element (mouse up): type in the callout number and then <up arrow>.
    To start a drag operation: type in the starting callout number and then <g>.
    To complete a drag operation (drop): type in the destination callout number and then <p>.
    To ctrl+drag: type in the destination callout number and then ctrl+<p>.
    To shift+drag: type in the destination callout number and then shift+<p>.
	To cancel callout entry:  type <esc>.
	To refresh the callouts: type F5.
	To show or hide the callouts: type F4.
	To show options: type F2.
	To show this help text:  type F1.

You can also use the following global hot keys (no need to press ctrl+shift+alt+" & GetBringToForegroundHotKeyString() & " first):

	To show flag options:  type ctrl+shift+alt+F2.
	To toggle the sticky keys setting:  type ctrl+shift+alt+F3.
	To show/hide the callouts:  type ctrl+shift+alt+F4.
	To refresh the callouts:  type ctrl+shift+alt+F5.

To change the 'Bring Click By Numbers to the foreground' global hotkey combination, see the following registry key:

    HKCU\" & REGISTRY_PATH_GLOBAL_SETTINGS

    End Sub
End Class