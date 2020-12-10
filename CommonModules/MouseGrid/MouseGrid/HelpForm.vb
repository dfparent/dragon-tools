Imports System.Threading
Imports System.Text

Public Class frmHelp

    Private Sub frmHelp_Load(sender As Object, e As EventArgs) Handles Me.Load

        webHelp.Document.OpenNew(False)
        webHelp.Document.Write(GetHelpHtml())

    End Sub

    Private Function GetHelpHtml() As String
        Dim text As New StringBuilder
        text.Append("<html><head>
<style>
table, th, td {
  border: 1px solid black;
}
</style>
</head>
<body>")
        text.Append("<h1>Mouse Grid Help Manual</h1>")
        text.Append("<p>This application allows you to easily click anywhere on a workspace using only the keyboard.  
It overlays a grid onto any area of the computer screen that you wish and then allows you to apply a mouse click to any area of the grid using simple keystrokes.
</p>")
        text.Append("<p>The script is very customizable, allowing you to specify:
<ul>
    <li>the grid placement and size</li>
    <li>how wide grid columns and how high grid rows are</li>
    <li>whether the grid remains displayed or is hidden after a mouse click</li>
    <li>a single mouse click, double-click, or right-click</li>
    <li>drag and drop</li>
    <li>grid transparency</li>
    <li>&quot;always on top&quot; behavior</li>
    <li>screen (absolute) or client (relative) window coordinates</li>
</ul></p>")
        text.Append("<h2>Usage Examples</h2>")
        text.Append("<p>The grid was primarily intended to be used with macros and so can be completely
customized using the command line.</p>")
        text.Append("<p>For example, to display the grid at a specific location, with a specific width and height, and a customized 
row height and column with, you might use this command line:
MouseGrid.exe --width 1658 --height 961 --locationx 6 --locationy 22 --rowheight 20 --columnwidth 20
</p>")
        text.Append("<p>To display the grid over the client area of a window with a certain HWND (note no width or height specified):
MouseGrid.exe --sizeToClient 1234567 </p>")
        text.Append("<p>To display the grid over the client area of the foreground window (note no width, height, or HWND specified):
MouseGrid.exe --sizeToClient </p>")
        text.Append("<p>To display the grid over the client area of the foreground window docked to the top of the client window (note no width, height, or HWND specified):
MouseGrid.exe --dock-to-client top</p>")
        text.Append("<h2>Position Mode</h2>")
        text.Append("<p>To determine the parameter values needed to place a grid at a precise location, run the grid in position mode.
When the application is started without any commandline parameters, it automatically enters position mode:
MouseGrid.exe</p>")
        text.Append("<p>The grid will appear with a draggable title bar and can be resized.  
Use the menus that appear during position mode to customize how the grid will appear.  The parameters needed 
to start the grid up in the exact location and size it is currently in, with all selected options is shown in the status bar.
Type ctrl+c to copy these parameters onto the clipboard.    
You can then use these parameters to start the grid from the command line, or from a marco.</p>")
        text.Append("<h2>Normal Mode</h2>")
        text.Append("<p>Normal mode is the mode that you will want to use for performing your click actions using voice macros.  
When you start Mouse Grid with any command line parameters, it will run in normal mode.</p>")
        text.Append("<p>Use the following keystrokes to perform the indicated actions when Mouse Grid has the focus:</p>")
        text.Append("<p>r##: move the pointer to the indicated row.  To move to row 23, type 'r23' followed by &lt;enter&gt;.  You can also do other actions without pressing &lt;enter&gt;.<br>
    c##: move the pointer to the indicated column.  To move to column 45, type 'c45' followed by &lt;enter&gt;.  You can also do other actions without pressing &lt;enter&gt;.<br>
    s: single click the mouse (ctrl+s to control click, shift+s to shift click)<br>
	d: double click the mouse (ctrl+d to control double click, shift+d to shift double click)<br>
    t: right click the mouse (ctrl+t to control right click, shift+t to shift right click)<br>
    m: move the mouse. Grid will close after the move.<br>
	f: move the mouse in preparation for a drag action.  Grid remains open after move. 
    g: drag the mouse from its current location to currently selected cell in the grid (ctrl+g to control drag, shift+g to shift drag)
    y: toggle the sticky setting.  If sticky, the grid will remain displayed after issuing a click.  If not sticky, the grid will close after issuing a click.
	x: close the application.
    F1: Show help
    <br>
    &lt;up arrow&gt;: increase row height in position mode<br>
    &lt;down arrowgt;: decrease row height in position mode<br>
    &lt;left arrowgt;: decrease column width in position mode<br>
    &lt;right arrowgt;: increase column width in position mode<br>
	ctrl+c: copy the current parameters into the clipboard as the commandline switches needed to run the grid</p>")
        text.Append("<h2>Command Line Reference</h2>")
        text.Append("<p><table style='width:100%'><tr><th>Parameter</th><th>Description</th></tr>
<tr><td>--width</td><td>Specify the width Of the grid in pixels.</td></tr>
<tr><td>--height</td><td>Specify the height Of the grid in pixels.</td></tr>
<tr><td>--location-x</td><td>Specify the screen coordinate Of the left edge of the grid.  Default Is 0.  Coordinates are relative to the screen, by default.  See --size-to-client for more details.</td></tr>
<tr><td>--location-y</td><td>Specify the screen coordinates Of the top edge of the grid.  Default Is 0.  Coordinates are relative to the screen, by default.  See --size-to-client for more details.</td></tr>
<tr><td>--size-to-client</td><td>Use this switch to size and position the grid relative to the client window of an application, rather than relative to the screen.
<ul>
    <li>When no argument is provided, the grid will be sized to the client area of the application currently in the foreground.</li>
    <li>You can provide a window handle (integer) as an argument to size the grid over the client area of the window belonging to that handle.</li>
    <li>The location-x And location-y parameters (if provided) are interpreted relative To the client window.  When location-X and location-y are omitted, The top left corner of the client area will be used for the location. </li>
    <li>When width and height are omitted, Then the grid will have the same width and height as the client window.</li>
</ul></td></tr>
<tr><td>--dock-to-client</td><td>Use this switch to size the grid over a portion of the client area.  This argument must be followed by one of the following arguments:
<ul>
    <li>left</li>
    <li>right</li>
    <li>top</li>
    <li>bottom</li>
    <li>center</li>
</ul></td></tr>
<tr><td>--row-height</td><td>Specify the height of a grid row in pixels.</td></tr>
<tr><td>--column-width</td><td>Specify the width of a grid column in pixels.</td></tr>
<tr><td>--sticky</td><td>Include this switch to make the grid remain open after issuing a mouse action.  Normally, the grid automatically closes after issuing a mouse action.</td></tr>
<tr><td>--opacity</td><td>Value from 0 to 1.  0 makes the grid completely transparent.  1 makes the grid completely opaque.</td></tr>
<tr><td>--always-on-top</td><td>Include this switch to make the grid stay on top of all windows.</td></tr></table></p>")

        text.Append("<p></p>")
        text.Append("</body></html>")
        Return text.ToString()
    End Function

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
End Class