This script allows you to easily click anywhere on a workspace using only the keyboard.  It overlays a grid onto any area of the computer screen that you wish and then allows you to apply a mouse click to any area of the grid using simple keystrokes.

The script is very customizable, allowing you to specify
    * the grid placement and size
    * how wide grid columns and how high grid rows are
    * whether the grid remains displayed or is hidden after a mouse click
    * a single mouse click, double-click, or right-click
    * drag and drop
    * grid transparency
    * "always on top" behavior
    * screen (absolute) or client (relative) window coordinates

The grid was primarily intended to be used with macros and so can be completely
customized using the command line.  For better performance, new "command line" parameters
can be applied without having to restart the grid, so that the grid window can be minimized
in between grid usages.

For example, to display the grid over your entire desktop, you might use this command line:
pythonw <mypath>CustomGrid.py --width 1658 --height 961 --locationx 6 --locationy 22 --rowheight 20 --columnwidth 20 --numrows 20 --numcolumns 20

To display the grid over the client area of a certain application:
pythonw <mypath>CustomGrid.py --sizeToClient winword.exe --rowheight 20 --columnwidth 20 --numrows 20 --numcolumns 20

To display the grid over the client area of a window with a certain HWND:
pythonw <mypath>CustomGrid.py --sizeToClient 1234567 --rowheight 20 --columnwidth 20 --numrows 20 --numcolumns 20

For a complete description of all command line switches, use:
pythonw <mypath>CustomGrid.py ?

To make it easier to figure out what the desired parameter values are,
run the grid in position mode:
pythonw <mypath>CustomGrid.py --positionMode

The grid will appear with a draggable title bar and can be resized.  Exit the grid by typing 'x'
and you'll be prompted to copy onto the clipboard the required command line parameters that could be used to display the grid
exactly the same way.  You can then go and simply paste commandline parameters in a macro
and call the macro to display the grid.

'''Once the grid is displayed, you can control the application using the following commands:
    ?: show this help text.

    r##: move the pointer to the indicated row.  To move to row 23, type 'r23' followed by <enter>.
    c##: move the pointer to the indicated column.  To move to column 45, type 'c45' followed by <enter>.
    You can also combine a row and column in a single command:  'r23c45' followed by <enter>  does a single move.

    s: single click the mouse.
    d: double click the mouse.
    t: right click the mouse.
    f: mark the drag from position
    g: drag the mouse from the drag start location to the current location
    y: toggle the sticky setting.  If sticky, the grid will remain displayed after issuing a click.  If not sticky, the grid will close after issuing a click.
    h: refresh the display. This is especially useful with the "sticky" setting turned on.
    l: allow you to enter new command line arguments and reconfigure the grid without restarting the application.

    <up arrow>: increase row height in position mode
    <down arrow>: decrease row height in position mode
    <left arrow>: decrease column width in position mode
    <right arrow>: increase column width in position mode
    <ctrl>c: copy the current parameters into the clipboard as the commandline switches needed to run the grid
    <esc>: hide the grid.
    x: close the application.

    In position mode, the origin of the grid as indicated by a small circle at the top left corner of the grid.  This point will be the locationx and location y values.''')

        parser.add_argument("-p", "--positionMode", action="store_true",
                help="Allows you to resize the grid and determine the desired parameters.  When the window is closed, parameters are printed to the command line.")
        parser.add_argument("--width", action="store", type=int,
                help="Specify the width of the grid in pixels.  You may omit this switch if you specify both --columnwidth and --numcolumns.")
        parser.add_argument("--height", action="store", type=int,
                help="Specify the height of the grid in pixels.  You may omit this switch if you specify both --rowheight and --numrows.")
        parser.add_argument("--locationx", action="store", type=int, default='0',
                help="Specify the x pixel screen coordinates of the upper left corner of the grid.  Default is 0.")
        parser.add_argument("--locationy", action="store", type=int, default='0',
                help="Specify the y pixel screen coordinates of the upper left corner of the grid.  Default is 0.")
        parser.add_argument("--sizeToClient", action="store",
                help="Use this switch to size and position the grid over the client window of the given application.  " + \
                     "Specify the process file name of the target application or the HWND of the application window. " + \
                     "If locationx and locationy switches are provided then they are interpreted relative to the client window.  " + \
                     "If width, height, locationx and locationy switches are not provided then the grid will cover the entire client window.  " + \
                     "It is ignored in position mode.")
        parser.add_argument("--rowheight", action="store", type=int, help ='Specify the height of a grid row in pixels.')
        parser.add_argument("--columnwidth", action="store", type=int, help ='Specify the width of a grid column in pixels.')
        parser.add_argument("--numrows", action="store", type=int, help ='Specify the number of rows in the grid.')
        parser.add_argument("--numcolumns", action="store", type=int, help ='Specify the number of columns in the grid.')
        parser.add_argument("--sticky", action="store_true", help='Include this switch to make the grid sticky.')
        parser.add_argument("--opacity", action="store", type=int, help='Value from 0 to 100.  0 makes the grid completely transparent.  100 makes the grid completely opaque.')
        parser.add_argument("--alwaysontop", action="store_true", help='Include this switch to make the grid stay on top of all windows.')
        return parser
