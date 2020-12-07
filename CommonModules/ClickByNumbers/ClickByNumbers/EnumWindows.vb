Imports System.Runtime.InteropServices
Imports System.Windows

Module EnumWindows

    Public Structure RECT
        Dim Left As Integer
        Dim Top As Integer
        Dim Right As Integer
        Dim Bottom As Integer
    End Structure

    Private Delegate Function EnumWindowsProc(ByVal hWnd As IntPtr, ByRef lParam As IntPtr) As Boolean

    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Function EnumChildWindows(ByVal hWndParent As System.IntPtr, ByVal lpEnumFunc As EnumWindowsProc, ByVal lParam As Integer) As Boolean
    End Function

    <DllImport("user32.dll")>
    Public Function GetWindowRect(ByVal hWnd As IntPtr, ByRef lpRect As RECT) As Boolean
    End Function

    <DllImport("user32.dll", EntryPoint:="GetWindowText")>
    Public Function GetWindowText(ByVal hwnd As IntPtr, ByVal lpString As System.Text.StringBuilder, ByVal cch As Integer) As Integer
    End Function

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Function GetWindowTextLength(ByVal hwnd As IntPtr) As Integer
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Function IsWindowVisible(ByVal hWnd As IntPtr) As Boolean
    End Function

    Public Sub ProcessWindowChildren(parentHandle As System.IntPtr)
        EnumChildWindows(parentHandle, AddressOf EnumWindow, 0)
    End Sub

    Private Function EnumWindow(ByVal Handle As IntPtr, ByRef Parameter As IntPtr) As Boolean

        If Handle.ToInt32 = 0 Then
            Return False
        End If


        Dim text = GetWindowTextByHandle(Handle)

        Dim theRect As RECT
        Dim boundingRect As System.Windows.Rect
        If GetWindowRect(Handle, theRect) Then
            Dim width As Long
            Dim height As Long
            width = theRect.Right - theRect.Left
            height = theRect.Bottom - theRect.Top
            If width > 0 And height > 0 Then
                boundingRect.X = theRect.Left
                boundingRect.Y = theRect.Top
                boundingRect.Width = width
                boundingRect.Height = height
                If IsWindowVisible(Handle) Then
                    'controls.Add(frmMain.MakeCallout(boundingRect, text))
                    frmMain.MakeCallout(boundingRect, text)
                End If
            End If
        End If

        'frmMain.Controls.AddRange(controls.ToArray())
        'frmMain.RenderCallouts()

        Return True

    End Function

    Public Function GetWindowTextByHandle(handle As IntPtr) As String
        Dim length As Integer
        Dim text As String = ""
        length = GetWindowTextLength(handle)
        If length > 0 Then
            Dim sb As New System.Text.StringBuilder("", length)
            GetWindowText(handle, sb, sb.Capacity + 1)
            text = sb.ToString()
        End If

        Return text
    End Function
End Module
