Imports System.Runtime.InteropServices
Imports Accessibility
Imports System.Windows.Forms

Module MSAAUtils
    <DllImport("oleacc.dll")>
    Private Function AccessibleObjectFromWindow(ByVal Hwnd As Int32,
                                                 ByVal dwId As Int32,
                                                 ByRef riid As Guid,
                                                 <MarshalAs(UnmanagedType.IUnknown)> ByRef ppvObject As Object) As Int32
    End Function

    <DllImport("oleacc.dll")>
    Function AccessibleChildren(ByVal paccContainer As IAccessible, ByVal iChildStart As Int32, ByVal cChildren As Int32, <[Out]()> ByVal rgvarChildren() As Object, ByRef pcObtained As Int32) As UInteger
    End Function

    Private Const S_OK As Int32 = 0
    Private Const S_FALSE As Int32 = 1
    Private Const CHILDID_SELF As Int32 = 0

    Private Const STATE_SYSTEM_NORMAL As Int32 = 0
    Private Const STATE_SYSTEM_UNAVAILABLE As Int32 = &H1
    Private Const STATE_SYSTEM_SELECTED As Int32 = &H2
    Private Const STATE_SYSTEM_FOCUSED As Int32 = &H4
    Private Const STATE_SYSTEM_PRESSED As Int32 = &H8
    Private Const STATE_SYSTEM_CHECKED As Int32 = &H10
    Private Const STATE_SYSTEM_MIXED As Int32 = &H20
    Private Const STATE_SYSTEM_INDETERMINATE As Int32 = STATE_SYSTEM_MIXED
    Private Const STATE_SYSTEM_READONLY As Int32 = &H40
    Private Const STATE_SYSTEM_HOTTRACKED As Int32 = &H80
    Private Const STATE_SYSTEM_DEFAULT As Int32 = &H100
    Private Const STATE_SYSTEM_EXPANDED As Int32 = &H200
    Private Const STATE_SYSTEM_COLLAPSED As Int32 = &H400
    Private Const STATE_SYSTEM_BUSY As Int32 = &H800
    Private Const STATE_SYSTEM_FLOATING As Int32 = &H1000
    Private Const STATE_SYSTEM_MARQUEED As Int32 = &H2000
    Private Const STATE_SYSTEM_ANIMATED As Int32 = &H4000
    Private Const STATE_SYSTEM_INVISIBLE As Int32 = &H8000
    Private Const STATE_SYSTEM_OFFSCREEN As Int32 = &H10000
    Private Const STATE_SYSTEM_SIZEABLE As Int32 = &H20000
    Private Const STATE_SYSTEM_MOVEABLE As Int32 = &H40000
    Private Const STATE_SYSTEM_SELFVOICING As Int32 = &H80000
    Private Const STATE_SYSTEM_FOCUSABLE As Int32 = &H100000
    Private Const STATE_SYSTEM_SELECTABLE As Int32 = &H200000
    Private Const STATE_SYSTEM_LINKED As Int32 = &H400000
    Private Const STATE_SYSTEM_TRAVERSED As Int32 = &H800000
    Private Const STATE_SYSTEM_MULTISELECTABLE As Int32 = &H1000000
    Private Const STATE_SYSTEM_EXTSELECTABLE As Int32 = &H2000000
    Private Const STATE_SYSTEM_ALERT_LOW As Int32 = &H4000000
    Private Const STATE_SYSTEM_ALERT_MEDIUM As Int32 = &H8000000
    Private Const STATE_SYSTEM_ALERT_HIGH As Int32 = &H10000000
    Private Const STATE_SYSTEM_PROTECTED As Int32 = &H20000000
    Private Const STATE_SYSTEM_VALID As Int32 = &H7FFFFFFF


    Public Sub ProcessAccessibleWindow(handle As IntPtr)
        Dim int32Handle As Int32
        int32Handle = handle.ToInt32()
        Dim IID_IAccessible As Guid = New Guid("618736E0-3C3D-11CF-810C-00AA00389B71")
        Dim iaccWindowObj As Object
        If AccessibleObjectFromWindow(int32Handle, 0, IID_IAccessible, iaccWindowObj) <> S_OK Then
            frmMain.ShowPrompt("Error obtaining accessible object for '" & EnumWindows.GetWindowTextByHandle(handle) & "' window.", 5000)
            Exit Sub
        End If

        If TypeOf iaccWindowObj IsNot Accessibility.IAccessible Then
            frmMain.ShowPrompt("Not an accessible window: '" & EnumWindows.GetWindowTextByHandle(handle) & "'.", 5000)
            Exit Sub
        End If

        Dim iaccWindow As Accessibility.IAccessible
        iaccWindow = iaccWindowObj

        Dim madeCallout As Boolean
        ProcessAccessibleChildren(iaccWindow, madeCallout)

    End Sub

    ' Returns true on success
    ' Returns fasle if an error occurred.
    Public Function ProcessAccessibleChildren(parentAcc As Accessibility.IAccessible, ByRef madeCallout As Boolean) As Boolean
        Dim childCount As Int32
        childCount = parentAcc.accChildCount()
        If childCount = 0 Then
            ' No Children.  Make a callout for this?
            If MaybeMakeCallout(parentAcc, CHILDID_SELF) Then
                madeCallout = True
            Else
                madeCallout = False
            End If
            Return True
        End If

        ' Has children.  Remember the default action
        Dim parentDefAction As String

        Try
            parentDefAction = parentAcc.accDefaultAction(CHILDID_SELF)
        Catch ex As Exception
            parentDefAction = ""
        End Try

        Dim childHasCallout As Boolean = False

        ' Call process on each child
        ' Get children
        Dim returnedCount As Int32
        Dim children(childCount) As Object
        Dim retVal As UInteger
        retVal = AccessibleChildren(parentAcc, 0, childCount, children, returnedCount)
        If retVal <> S_OK And retVal <> S_FALSE Then
            frmMain.ShowPrompt("An error occurred while obtaining accessibility children.", 5000)
            Application.DoEvents()
            Return False
        End If
        childCount = returnedCount
        Dim child As IAccessible
        Dim childID As Int32
        Dim innerMadeCallout As Boolean
        For i = 0 To childCount - 1
            If TypeOf children(i) Is IAccessible Then
                ' An actual accessible object and not a "simple element"
                child = children(i)
                ProcessAccessibleChildren(child, innerMadeCallout)
                If innerMadeCallout Then
                    childHasCallout = True
                End If
            ElseIf TypeOf children(i) Is Int32 Then
                ' Simple element.  It's accessiblilty properties come from the parent
                ' Make callout?
                childID = children(i)
                If MaybeMakeCallout(parentAcc, childID) Then
                    childHasCallout = True
                End If
            Else
                frmMain.ShowPrompt("Unknown child type: " & children(i).GetType().ToString(), 5000)
                Application.DoEvents()
                Return False
            End If
        Next

        madeCallout = False
        If childHasCallout Then
            madeCallout = True
        Else
            ' Does this parent object have a default action?
            If parentDefAction <> "" Then
                ' Need to make a callout for this parent
                If MaybeMakeCallout(parentAcc, CHILDID_SELF) Then
                    madeCallout = True
                End If
            End If
        End If
        Return True
    End Function

    ' Returns true if a callout was made, false if a callout was not made
    Private Function MaybeMakeCallout(iaccObject As IAccessible, childID As Int32) As Boolean
        ' Make a callout if 
        '   state Is focusable
        '   Def Action is not null
        Dim state As Int32
        Try
            state = iaccObject.accState(childID)
        Catch ex As Exception
            state = STATE_SYSTEM_UNAVAILABLE
        End Try

        Dim defAction As String = ""

        Try
            defAction = iaccObject.accDefaultAction(childID)
        Catch ex As Exception
            defAction = ""
        End Try

        ' Get bounding rect
        Dim left, top, width, height As Int32
        left = 0
        top = 0
        width = 0
        height = 0
        Try
            iaccObject.accLocation(left, top, width, height, childID)
        Catch ex As Exception
            left = 0
            top = 0
            width = 0
            height = 0
        End Try

        If (state And STATE_SYSTEM_FOCUSABLE) Or (defAction <> "") Then
            ' Make callout
            If left >= 0 And top >= 0 And width >= 0 And height >= 0 Then
                Dim boundsRect As New System.Windows.Rect(left, top, width, height)
                frmMain.MakeCallout(boundsRect, "")
                Return True
            Else
                Return False
            End If
        Else
            ' No callout
            Return False
        End If
    End Function

End Module
