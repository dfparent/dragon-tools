Imports ClickByNumbers
Imports Microsoft.Win32

Public Class AppSettings
    Private Const DEFAULT_OPACITY = 75
    Private isStickyValue As Boolean = True
    Private isDisabledValue As Boolean = False
    Private opacityValue As Integer = DEFAULT_OPACITY
    Private usesWindowHandleDiscoveryValue As Boolean = True
    Private usesUIAutomationDiscoveryValue As Boolean = True
    Private usesMSAADiscoveryValue As Boolean = False
    Private processNameValue As String = ""

    Public Sub New()

    End Sub

    Public Sub New(processName As String)
        ' Read in from registry
        Load(processName)
    End Sub

    Public Function Load(windowHandle As IntPtr) As Boolean
        Return Load(GetProcessNameForWindow(windowHandle))
    End Function

    Public Function Load(processName As String) As Boolean
        If processName = "" Then
            processName = DEFAULT_PROCESS_NAME
        End If

        processNameValue = processName


        Dim key As RegistryKey = My.Computer.Registry.CurrentUser.OpenSubKey(REGISTRY_PATH & "\" & processNameValue, False)
        If key Is Nothing Then
            ' Key does not exist.  Try the default settings
            key = My.Computer.Registry.CurrentUser.OpenSubKey(REGISTRY_PATH & "\" & DEFAULT_PROCESS_NAME, False)
        End If

        If key IsNot Nothing Then
            Try
                isStickyValue = key.GetValue(REGISTRY_VALUE_STICKY, True)
                isDisabledValue = key.GetValue(REGISTRY_VALUE_DISABLED, False)
                opacityValue = key.GetValue(REGISTRY_VALUE_OPACITY, DEFAULT_OPACITY)
                usesWindowHandleDiscoveryValue = key.GetValue(REGISTRY_VALUE_USES_WINDOW_HANDLE_DISCOVERY, True)
                usesUIAutomationDiscoveryValue = key.GetValue(REGISTRY_VALUE_USES_UI_AUTOMATION_DISCOVERY, True)
                usesMSAADiscoveryValue = key.GetValue(REGISTRY_VALUE_USES_MSAA_DISCOVERY, False)
            Catch ex As Exception
                ' Use defaults
                isStickyValue = True
                isDisabledValue = False
                opacityValue = DEFAULT_OPACITY
                usesWindowHandleDiscoveryValue = True
                usesUIAutomationDiscoveryValue = True
                usesMSAADiscoveryValue = False
                Return False
            Finally
                key.Close()
            End Try
        Else
            isStickyValue = True
            isDisabledValue = False
            opacityValue = DEFAULT_OPACITY
            usesWindowHandleDiscoveryValue = True
            usesUIAutomationDiscoveryValue = True
            usesMSAADiscoveryValue = False
        End If

        Return True
    End Function

    Public Function Save() As Boolean
        If processNameValue = "" Then
            Return False
        End If

        Dim key As RegistryKey = My.Computer.Registry.CurrentUser.OpenSubKey(REGISTRY_PATH & "\" & processNameValue, True)
        If key Is Nothing Then
            key = My.Computer.Registry.CurrentUser.CreateSubKey(REGISTRY_PATH & "\" & processNameValue, True)
            If key Is Nothing Then
                Return False
            End If
        End If

        Try
            key.SetValue(REGISTRY_VALUE_STICKY, isStickyValue)
            key.SetValue(REGISTRY_VALUE_DISABLED, isDisabledValue)
            key.SetValue(REGISTRY_VALUE_OPACITY, opacityValue)
            key.SetValue(REGISTRY_VALUE_USES_WINDOW_HANDLE_DISCOVERY, usesWindowHandleDiscoveryValue)
            key.SetValue(REGISTRY_VALUE_USES_UI_AUTOMATION_DISCOVERY, usesUIAutomationDiscoveryValue)
            key.SetValue(REGISTRY_VALUE_USES_MSAA_DISCOVERY, usesMSAADiscoveryValue)
        Catch ex As Exception
            Return False
        Finally
            key.Close()
        End Try

        Return True
    End Function

    Public Function Delete() As Boolean
        If processNameValue = "" Then
            Return False
        End If

        Try
            My.Computer.Registry.CurrentUser.DeleteSubKey(REGISTRY_PATH & "\" & processNameValue, False)
        Catch ex As Exception
            Return False
        End Try

        Return True
    End Function

    Public Property IsSticky As Boolean
        Get
            Return isStickyValue
        End Get
        Set(value As Boolean)
            isStickyValue = value
        End Set
    End Property

    Public Property IsDisabled As Boolean
        Get
            Return isDisabledValue
        End Get
        Set(value As Boolean)
            isDisabledValue = value
        End Set
    End Property

    Public Property Opacity As Integer
        Get
            Return opacityValue
        End Get
        Set(value As Integer)
            opacityValue = value
        End Set
    End Property

    Public Property UsesWindowHandleDiscovery As Boolean
        Get
            Return usesWindowHandleDiscoveryValue
        End Get
        Set(value As Boolean)
            usesWindowHandleDiscoveryValue = value
        End Set
    End Property

    Public Property UsesUIAutomationDiscovery As Boolean
        Get
            Return usesUIAutomationDiscoveryValue
        End Get
        Set(value As Boolean)
            usesUIAutomationDiscoveryValue = value
        End Set
    End Property

    Public Property UsesMSAADiscovery As Boolean
        Get
            Return usesMSAADiscoveryValue
        End Get
        Set(value As Boolean)
            usesMSAADiscoveryValue = value
        End Set
    End Property

End Class
