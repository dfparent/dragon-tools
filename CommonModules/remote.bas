'#Uses "cache.bas"
'#Language "WWB.NET"
Option Explicit On

Public Const CURRENT_RDP_APP = "CurrentRdpApplication"

Public Sub SetActiveApplication(appName As String)
    UpdateCacheValueSingle(CURRENT_RDP_APP, appName)
	TTSPlayString("App is " & appName)
End Sub

Public Function GetActiveApplication() As String
    Return GetCacheValueSingle(CURRENT_RDP_APP)
End Function

Public Function CheckActiveApplication(appName As String) As Boolean
    Dim activeApp As String
    activeApp = GetCacheValueSingle(CURRENT_RDP_APP)
    If LCase(activeApp) = LCase(appName) Then
        Return True
    Else
		TTSPlayString("Current app is " & activeApp)
        Return False
    End If
End Function