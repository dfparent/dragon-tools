'#Language "WWB.NET"

Public Enum CompassDirection
    N
    S
    E
    W
    NW
    NE
    SW
    SE
    Invalid
End Enum

Public Function GetCompassDirectionEnum(theDirection As String) As CompassDirection
    Select Case theDirection
        Case "North"
            Return CompassDirection.N
        Case "South"
            Return CompassDirection.S
        Case "East"
            Return CompassDirection.E
        Case "West"
            Return CompassDirection.W
        Case "Northwest"
            Return CompassDirection.NW
        Case "Northeast"
            Return CompassDirection.NE
        Case "Southwest"
            Return CompassDirection.SW
        Case "Southeast"
            Return CompassDirection.SE
        Case Else
            Return CompassDirection.Invalid
    End Select
End Function
