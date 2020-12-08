Imports DNSTools

Module Dragon

    Public Sub EmulateCommand(command As String)
        Dim dragon As New DgnEngineControl
        dragon.RecognitionMimicEx(DgnMimicTypeConstants.dgnmimictypeCommand,
                                  DgnMimicFormatConstants.dgnmimicformatPlain,
                                  command)
    End Sub
End Module
