﻿'****************************** Module Header ******************************'
' Module Name:  MainModule.vb
' Project:      VBExeCOMServer
' Copyright (c) Microsoft Corporation.
' 
' The main entry point for the application. It is responsible for starting  
' the out-of-proc COM server registered in the executable.
' 
' This source is subject to the Microsoft Public License.
' See http://www.microsoft.com/en-us/openness/licenses.aspx#MPL.
' All other rights reserved.
' 
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
' EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED 
' WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
'***************************************************************************'

Imports System.Threading

Module MainModule
    Dim memory As Object

    Sub Main()
        ' Run the out-of-process COM server
        ExeCOMServer.Instance.Run()
    End Sub

End Module
