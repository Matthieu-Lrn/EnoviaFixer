Attribute VB_Name = "SFB_WebServiceObject"
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'                                   WebService Access Tool Wrapper
'
' Recipe for success:
' 0- Copy this moduile and class SFB_clsShellAndWait to the required CATVBA library
' 1- Update strProgramName to pointe to the required version of the wrapper tool
' 2- Update SFB_strArgument to match the path and module of the new CATVBA
' 3- Use the syntax from Sub Test and ensure to assign your application's name to SFB_WebServiceAccessTool.UsedbyApplication
' 4- line Call SFB_WebServiceExecute("Stop") should be placed at the end of the execution of the main code/class
' 5- Ensure the WebServicePayloadScript has executed before trying to access the SFB_WebServiceAccessTool.exe from the User Temp folder.
'
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

' Manages the connection with SFB_WebServiceAccessTool.exe
' Argument can be either the actual SFB_WebServiceAccessTool object
' as returned from SFB_WebServiceAccessTool.exe or SFB_a string used
' to start/stop the process
Public Sub SFB_WebServiceExecute(Optional ByRef obj = Nothing)
    
    Dim SFB_strArgument As String
    Dim SFB_ShellCommand As String
    
    'Concatenate SFB_strArgument
    SFB_strArgument = "/ScriptLibraryType=2 /VBAToolFullName="
    SFB_strArgument = SFB_strArgument & Chr(34) & SFB_sActiveToolbarPath & Chr(34)
    SFB_strArgument = SFB_strArgument & " /VBAModuleName=SFB_WebServiceObject /VBAFunctionName=SFB_WebServiceExecute"
    
    
    ' Select action based on argument type
    Select Case True

        ' sub called externally by SFB_WebServiceAccessTool.exe
        ' to pass the SFB_WebServiceAccessTool object back to VBA
        Case LCase(TypeName(obj)) = "object", LCase(TypeName(obj)) = "webserviceaccesstool"
            Set SFB_WebServiceAccessTool = obj
            
        Case LCase(TypeName(obj)) = "string"
            Select Case LCase(obj)
                
                ' Start the SFB_WebServiceAccessTool object from the .exe
                Case "start"
                    
                    'Define Shell command
                    SFB_ShellCommand = SFB_sWebServiceAccessToolName & " " & SFB_strArgument
                    
                    ' Launch and wait for process to return
                    Dim SFB_oShell As New SFB_clsShellAndWait
                    iResult = SFB_oShell.ShellAndWait(SFB_ShellCommand, 5000, VbAppWinStyle.vbNormalNoFocus, ActionOnBreak.AbandonWait)
                    Set SFB_oShell = Nothing
                    Set fso = Nothing

                ' Stop the SFB_WebServiceAccessTool.exe process
                Case "stop"
                
                    ' purge any remaining silent log and kill the process
                    SFB_WebServiceAccessTool.Close
                    Set SFB_WebServiceAccessTool = Nothing
                    
            End Select
    End Select
End Sub



