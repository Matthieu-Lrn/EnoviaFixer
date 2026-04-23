Attribute VB_Name = "WebServiceObject"
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'                                   WebService Access Tool Wrapper
'
' Recipe for success:
' 0- Copy this moduile and class clsShellAndWait to the required CATVBA library
' 1- Update strProgramName to pointe to the required version of the wrapper tool
' 2- Update strArgument to match the path and module of the new CATVBA
' 3- Use the syntax from Sub Test and ensure to assign your application's name to WebServiceAccessTool.UsedbyApplication
' 4- line Call WebServiceExecute("Stop") should be placed at the end of the execution of the main code/class
' 5- Ensure the WebServicePayloadScript has executed before trying to access the WebServiceAccessTool.exe from the User Temp folder.
'
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

' Manages the connection with WebServiceAccessTool.exe
' Argument can be either the actual WebServiceAccessTool object
' as returned from WebServiceAccessTool.exe or a string used
' to start/stop the process
Public Sub WebServiceExecute(Optional ByRef obj = Nothing)
    
    Dim strArgument As String
    Dim ShellCommand As String
    
    'Concatenate strArgument
    strArgument = "/ScriptLibraryType=2 /VBAToolFullName="
    strArgument = strArgument & Chr(34) & sActiveToolbarPath & Chr(34)
    strArgument = strArgument & " /VBAModuleName=WebServiceObject /VBAFunctionName=WebServiceExecute"
    
    
    ' Select action based on argument type
    Select Case True

        ' sub called externally by WebServiceAccessTool.exe
        ' to pass the WebServiceAccessTool object back to VBA
        Case LCase(TypeName(obj)) = "object", LCase(TypeName(obj)) = "webserviceaccesstool"
            Set WebServiceAccessTool = obj
            
        Case LCase(TypeName(obj)) = "string"
            Select Case LCase(obj)
                
                ' Start the WebServiceAccessTool object from the .exe
                Case "start"
                    
                    'Define Shell command
                    ShellCommand = sWebServiceAccessToolName & " " & strArgument
                    
                    ' Launch and wait for process to return
                    Dim oShell As New clsShellAndWait
                    iResult = oShell.ShellAndWait(ShellCommand, 5000, VbAppWinStyle.vbNormalNoFocus, ActionOnBreak.AbandonWait)
                    Set oShell = Nothing
                    Set fso = Nothing

                ' Stop the WebServiceAccessTool.exe process
                Case "stop"
                
                    ' purge any remaining silent log and kill the process
                    WebServiceAccessTool.Close
                    Set WebServiceAccessTool = Nothing
                    
            End Select
    End Select
End Sub



