'VST Tool Caller
'Open VST Tool, build a VST, close tool.
'Example: cscript buildVST.vbs <vstToolPath> <a2lPath> <h32Path> <vstPath> <Optional:mapPath>
'Note: File paths may be specified in any order
'mwaterbu@ford.com
'Copyright: Ford Motor Company Limited

Option Explicit

Main

Sub Main
    Dim xl
    Dim wb
    Dim toolPath, a2lPath, mapPath, h32Path, vstPath
    Dim file_names
    Dim result
    Dim resultOut
    Dim pos 
	
    
    

    If WScript.Arguments.Count < 4 Or WScript.Arguments.Count > 5 Then
        WScript.Echo "Invalid number of parameters!"
        WScript.Quit(1)
        Exit Sub
    End If
    toolPath = findArg("xlsm")
    a2lPath = findArg("a2l")
    h32Path = findArg("h32")
    vstPath = findArg("vst")
    If WScript.Arguments.Count = 5 Then mapPath = findArg("map") Else mapPath = ""
    If IsEmpty(toolPath) Or IsEmpty(a2lPath) Or IsEmpty(h32Path) Or IsEmpty(vstPath) Or IsEmpty(mapPath) Then Exit Sub
    On Error Resume Next

    Set xl = CreateObject("Excel.Application")
    If Err.number <> 0 Then
         ShowError("Error creating Excel Object. User may not have permission to open Excel")
    End if 
    xl.Visible = True 'Not necessary, but allows user to see some progress
    xl.DisplayAlerts = False 'Do not display errors or other messages that could halt execution
    xl.AutomationSecurity = 1
    xl.EnableEvents = False

    Set wb = xl.Workbooks.Open(toolPath)
    If Err.number <> 0 Then
         xl.Quit
         ShowError("File Path " & toolPath & " may be incorrect")
    End if 
    WScript.Sleep 500
    file_names = Array(a2lPath, mapPath, h32Path, vstPath)

    result = xl.Run("ThisWorkbook.StartupTasks", True)
    if Err.number <> 0 Then
        wb.Close False 'Don't save workbook
        xl.Quit
        WScript.Echo result
        ShowError("Error in StartUp Tasks")
    End If
    xl.EnableEvents = True
    
    result = xl.Run("BuildVST.BuildVSTFile", False, True, file_names)
    If Err.number <> 0 Then
        wb.Close False 'Don't save workbook
        xl.Quit
        WScript.Echo result
        ShowError("Error in BuildVst" )
    End If
        
    

'[variable =] InStr(Start, String1, String2)    
'Start (Optional)The starting position of the search. Matches before Start are ignored. If omitted, then Start is 1. The position is one-based.
'String1 The string to be searched in String2 The string to search for. Also known as needle.
'String2 The string to search for. Also known as needle.
'Condition	                                    InStr returns
'1. String2 is found within String1                 (Position where the first match begins (one-based))
'2. String2 is not found                            0
'3. String1 is empty                                0
'4. Start > length of String1                       0
'5. String2 is empty, but String1 is not empty      Start
'6. Start < 1                                       Exception. Results in runtime error if not handled.
	

   
   	pos = InStr(result,"WScriptOutputAsSuccess") 
    If pos > 0 Then
        resultOut=0
    Else  'for  "WScriptOutputAsFailed"
        resultOut=1
    End If
	
	
    WScript.Sleep 500
    wb.Close False 'Don't save workbook
    xl.Quit
    WScript.Echo result
    WScript.Quit(resultOut)
    Exit Sub

End Sub

Sub ShowError(strMessage)
    WScript.Echo strMessage
    WScript.Echo "Error#: " & Err.number  & " Src: " & Err.Source
    WScript.Echo "Description: " & Err.Description
    Err.Clear
    WScript.Quit(1)
    Exit Sub
 End Sub

Function findArg(ByVal ext)
    Dim fso
    Dim arg
    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each arg In WScript.Arguments
        If fso.GetExtensionName(LCase(arg)) = ext Then
            findArg = arg
            Exit Function
        End If
    Next
    WScript.Echo "No ." & ext & " file specified in input args!"
End Function
