Attribute VB_Name = "Enable_Trust_VBA_objectModel"
Option Explicit

'============================================================================
' Note: Using the code provided below is at your own risk.
'
'============================================================================

Sub CheckIfVBAAccessIsOn()

' Excel Registry Settings :[HKEY_LOCAL_MACHINE/Software/Microsoft/Office/10.0/Excel/Security]
'                           "AccessVBOM"=dword:00000001

Dim strRegPath As String
strRegPath = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & Application.Version & "\Excel\Security\AccessVBOM"

If TestKeyValue(strRegPath) = 0 Then
    MsgBox "A change has been introduced into your registry configuration. The Excel application will now restart."
   
    WriteVBS
End If


Dim VBProj    As Object     'VBIDE.VBProject
Dim VBComp    As Object     'VBIDE.VBComponent
Dim Counter As Long

Set VBProj = ThisWorkbook.VBProject

' loop through all References and show them in the Immediate window
For Counter = 1 To VBProj.References.Count
    Debug.Print VBProj.References(Counter).FullPath
    'Debug.Print VBProj.References(counter).Name
    Debug.Print VBProj.References(Counter).Description
    Debug.Print "—————————————————"
Next

End Sub
 
Function TestKeyValue(ByVal path As String) As Integer

Dim WSHShell As Object
Set WSHShell = CreateObject("WScript.Shell")

On Error Resume Next
WSHShell.regread path

If Err.Number <> 0 Then
   Err.Clear
   TestKeyValue = ""
Else
   TestKeyValue = CInt(WSHShell.regread(path))
End If
On Error GoTo 0
 
End Function

Sub WriteVBS()

' This sub-routine creaes a new VBScript file, writes to it
' then it runs the VBScript
' The VBScript closes the Excel Application (in order for the check-mark effect to take effect)
' Then it write to the Excel Setting registry the value for "Trust access to the VBA object model"
' afterwards it will delete the VBScipt file that was created from the Desktop

Dim wsh As Object
Dim objFile     As Object
Dim objFSO      As Object
Dim codePath    As String
Dim waitOnReturn As Boolean, windowStyle As Integer
 
codePath = "C:\Users\" & Environ$("USERNAME") & "\Desktop\Excel_Reg_Setting.vbs"
'codePath = ThisWorkbook.path & "\RestartExcel.vbs"

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(codePath, 8, True)

' --- Headers and Calling Subs Section ---
objFile.WriteLine ("Option Explicit")
objFile.WriteLine ("")
objFile.WriteLine ("TrustAccessVBProjectModule")
objFile.WriteLine ("KillProcesses")
objFile.WriteLine ("ExcelRestart")
objFile.WriteLine ("WScript.Quit(0)")
objFile.WriteLine ("")
objFile.WriteLine ("")

' --- Sub ExcelRestart Section (open Excel and this Workbook after shutting down Excel) ---
objFile.WriteLine ("Sub ExcelRestart()")
objFile.WriteLine ("")
objFile.WriteLine ("Dim xlApp")
objFile.WriteLine ("Dim xlBook")
objFile.WriteLine ("")
objFile.WriteLine ("Set xlApp = CreateObject(""Excel.Application"")")
objFile.WriteLine ("")
objFile.WriteLine ("xlApp.DisplayAlerts = False")
'objFile.WriteLine ("xlApp.Application.Run " & Chr(34) & "'" & ThisWorkbook.FullName & "'!Module1.CheckIfVBAAccessIsOn" & Chr(34))
objFile.WriteLine ("xlApp.Application.Run " & Chr(34) & "'" & ThisWorkbook.FullName & "'!Module1.Test" & Chr(34) & ",1, True")
objFile.WriteLine ("KillProcesses")
objFile.WriteLine ("Set xlApp = CreateObject(""Excel.Application"")")
objFile.WriteLine ("")
objFile.WriteLine ("xlApp.DisplayAlerts = False")
'
objFile.WriteLine ("xlApp.Workbooks.Open(" & Chr(34) & ThisWorkbook.FullName & Chr(34) & ")")
objFile.WriteLine ("")
' TEST
'objFile.WriteLine ("xlApp.Run ""CheckIfVBAAccessIsOn""")
objFile.WriteLine ("xlApp.visible = True")
objFile.WriteLine ("xlBook.activate")
objFile.WriteLine ("Set xlBook = Nothing")
objFile.WriteLine ("Set xlApp = Nothing")
objFile.WriteLine ("")
objFile.WriteLine ("End sub")
objFile.WriteLine ("")

' --- Sub KillProcesses Section (Closes Excel) ---
objFile.WriteLine ("Sub KillProcesses()")
objFile.WriteLine ("")
objFile.WriteLine ("On error resume next")
objFile.WriteLine ("")
objFile.WriteLine ("Dim objWMIService, WshShell")
objFile.WriteLine ("Dim proc, procList")
objFile.WriteLine ("Dim strComputer, strCommand")
objFile.WriteLine ("")
objFile.WriteLine ("strCommand = ""taskkill /F /IM excel.exe""")
objFile.WriteLine ("strComputer = "".""")
objFile.WriteLine ("")
objFile.WriteLine ("Set WshShell = WScript.CreateObject(""WScript.Shell"") ")
objFile.WriteLine ("Set objWMIService = GetObject(""winmgmts:""& ""{impersonationLevel=impersonate}!\\""& strComputer & ""\root\cimv2"")")
objFile.WriteLine ("Set procList = objWMIService.ExecQuery(""SELECT * FROM Win32_Process WHERE Name = 'excel.exe'"")")
objFile.WriteLine ("For Each proc In procList")
objFile.WriteLine ("   WshShell.run strCommand, 0, TRUE")
objFile.WriteLine ("Next")
objFile.WriteLine ("")
objFile.WriteLine ("End sub")
objFile.WriteLine ("")

' --- Sub TrustAccessVBProjectModule Section (moidfy the Registry setting) ---
objFile.WriteLine ("Sub TrustAccessVBProjectModule()")
objFile.WriteLine ("")
objFile.WriteLine ("On Error Resume Next")
objFile.WriteLine ("")
objFile.WriteLine ("Dim WshShell")
objFile.WriteLine ("Set WshShell = CreateObject(""WScript.Shell"")")
objFile.WriteLine ("")
'objFile.WriteLine ("MsgBox ""Click OK to complete the setup process.""")
'objFile.WriteLine ("")
objFile.WriteLine ("Dim strRegPath")
objFile.WriteLine ("Dim Application_Version")
objFile.WriteLine ("Application_Version = """ & Application.Version & """")
objFile.WriteLine ("strRegPath = ""HKEY_CURRENT_USER\Software\Microsoft\Office\"" & Application_Version & ""\Excel\Security\AccessVBOM""")
objFile.WriteLine ("WScript.echo strRegPath")
objFile.WriteLine ("WshShell.RegWrite strRegPath, 1, ""REG_DWORD""")
objFile.WriteLine ("")
objFile.WriteLine ("If Err.Code <> 0 Then")
objFile.WriteLine ("   MsgBox ""Error"" & Chr(13) & Chr(10) & Err.Source & Chr(13) & Chr(10) & Err.Message")
objFile.WriteLine ("End If")
objFile.WriteLine ("")
objFile.WriteLine ("MsgBox ""Successful Excel Registry edit""")
objFile.WriteLine ("End sub")

objFile.Close
Set objFile = Nothing
Set objFSO = Nothing

' --- Run Shell section ---
Set wsh = VBA.CreateObject("WScript.Shell")
   
waitOnReturn = True
windowStyle = 1

' run Shell command to modify the Excel registry
' wait for Shell to finish before going to the next line, to avoid getting errors on checking the VBProject References
wsh.Run "cscript " & Chr(34) & codePath & Chr(34), windowStyle, waitOnReturn
'
' To fix this issue, we add a pair of double quotes (" ") around [codepath]
'Shell "cscript " & Chr(34) & codePath & Chr(34), vbNormalFocus

' --- Delete the VB Script file after it ran ---

' check if VBS file name exits in folder
If Dir(codePath) <> "" Then
   ' remove read-only attribute (if set)
   SetAttr codePath, vbNormal
   '  delete the file
   Kill codePath
End If

End Sub
