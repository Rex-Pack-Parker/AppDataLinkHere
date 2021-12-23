
'需要设置你的转接文件夹(绝对路径,并带\号结尾) 空则以该工具所在的文件夹下创建
Dim NewPath:NewPath = ""




Set WSS = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")
Set SD  = CreateObject("Scripting.Dictionary")
Set SA = CreateObject("Shell.Application")

Dim NowSHE:NowSHE = LCase(USplit(WScript.FullName, "\"))

If Not NowSHE = "cscript.exe" Then
	Const EXE = "CMD"
	Const Color = "Color F0"
	Const Title = "Title AppDataLinkHere"
	Dim CMS:CMS = "CScript //nologo " & PathC34(WSH.ScriptFullName)
	SA.ShellExecute EXE, " /k " & Color & " & " & Title & " & " & CMS, "", "runas", ""
	Call EndVBS
End If

Set STD = New BaseSTD

Set SES_P = WSS.Environment("Process")

Dim ScriptPath:ScriptPath = Replace(WSH.ScriptFullName, WSH.ScriptName, "")
Dim RootPath:RootPath = SES_P.Item("USERPROFILE") & "\AppData\"
Dim MidPath:MidPath = ""
Dim TargetPath:TargetPath = ""

'----
If NewPath = "" Then NewPath = ScriptPath & "AppData\"

If Not FSO.FolderExists(NewPath) Then SA.NameSpace(PathCC(ScriptPath)).NewFolder "AppData"
If Not FSO.FolderExists(NewPath & "Local") Then SA.NameSpace(PathCC(NewPath)).NewFolder "Local"
If Not FSO.FolderExists(NewPath & "LocalLow") Then SA.NameSpace(PathCC(NewPath)).NewFolder "LocalLow"
If Not FSO.FolderExists(NewPath & "Roaming") Then SA.NameSpace(PathCC(NewPath)).NewFolder "Roaming"
With STD
	.PrintL "	"
	.PrintL "	[AppData文件夹转接工具]"
	.PrintL "	 选择分支 > 选择文件夹序号 = 转接"
	.PrintL "	 存储路径可修改本工具代码最上方 NewPath 右侧值"
	.PrintL "	----"
	.PrintL "	AppData路径: " & RootPath
	.PrintL "	存储路径: " & NewPath
	.PrintL "	----"
End With

Call Main
'----
Sub Main()
	STD.PrintL "	[方式]"
	STD.PrintL "	1	设定链接到源"
	STD.PrintL "	2	源文件夹里所有文件夹链接到AppData"
	STD.PrintI "	选择> "
	Dim CM:CM = STD.Input
	If Not CM = "" Then
		Select Case CM
		Case 1
			Call SelectMidFolder
			Call SelectFolder
		Case 2
			Call MKSs
		'Case 3
		'	Call SelectMidFolder
		End Select
	End If
	Call Main
End Sub

'----0s
Sub SelectMidFolder()
	STD.PrintL "	[分支]"
	Set GEOM_TFO = FSO.GetFolder(RootPath).SubFolders
	Dim I:I = 0
	SD.RemoveAll
	For Each GEOM_TempDataVar In GEOM_TFO
		SD.Add CStr(I), GEOM_TempDataVar.Name
		STD.PrintL "	" & I & "	" & GEOM_TempDataVar.Name
		I = I + 1
	Next
	
	STD.PrintI "	选择> "
	Dim CM:CM = STD.Input
	If Not CM = "" Then
		If SD.Exists(CM) Then
			MidPath = SD.Item(CM) & "\"
		Else 
			STD.PrintL "错误."
		End If
	Else
		
	End If
End Sub

'----1s
Sub SelectFolder()
	STD.PrintL "	[目标]"
	Set GEOM_TFO = FSO.GetFolder(RootPath & MidPath).SubFolders
	Dim I:I = 0
	SD.RemoveAll
	For Each GEOM_TempDataVar In GEOM_TFO
		Select Case GEOM_TempDataVar.Name
		Case "History", "Application Data", "Temporary Internet Files"
		Case Else
			SD.Add CStr(I), GEOM_TempDataVar.Name
			STD.PrintL "	" & IIf(SA.NameSpace(PathCC(GEOM_TempDataVar.Path)).Self.IsLink, "已链接 ", "       ") & I & "	" & GEOM_TempDataVar.Name
			I = I + 1
		End Select
	Next
	STD.PrintI "	选择> "
	Dim CM:CM = STD.Input
	If Not CM = "" Then
		If SD.Exists(CM) Then
			TargetPath = SD.Item(CM)
			Call MKlink
			Call SelectFolder
		Else 
			STD.PrintL "错误."
		End If
	Else
		Call Main
	End If
End Sub

'----2s
Sub MKSs()
	For Each NowMidPath In Array("Local\", "LocalLow\", "Roaming\")
		MidPath = NowMidPath
		For Each NowTargetPath In FSO.GetFolder(NewPath & NowMidPath).SubFolders
			TargetPath = NowTargetPath.Name
			Set TempExec = Exec("cmd /c rd /q " & """" & RootPath & MidPath & TargetPath & """")
			'STD.PrintL TempExec.StdOut.ReadAll
			Call MKlink
		Next
	Next
End Sub

Sub MKS()
	
End Sub


'
'----
Sub MKlink()
	Dim TempVal:TempVal = "/c mklink /j """ & RootPath & MidPath & TargetPath & """ """ & NewPath & MidPath & TargetPath & """"
	STD.PrintL "	目标: " & PathC34(RootPath & MidPath & TargetPath)
	STD.PrintL "	转接: " & PathC34(NewPath & MidPath & TargetPath)
	Call SA.NameSpace(PathCC(NewPath & MidPath)).MoveHere(PathCC(RootPath & MidPath & TargetPath))
	SA.ShellExecute "CMD", TempVal
	WSH.Sleep 1000
End Sub

Function Exec(ByVal ProgramPath)
	Set ExeCM = WSS.Exec(ProgramPath)
	Set Exec = ExeCM
 End Function
Function IIf(ByVal TF, ByVal T, ByVal F) '标准IIf函数
	If TF = True Then IIf = T Else IIf = F
 End Function
Function USplit(ByVal GEOM_String, ByVal GEOM_Delimiter)
	GEOM_TempArray = Split(GEOM_String, GEOM_Delimiter)
	USplit = GEOM_TempArray(UBound(GEOM_TempArray))
 End Function

Function PathC34(ByVal FP) '判断路径中是否含有空格,没有则去除"号 有则添加上""
	If InStr(FP, " ") = 0 Then
		PathC34 = Replace(FP, Chr(34), "")
	Else
		PathC34 = IIf(Left(FP, 1) = Chr(34), "", Chr(34)) & FP & IIf(Right(FP, 1) = Chr(34), "", Chr(34))
	End If
 End Function

Function PathCC(ByVal FP) '去除路径的""号,无论有没有空格
	PathCC = Replace(FP, Chr(34), "")
 End Function

Sub EndVBS() '退出脚本
	WScript.Quit
 End Sub

Function IIf(ByVal TF, ByVal T, ByVal F) '标准IIf函数
	If TF = True Then IIf = T Else IIf = F
 End Function

Class BaseSTD 'CScript I\O
	Sub PrintI(ByVal Texts)
		WScript.StdOut.Write Texts
	End Sub
	
	Sub PrintC(ByVal Texts, ByVal LenNum)
		WScript.StdOut.Write Chr(13) & Texts & String(LenNum, " ")
	End Sub
	
	Sub PrintL(ByVal Texts)
		WScript.StdOut.WriteLine(Texts)
	End Sub
	
	Function InPut()
		InPut = Trim(WScript.StdIn.ReadLine)
	End Function
 End Class
