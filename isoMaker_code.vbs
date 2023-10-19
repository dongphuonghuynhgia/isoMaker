'---------- SUPPORT FUNCTION ----------------------------------
'----- FOLDER BROWSE DIALOG -----------------------------------
Function getFolder(title, isPC, newFolder, showFiles)
	Set oShell = CreateObject("Shell.Application")
	Dim vPC, vOpt
	If isPC Then vPC=0 Else vPC=17
	vOpt = 0
	If (Not newFolder) Then vOpt = vOpt + 512
	If showFiles Then vOpt = vOpt + 16384
	Set dDlg = oShell.BrowseForFolder(0, title, vOpt, vPC)
	If (Not dDlg Is Nothing) Then
		tPath=dDlg.items.item.path
		If Left(tPath,2)="::" Then
			getFolder=""
		Else
			getFolder=tPath
		End If
	Else
		getFolder=""
	End If
End Function
'----- OPEN FILE DIALOG ---------------------------------------
Function getFile(fTitle,fFilter)
	Set FSO = CreateObject("Scripting.FileSystemObject")
	sIniDir = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"'FSO.GetSpecialFolder(Desktop)
	Set oShell = CreateObject("WScript.Shell").Exec("mshta.exe ""about:<object id=d classid=clsid:3050f4e1-98b5-11cf-bb82-00aa00bdce0b></object><script>moveTo(0,-9999);eval(new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(0).Read("&Len(sInitDir)+Len(fFilter)+Len(fTitle)+41&"));function window.onload(){var p=/[^\16]*/;new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).Write(p.exec(d.object.savefiledlg(iniDir,null,filter,title)));close();}</script><hta:application showintaskbar=no />""")
	oShell.StdIn.Write "var iniDir='" & sInitDir & "';var filter='" & fFilter & "';var title='" & fTitle & "';"
	res=oShell.StdOut.ReadAll
	If res<>"" Then
		fInd=Instr(1,res,Chr(0))
		getFile = Left(res,fInd-1)
	Else
		getFile=""
	End If
End Function
'----- CHECK FILE NAME ----------------------------------------
Function checkName(name)
	checkName=((InStr(1, name, Chr(&H2F)) + InStr(1, name, Chr(&H5C)) + InStr(1, name, Chr(&H3A)) +_
				InStr(1, name, Chr(&H2A)) + InStr(1, name, Chr(&H3F)) + InStr(1, name, Chr(&H22)) +_
				InStr(1, name, Chr(&H3C)) + InStr(1, name, Chr(&H3E)) + InStr(1, name, Chr(&H7C)))=0)
End Function
'----- GET COMMAND OUTPUT -------------------------------------
Function readCmd(cmdLine)
	Set objShell = CreateObject("WScript.Shell")
	Set objExecObject = objShell.Exec("cmd /c " & cmdLine)
	strText = ""
	Do While Not objExecObject.StdOut.AtEndOfStream
		strText = strText & chr(13) & objExecObject.StdOut.ReadLine()
	Loop
	readCmd=strText
End Function
'--------------------------------------------------------------

'---------- ALL CONTROLS EVENTS -------------------------------
Sub formInit()
	Dim frmW, frmH: frmW=400: frmH=200
	window.resizeTo frmW, frmH
	window.moveTo screen.width/2 - frmW/2, screen.height/2 - frmH/2
End Sub

Sub selectDir()
	Dim dirText:dirText = document.getElementById("selDir").innerText
	If dirText="Double click to select captured folder" Then 
		Dim getDir: getDir=getFolder("Select folder to capture", true, false, false)
		If getDir<>"" Then 
			document.getElementById("selDir").innerText=getDir
		Else
			MsgBox "User cancel :)", 64, "Info"
		End If
	Else
		document.getElementById("selDir").innerText="Double click to select captured folder"
	End If
End Sub

Sub saveIso()
	Dim defName: defName="New ISO file"
	Set fso=CreateObject("Scripting.FileSystemObject")
	Dim capDir: capDir=document.getElementById("selDir").innerText
	If fso.FolderExists(capDir) Then defName=fso.GetBaseName(capDir)
	Dim isoText: isoText = document.getElementById("selIso").innerText
	If isoText="Double click to save ISO file" Then
		Dim isoPath: isoPath=getFolder("Select folder to save ISO file", true, true, false)
		If isoPath<>"" Then
			Dim isoFile: Dim condFile: condFile=False
			Do While (Not condFile)
				isoFile=InputBox("Enter your ISO file name", "Enter name", defName)
				If isoFile="" Then
					MsgBox "User cancel :)", 64, "Info"
					condFile=True
					Exit Sub
				Else
					If checkName(isoFile) Then
						condFile=True
					Else
						MsgBox "Invalid file name!", 16, "Warning"
						condFile=False
					End If
				End If
			Loop
			document.getElementById("selIso").innerText=isoPath & "\" & isoFile & ".iso"
		Else
			MsgBox "User cancel :)", 32, "Info"
		End If
	Else
		document.getElementById("selIso").innerText="Double click to save ISO file"
	End If
End Sub

Sub makeIso()
	Set fso=CreateObject("Scripting.FileSystemObject")
	Dim selDir: selDir=document.getElementById("selDir").innerText
	Dim savIso: savIso=document.getElementById("selIso").innerText
	If fso.FolderExists(selDir) Then
		cmdStr="oscdimg.exe -n -m -d -h -l" & Trim(fso.GetBaseName(selDir)) & " " & Chr(34) & selDir & Chr(34) & " " & Chr(34) & savIso & Chr(34)
		Dim makIso: makIso = readCmd(cmdStr)
		'MsgBox cmdStr, 32, "Info"
	Else
		MsgBox "Select folder to capture, please :)", 48, "Warning"
		selectDir()
	End If
End Sub
'---------- END OF FILE ---------------------------------------
