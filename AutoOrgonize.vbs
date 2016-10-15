'*****************************************************************************'
'
'	This program is free software: you can redistribute it and/or modify
'	it under the terms of the GNU General Public License as published by
'	the Free Software Foundation, either version 3 of the License, or
'	(at your option) any later version.
'
'	This program is distributed in the hope that it will be useful,
'	but WITHOUT ANY WARRANTY; without even the implied warranty of
'	MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'	GNU General Public License for more details.
'
'	You should have received a copy of the GNU General Public License
'	along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
'*****************************************************************************'
'	File Name:	AutoOrgonize.vbs
'	Authors:	Seinu
'	Purpose:	Auto Organize Files in script directory and sub directory
'	History:
'		10/15/2016 - v1.0a	Initial creation of the script by Seinu
'		10/15/2016 - v1.0b	added comments
'		10/15/2016 - v1.1	Refactored the code
'
'*****************************************************************************'
Option Explicit

'Classes
Class File
	Public strPath
	Public strData
End Class

'Globals
Dim oExt, oFSO, oFile, ExtList, SubFldr
Dim fList()
Dim i, j, b_Fldr

'Create File System Object
Set oFSO = CreateObject("Scripting.FileSystemObject")

'Get and Set Root Folder
Dim o_rFolder
Set o_rFolder = oFSO.GetFolder(oFSO.GetAbsolutePathName("."))

'Check for arguments 1 and/or 2
if WScript.Arguments.Count <= 2 And Not WScript.Arguments.Count <= 0 then
	if WScript.Arguments.Count = 2 Then
		ExtList = Wscript.Arguments(0)
		SubFldr = Wscript.Arguments(1)
		Call GetFileExt(ExtList, True)
	ElseIf WScript.Arguments.Count = 1 Then
		Call GetFileExt(Wscript.Arguments(0), True)
		SubFldr = "organized"
		b_Fldr = False
	End If
'if no arguments
ElseIf WScript.Arguments.Count = 0 Then
	Call GetFileExt(ExtList, False)
	b_Fldr = False
End If

Call GetFileList()
Call CheckFileDupes()
Call MoveFiles(SubFldr, b_Fldr)
'*****************************************************************************'
Sub GetFileExt(p_ExtList, b_ListExist)
	if b_ListExist = True Then
		Set oExt = CreateObject("Scripting.Dictionary")
		oExt.CompareMode = vbTextCompare
		Dim ExtArray
		ExtArray = GetLineArray(p_ExtList)
		For i = 0 To UBound(ExtArray)
			oExt.Add ExtArray(i), True
		Next
	ElseIf b_ListExist = False Then
		Set oExt = CreateObject("Scripting.Dictionary")
		oExt.CompareMode = vbTextCompare
		oExt.Add "doc", True
		oExt.Add "docx", True
		oExt.Add "pdf", True
		oExt.Add "xlsx", True
		oExt.Add "pptx", True
		oExt.Add "ppsx", True
		oExt.Add "rtf", True
		oExt.Add "txt", True
	End If
End Sub
'*****************************************************************************'
Sub GetFileList()
	i = -1
	For Each oFile in o_rFolder.Files
		If oExt.Exists(oFSO.GetExtensionName(oFile.Name)) Then
			i = i + 1
			Dim oRead
			ReDim Preserve fList(i) 'Dynamically add to array
			Set fList(i) = New File 'Set Array index (i) to the Class File
			fList(i).strPath = oFile.Path 'get file paths and set to the array
			'get file contents and save to the array
			Set oRead = oFSO.OpenTextFile(oFile.Path, 1)
			fList(i).strData = oRead.ReadAll
			oRead.Close
			'close file
			Set oRead = Nothing
		End If
	Next
End Sub
'*****************************************************************************'
Sub CheckFileDupes()
	'Create temporary text file to hold non duplicate file names
	Dim fListTmp
	Set fListTmp = oFSO.CreateTextFile(o_rFolder & "\fListTmp.txt")
	
	'loop through each file and check for duplicates
	For i = 0 To UBound(fList)
		If Not fList(i) Is Nothing Then
			For j = i + 1 To UBound(fList)
				If not fList(i) Is Nothing Then
					If not fList(j) Is Nothing Then
						If (fList(i).strData = fList(j).strData) And (oFSO.GetExtensionName(fList(i).strPath) = oFSO.GetExtensionName(fList(j).strPath)) Then
							'MsgBox "(" & fList(j).strPath & ") duplicates (" & fList(i).strPath & ")"
						
							oFSO.DeleteFile fList(j).strPath
							Set fList(j) = Nothing
						End If
					End If
				End If
			Next
			fListTmp.WriteLine(fList(i).strPath) 'write non duplicate file to \fListTmp.txt
		End If
	Next
	fListTmp.Close
End Sub
'*****************************************************************************'
Sub MoveFiles(N_subfolder, b_fldr)
	'Locals
	Dim strContents,strFolder, flKeep
	
	'check that a sub folder name was given
	if b_Fldr = True Then
		strFolder = o_rFolder & "\" & N_subfolder & "\"'create subfolder string
	ElseIf b_Fldr = false Then
		strFolder = o_rFolder & "\organized\"
	End If
	
	'Set flKeep to File List Line Array
	flKeep = GetLineArray(o_rFolder & "\fListTmp.txt")

	'check that folder exists if not create subfolder
	if not oFSO.FolderExists(strFolder) Then
		oFSO.CreateFolder o_rFolder & "\organized" 'create subfolder \organized\
	End If
	
	'move files to new subfolder
	For i = 0 To UBound(flKeep)
		'check that the same files don't already exist
		if oFSO.FileExists(flKeep(i)) Then
			'MsgBox flKeep(i)
			oFSO.MoveFile flKeep(i), strFolder
		End If
	Next
	oFSO.DeleteFile(o_rFolder & "\fListTmp.txt")
End Sub
'*****************************************************************************'
function GetLineArray(path)
	Dim InputFile, fLine()
	i = -1
	Set InputFile = oFSO.OpenTextFile(path)
	Do Until InputFile.AtEndOfStream
		i = i + 1
		ReDim Preserve fLine(i)
		fLine(i) = InputFile.ReadLine
		'MsgBox flKeep(i)
	Loop
	GetLineArray = fLine
End Function
'*****************************************************************************'