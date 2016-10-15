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
'		10/15/2016 - v1.0	Initial creation of the script by Seinu
'
'*****************************************************************************'
Option Explicit

'Classes
Class File
	Public strPath
	Public strData
End Class

'Globals
Dim oExt, oFSO, o_rFolder, oFile, fListTmp
Dim fList()

'Create File System Object
Set oFSO = CreateObject("Scripting.FileSystemObject")

'Get and Set Root Folder
Set o_rFolder = oFSO.GetFolder(oFSO.GetAbsolutePathName("."))

	
Set oExt = CreateObject("Scripting.Dictionary")
oExt.CompareMode = vbTextCompare
oExt.Add "doc", True
oExt.Add "docx", True
oExt.Add "pdf", True
oExt.Add "xlsx", True
oExt.Add "pptx", True
oExt.Add "ppsx", True

Dim i, j
i = -1

For Each oFile in o_rFolder.Files
	If oExt.Exists(oFSO.GetExtensionName(oFile.Name)) Then
		i = i + 1
		Dim oRead, strTemp
		ReDim Preserve fList(i)
		Set fList(i) = New File
	
		fList(i).strPath = oFile.Path
	
		Set oRead = oFSO.OpenTextFile(oFile.Path, 1)
		strTemp = oRead.ReadAll
		oRead.Close
		Set oRead = Nothing
		fList(i).strData = strTemp
	End If
Next

Set fListTmp = oFSO.CreateTextFile(o_rFolder & "\fListTmp.txt")

For i = 0 To UBound(fList)
	If Not fList(i) Is Nothing Then
		For j = i + 1 To UBound(fList)
			If not fList(i) Is Nothing Then
				If not fList(j) Is Nothing Then
					If (fList(i).strData = fList(j).strData) And (oFSO.GetExtensionName(fList(i).strPath) = oFSO.GetExtensionName(fList(j).strPath)) Then
						MsgBox "(" & fList(j).strPath & ") duplicates (" & fList(i).strPath & ")"
						
						oFSO.DeleteFile fList(j).strPath
						Set fList(j) = Nothing
					End If
				End If
			End If
		Next
		fListTmp.WriteLine(fList(i).strPath)
	End If
Next

fListTmp.Close

Dim strContents,strFolder, flKeep()

Set fListTmp = oFSO.OpenTextFile(o_rFolder & "\fListTmp.txt")
i = -1
Do Until fListTmp.AtEndOfStream
	i = i + 1
	ReDim Preserve flKeep(i)
	flKeep(i) = fListTmp.ReadLine
	MsgBox flKeep(i)
Loop

strFolder = o_rFolder& "\organized\"
if not oFSO.FolderExists(strFolder) Then
	oFSO.CreateFolder o_rFolder& "\organized"
End If
For i = 0 To UBound(flKeep)
	MsgBox flKeep(i)
	oFSO.MoveFile flKeep(i), strFolder
Next	