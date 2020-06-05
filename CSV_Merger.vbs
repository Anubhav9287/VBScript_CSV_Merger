Dim objExcel, objWorkbook, wbSrc
Dim strFileName, strDirectory, extension, Filename
Dim objFSO, objFolder, objFile
Dim strPath
Dim output_path 

strPath = SelectFolder( "" )
If strPath = vbNull Then
    WScript.Echo "Cancelled"
Else
    WScript.Echo "Selected Folder: """ & strPath & """"
End If

Function SelectFolder( myStartFolder )
    Dim objFolder, objItem, objShell
    SelectFolder = vbNull

    ' Create a dialog object
    Set objShell  = CreateObject( "Shell.Application" )
    Set objFolder = objShell.BrowseForFolder( 0, "Select Folder", 0, myStartFolder )

    ' Return the path of the selected folder
    If IsObject( objfolder ) Then SelectFolder = objFolder.Self.Path

    ' Standard housekeeping
    Set objFolder = Nothing
    Set objshell  = Nothing
    On Error Goto 0
End Function
Set WshShell = CreateObject("WScript.Shell")
strCurDir    = WshShell.CurrentDirectory
output_path =  strCurDir & "\merger_output.xlsx"
Set WshShell = Nothing
Set obj = CreateObject("Scripting.FileSystemObject")
If obj.FileExists(output_path ) Then
obj.DeleteFile(output_path )
End If

strFileName = output_path 

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False

Set objWorkbook = objExcel.Workbooks.Add()

extension = "csv"

strDirectory = strPath

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(strDirectory)

For Each objFile In objFolder.Files
    If LCase((objFSO.GetExtensionName(objFile))) = LCase(extension) Then
        Filename = objFile.Name
        Filename = strDirectory & "\" & Filename
        Set wbSrc = objExcel.Workbooks.Open(Filename)
        wbSrc.Sheets(1).Copy objWorkbook.Sheets(objWorkbook.Sheets.Count)
        wbSrc.Close
    End If
Next

'~~> Close and Cleanup
objWorkbook.SaveAs (strFileName)
objWorkbook.Close
objExcel.Quit

Set wbSrc = Nothing
MsgBox "Merger Completed!! Please check your file."