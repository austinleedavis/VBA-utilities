Attribute VB_Name = "Installation"
Sub SaveAndReloadFromDisk()
    ExportAllModules
    ForceImportAllModules
End Sub


Sub ExportAllModules()
    Dim strPath, localPath As String
    Dim strFileName As String
    Dim cmp As VBComponent
    Dim intCount As Integer
    Dim objFSO As Object
    Dim objTextFile As Object
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    localPath = objFSO.GetParentFolderName(objFSO.GetAbsolutePathName(Application.ActiveWorkbook.Name))
    strPath = localPath & "\"
    Set objTextFile = objFSO.CreateTextFile(strPath & "modulelist.txt", True)

    intCount = 0

    For Each cmp In ThisWorkbook.VBProject.VBComponents
        If cmp.Type = vbext_ct_StdModule Then
            strFileName = strPath & cmp.Name & ".bas"
            cmp.Export strFileName
            objTextFile.WriteLine cmp.Name & ".bas"
            intCount = intCount + 1
        ElseIf cmp.Type = vbext_ct_ClassModule Then
            strFileName = strPath & cmp.Name & ".cls"
            cmp.Export strFileName
            objTextFile.WriteLine cmp.Name & ".cls"
            intCount = intCount + 1
        End If
    Next cmp

    objTextFile.Close
    Set objTextFile = Nothing
    Set objFSO = Nothing

    'MsgBox intCount & " modules exported successfully to " & strPath
End Sub

Sub ForceImportAllModules()
    Dim strPath As String
    Dim strFileName As String
    Dim cmp As VBComponent
    Dim intCount As Integer
    Dim objFSO As Object
    Dim objTextFile As Object
    Dim strModuleName As String
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    localPath = objFSO.GetParentFolderName(objFSO.GetAbsolutePathName(Application.ActiveWorkbook.Name))
    strPath = localPath & "\"

    Set objTextFile = objFSO.OpenTextFile(strPath & "modulelist.txt")
    intCount = 0

    Do Until objTextFile.AtEndOfStream
        strLine = objTextFile.ReadLine
        isBas = Right(strLine, 4) = ".bas"
        isCls = Right(strLine, 4) = ".cls"
        If isBas Or isCls Then
            For Each VBComponent In ThisWorkbook.VBProject.VBComponents
                If VBComponent.Name = left(strLine, Len(strLine) - 4) Then
                    VBComponent.Name = VBComponent.Name & "_REMOVED"
                    ThisWorkbook.VBProject.VBComponents.remove VBComponent
                    Exit For
                End If
            Next VBComponent
            
            If isBas Then
                ThisWorkbook.VBProject.VBComponents.Import strPath & strLine
            Else 'isCls
                Set objFSO = GetFileObject(strPath & strLine)
                ThisWorkbook.VBProject.VBComponents.add vbext_ct_ClassModule
                ThisWorkbook.VBProject.VBComponents("Class1").CodeModule.AddFromFile strPath & strLine
                ThisWorkbook.VBProject.VBComponents(left(strLine, Len(strLine) - 4)).CodeModule.DeleteLines 1, 4
            End If

        End If
    Loop
    
    objTextFile.Close
    Set objTextFile = Nothing
    Set objFSO = Nothing
End Sub

Private Function GetFileObject(filePath As String) As Object
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(filePath) Then
        Set GetFileObject = fso.getFile(filePath)
    Else
        MsgBox "The file " & filePath & " does not exist."
    End If
End Function

