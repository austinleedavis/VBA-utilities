Attribute VB_Name = "Installer"
'In order to install modules, access to the VBA project object model must be trusted.
'To enable go to:
' Options > Trust Center > TrustCenterSettings > Macro Settings > Trust access to the VBA project object model



Public Enum componentType
    ModuleType = 1
    ClassType = 2
End Enum

Public Sub installVBAComponents()
    Dim packageList As Collection
    Dim packagePath As Variant
    Dim vbcomp As Object
    Set packageList = getPackageList()
    Set vbcomp = Application.ThisWorkbook.VBProject.VBComponents
    Dim compCount As Long
    
    
    For Each packagePath In packageList
        vbcomp.Import packagePath
        
        Debug.Print "Imported: " & packagePath
    Next packagePath
    
    'err.Raise 10022, "InstallerProject.InstallModules", _
            "In order to install modules, access to the VBA project object model must be trusted." & _
            vbNewLine & vbNewLine & _
            "To enable:" & vbNewLine & _
            "Options > Trust Center > TrustCenterSettings > Macro Settings > " & vbNewLine & _
            "Trust access to the VBA project object model"
    
End Sub


Public Sub exportVBAComponents()
    Dim appPath As String
    Dim vbcomps As Object
    Dim vbcomp As Object
    Dim path As String
    
    appPath = Application.ThisWorkbook.path & "\"
    Set vbcomps = Application.ThisWorkbook.VBProject.VBComponents
    
    For Each vbcomp In vbcomps
        path = ""
        If vbcomp.Type = componentType.ModuleType Then
            path = appPath & vbcomp.Name & ".bas"
        ElseIf vbcomp.Type = componentType.ClassType Then
            path = appPath & vbcomp.Name & ".cls"
        End If
        
        If path <> "" Then
            vbcomp.Export path
        End If
            
    Next vbcomp
End Sub

Public Sub removeVBAComponents()
    Dim vbcomps As Object
    Dim vbcomp As Object

    Set vbcomps = Application.ThisWorkbook.VBProject.VBComponents
    
    For Each vbcomp In vbcomps
        path = ""
        If (vbcomp.Type = componentType.ClassType Or _
        vbcomp.Type = componentType.ModuleType) And _
        vbcomp.Name <> "Installer" Then
            Debug.Print "Removing: " & vbcomp.Name
            vbcomps.remove vbcomp
        End If
            
    Next vbcomp
End Sub

Private Function getPackageList() As Collection
    Dim extensions As Collection
    Dim packages As Collection
    Dim installerPath As String
    Dim extension As Variant
    Dim varDirectory As Variant
    
    Set packages = New Collection
    Set extensions = New Collection
    extensions.add "*.cls"
    extensions.add "*.bas"
    
    'path to the installer xlam
    installerPath = Application.ThisWorkbook.path & "\"
    Debug.Print "Installer Located at: " & installerPath
    
    For Each extension In extensions
        varDirectory = Dir(installerPath & extension)
        Do While True
            If varDirectory = "" Then
                Exit Do
            End If
            packages.add (installerPath & varDirectory)
            varDirectory = Dir
        Loop
    Next extension
    
    Debug.Print "Found " & packages.Count & " packages."
    
    Set getPackageList = packages
    
End Function
