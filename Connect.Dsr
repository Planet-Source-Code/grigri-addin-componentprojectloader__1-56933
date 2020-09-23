VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   7185
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   10050
   _ExtentX        =   17727
   _ExtentY        =   12674
   _Version        =   393216
   Description     =   $"Connect.dsx":0000
   DisplayName     =   "ComponentProjectLoader"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private VBInstance As VBIDE.VBE

Private WithEvents PrjEvts As VBProjectsEvents
Attribute PrjEvts.VB_VarHelpID = -1

' ······························
' · AddinInstance_OnConnection ·
' ······························
'
' This method is called when VB loads the add-in
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo error_handler
    
    'save the vb instance
    Set VBInstance = Application
    
    ' sink the event handler
    Set PrjEvts = VBInstance.Events.VBProjectsEvents
    
    Exit Sub
    
error_handler:
    
    MsgBox Err.Description
    
End Sub

' ·····················
' · PrjEvts_ItemAdded ·
' ·····················
'
' This method is called when a project has been added the the VB IDE.
'
Private Sub PrjEvts_ItemAdded(ByVal VBProject As VBIDE.VBProject)
    ' Project has been added
    
    ' Is it unnamed?
    If Len(VBProject.FileName) <> 0 Then Exit Sub
    ' Does it have exactly one component?
    If VBProject.VBComponents.Count <> 1 Then Exit Sub
    ' Does that component have a file name?
    If Len(VBProject.VBComponents(1).FileNames(1)) = 0 Then Exit Sub
    
    ' The project has been created for an existing component
    ' so try and find the project[s] that use it
    
    Dim colProjects As New Collection           ' To hold any projects we find
    Dim sCompPath As String                     ' Full path of component
    Dim sDirPath As String                      ' Directory path of component
    Dim sProjectPath As String                  ' Full path of the new project to load
        
    ' VB has a nasty habit of using 8.3 filenames
    ' Convert to the long format...
    sCompPath = LongPathFromShort(VBProject.VBComponents(1).FileNames(1))
    ' Get the parent directory
    sDirPath = ParentDirectory(sCompPath)
    
    ' List all projects containing the component
    FindProjectsForComponent sCompPath, sDirPath, colProjects
    
    If colProjects.Count = 0 Then
        ' No projects were found - boo hoo :(
        sProjectPath = vbNullString
    ElseIf colProjects.Count = 1 Then
        ' Yippee! Exactly one project found
        ' We'll load that one
        sProjectPath = colProjects.Item(1)
    Else
        ' Eek! Several projects found
        ' Best to ask the user what to do...
        sProjectPath = ProjectChooser.SelectProject(sCompPath, colProjects)
    End If
    
    ' At this stage, if we have a project to load then its full path
    ' Will be in sProjectPath. If we don't, it'll be a null string.
    If Len(sProjectPath) <> 0 Then
        Dim cmps As VBComponents, cmp As VBComponent
        ' Delete the auto-generated project
        VBInstance.VBProjects.Remove VBProject
        
        ' Add the project we've found to the workspace and get its components
        Set cmps = VBInstance.VBProjects.AddFromFile(sProjectPath, True).Item(1).VBComponents
        
        ' Now iterate through the components to find the one that was loaded first
        For Each cmp In cmps
            If cmp.FileNames(1) = sCompPath Then
                ' Found it! Activate it...
                cmp.Activate
                ' No need to hang around...
                Exit For
            End If
        Next
    End If
End Sub

' ······························
' · FindProjectsForComponent() ·
' ······························
'
' Searches local paths for project files containing a reference to a given component,
' Returns all matching project files via the colProjects member
'
Private Sub FindProjectsForComponent(ByVal sCmpName As String, ByVal sDirPath As String, colProjects As Collection, Optional ByVal DoParent As Boolean = True)
    Dim sProjFile As String
    On Error Resume Next
    
    ' Start by searching in the same directory as the component file
    sProjFile = Dir$(sDirPath & "*.vbp")
    While Len(sProjFile) <> 0
        If ProjectContainsComponent(sDirPath, sProjFile, sCmpName) Then
            colProjects.Add sDirPath & sProjFile
        End If
        sProjFile = Dir$()
    Wend
    If colProjects.Count > 0 Then Exit Sub
    
'    ' Search subdirectories
'    Dim sSubDir As String
'    sSubDir = Dir$(sDirPath, vbDirectory)
'    While Len(sSubDir) <> 0
'        If (GetAttr(sDirPath & sSubDir) And vbDirectory) = vbDirectory Then
'            If sSubDir <> "." And sSubDir <> ".." Then
'                ' ERROR WITH RECURSING
'                ' DAMN DIR$() DOESN'T WORK PROPERLY
'                FindProjectsForComponent sCmpName, sDirPath & sSubDir & "\", colProjects, False
'            End If
'        End If
'        sSubDir = Dir$()
'    Wend
'    If colProjects.Count > 0 Then Exit Sub
    
    If DoParent Then
        ' Search parent directory
        FindProjectsForComponent sCmpName, ParentDirectory(sDirPath), colProjects
    End If
End Sub

' ······························
' · ProjectContainsComponent() ·
' ······························
'
' Parses a .VBP project file to ascertain whether or not it contains a reference
' to a given component file.
'
Private Function ProjectContainsComponent(ByVal sProjDir As String, ByVal sProjName As String, ByVal sCompPath As String) As Boolean
    Dim nFile As Integer
    
    On Error GoTo CANNOT_OPEN_FILE
    nFile = FreeFile
    Open sProjDir & sProjName For Input Access Read Shared As nFile
    Debug.Print "File Opened : " & sProjDir & sProjName
    
    On Error GoTo ERROR_IN_FILE
    
    Dim sLine As String, i As Integer
    Dim sTemp As String
    Do While Not EOF(nFile)
        Line Input #nFile, sLine
        i = InStr(sLine, "=")
        If i <> 0 Then
            ' Quick check to avoid parsing the whole file
            ' [The "Title" attribute is always after the component list]
            If Left$(sLine, i - 1) = "Title" Then Exit Do
            ' Assume it's a component
            ' Strip off the first part
            sLine = Mid$(sLine, i + 1)
            ' Check for name prefix and remove if present [not present for Forms]
            i = InStr(sLine, "; ")
            If i <> 0 Then sLine = Mid$(sLine, i + 2)
            ' Assuming it's a file name relative to the project's path
            ' check if it's the component path
            sTemp = CanonicalizePath(sProjDir & sLine)
            If UCase$(sTemp) = UCase$(sCompPath) Then
                ProjectContainsComponent = True
                Exit Do
            End If
        End If
    Loop
ERROR_IN_FILE:
    Close nFile
    Debug.Print "File Closed : " & sProjDir & sProjName
CANNOT_OPEN_FILE:
End Function

