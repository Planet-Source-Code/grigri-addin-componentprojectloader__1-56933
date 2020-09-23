VERSION 5.00
Begin VB.Form ProjectChooser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Please Select Project"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton btnLoad 
      Caption         =   "Load"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   3240
      Width           =   975
   End
   Begin VB.ListBox lstProjects 
      Height          =   2205
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   4455
   End
   Begin VB.Label lblComponentPath 
      Caption         =   "X:\Path\To\Component.ext"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label Label2 
      Caption         =   "were found. Please choose one to load:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "More than one project containing the component"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "ProjectChooser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1

Private m_sSelectedProject As String

Private Sub btnCancel_Click()
    m_sSelectedProject = vbNullString
    Unload Me
End Sub

Private Sub btnLoad_Click()
    m_sSelectedProject = lstProjects.List(lstProjects.ListIndex)
    Unload Me
End Sub

Public Function SelectProject(ByVal sComponent As String, colProjects As Collection) As String
    Load Me
    
    lblComponentPath.Caption = sComponent
    Dim v As Variant
    For Each v In colProjects
        lstProjects.AddItem v
    Next
    lstProjects.ListIndex = 0
    
    Me.Show vbModal
    
    SelectProject = m_sSelectedProject
End Function

Private Sub Form_Load()
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub lstProjects_DblClick()
    btnLoad_Click
End Sub
