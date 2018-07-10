VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FolderIconsEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents g_explorers As Outlook.Explorers
Attribute g_explorers.VB_VarHelpID = -1
Private Sub Class_Initialize()
    'Debug.Print "FolderIcons::Initialize"
    Set g_explorers = Application.Explorers
End Sub
Private Sub g_explorers_NewExplorer(ByVal expl As Outlook.Explorer)
    'Debug.Print "g_explorers_NewExplorer"
    FolderIcons.Initialize
End Sub
