VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   9948
   ClientLeft      =   1740
   ClientTop       =   1548
   ClientWidth     =   6588
   _ExtentX        =   11621
   _ExtentY        =   17547
   _Version        =   393216
   Description     =   "Delete *.vbw files on project loads"
   DisplayName     =   "Delete *.vbw files on project loads"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 98 (ver 6.0)"
   LoadName        =   "Command Line / Startup"
   LoadBehavior    =   5
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mVBInstance             As VBIDE.VBE
Private WithEvents mProjHandler As VBIDE.VBProjectsEvents
Attribute mProjHandler.VB_VarHelpID = -1

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo error_handler
    
    Set mVBInstance = Application
    Set mProjHandler = Nothing
    Set mProjHandler = mVBInstance.Events.VBProjectsEvents
    
    Exit Sub
    
error_handler:
    MsgBox Err.Description
End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    
    Set mVBInstance = Nothing
    Set mProjHandler = Nothing
End Sub

Private Sub mProjHandler_ItemAdded(ByVal VBProject As VBIDE.VBProject)
    Dim vbwFile As String
    
    On Error Resume Next
    
    vbwFile = Left$(VBProject.FileName, Len(VBProject.FileName) - 3) & "vbw"
    If GetAttr(vbwFile) And vbReadOnly <> 0 Then
        SetAttr vbwFile, GetAttr(vbwFile) And Not vbReadOnly
    End If
    Kill vbwFile
End Sub

Private Sub mProjHandler_ItemRemoved(ByVal VBProject As VBIDE.VBProject)
    mProjHandler_ItemAdded VBProject
End Sub
