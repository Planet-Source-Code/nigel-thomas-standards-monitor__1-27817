VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   8775
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   6585
   _ExtentX        =   11615
   _ExtentY        =   15478
   _Version        =   393216
   Description     =   "Allows you to check that you are meeting company standards for controls."
   DisplayName     =   "Standards Monitor"
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

Public FormDisplayed          As Boolean
Public VBInstance             As VBIDE.VBE

Private mobjStandards As Office.CommandBarControl
Public WithEvents Standards_ToolBarHandler As CommandBarEvents
Attribute Standards_ToolBarHandler.VB_VarHelpID = -1

'***************************************************************************
'*  Name         : AddinInstance_OnConnection
'*  Description  : Puts addin in IDE
'*  Parameters   : Application As Object, ConnectMode As AddInDesignerObjects.ext_ConnectMode,
'*               : AddInInst As Object, custom() As Variant
'*  Returns      : Nothing
'*  Author       : nthomas
'*  Date         : 09 Oct 2001
'***************************************************************************

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    
On Error GoTo errHandler

    'save the vb instance
    Set VBInstance = Application
       
    Put_Standards_Monitor_Into_IDE
    
Exit Sub

errHandler:

MsgBox Err.Description
    
End Sub

'***************************************************************************
'*  Name         : Put_Standards_Monitor_Into_IDE
'*  Description  : Put addin in IDE, with image, and contect events
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       : nigel
'*  Date         : 27 Jan 2001
'***************************************************************************

Sub Put_Standards_Monitor_Into_IDE()

Dim enumOSType As OSTypes

On Error Resume Next

    '******************************************************
    '* Add to the new Toolbar.
    '******************************************************
            
    Set mobjStandards = VBInstance.CommandBars("Standard").Controls.Add
    
    ' The default backgrounds are different for NT
    ' and below and 2000 and above
    
    enumOSType = GetOSType
    
    ' Select which image to use
    Select Case enumOSType
    
        Case osWindows2000AdvancedServer, _
             osWindows2000DataCenter, _
             osWindows2000Professional, _
             osWindows2000Server
                 
            Clipboard.SetData frmStandards.Image2000.Picture
        
        Case Else
        
            Clipboard.SetData frmStandards.ImageNT.Picture
    End Select
    
    mobjStandards.PasteFace
    
    '******************************************************
    ' Connect the event handler to receive the
    ' events for the new command bar control.
    '******************************************************
    
    Set Me.Standards_ToolBarHandler = VBInstance.Events.CommandBarEvents(mobjStandards)

End Sub

'***************************************************************************
'*  Name         : AddinInstance_OnDisconnection
'*  Description  : Remove from IDE
'*  Parameters   : RemoveMode As AddInDesignerObjects.extDisconnectMode, custom(
'*  Returns      : Nothing
'*  Author       : nigel
'*  Date         : 27 Jan 2001
'***************************************************************************

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)

On Error Resume Next
  
    mobjStandards.Delete

End Sub

'***************************************************************************
'*  Name         : Standards_ToolBarHandler_Click
'*  Description  : Show form
'*  Parameters   : CommandBarControl As Object, handled As Boolean,
'*               : CancelDefault As Boolean
'*  Returns      : Nothing
'*  Author       : nthomas
'*  Date         : 09 Oct 2001
'***************************************************************************

Private Sub Standards_ToolBarHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)

On Error Resume Next

    Set frmStandards.VBInstance = VBInstance

    frmStandards.SetupForm
    
    frmStandards.Show
    
End Sub
