VERSION 5.00
Begin VB.Form frmStandards 
   Caption         =   "Standards"
   ClientHeight    =   765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4830
   Icon            =   "frmStandards.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   765
   ScaleWidth      =   4830
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReport 
      Caption         =   "Report"
      Height          =   390
      Left            =   1710
      TabIndex        =   2
      Top             =   135
      Width           =   1470
   End
   Begin VB.CommandButton cmdSetup 
      Caption         =   "Setup Standards"
      Height          =   420
      Left            =   210
      TabIndex        =   1
      Top             =   135
      Width           =   1455
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   390
      Left            =   3240
      TabIndex        =   0
      Top             =   135
      Width           =   1410
   End
End
Attribute VB_Name = "frmStandards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public VBInstance As VBIDE.VBE
Public Connect As Connect

'***************************************************************************
'*  Name         : cmdClose_Click
'*  Description  : Close addin
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       : nthomas
'*  Date         : 05 Oct 2001
'***************************************************************************

Private Sub cmdClose_Click()

    Unload Me
    
End Sub

'***************************************************************************
'*  Name         : cmdReport_Click
'*  Description  : Show report
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       : nthomas
'*  Date         : 05 Oct 2001
'***************************************************************************

Private Sub cmdReport_Click()

    Set frmStandardsMaintenance.VBInstance = VBInstance
    
    frmStandardsMaintenance.Mode = "Report"
    
    frmStandardsMaintenance.Show

End Sub

'***************************************************************************
'*  Name         : cmdSetup_Click
'*  Description  : Show Setup
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       : nthomas
'*  Date         : 05 Oct 2001
'***************************************************************************

Private Sub cmdSetup_Click()

    Set frmStandardsMaintenance.VBInstance = VBInstance

    frmStandardsMaintenance.Mode = "Setup"
    
    frmStandardsMaintenance.Show
    
End Sub
