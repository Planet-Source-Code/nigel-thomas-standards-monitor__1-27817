VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmStandards 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Standards Monitor"
   ClientHeight    =   7845
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   9840
   Icon            =   "frmStandardsReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   9840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   8535
      TabIndex        =   14
      Top             =   7335
      Width           =   1215
   End
   Begin TabDlg.SSTab tabStandards 
      Height          =   7155
      Left            =   105
      TabIndex        =   0
      Top             =   90
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   12621
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Setup"
      TabPicture(0)   =   "frmStandardsReport.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraSetup"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Report"
      TabPicture(1)   =   "frmStandardsReport.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraReport"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame fraReport 
         Caption         =   "Report"
         Height          =   6510
         Left            =   -74835
         TabIndex        =   11
         Top             =   450
         Width           =   9300
         Begin VB.CommandButton cmdReport 
            Caption         =   "Report"
            Height          =   375
            Left            =   7905
            TabIndex        =   12
            Top             =   6030
            Width           =   1215
         End
         Begin MSComctlLib.ListView lvwReport 
            Height          =   5685
            Left            =   105
            TabIndex        =   13
            Top             =   270
            Width           =   9030
            _ExtentX        =   15928
            _ExtentY        =   10028
            View            =   3
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            Icons           =   "ilsControlIcons"
            SmallIcons      =   "ilsControlIcons"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Control Type"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Control Name"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Property"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Standard Value"
               Object.Width           =   2602
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Property Value"
               Object.Width           =   2602
            EndProperty
         End
         Begin MSComctlLib.ImageCombo cboFormType 
            Height          =   330
            Left            =   5835
            TabIndex        =   15
            Top             =   6030
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Locked          =   -1  'True
         End
      End
      Begin VB.Frame fraSetup 
         Caption         =   "Setup"
         Height          =   6060
         Left            =   150
         TabIndex        =   1
         Top             =   435
         Width           =   9285
         Begin VB.CommandButton cmdUpdate 
            Caption         =   "Update"
            Enabled         =   0   'False
            Height          =   375
            Left            =   8070
            TabIndex        =   16
            Top             =   5490
            Width           =   1065
         End
         Begin VB.TextBox txtValue 
            Height          =   345
            Left            =   4815
            TabIndex        =   6
            Top             =   5490
            Width           =   1950
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add"
            Enabled         =   0   'False
            Height          =   375
            Left            =   6900
            TabIndex        =   3
            Top             =   5490
            Width           =   1065
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Delete"
            Enabled         =   0   'False
            Height          =   375
            Left            =   8070
            TabIndex        =   2
            Top             =   4755
            Width           =   1065
         End
         Begin MSComctlLib.ImageCombo cboProperty 
            Height          =   330
            Left            =   2505
            TabIndex        =   4
            Top             =   5490
            Width           =   2220
            _ExtentX        =   3916
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Locked          =   -1  'True
         End
         Begin MSComctlLib.ImageCombo cboControl 
            Height          =   330
            Left            =   195
            TabIndex        =   5
            Top             =   5490
            Width           =   2220
            _ExtentX        =   3916
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Locked          =   -1  'True
         End
         Begin MSComctlLib.ListView lvwSetup 
            Height          =   4305
            Left            =   195
            TabIndex        =   7
            Top             =   375
            Width           =   8925
            _ExtentX        =   15743
            _ExtentY        =   7594
            View            =   3
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            Icons           =   "ilsControlIcons"
            SmallIcons      =   "ilsControlIcons"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Control Type"
               Object.Width           =   5027
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Property Name"
               Object.Width           =   5027
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Value"
               Object.Width           =   4586
            EndProperty
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Control:"
            Height          =   195
            Left            =   195
            TabIndex        =   10
            Top             =   5220
            Width           =   540
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Property:"
            Height          =   195
            Left            =   2505
            TabIndex        =   9
            Top             =   5220
            Width           =   630
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Value:"
            Height          =   195
            Left            =   4815
            TabIndex        =   8
            Top             =   5220
            Width           =   450
         End
      End
   End
   Begin MSComctlLib.ImageList ilsControlIcons 
      Left            =   2175
      Top             =   7185
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   29
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStandardsReport.frx":0342
            Key             =   "CheckBox"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStandardsReport.frx":065C
            Key             =   "Unknown"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStandardsReport.frx":0976
            Key             =   "ComboBox"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStandardsReport.frx":0C90
            Key             =   "CommandButton"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStandardsReport.frx":0FAA
            Key             =   "Frame"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStandardsReport.frx":12C4
            Key             =   "Grid"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStandardsReport.frx":15DE
            Key             =   "Image"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStandardsReport.frx":18F8
            Key             =   "ImageCombo"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStandardsReport.frx":1C12
            Key             =   "Label"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStandardsReport.frx":1F2C
            Key             =   "ListBox"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStandardsReport.frx":2246
            Key             =   "ListView"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStandardsReport.frx":2560
            Key             =   "OptionButton"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStandardsReport.frx":287A
            Key             =   "PictureBox"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStandardsReport.frx":2B94
            Key             =   "TextBox"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStandardsReport.frx":2EAE
            Key             =   "Drive"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStandardsReport.frx":31CA
            Key             =   "DTPicker"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStandardsReport.frx":34E6
            Key             =   "File"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStandardsReport.frx":3802
            Key             =   "Folder"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStandardsReport.frx":3B1E
            Key             =   "HScrollBar"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStandardsReport.frx":3E3A
            Key             =   "MonthView"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStandardsReport.frx":4156
            Key             =   "MSFlexGrid"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStandardsReport.frx":4472
            Key             =   "MaskEdBox"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStandardsReport.frx":478E
            Key             =   "RichTextBox"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStandardsReport.frx":4AAA
            Key             =   "WebBrowser"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStandardsReport.frx":4DC6
            Key             =   "Slider"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStandardsReport.frx":50E2
            Key             =   "Tab"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStandardsReport.frx":53FE
            Key             =   "UpDown"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStandardsReport.frx":571A
            Key             =   "VScrollBar"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStandardsReport.frx":5A36
            Key             =   "vsFlexArray"
         EndProperty
      EndProperty
   End
   Begin VB.Image Image2000 
      Height          =   480
      Left            =   1425
      Picture         =   "frmStandardsReport.frx":5D52
      Top             =   7350
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageNT 
      Height          =   480
      Left            =   795
      Picture         =   "frmStandardsReport.frx":6994
      Top             =   7320
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmStandards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public VBInstance As VBIDE.VBE

' API Declarations for Highlight_Control_Text
Private Declare Function GetKeyState Lib "user32" (ByVal lVirtKey As Long) As Integer

'***************************************************************************
'*  Name         : cboControl_Click
'*  Description  : Load into the properties combo the properties for the selected
'*               : control
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       : nthomas
'*  Date         : 05 Oct 2001
'***************************************************************************

Private Sub cboControl_Click()

On Error GoTo errHandler

Dim lngPos As Long, strName As String, objForm As VBForm
Dim objControl As VBControl, objProperty As Property

    ' Find classname of control
    lngPos = InStr(1, cboControl.ComboItems(cboControl.SelectedItem.Key).Key, " <> ")
    
    strName = Mid$(cboControl.ComboItems(cboControl.SelectedItem.Index).Key, 1, lngPos - 1)
    
    ' Get active form
    Set objForm = VBInstance.SelectedVBComponent.Designer
    
    ' Loop thu. Controls of form
    For Each objControl In objForm.ContainedVBControls
           
        ' Found type of control ?
        If objControl.ClassName = strName Then
        
            ' Clear old items
            cboProperty.ComboItems.Clear
            
            ' Add Properties for the control
        
            For Each objProperty In objControl.Properties
                
                cboProperty.ComboItems.Add , , objProperty.Name
            Next
            
            Exit For
        End If
    Next
    
    SortImageCombos cboProperty
    
    txtValue = ""
    
    cboProperty.SetFocus
    cboControl.SetFocus
    
Exit Sub
 
errHandler:
 
    Err.Source = Err.Source & "." & TypeName(Me) & ".cboControl_Click"
 
    MsgBox Err.Number & " - " & Err.Source & " - " & Err.Description
 
End Sub

'***************************************************************************
'*  Name         : cboControl_GotFocus
'*  Description  : Highlight
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       : nthomas
'*  Date         : 08 Oct 2001
'***************************************************************************

Private Sub cboControl_GotFocus()

On Error Resume Next
    
    HighlightControlText Me, cboControl
    
End Sub

'***************************************************************************
'*  Name         : cboControl_LostFocus
'*  Description  : Set text
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       : nthomas
'*  Date         : 08 Oct 2001
'***************************************************************************

Private Sub cboControl_LostFocus()

On Error GoTo errHandler

    If cboProperty.ComboItems.Count > 0 Then
    
        ' Highlight text
        cboProperty.Text = cboProperty.ComboItems(1).Text
    End If
    
Exit Sub
 
errHandler:
 
    If Err.Number = 5 Then Resume Next
 
    Err.Source = Err.Source & "." & TypeName(Me) & ".cboControl_Click"
    
    MsgBox Err.Number & " - " & Err.Source & " - " & Err.Description
    
End Sub

'***************************************************************************
'*  Name         : cboProperty_GotFocus
'*  Description  : Highlight
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       : nthomas
'*  Date         : 08 Oct 2001
'***************************************************************************

Private Sub cboProperty_GotFocus()

On Error Resume Next

    HighlightControlText Me, cboProperty
    
End Sub

'***************************************************************************
'*  Name         : cmdAdd_Click
'*  Description  : Add the standard to the listview
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       : nthomas
'*  Date         : 05 Oct 2001
'***************************************************************************

Private Sub cmdAdd_Click()

On Error GoTo errHandler

Dim objListItem As ListItem

    ' Remove any existing standard with same values as 1 to be added
    
    For Each objListItem In lvwSetup.ListItems
    
        If objListItem.Text = cboControl.Text Then
        
            If objListItem.ListSubItems(1).Text = cboProperty.Text Then
            
                ' If same classname and property - remove
                lvwSetup.ListItems.Remove (objListItem.Index)
                Exit For
            End If
        End If
    Next
    
    ' Add new standard
    Set objListItem = lvwSetup.ListItems.Add(, , cboControl.SelectedItem.Text, , GetIcon(cboControl.SelectedItem.Text))
    
    objListItem.SubItems(1) = cboProperty.Text
    objListItem.SubItems(2) = txtValue
    
    Set objListItem = Nothing
    
    cmdDelete.Enabled = True
    cmdUpdate.Enabled = True
    
    txtValue = ""
    
Exit Sub
 
errHandler:
 
    Err.Source = Err.Source & "." & TypeName(Me) & ".cmdAdd_Click"
 
    MsgBox Err.Number & " - " & Err.Source & " - " & Err.Description
 
End Sub

'***************************************************************************
'*  Name         : cmdClose_Click
'*  Description  : Close form
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       : nthomas
'*  Date         : 05 Oct 2001
'***************************************************************************

Private Sub cmdClose_Click()

On Error GoTo errHandler

    Unload Me
    
Exit Sub
 
errHandler:
 
    Err.Source = Err.Source & "." & TypeName(Me) & ".cmdClose_Click"
 
    MsgBox Err.Number & " - " & Err.Source & " - " & Err.Description
 
End Sub

'***************************************************************************
'*  Name         : LoadStandards
'*  Description  : Load in the current standards from the registry
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       : nthomas
'*  Date         : 04 Oct 2001
'***************************************************************************

Private Sub LoadStandards()

On Error GoTo errHandler

Dim objListItem As ListItem, varSettings As Variant, lngCount As Long
Dim lngCount2 As Long

    ' Go thu. each control type of current form, adding any standards found
    For lngCount = 1 To cboControl.ComboItems.Count
    
        ' Get standards
        varSettings = GetAllSettings(App.Title, cboControl.ComboItems(lngCount).Text)
    
        If Not IsEmpty(varSettings) Then
        
            For lngCount2 = LBound(varSettings, 1) To UBound(varSettings, 1)
    
                ' Add to listview
                Set objListItem = lvwSetup.ListItems.Add(, , cboControl.ComboItems(lngCount).Text, , GetIcon(cboControl.ComboItems(lngCount).Text))
                
                objListItem.ListSubItems.Add , , varSettings(lngCount2, 0)
                objListItem.ListSubItems.Add , , varSettings(lngCount2, 1)
            Next
        End If
    Next

    ' Enable/ Disable delete/edit
    If lvwSetup.ListItems.Count > 0 Then
    
        cmdDelete.Enabled = True
    Else
    
        cmdDelete.Enabled = False
    End If
    
Exit Sub
 
errHandler:
 
    Err.Source = Err.Source & "." & TypeName(Me) & ".LoadStandards"
 
    Err.Raise Err.Number, Err.Source, Err.Description
 
End Sub

'***************************************************************************
'*  Name         : LoadControls
'*  Description  : Load in the control types of the current form
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       : nthomas
'*  Date         : 04 Oct 2001
'***************************************************************************

Private Sub LoadControls()

Dim objForm As VBForm, objProperty As Property, objControl As VBControl
    
On Error Resume Next
    
    Set objForm = VBInstance.SelectedVBComponent.Designer
    
    ' Add Controls
    For Each objControl In objForm.ContainedVBControls
           
        cboControl.ComboItems.Add , objControl.Properties("Name") & " : " & objControl.ClassName, objControl.ClassName
    Next
    
    ' Sort controls imagecombo
    
    SortImageCombos cboControl
    
    ' Highlight 1st control load in properties
    
    If cboControl.ComboItems.Count > 0 Then
    
        cboControl.SelectedItem = cboControl.ComboItems(1)
    
        ' Add Properties for 1st control
        
        For Each objProperty In objForm.ContainedVBControls(1).Properties
            
            cboProperty.ComboItems.Add , , objProperty.Name
        Next
        
        cboProperty.SelectedItem = cboProperty.ComboItems(1)
        cboProperty.Text = cboProperty.SelectedItem.Text
        
        ' Sort the properties combo
        
        SortImageCombos cboProperty
    End If
    
    ' Load current standards into listview
    
    LoadStandards
    
End Sub

'***************************************************************************
'*  Name         : cmdDelete_Click
'*  Description  : Delete listitem from listview
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       : nthomas
'*  Date         : 05 Oct 2001
'***************************************************************************

Private Sub cmdDelete_Click()

On Error GoTo errHandler

    lvwSetup.ListItems.Remove (lvwSetup.SelectedItem.Index)
    
    If lvwSetup.ListItems.Count = 0 Then
    
        cmdDelete.Enabled = False
    Else
    
        cmdDelete.Enabled = True
    End If
    
    cmdUpdate.Enabled = True
      
Exit Sub
 
errHandler:
 
    Err.Source = Err.Source & "." & TypeName(Me) & ".cmdDelete_Click"
 
    MsgBox Err.Number & " - " & Err.Source & " - " & Err.Description
 
End Sub

'***************************************************************************
'*  Name         : cmdReport_Click
'*  Description  : Do the checking of the control properties
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       : nthomas
'*  Date         : 05 Oct 2001
'***************************************************************************

Private Sub cmdReport_Click()

On Error GoTo errHandler

Dim objForm As VBForm, objComponent As VBComponent
    
    ' Remove any old items
    lvwReport.ListItems.Clear
    
    ' Check for the users selection
    Select Case cboFormType.SelectedItem.Key
    
        Case "CURFORM"
        
            ' Current form
            Set objForm = VBInstance.SelectedVBComponent.Designer
    
            CheckForm objForm
        
        Case "ALLFORMS"
            
            ' Go thu. each component if form type send in
            
            For Each objComponent In VBInstance.ActiveVBProject.VBComponents
            
                If objComponent.Type = vbext_ct_VBForm Or objComponent.Type = vbext_ct_VBMDIForm Then
                
                    Set objForm = objComponent.Designer
                    
                    CheckForm objForm
                End If
                
            Next
        
    End Select
    
Exit Sub
 
errHandler:
 
    Err.Source = Err.Source & "." & TypeName(Me) & ".cmdReport_Click"
 
    MsgBox Err.Number & " - " & Err.Source & " - " & Err.Description
 
End Sub

'***************************************************************************
'*  Name         : cmdUpdate_Click
'*  Description  : Save changes to registry
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       : nthomas
'*  Date         : 05 Oct 2001
'***************************************************************************

Private Sub cmdUpdate_Click()

On Error GoTo errHandler

    SaveStandards
    
    cmdUpdate.Enabled = False
Exit Sub
 
errHandler:
 
    Err.Source = Err.Source & "." & TypeName(Me) & ".cmdUpdate_Click"
 
    MsgBox Err.Number & " - " & Err.Source & " - " & Err.Description
 
End Sub

'***************************************************************************
'*  Name         : lvwSetup_ItemClick
'*  Description  : Enable edit of standard
'*  Parameters   : Item As MSComctlLib.ListItem
'*  Returns      : Nothing
'*  Author       : nthomas
'*  Date         : 05 Oct 2001
'***************************************************************************

Private Sub lvwSetup_ItemClick(ByVal Item As MSComctlLib.ListItem)

On Error GoTo errHandler

Dim objListItem As ListItem
    
    Set objListItem = lvwSetup.SelectedItem
    
    If objListItem Is Nothing Then Exit Sub
    
    ' Copy values
    cboControl.Text = objListItem.Text
    cboProperty.Text = objListItem.ListSubItems(1).Text
    txtValue = objListItem.ListSubItems(2).Text
    
Exit Sub
 
errHandler:
 
    Err.Source = Err.Source & "." & TypeName(Me) & ".lvwSetup_ItemClick"
 
    MsgBox Err.Number & " - " & Err.Source & " - " & Err.Description
 
End Sub

'***************************************************************************
'*  Name         : tabStandards_Click
'*  Description  : Tab orders
'*  Parameters   : PreviousTab As Integer
'*  Returns      : Nothing
'*  Author       : nthomas
'*  Date         : 05 Oct 2001
'***************************************************************************

Private Sub tabStandards_Click(PreviousTab As Integer)

On Error GoTo errHandler

    ' Change tab order on chnage of tab
    SetTabOrder
    
    ' Set starting point
    If tabStandards.Tab = 0 Then
    
        lvwSetup.SetFocus
    Else
    
        cmdReport.SetFocus
    End If
    
Exit Sub
 
errHandler:
 
    If Err.Number = 5 Then Resume Next
    
    Err.Source = Err.Source & "." & TypeName(Me) & ".tabStandards_Click"
 
    MsgBox Err.Number & " - " & Err.Source & " - " & Err.Description
 
End Sub

'***************************************************************************
'*  Name         : txtValue_Change
'*  Description  : Enable/Disenable txtValue
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       : nthomas
'*  Date         : 05 Oct 2001
'***************************************************************************

Private Sub txtValue_Change()

On Error GoTo errHandler

    If Trim$(txtValue) <> "" And Trim$(cboControl.Text) <> "" And Trim$(cboProperty.Text) <> "" Then
    
        cmdAdd.Enabled = True
    Else
    
        cmdAdd.Enabled = False
    End If
    
Exit Sub
 
errHandler:
 
    Err.Source = Err.Source & "." & TypeName(Me) & ".txtValue_Change"
 
    MsgBox Err.Number & " - " & Err.Source & " - " & Err.Description
 
End Sub

'***************************************************************************
'*  Name         : SaveStandards
'*  Description  : Save the standards for the controls to the registry
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       : nthomas
'*  Date         : 05 Oct 2001
'***************************************************************************

Private Sub SaveStandards()

Dim objListItem As ListItem, objComboItem As ComboItem

On Error Resume Next

    ' Remove any previous standards
    
    For Each objComboItem In cboControl.ComboItems
    
        DeleteSetting App.Title, objComboItem.Text
    Next
    
    ' Save the new ones, if any
    
    For Each objListItem In lvwSetup.ListItems
    
        SaveSetting App.Title, objListItem.Text, objListItem.ListSubItems(1).Text, objListItem.ListSubItems(2).Text
    Next

End Sub

'***************************************************************************
'*  Name         : SetTabOrder
'*  Description  : Tab Order
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       : nthomas
'*  Date         : 05 Oct 2001
'***************************************************************************

Private Sub SetTabOrder()

On Error GoTo errHandler
    
    If tabStandards.Tab = 0 Then
        
        lvwReport.TabStop = False
        cboFormType.TabStop = False
        cmdReport.TabStop = False
        
        lvwSetup.TabStop = True
        cmdDelete.TabStop = True
        cboControl.TabStop = True
        cboProperty.TabStop = True
        txtValue.TabStop = True
        cmdAdd.TabStop = True
        cmdUpdate.TabStop = True
        
        lvwSetup.TabIndex = 1
        cmdDelete.TabIndex = 2
        cboControl.TabIndex = 3
        cboProperty.TabIndex = 4
        txtValue.TabIndex = 5
        cmdAdd.TabIndex = 6
        cmdUpdate.TabIndex = 7
        
        cmdClose.TabIndex = 8
        
    Else
    
        lvwReport.TabStop = True
        cmdReport.TabStop = True
        cboFormType.TabStop = True
        
        lvwSetup.TabStop = False
        cmdDelete.TabStop = False
        cboControl.TabStop = False
        cboProperty.TabStop = False
        txtValue.TabStop = False
        cmdAdd.TabStop = False
        cmdUpdate.TabStop = False
        
        lvwReport.TabIndex = 1
        cboFormType.TabIndex = 2
        cmdReport.TabIndex = 3
        
        cmdClose.TabIndex = 4
    End If

Exit Sub
 
errHandler:
 
    Err.Source = Err.Source & "." & TypeName(Me) & ".SetTabOrder"
 
    MsgBox Err.Number & " - " & Err.Source & " - " & Err.Description
 
End Sub

'***************************************************************************
'*  Name         : QuickSort
'*  Description  : Sort Array
'*  Parameters   : strValues() As String, intMin As Integer,
'*               : intMax As Integer
'*  Returns      : Nothing
'*  Author       : nthomas
'*  Date         : 08 Oct 2001
'***************************************************************************

Public Sub QuickSort(ByRef strValues() As String, ByVal intMin As Integer, ByVal intMax As Integer)

On Error GoTo errHandler

Dim strMed_Value As String
Dim intHigh As Integer, intLow As Integer
Dim intRandomValue As Integer
    
    ' If the list has only 1 item, it's sorted.
    If intMin >= intMax Then Exit Sub
    
    ' Pick a dividing item randomly
    intRandomValue = intMin + Int(Rnd(intMax - intMin + 1))
    
    strMed_Value = strValues(intRandomValue)
      
    ' Swap the dividing item to the front of the list.
    strValues(intRandomValue) = strValues(intMin)
      
    ' Separate the list into sublists.
    intLow = intMin
    intHigh = intMax
    
    Do
        ' Look down from High for a value < med_value.
        
        Do While strValues(intHigh) >= strMed_Value
            intHigh = intHigh - 1
            
            If intHigh <= intLow Then Exit Do
        Loop
        
        If intHigh <= intLow Then
            
            ' The list is separated.
            strValues(intLow) = strMed_Value
            Exit Do
        End If
        ' Swap the Low and High values.
        
        strValues(intLow) = strValues(intHigh)
        ' Look up from Low for a value >= med_value.
        
        intLow = intLow + 1
        
        Do While strValues(intLow) < strMed_Value
            
            intLow = intLow + 1
            If intLow >= intHigh Then Exit Do
                
        Loop
        
        If intLow >= intHigh Then
            
            ' The list is separated.
            intLow = intHigh
            strValues(intHigh) = strMed_Value
            Exit Do
        End If
                    
        ' Swap the Low and High values.
        strValues(intHigh) = strValues(intLow)
    Loop
               
    ' Loop until the list is separated.
      
    ' Recursively sort the sublists.
                
    QuickSort strValues, intMin, intLow - 1
    QuickSort strValues, intLow + 1, intMax
                
Exit Sub
 
errHandler:
 
    Err.Source = Err.Source & "." & TypeName(Me) & ".QuickSort"
 
    MsgBox Err.Number & " - " & Err.Source & " - " & Err.Description
                
End Sub

'***************************************************************************
'*  Name         : SortImageCombos
'*  Description  : Allows the sorting of the contents of an imagecombo
'*  Parameters   : cboImageCombo - ImageCombo
'*  Returns      : Nothing
'*  Author       : nthomas
'*  Date         : 08 Oct 2001
'***************************************************************************

Private Sub SortImageCombos(ByRef cboImageCombo As ImageCombo)

On Error GoTo errHandler

Dim strValues() As String, intCount As Integer, blnKeyFound As Boolean
    
    ' Create array of the contents so that it can be sorted

    ReDim strValues(0 To cboImageCombo.ComboItems.Count)
    
    For intCount = 0 To cboImageCombo.ComboItems.Count - 1
    
        If Trim$(cboImageCombo.ComboItems(intCount + 1).Key) <> "" Then
        
            ' Add any key to the temp array
            strValues(intCount) = cboImageCombo.ComboItems(intCount + 1).Text & " <> " & cboImageCombo.ComboItems(intCount + 1).Key
            blnKeyFound = True
        Else
        
            strValues(intCount) = cboImageCombo.ComboItems(intCount + 1).Text
        End If
    Next

    ' Now do the sort
    
    QuickSort strValues, LBound(strValues), UBound(strValues)
    
    ' Now clear the imagecombo and add the sorted array to it
    
    cboImageCombo.ComboItems.Clear
    
    ' Do the add
    
    For intCount = 1 To UBound(strValues)
    
        ' Reconstruct Comboitem
        
        If blnKeyFound Then
        
            cboImageCombo.ComboItems.Add , Mid$(strValues(intCount), 1, InStr(1, strValues(intCount), " : ")), _
                                           Mid$(strValues(intCount), InStr(1, strValues(intCount), " : ") + 3)
        Else
        
            cboImageCombo.ComboItems.Add , , strValues(intCount)
        End If
    Next

Exit Sub
 
errHandler:
 
    Err.Source = Err.Source & "." & TypeName(Me) & ".SortImageCombos"
 
    MsgBox Err.Number & " - " & Err.Source & " - " & Err.Description

End Sub

'***************************************************************************
'*  Name         : HighlightControlText
'*  Description  : Highlight the text of a control on GotFocus
'*  Parameters   : robjParentForm As Object, robjCtl As Object
'*  Returns      : Nothing
'*  Author       : nthomas
'*  Date         : 08 Oct 2001
'***************************************************************************

Public Sub HighlightControlText(ByRef robjParentForm As Object, ByRef robjCtl As Object)

On Error Resume Next
    
    ' If you got to the "Ctl" field via either TAB or an
    ' Alt-Key, highlight the whole field. Otherwise select
    ' no text, since it must have received focus using a mouse-click.

    ' Note difference between vbTab (character) and vbKeyTab
    ' (numeric constant). If vbTab were used, we'd have to
    ' Asc() it to get a number as an argument.
    
    With robjCtl
    
        If (GetKeyState(vbKeyTab) < 0) Or _
            (GetKeyState(vbKeyMenu) < 0) Then
            ' We tabbed or used a hotkey - select all text. In
            ' the case of a long field, use Sendkeys so we
            ' see the beginning of the selected text.

            ' TextWidth Method tells how much width a string
            ' takes up to display (default target object is
            ' the Form).
            If robjParentForm.TextWidth(.Text) > .Width Then
                
                SendKeys "{End}", True
                SendKeys "+{Home}", True
            Else
                
                .SelStart = 0
                .SelLength = Len(.Text)
            End If
        Else
        
            .SelLength = 0
        
        End If
    
    End With
    
Exit Sub
 
errHandler:
 
    Err.Source = Err.Source & "." & TypeName(Me) & ".HighlightControlText"
 
    MsgBox Err.Number & " - " & Err.Source & " - " & Err.Description

End Sub

'***************************************************************************
'*  Name         : txtValue_GotFocus
'*  Description  : Highlight
'*  Parameters   : None
'*  Returns      : Nothing
'*  Author       : nthomas
'*  Date         : 08 Oct 2001
'***************************************************************************

Private Sub txtValue_GotFocus()

On Error Resume Next

    HighlightControlText Me, txtValue
    
End Sub

'***************************************************************************
'*  Name         : CheckForm
'*  Description  : Check the supplied form for the standards setup
'*  Parameters   : objForm as VBForm
'*  Returns      : Nothing
'*  Author       : nthomas
'*  Date         : 08 Oct 2001
'***************************************************************************

Private Sub CheckForm(ByVal objForm As VBForm)

On Error Resume Next

Dim varStandards As Variant, objListItem As ListItem, blnAdd As Boolean
Dim objControl As VBControl, lngCount As Long, objFont As IFontDisp
Dim strPropValue As String
    
    For Each objControl In objForm.ContainedVBControls
           
        ' Get standards for this type of control
        varStandards = GetAllSettings(App.Title, objControl.ClassName)
        
        If Not IsEmpty(varStandards) Then
        
            ' We have standards to check
        
            For lngCount = LBound(varStandards, 1) To UBound(varStandards, 1)
    
                ' Test property value
                
                If UCase$(objControl.Properties(varStandards(lngCount, 0)).Value) <> UCase$(varStandards(lngCount, 1)) Then
                    
                    ' Get the value of the property - This value may
                    ' be changed by the font property
                    
                    strPropValue = objControl.Properties(varStandards(lngCount, 0)).Value
                    
                    ' By default we add has the properties are different but we
                    ' need to handle the font object property
                    blnAdd = True
                    
                    ' Check if property is 'Font' - this is an object type
                    
                    If objControl.Properties(varStandards(lngCount, 0)).Name = "Font" Then
                    
                        Set objFont = objControl.Properties(varStandards(lngCount, 0)).Object
                        
                        ' Check Font Charset
                        If UCase$(objFont.Name) = UCase$(varStandards(lngCount, 1)) Then
                        
                            blnAdd = False
                        Else
                        
                            ' Get the value of the font property
                            strPropValue = objFont.Name
                        End If
                    End If
                    
                    If blnAdd Then
                    
                        ' Add to report listitem if different
                        Set objListItem = lvwReport.ListItems.Add(, , objControl.ClassName, , GetIcon(objControl.ClassName))
                        
                        objListItem.SubItems(1) = objControl.Properties("Name").Value
                        objListItem.SubItems(2) = varStandards(lngCount, 0)
                        objListItem.SubItems(3) = varStandards(lngCount, 1)
                        objListItem.SubItems(4) = strPropValue
                    End If
                End If
            Next
        
        End If
    Next

End Sub

'***************************************************************************
'*  Name         : GetIcons
'*  Description  : Return the index for the icon for the classname of the control
'*  Parameters   : strClassname - String
'*  Returns      : Long
'*  Author       : nthomas
'*  Date         : 08 Oct 2001
'***************************************************************************

Private Function GetIcon(ByVal strClassName As String) As Long

On Error Resume Next

    GetIcon = ilsControlIcons.ListImages(strClassName).Index

    ' Check for an unknown control type - If found return
    ' the index for the 'Unknown' control icon
    
    If GetIcon = 0 Then GetIcon = 2

End Function

'***************************************************************************
'*  Name         : SetupForm
'*  Description  : Set up the form with taborder, controls etc
'*  Parameters   : none
'*  Returns      : Nothing
'*  Author       : nthomas
'*  Date         : 08 Oct 2001
'***************************************************************************

Public Sub SetupForm()

On Error Resume Next

    cboFormType.ComboItems.Clear
    cboControl.ComboItems.Clear
    cboProperty.ComboItems.Clear
    
    lvwSetup.ListItems.Clear
    lvwReport.ListItems.Clear
    
    ' Report types
    cboFormType.ComboItems.Add , "CURFORM", "Current Form"
    cboFormType.ComboItems.Add , "ALLFORMS", "All Forms in Project"
    
    cboFormType.SelectedItem = cboFormType.ComboItems(1)
    
    ' Start Tab
    tabStandards.Tab = 0
    
    ' Load the controls from the active form into the imagecombos
    LoadControls
    
    ' Tab order
    SetTabOrder

End Sub
