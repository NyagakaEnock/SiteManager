VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmODASMOrder 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7200
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11910
   Icon            =   "frmODASMOrder1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Caption         =   "List of Valid Products"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1935
      Left            =   120
      TabIndex        =   37
      Top             =   2880
      Width           =   5895
      Begin MSComctlLib.ListView ListView3 
         Height          =   1575
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   2778
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Purchase Order Items"
      Height          =   2655
      Left            =   6120
      TabIndex        =   13
      Top             =   720
      Width           =   5655
      Begin VB.TextBox txtExchangeRate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   1320
         TabIndex        =   30
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtCurrencySymbol 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   960
         TabIndex        =   29
         Top             =   1680
         Width           =   375
      End
      Begin VB.ComboBox cboCurrencyCode 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   3840
         TabIndex        =   28
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtRequisitionDate 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   3840
         TabIndex        =   25
         Top             =   840
         Width           =   1575
      End
      Begin MSComCtl2.UpDown UpDownQuantity 
         Height          =   255
         Left            =   2280
         TabIndex        =   24
         Top             =   810
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         _Version        =   393216
         Max             =   1000
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtTotalItemCost 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   3840
         TabIndex        =   21
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox txtVATAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   3840
         TabIndex        =   19
         Top             =   1635
         Width           =   1575
      End
      Begin VB.TextBox txtUnitCost 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   960
         TabIndex        =   17
         Top             =   1215
         Width           =   1575
      End
      Begin VB.TextBox txtQuantityOrdered 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   960
         TabIndex        =   15
         Top             =   795
         Width           =   1335
      End
      Begin VB.TextBox txtProductCode 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   960
         TabIndex        =   14
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label Label2 
         Caption         =   "Product"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label28 
         Caption         =   "Currency"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label27 
         Caption         =   "Curr Code"
         Height          =   255
         Left            =   2880
         TabIndex        =   32
         Top             =   1230
         Width           =   735
      End
      Begin VB.Label Label26 
         Caption         =   "Req Date"
         Height          =   255
         Left            =   2880
         TabIndex        =   31
         Top             =   855
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Total Cost"
         Height          =   255
         Left            =   2880
         TabIndex        =   22
         Top             =   2055
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "VAT"
         Height          =   255
         Left            =   2880
         TabIndex        =   20
         Top             =   1650
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Cost"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1230
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Quantity"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   810
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "All Purchase Orders"
      Height          =   3615
      Left            =   6120
      TabIndex        =   12
      Top             =   3360
      Width           =   5655
      Begin MSComctlLib.ListView ListView2 
         Height          =   3255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   5741
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2175
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   5895
      Begin VB.TextBox txtRemarks 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   35
         Top             =   1680
         Width           =   4455
      End
      Begin VB.TextBox txtTotalCost 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4440
         TabIndex        =   26
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtOrderNo 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtOrderDate 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4440
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtDeadlineDate 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtOrderDescription 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   720
         Width           =   4455
      End
      Begin VB.Label Label22 
         Caption         =   "Remarks"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "Total Cost"
         Height          =   255
         Left            =   3360
         TabIndex        =   27
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Order No"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Order Date"
         Height          =   255
         Left            =   3360
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label23 
         Caption         =   "Deadline Date"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label25 
         Caption         =   "Description"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Purchase Order Irens"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   4920
      Width           =   5895
      Begin MSComctlLib.ListView ListView1 
         Height          =   1575
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   2778
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7080
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMOrder1.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMOrder1.frx":0ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMOrder1.frx":0F0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMOrder1.frx":1228
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMOrder1.frx":18A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMOrder1.frx":1F1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMOrder1.frx":236E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   1164
      ButtonWidth     =   3069
      ButtonHeight    =   1005
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New &Record "
            Key             =   "N"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Edit/Change "
            Key             =   "E"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Search/Find "
            Key             =   "S"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Refresh "
            Key             =   "R"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print &Preview "
            Key             =   "P"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Help"
            Key             =   "F"
            ImageIndex      =   1
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComDlg.CommonDialog HelpCommonDialog 
         Left            =   10560
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         HelpCommand     =   11
         HelpContext     =   72
         HelpFile        =   "REGHELP.HLP"
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuClear 
         Caption         =   "Clear the &Screen"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnumm 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuShow 
      Caption         =   "&Show/View"
      Begin VB.Menu mnuClosedJobs 
         Caption         =   "Closed Jobs"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuKJHGFDGFVHJ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFullInventory 
         Caption         =   "Full Inventory"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help System"
      Begin VB.Menu mnuHow 
         Caption         =   "How to use this System"
         Shortcut        =   ^{F1}
      End
   End
End
Attribute VB_Name = "frmODASMOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsREQ As clsODASRequisition
Dim rsOPERATION As clsODASOperation

Private Sub cboCurrencyCode_GotFocus()
        SelectCurrencyGotFocus
End Sub

Private Sub cboCurrencyCode_KeyPress(KeyAscii As Integer)
        KeyAscii = 0
End Sub

Private Sub cboCurrencyCode_LostFocus()
        selectCurrencyLostFocus
End Sub
Private Sub cboCostingCode_GotFocus()
        SelectCostingGotFocus
End Sub

Private Sub cboCostingCode_KeyPress(KeyAscii As Integer)
        KeyAscii = 0
End Sub

Private Sub cboCostingCode_LostFocus()
        selectCostingLostFocus
End Sub


Private Sub Form_Activate()
        Set rsREQ = New clsODASRequisition
        rsREQ.LoadDEFAULT
        rsREQ.obtainBASECURRENCY
        rsREQ.calculateTOTALS
        Set rsREQ = Nothing
        showALLREQUISITIONSRAISED
'        disableFRAME
End Sub


Private Sub Form_Load()
        OpenODBCConnection
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo err
    ListView1.SortKey = ColumnHeader.Index - 1
    ListView1.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ListView1_DblClick()
On Error GoTo err
            With SchedulingMain
                    
                    CurrentRecord = Trim(Me.ListView1.SelectedItem.Text)
                    GlobalDepartmentCode = ""
                    bRequisitionAPPROVAL = False
                    frmODASMOperation.txtApplicationNo.Text = CurrentRecord
                    Set rsOPERATION = New clsODASOperation
                    bRequisitionAPPROVAL = True
                    GlobalDepartmentCode = .ListView1.SelectedItem.SubItems(1)
                    rsOPERATION.checkAPPROVEDDISCHARGE
                    If bRequisitionAPPROVAL = False Then Exit Sub
                    
                    rsOPERATION.approveOPERATION
                    Set rsOPERATION = Nothing
                    bRequisitionAPPROVAL = False
            End With

Exit Sub

err:
    ErrorMessage
End Sub


Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo err
        Dim i, j As Double
        
        'If Item.Checked = True And baddRECORD = True Or beditRECORD = True Or bsearchRECORD = True Then

        If Item.Checked = True Then
            
            j = Screen.ActiveForm.ListView1.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                If Screen.ActiveForm.ListView1.ListItems(i) <> Item Then
                   Screen.ActiveForm.ListView1.ListItems(i).Checked = False
                End If
            Next i
            
            Screen.ActiveForm.txtLandLordNo.Text = Item.Text
            Screen.ActiveForm.txtNames.Text = Item.SubItems(1)
            showALLLandLORDSites

        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub
Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo err
    ListView2.SortKey = ColumnHeader.Index - 1
    ListView2.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo err
        Dim i, j As Double
        
        'If Item.Checked = True And baddRECORD = True Or beditRECORD = True Or bsearchRECORD = True Then

        If Item.Checked = True Then
            
            j = Screen.ActiveForm.ListView2.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                If Screen.ActiveForm.ListView2.ListItems(i) <> Item Then
                   Screen.ActiveForm.ListView2.ListItems(i).Checked = False
                End If
            Next i
            
            Screen.ActiveForm.txtItemCode.Text = Item.Text
            Set rsREQ = New clsODASRequisition
            rsREQ.LoadItemCOST
            Set rsREQ = Nothing
            
            Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub disableFRAME()
On Error GoTo err
    
    With Screen.ActiveForm
        .Frame1.Enabled = False
        .Frame2.Enabled = False
        .Frame3.Enabled = False
    End With
Exit Sub

err:
    ErrorMessage
End Sub
Private Sub enableFRAME()
On Error GoTo err
    
    With Screen.ActiveForm
        .Frame1.Enabled = False
        .Frame2.Enabled = True
        .Frame3.Enabled = True
    End With
Exit Sub

err:
    ErrorMessage
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo Err
        
        With Screen.ActiveForm
        
        Select Case Button.Key
                Case "N"
                    Select Case Button.Caption
                    
                    Case "New &Record "
                        If EditRecord Then Exit Sub
                        .ListView2.ListItems.Clear:
                        NewRecord = True: Button.Caption = "&Save Record ": Button.Image = 4
                        enableFRAME
                    Case "&Save Record "
                        If NewRecord Then
                            Set rsREQ = New clsODASRequisition
                            rsREQ.updateRECORD
                            Set rsREQ = Nothing

                            .Toolbar1.Buttons(3).Caption = "FINISH"
                            
                        End If
                        
                    Case Else
                        Exit Sub
                    End Select
    
        Case "E"
            Select Case Button.Caption
                Case "&Edit/Change "
                Case Else
            End Select
        
        Case "S"
        
        Case "R"
            If MsgBox(RefreshMessage, vbExclamation + vbOKCancel + vbDefaultButton2, "Refresh The Screen") = vbCancel Then Exit Sub
                .Toolbar1.Buttons(2).Caption = "New &Record "
                .Toolbar1.Buttons(2).Image = 2
                .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                .Toolbar1.Buttons(3).Image = 5
                NewRecord = False: EditRecord = False: clearALLRECORD
        Case "P"
        Case "F"
     
     
        Case Else
            Exit Sub
        End Select
End With
Exit Sub
err:
    ErrorMessage

End Sub


Private Sub UpDownQuantity_Change()
On Error GoTo err
        
        
        With frmODASMOpenJobCard
            If .txtExchangeRate.Text <= Empty Or CDbl(.txtItemCost.Text) <= 0 Then Exit Sub
            .txtItemQuantity.Text = .UpDownQuantity.Value
            .txtTotalItemCost.Text = CCur(.txtItemCost.Text) * CDbl(.txtItemQuantity.Text) * CDbl(.txtExchangeRate.Text)
            .txtTotalVATAmount.Text = CCur(.txtVATAmount.Text) * CDbl(.txtItemQuantity.Text) & CDbl(.txtExchangeRate.Text)
        
        End With

Exit Sub

err:
    ErrorMessage
End Sub
