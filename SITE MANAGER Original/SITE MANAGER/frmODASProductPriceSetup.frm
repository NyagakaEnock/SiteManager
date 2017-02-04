VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmODASPProductPriceSetup 
   BackColor       =   &H00FFC0C0&
   Caption         =   "PRODUCTS' COSTS AND PRICES SETUP"
   ClientHeight    =   8010
   ClientLeft      =   3270
   ClientTop       =   3510
   ClientWidth     =   11880
   Icon            =   "frmODASProductPriceSetup.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   Picture         =   "frmODASProductPriceSetup.frx":0442
   ScaleHeight     =   8010
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtTOTAL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   390
      Left            =   6720
      TabIndex        =   24
      Top             =   1440
      Width           =   1095
   End
   Begin VB.ComboBox cboCategoryCode 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   3135
   End
   Begin VB.TextBox txtCategoryName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   390
      Left            =   3240
      TabIndex        =   1
      Top             =   960
      Width           =   4575
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6615
      Left            =   7920
      TabIndex        =   14
      Top             =   960
      Width           =   3855
      Begin VB.CheckBox chkSearchEdit 
         Caption         =   "Activate Search and Edit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   3945
         Width           =   3495
      End
      Begin VB.CommandButton cmdCHANGE 
         BackColor       =   &H80000000&
         Caption         =   "&CHANGE"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   5160
         Width           =   3015
      End
      Begin VB.CommandButton cmdREFRESH 
         BackColor       =   &H80000000&
         Caption         =   "&REFRESH"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   5880
         Width           =   3015
      End
      Begin VB.CommandButton cmdNEW 
         BackColor       =   &H80000000&
         Caption         =   "&NEW"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4440
         Width           =   3015
      End
      Begin VB.TextBox txtDosagePrice 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   1200
         TabIndex        =   7
         Top             =   3480
         Width           =   2535
      End
      Begin VB.TextBox txtPriceMarkup 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   1200
         TabIndex        =   6
         Top             =   2982
         Width           =   2535
      End
      Begin VB.TextBox txtDosageCost 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   1200
         TabIndex        =   5
         Top             =   2485
         Width           =   2535
      End
      Begin VB.TextBox txtDrugName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1491
         Width           =   2655
      End
      Begin VB.TextBox txtDrugCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   3
         Top             =   994
         Width           =   2655
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00404000&
         BorderWidth     =   2
         X1              =   120
         X2              =   3720
         Y1              =   857
         Y2              =   857
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C00000&
         X1              =   120
         X2              =   3720
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Label Label8 
         Caption         =   "Price Markup"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   2982
         Width           =   975
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C00000&
         X1              =   120
         X2              =   3720
         Y1              =   3915
         Y2              =   3915
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   120
         X2              =   3720
         Y1              =   2348
         Y2              =   2348
      End
      Begin VB.Label Label9 
         Caption         =   "Product Price"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   3480
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Product Cost"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   2485
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Product Name"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   1491
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Product Code"
         Height          =   375
         Left            =   0
         TabIndex        =   19
         Top             =   990
         Width           =   975
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5775
      Left            =   0
      TabIndex        =   2
      Top             =   1815
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   10186
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777152
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   7635
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   4304
            MinWidth        =   4304
            Picture         =   "frmODASProductPriceSetup.frx":5B80
            Text            =   "COST AND PRICES"
            TextSave        =   "COST AND PRICES"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   13829
            MinWidth        =   13829
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            Object.Width           =   3246
            MinWidth        =   3246
            TextSave        =   "15/11/2004"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
      BorderStyle     =   1
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   585
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "*** PRODUCTS' COSTS AND PRICES SETUP ***"
         Top             =   0
         Width           =   11895
      End
   End
   Begin VB.Label lblTask 
      BackStyle       =   0  'Transparent
      Caption         =   "Tasks / Setup Options"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   7920
      TabIndex        =   18
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "New Products or Settings in Pre-Specified Dosages"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   1440
      Width           =   5775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Category Code"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   600
      Width           =   3135
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404000&
      BorderWidth     =   2
      X1              =   0
      X2              =   7800
      Y1              =   1380
      Y2              =   1380
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Category Name"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3240
      TabIndex        =   15
      Top             =   600
      Width           =   3135
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuClearScreen 
         Caption         =   "Clear &Screen"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnujjjh 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuShow 
      Caption         =   "&Show/View"
      Begin VB.Menu mnuAllNewProductsDrugs 
         Caption         =   "All N&ew Products"
      End
      Begin VB.Menu mnuDFDS 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowAllSettings 
         Caption         =   "&All Existing Cost and Price Settings"
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu mnuHelpSystem 
      Caption         =   "&Help System"
      Begin VB.Menu mnuHelpCostsSetup 
         Caption         =   "Using Costs and Prices &Setup"
         Shortcut        =   {F1}
      End
   End
   Begin VB.Menu mnuPrintPreview 
      Caption         =   "&Print Preview"
      Begin VB.Menu mnuPrint 
         Caption         =   "Print &All"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnujkhsd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintByProductCode 
         Caption         =   "Print By P&roduct Code"
         Shortcut        =   {F8}
      End
   End
End
Attribute VB_Name = "frmODASPProductPriceSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MyCostPrice As clsPharmacySettings

Private Sub cboCategoryCode_Click()
    Me.ListView1.SetFocus
End Sub
Private Sub cboCategoryCode_GotFocus()
    MyCostPrice.AttachDrugCategories
End Sub

Private Sub cboCategoryCode_LostFocus()
    MyCostPrice.GetCategoryByName
End Sub

Private Sub chkSearchEdit_Click()
On Error GoTo err
With Me
If NewRecord Or EditRecord Then .chkSearchEdit.Value = 0: Exit Sub
    
    If .chkSearchEdit.Value = 1 Then
        'MyCostPrice.ClearTheScreen
        Me.cmdNEW.Enabled = False
    ElseIf .chkSearchEdit.Value = 0 Then
        Me.cmdNEW.Enabled = True
    End If
    
End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub cmdChange_Click()
On Error GoTo err
If NewRecord Then Exit Sub
If Not chkSearchEdit.Value = 1 Then Exit Sub
Select Case cmdCHANGE.Caption
    Case "&CHANGE"
        MyCostPrice.CheckForRecord
        If Not AllowEdit Then Exit Sub
'        MyCostPrice.GetDosageRecords
        cmdCHANGE.Caption = "SAVE &CHANGES"
        EditRecord = True
    Case "SAVE &CHANGES"
        If EditRecord Then
            If ValidCostPrice Then
                MyCostPrice.EditCurrentRecord
            End If
        End If
    Case Else
        Exit Sub
    End Select
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub cmdNew_Click()
On Error GoTo err
With Me
If EditRecord Then Exit Sub
    Select Case .cmdNEW.Caption
    Case "&NEW"
    
        NewRecord = True
        
        If Me.cboCategoryCode.Text = Empty Then
            MyCostPrice.ShowDrugsDosages
        Else
            MyCostPrice.ShowDrugsDosagesByCategory
        End If
        
        MyCostPrice.AddNewCostSetup
        
    Case "&SAVE RECORD"
        If NewRecord Then
            If ValidCostPrice Then
                MyCostPrice.SaveCostPriceSetup
            End If
        End If
    Case Else
        Exit Sub
    End Select
End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Function ValidCostPrice() As Boolean
On Error GoTo err
With Me
    If .txtDosageCost.Text = Empty Then
        strMessage = "Required Cost of Dosage!!"
        .txtDosageCost.SetFocus
    ElseIf .txtPriceMarkup.Text = Empty Then
        strMessage = "Required Price Markup!!"
        .txtPriceMarkup.SetFocus
    ElseIf .txtDosagePrice.Text = Empty Then
        strMessage = "Required Price of Dosage!!"
        .txtDosagePrice.SetFocus
    Else
        ValidCostPrice = True
    End If
    If Not ValidCostPrice Then
        MsgBox strMessage, vbCritical + vbOKOnly, "Data Validation"
    End If
End With
Exit Function
err:
    ErrorMessage
End Function

Private Sub cmdPrintPReview_Click()

End Sub

Private Sub cmdRefresh_Click()
On Error GoTo err
With Me

If MsgBox(RefreshMessage, vbExclamation + vbOKCancel + vbDefaultButton2, "Screen Refresher") = vbCancel Then Exit Sub
    NewRecord = False: EditRecord = False
    MyCostPrice.ClearTheScreen
    
    .cmdNEW.Enabled = True
    .cmdNEW.Caption = "&NEW"
    .cmdCHANGE.Enabled = True
    .cmdCHANGE.Caption = "&CHANGE"
    .cmdRefresh.Enabled = True
        
End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub Form_Activate()
    MyCostPrice.GetMainStructure
End Sub

Private Sub FormatFields()
On Error GoTo err
With Me
    .txtDosageCost.Text = FormatNumber(.txtDosageCost.Text, 2, vbUseDefault, vbUseDefault, vbTrue)
    .txtDosagePrice.Text = FormatNumber(.txtDosagePrice.Text, 2, vbUseDefault, vbUseDefault, vbTrue)
    .txtPriceMarkup.Text = FormatNumber(.txtPriceMarkup.Text, 2, vbUseDefault, vbUseDefault, vbTrue)
End With
Exit Sub
err:
If err.Number = 5 Then Resume Next
    ErrorMessage
End Sub

Private Sub Form_Initialize()
    Set MyCostPrice = New clsPharmacySettings
End Sub

Private Sub Form_Resize()
On Error GoTo err
With Me
    .ListView1.Width = .Width - (12000 - 7815)
    .ListView1.Height = .Height - (8700 - 5775)
    .txtCategoryName.Width = .Width - (12000 - 4575)
    .txtTotal.Left = .ListView1.Width - .txtTotal.Width
    .Text5.Width = .Width - (12000 - 11895)
    .Frame1.Height = .Height - (8700 - 6615)
    .Frame1.Left = .ListView1.Width + 125
    .Line1.X2 = .Width - (12000 - 7800)
    .lblTask.Left = .Frame1.Left
End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub Form_Terminate()
    Set MyCostPrice = Nothing
End Sub

Private Sub ListView1_DblClick()
On Error GoTo err
With Me
If .ListView1.ListItems.Count = 0 Or .ListView1.View <> lvwReport Then Exit Sub

If .chkSearchEdit.Value = 0 Then

    If Not NewRecord And Not EditRecord Then: Exit Sub
    
    Dim i, j, k
    j = .ListView1.ListItems.Count
    
    .ListView1.SelectedItem.Checked = True
    
    Me.txtProductCode.Text = .ListView1.SelectedItem.Text
'    Me.txtDosageCode.Text = .ListView1.SelectedItem.Text
    Me.txtDrugName.Text = .ListView1.SelectedItem.SubItems(1)
'    Me.txtDosageType.Text = .ListView1.SelectedItem.SubItems(3)
    Me.txtPriceMarkup.Text = MyCostPrice.GetPriceMarkup
    
    For i = 1 To j
        If .ListView1.ListItems(i).Text <> Trim(.txtProductCode.Text) Then
            .ListView1.ListItems(i).Checked = False
        End If
    Next i
    
    .txtDosageCost.SetFocus
    
ElseIf .chkSearchEdit.Value = 1 Then
    
    .ListView1.SelectedItem.Checked = True
    
    With Me
'        .txtDosageCode.Text = .ListView1.SelectedItem.Text
'        .txtDosageType.Text = .ListView1.SelectedItem.SubItems(1)
        .txtProductCode.Text = .ListView1.SelectedItem
        .txtDosageCost.Text = .ListView1.SelectedItem.SubItems(2)
        .txtDosagePrice.Text = .ListView1.SelectedItem.SubItems(3)
        .txtPriceMarkup.Text = .ListView1.SelectedItem.SubItems(4)
        .txtDrugName.Text = .ListView1.SelectedItem.SubItems(1)
        
        j = .ListView1.ListItems.Count
        
        For i = 1 To j
            If .ListView1.ListItems(i).Text <> Trim(.txtProductCode.Text) Then
                .ListView1.ListItems(i).Checked = False
            End If
        Next i
        
    End With
    
End If

End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo err
If Me.ListView1.ListItems.Count = 0 Or Me.ListView1.View <> lvwReport Then Exit Sub

If Me.chkSearchEdit.Value = 0 Then
    If Not NewRecord And Not EditRecord Then Item.Checked = False: Exit Sub
    
    Dim i, j, k
    j = Me.ListView1.ListItems.Count
    
    If j = 0 Then Exit Sub
    
    For i = 1 To j
        If Me.ListView1.ListItems(i).Text <> Item Then
            Me.ListView1.ListItems(i).Checked = False
        End If
    Next i
    
    If Item.Checked = True Then
    
        MyCostPrice.ClearForNewRecord
'        Me.txtDosageCode.Text = Item
        Me.txtProductCode.Text = Item
        Me.txtDrugName.Text = Item.SubItems(1)
'        Me.txtDosageType.Text = Item.SubItems(3)
        Me.txtPriceMarkup.Text = MyCostPrice.GetPriceMarkup
        
        MyCostPrice.GetCategoryByName
        
        Me.txtDosageCost.SetFocus
        
    ElseIf Item.Checked = False Then
    
        MyCostPrice.ClearForNewRecord
        
    End If
    
ElseIf Me.chkSearchEdit.Value = 1 Then
    
    With Me
'        .txtDosageCode.Text = Item
'        .txtDosageType.Text = Item.SubItems(1)
        .txtProductCode.Text = Item
        .txtDosageCost.Text = Item.SubItems(2)
        .txtDosagePrice.Text = Item.SubItems(3)
        .txtPriceMarkup.Text = Item.SubItems(4)
        .txtDrugName.Text = Item.SubItems(1)
    End With
    
    Dim l, m, n
        m = Me.ListView1.ListItems.Count: n = 0
        
        For l = 1 To m
            If Me.ListView1.ListItems(l).Text <> Item Then
                Me.ListView1.ListItems(l).Checked = False
            End If
        Next l

End If

Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnuAllNewProductsDrugs_Click()
MyCostPrice.ShowDrugsDosages
End Sub

Private Sub mnuClearScreen_Click()
    MyCostPrice.ClearTheScreen
End Sub

Private Sub mnuClose_Click()
    Unload Me
End Sub

Private Sub mnuPrint_Click()
Load frmRPTCostPriceSetup
frmRPTCostPriceSetup.Show 1, Me
End Sub

Private Sub mnuPrintByProductCode_Click()
On Error GoTo err
With Me
    INPQRY = InputBox("Please Enter the Product Code To Print A Price List ", "Cost And Price Setup", Trim(.txtProductCode.Text))
    
    If Len(INPQRY) = 0 Then
        MsgBox "No Values Entered or Operation Was Cancelled! No Work Will Be Done!!"
        Exit Sub
    Else
        Set rsFindRecord = cnCOMMON.Execute("SELECT DrugCode FROM ProductsCostPriceSetup WHERE DrugCode='" & Trim(INPQRY) & "';")
        
        If rsFindRecord.EOF And rsFindRecord.BOF Then
            MsgBox "No Products Price Settings Have Been Made Or Value Entered is Not Correct!!", vbCritical + vbOKOnly, "Invalid Transfer No"
            Set rsFindRecord = Nothing: Exit Sub
        
        Else
            
            SelectedTransfer = Trim(INPQRY)
            
            Load frmRPTCostPriceSetupByNo
            frmRPTCostPriceSetupByNo.Show 1, Me
            
        End If
     End If
End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnuShow_Click()
On Error GoTo err
With Me

If .chkSearchEdit.Value = 0 Then

    .mnuShowAllSettings.Enabled = False
'    .mnuShowByDoseCode.Enabled = False
'    .mnuShowByDrugCode.Enabled = False
'    .mnuShowSettingsByDate.Enabled = False
'    .mnuShowSettingsByName.Enabled = False
'    .mnuShowSettingsForToday.Enabled = False
    
ElseIf .chkSearchEdit.Value = 1 Then

    .mnuShowAllSettings.Enabled = True
'    .mnuShowByDoseCode.Enabled = True
'    .mnuShowByDrugCode.Enabled = True
'    .mnuShowSettingsByDate.Enabled = True
'    .mnuShowSettingsByName.Enabled = True
'    .mnuShowSettingsForToday.Enabled = True
    
End If

End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnuShowAllSettings_Click()
On Error GoTo err
    'MyCostPrice.ClearTheScreen
    MyCostPrice.FindAllSettings
Exit Sub
err:
    ErrorMessage
End Sub


Private Sub mnuShowSettingsByDate_Click()
On Error GoTo err
    MyCostPrice.FindSettingsBySpecifiedDate
Exit Sub
err:
    ErrorMessage
End Sub


Private Sub mnuShowSettingsForToday_Click()
On Error GoTo err
    MyCostPrice.FindSettingsForDateToday
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub Text5_GotFocus()
    Me.cboCategoryCode.SetFocus
End Sub

Private Sub txtDosageCost_LostFocus()
If Me.txtDosageCost.Text = Empty Or Me.txtPriceMarkup.Text = Empty Then Exit Sub
    Me.txtDosagePrice.Text = CDbl(Me.txtDosageCost.Text) * CDbl(Me.txtPriceMarkup.Text)
    Call FormatFields
    If Me.chkSearchEdit.Value = 0 Then
        Me.cmdNEW.SetFocus
    ElseIf Me.chkSearchEdit.Value = 1 Then
        Me.cmdCHANGE.SetFocus
    End If
End Sub

Private Sub txtDosagePrice_LostFocus()
    Call FormatFields
End Sub
