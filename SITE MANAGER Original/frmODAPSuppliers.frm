VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmODASPSuppliers 
   Caption         =   "GENERAL INVENTORY-Suppliers Information"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmODAPSuppliers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame5 
      Caption         =   "List of All Suppliers"
      Height          =   1695
      Left            =   120
      TabIndex        =   40
      Top             =   5400
      Width           =   11535
      Begin MSComctlLib.ListView ListView3 
         Height          =   1335
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   2355
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
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
      Caption         =   "List of All Towns"
      Height          =   2055
      Left            =   120
      TabIndex        =   37
      Top             =   3360
      Width           =   3975
      Begin MSComctlLib.ListView ListView2 
         Height          =   1575
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   2778
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
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Delivery Methods"
      Height          =   2655
      Left            =   120
      TabIndex        =   36
      Top             =   720
      Width           =   3975
      Begin MSComctlLib.ListView ListView1 
         Height          =   2175
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   3836
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
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
      Caption         =   "Supply/Shipping Parameters"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4200
      TabIndex        =   29
      Top             =   4080
      Width           =   7455
      Begin VB.TextBox txtSupplyduration 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1440
         TabIndex        =   32
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox txtSupplierType 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1440
         TabIndex        =   31
         Top             =   360
         Width           =   2415
      End
      Begin VB.ComboBox cboDeliveryMethod 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5280
         TabIndex        =   30
         Top             =   315
         Width           =   1935
      End
      Begin VB.Label Label11 
         Caption         =   "Duration [Days]"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Supplier Type"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Delivery Method"
         Height          =   255
         Left            =   4080
         TabIndex        =   33
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   630
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      _Version        =   393216
      BorderStyle     =   1
      Begin VB.CommandButton cmdView 
         BackColor       =   &H80000000&
         Caption         =   "&VIEW"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9780
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Click to view list of current suppliers"
         Top             =   0
         Width           =   2415
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H80000000&
         Caption         =   "&PRINT"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7365
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   0
         Width           =   2415
      End
      Begin VB.CommandButton cmdEditRecord 
         BackColor       =   &H80000000&
         Caption         =   "E&DIT"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2535
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Change existing record"
         Top             =   0
         Width           =   2415
      End
      Begin VB.CommandButton cmdAddNew 
         BackColor       =   &H80000000&
         Caption         =   "&NEW"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Add New"
         Top             =   0
         Width           =   2535
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H80000000&
         Caption         =   "REFRE&SH"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4950
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Clear screen"
         Top             =   0
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   4200
      TabIndex        =   15
      Top             =   600
      Width           =   7455
      Begin VB.ComboBox cboContactTitle 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5280
         TabIndex        =   6
         Top             =   2040
         Width           =   1935
      End
      Begin VB.ComboBox cboTownCity 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5280
         TabIndex        =   4
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox txtfaxNo 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5280
         TabIndex        =   8
         Top             =   2520
         Width           =   1935
      End
      Begin VB.TextBox txtPhysicalAddress 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1440
         TabIndex        =   2
         Top             =   1080
         Width           =   5775
      End
      Begin VB.TextBox txtContactPerson 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1440
         TabIndex        =   5
         Top             =   2040
         Width           =   2415
      End
      Begin VB.TextBox txtSupplierName 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2760
         TabIndex        =   1
         Top             =   495
         Width           =   4455
      End
      Begin VB.TextBox txtPostalAddress 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1440
         TabIndex        =   3
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox txtTelephoneNo 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1440
         TabIndex        =   7
         Top             =   2520
         Width           =   2415
      End
      Begin VB.TextBox txtEmail 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1440
         TabIndex        =   9
         Top             =   3000
         Width           =   2415
      End
      Begin VB.TextBox txtMobileNo 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5280
         TabIndex        =   10
         Top             =   3000
         Width           =   1935
      End
      Begin VB.TextBox txtSupplierCode 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   0
         Top             =   495
         Width           =   2535
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   120
         X2              =   8760
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label15 
         Caption         =   "Fax No."
         Height          =   375
         Left            =   4080
         TabIndex        =   27
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Physical Address"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "Supplier Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label14 
         Caption         =   "Contact Person"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Official Tittle"
         Height          =   375
         Left            =   4080
         TabIndex        =   23
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Supplier Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   22
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "Postal Address"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Town City"
         Height          =   375
         Left            =   4080
         TabIndex        =   20
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Telephone"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "E-Mail"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Mobile Phone"
         Height          =   375
         Left            =   4080
         TabIndex        =   17
         Top             =   3000
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmODASPSuppliers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MySuppliers As clsODASSuppliers


Private Sub cboCurrencyCode_Click()
    Me.txtSupplyduration.SetFocus
End Sub

Private Sub cboCurrencyCode_GotFocus()
    MySuppliers.AttachCurrencies
End Sub

Private Sub cboCurrencyCode_LostFocus()
    MySuppliers.GetCurrencyCode
End Sub


Private Sub cboContactTitle_Change()
    MySuppliers.AttachOfficialTitle
End Sub

Private Sub cboDeliveryMethod_Click()
    Me.txtFreightCharge.SetFocus
End Sub

Private Sub cboDeliveryMethod_GotFocus()
    MySuppliers.AttachDeliveryMethods
End Sub

Private Sub cboDeliveryMethod_LostFocus()
    MySuppliers.GetDeliveryMethodID
End Sub




Private Sub cmdAddNew_Click()
    If Edit = False Then
        baddRECORD = True
        MySuppliers.AddNew
    End If
End Sub

Private Sub cmdEditRecord_Click()
    If Save = False Then
        MySuppliers.EditRecord
    End If
End Sub

Private Sub cmdPrint_Click()
'On Error GoTo err
    Load frmRptSuppliers
    frmRptSuppliers.Show 1, Me
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub cmdRefresh_Click()

    Set MySuppliers = New clsODASSuppliers
    MySuppliers.Refresh
    Found = False

End Sub


Private Sub cmdView_Click()
Load frmSuppliersList
frmSuppliersList.Show vbModal, Me
End Sub

Private Sub Form_Activate()
    ShowALLSuppliers
    showDeliveryMethodS
    showTOWNS
End Sub

Private Sub Form_Initialize()
    Set MySuppliers = New clsODASSuppliers
End Sub

Private Sub Form_Load()

    'MySuppliers.disableControls

'On Error Resume Next
    Me.Left = 2805
    Me.Top = 1425
End Sub

Private Sub Form_Terminate()
    Set MySuppliers = New clsODASSuppliers
End Sub

Private Sub Form_Unload(Cancel As Integer)

'On Error GoTo err
    If Save = True Or Edit = True Then
        MsgBox "Please there is Work going on, Refresh to continue", vbOKCancel + vbCritical
        Cancel = 1
    Else
        Found = False

    End If
Exit Sub

err:
ErrorMessage
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'On Error GoTo err
    ListView1.SortKey = ColumnHeader.Index - 1
    ListView1.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'On Error GoTo err
        Dim i, j As Double
        
        If Item.Checked = True And baddRECORD = True Or beditRECORD = True Or bsearchRECORD = True Then
            
            j = Screen.ActiveForm.ListView1.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                If Screen.ActiveForm.ListView1.ListItems(i) <> Item Then
                   Screen.ActiveForm.ListView1.ListItems(i).Checked = False
                End If
            Next i
                        
            frmODASPSuppliers.cboDeliveryMethod.Text = Item.Text
        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub
Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'On Error GoTo err
    ListView2.SortKey = ColumnHeader.Index - 1
    ListView2.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'On Error GoTo err
        Dim i, j As Double
        
        If Item.Checked = True And baddRECORD = True Or beditRECORD = True Or bsearchRECORD = True Then
            
            j = Screen.ActiveForm.ListView2.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                If Screen.ActiveForm.ListView2.ListItems(i) <> Item Then
                   Screen.ActiveForm.ListView2.ListItems(i).Checked = False
                End If
            Next i
                        
            frmODASPSuppliers.cboTownCity.Text = Item.Text
        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub
Private Sub ListView3_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'On Error GoTo err
    ListView3.SortKey = ColumnHeader.Index - 1
    ListView3.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ListView3_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'On Error GoTo err
        Dim i, j As Double
        
        If Item.Checked = True Then
            
            j = Screen.ActiveForm.ListView3.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                If Screen.ActiveForm.ListView3.ListItems(i) <> Item Then
                   Screen.ActiveForm.ListView3.ListItems(i).Checked = False
                End If
            Next i
                        
            frmODASPSuppliers.txtSupplierCode.Text = Item.Text
            MySuppliers.FindRecord
        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub


Private Sub txtSupplierCode_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    MySuppliers.SearchRecord
End If
End Sub

Private Sub cboSupplierType_GotFocus()
'On Error GoTo err

If Me.cboSupplierType.ListCount <> 0 Then Exit Sub

    Dim rssup As ADODB.Recordset
    Set rssup = New ADODB.Recordset
    
    rssup.Open "Select * from ParamSuppliertypes", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    While Not rssup.EOF
        cboSupplierType.AddItem rssup!descriptions
        rssup.MoveNext
    Wend
        
    Set rssup = Nothing

Exit Sub
err:
    ErrorMessage
End Sub

