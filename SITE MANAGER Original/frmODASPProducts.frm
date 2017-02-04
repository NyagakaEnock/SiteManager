VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmODASPProducts 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GENERAL INVENTORY-Product Details"
   ClientHeight    =   5955
   ClientLeft      =   2850
   ClientTop       =   1860
   ClientWidth     =   9735
   Icon            =   "frmODASPProducts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   9735
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      _Version        =   393216
      BorderStyle     =   1
      Begin VB.CommandButton Command1 
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
         Left            =   7800
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "View all products per supplier"
         Top             =   0
         Width           =   1935
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
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         Width           =   1815
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
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Change existing record"
         Top             =   0
         Width           =   1935
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
         TabIndex        =   3
         ToolTipText     =   "Add new record"
         Top             =   0
         Width           =   2055
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H80000000&
         Caption         =   "REFRES&H"
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
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Clear screen"
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5175
      Left            =   0
      TabIndex        =   6
      Top             =   600
      Width           =   9735
      Begin VB.TextBox txtCategoryCode 
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
         TabIndex        =   17
         Top             =   480
         Width           =   2415
      End
      Begin VB.Frame Frame2 
         Caption         =   "List of Current Product Details"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   120
         TabIndex        =   13
         Top             =   2160
         Width           =   9495
         Begin MSComctlLib.ListView ListView1 
            Height          =   2415
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   4260
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            GridLines       =   -1  'True
            HotTracking     =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
      Begin VB.TextBox txtCategoryName 
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
         Left            =   2520
         TabIndex        =   0
         Top             =   480
         Width           =   6135
      End
      Begin VB.TextBox txtProductCode 
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
         Left            =   1560
         TabIndex        =   1
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox txtProductName 
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
         Left            =   1560
         TabIndex        =   2
         Top             =   1680
         Width           =   7095
      End
      Begin VB.Label Label1 
         Caption         =   "Category Name"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   16
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Product Code"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Category Code"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Name of Product"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   120
         X2              =   9600
         Y1              =   960
         Y2              =   960
      End
   End
   Begin VB.Label Label7 
      Caption         =   "Economic Capital"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   4200
      Width           =   1335
   End
End
Attribute VB_Name = "frmODASPProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim MyProduct As clsODASProductInfo


Private Sub cboDrugType_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboDrugcategory_LostFocus()
    MyProduct.GetCategoryCode
End Sub

Private Sub cboPackageType_Click()
    Me.txtDrugName.SetFocus
End Sub

Private Sub cboPackageType_GotFocus()
    MyProduct.AttachPackageTypes
End Sub

Private Sub cboPackageType_LostFocus()
    MyProduct.GetPackageType
End Sub

Private Sub cboSupplierCode_Click()
    Me.txtProductCode.SetFocus
End Sub


Private Sub cmdAddNew_Click()
'On Error GoTo err
    If Edit = False Then
        MyProduct.AddNewDetails
    End If
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub cmdEditRecord_Click()
'On Error GoTo err
    If Save = False Then
        MyProduct.EditRecordDetails
    End If
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub cmdPrint_Click()
Load frmRPTGeneralProducts
frmRPTGeneralProducts.Show 1, Me

End Sub

Private Sub cmdRefresh_Click()
'On Error GoTo err

    MyProduct.RefreshDetails
    Found = False


Exit Sub

err:
ErrorMessage
End Sub

Private Sub Command1_Click()
MyProduct.GetDetailsAllProducts
End Sub

Private Sub Form_Activate()
        showALLProducts
End Sub

Private Sub Form_Initialize()
    Set MyProduct = New clsODASProductInfo
End Sub

Private Sub Form_Load()
Call OpenODBCConnection
    'MyProduct.disableControlsDetails
    'MyProduct.GetDetailsStructure
'On Error Resume Next
End Sub

Private Sub Form_Terminate()
    Set MyProduct = Nothing
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

Private Sub txtDrugId_KeyPress(KeyAscii As Integer)
'On Error GoTo err
If KeyAscii = vbKeyReturn Then
    MyProduct.SearchRecordDetails
Else
    Select Case KeyAscii
    Case Asc("0") To Asc("9"), vbKeyBack
    Case Else
        KeyAscii = 0
    End Select
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
        
        If Item.Checked = True Then
            
            j = Screen.ActiveForm.ListView1.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                If Screen.ActiveForm.ListView1.ListItems(i) <> Item Then
                   Screen.ActiveForm.ListView1.ListItems(i).Checked = False
                End If
            Next i
                        
            frmODASPProducts.txtProductCode.Text = Item.Text
            MyProduct.FindDetailsRecord
        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub

