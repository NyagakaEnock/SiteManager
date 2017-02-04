VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmUVouchersPrepared 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vouchers Prepared"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   12585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Update"
      Height          =   495
      Left            =   9000
      TabIndex        =   10
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000002&
      Height          =   495
      Left            =   5640
      TabIndex        =   9
      Top             =   120
      Width           =   3375
   End
   Begin VB.TextBox txtVoucherNo 
      Height          =   375
      Left            =   10800
      TabIndex        =   7
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00808000&
      Caption         =   "&Print Record"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4680
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808000&
      Caption         =   "Search List"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4560
      Width           =   3015
   End
   Begin MSComctlLib.ListView listView1 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   6376
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
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComCtl2.DTPicker DTPLastDate 
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   180
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   90243075
      CurrentDate     =   40961
   End
   Begin MSComCtl2.DTPicker DTPStartDate 
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   90243075
      CurrentDate     =   40961
   End
   Begin VB.Label Label3 
      Caption         =   "Voucher No:"
      Height          =   255
      Left            =   9840
      TabIndex        =   8
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Last Date:"
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Start Date:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   180
      Width           =   1095
   End
End
Attribute VB_Name = "frmUVouchersPrepared"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSearch_Click()
If Trim(Me.txtVoucherNo.Text) = "" Then
    MsgBox "Select a voucher before you proceed", vbExclamation
    Me.listView1.SetFocus
    Exit Sub
End If

CurrentRecord = Me.txtVoucherNo.Text
Load frmPayRequisition
frmPayRequisition.Show 1, Me
End Sub

Private Sub Command1_Click()
GetVouchersPrepared

End Sub

Private Sub Command2_Click()
            Set rsSAVE = New Recordset
            If txtVoucherNo = "" Then Exit Sub
            strSAVE = "Select * From ODASMVoucherItem where VoucherNo = '" & txtVoucherNo & "'"
            rsSAVE.Open strSAVE, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            If rsSAVE.EOF Or rsSAVE.BOF Then
            
            Else
            rsSAVE!ItemName = Me.Text1
            rsSAVE.Update
            MsgBox "Vourcher Details No '" & txtVoucherNo & " ' Update Success"
            End If
 
End Sub

Private Sub Form_Load()
Me.DTPStartDate.Value = Date
Me.DTPLastDate.Value = Date
End Sub

Private Sub listView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
checkOne Item, Me.listView1
Item.Selected = True
Me.txtVoucherNo.Text = Item.Text
Set rsSAVE = New Recordset
            strSAVE = "Select * From ODASMVoucherItem where VoucherNo = '" & txtVoucherNo & "'"
            rsSAVE.Open strSAVE, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            If rsSAVE.EOF Or rsSAVE.BOF Then
            
            Else
            Me.Text1 = rsSAVE!ItemName
           
            End If
End Sub

Private Sub listView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
Item.Checked = True
checkOne Item, Me.listView1

End Sub
