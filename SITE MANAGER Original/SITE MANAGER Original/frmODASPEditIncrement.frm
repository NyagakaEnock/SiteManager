VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmODASPEditIncrement 
   Caption         =   "                  Contract Increment Editing"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   8535
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   " "
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      Begin VB.TextBox txtContractNo 
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
         Left            =   720
         TabIndex        =   8
         Text            =   " "
         Top             =   480
         Width           =   2895
      End
      Begin VB.TextBox txtYearNo 
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Left            =   2400
         MaxLength       =   2
         TabIndex        =   7
         Text            =   " "
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox txtAmount 
         BackColor       =   &H00FFC0C0&
         Height          =   300
         Left            =   6120
         TabIndex        =   6
         Text            =   " "
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&Load Installments"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1200
         Picture         =   "frmODASPEditIncrement.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   4680
         Width           =   2775
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save Changes"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4680
         Picture         =   "frmODASPEditIncrement.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4680
         Width           =   2655
      End
      Begin VB.TextBox txtLandlord 
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
         Left            =   3600
         TabIndex        =   2
         Text            =   " "
         Top             =   480
         Width           =   3975
      End
      Begin VB.TextBox txtLocation 
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
         Left            =   720
         TabIndex        =   1
         Text            =   " "
         Top             =   1080
         Width           =   6975
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2295
         Left            =   120
         TabIndex        =   3
         Top             =   2280
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   4048
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label5 
         Caption         =   "LOCATION"
         Height          =   255
         Left            =   3240
         TabIndex        =   13
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "CONTRACT NUMBER"
         Height          =   255
         Left            =   960
         TabIndex        =   12
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Enter The Year No"
         Height          =   255
         Left            =   840
         TabIndex        =   11
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Amount"
         Height          =   255
         Left            =   5160
         TabIndex        =   10
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "OWNER"
         Height          =   255
         Left            =   5040
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmODASPEditIncrement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNew_Click()
  With Me
    .txtAmount.Text = ""
    .txtYearNo.Text = ""
    GetInstallmentToEdit
  End With
End Sub
Private Sub cmdSave_Click()
With Me
   If ValidRecord = True Then
       updateInstallmentRent
   Else
   End If
End With
End Sub
Public Sub updateInstallmentRent()
On Error GoTo err

       With frmODASPEditIncrement
            Set rsSAVE = New ADODB.Recordset
                strSQL = "select * from ODASMInstallment Where ContractNo = '" & .txtContractNo & "' and contractYear='" & .txtYearNo & "';"
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsSAVE.EOF Or rsSAVE.BOF Then Exit Sub
                    rsSAVE!TotalRent = .txtAmount.Text
                    rsSAVE!PaymentDue = .txtAmount.Text
                    rsSAVE!Balance = .txtAmount.Text
                    rsSAVE.Update
                
                Set rsSAVE = Nothing
                strSQL = Empty
        End With
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub Form_Load()
With Me
  .cmdSave.Enabled = True
  .cmdNew.Enabled = True
End With
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
With Me
  .txtAmount.Text = Item.SubItems(3)
  .txtYearNo.Text = Item.SubItems(1)
End With
End Sub

Private Sub txtAmount_Change()
With Me
  .cmdSave.Enabled = True
  .cmdNew.Enabled = True
  End With
End Sub

Private Sub txtYearNo_Change()
With Me
  .cmdSave.Enabled = True
  .cmdNew.Enabled = True
End With
End Sub
Private Function ValidRecord()
On Error GoTo err
    ValidRecord = False
    With frmODASPEditIncrement
      If .txtAmount.Text = " " Then
          strMessage = "Please Enter The Correct Rent ..........."
          .txtAmount.SetFocus
      ElseIf .txtYearNo.Text = Empty Then
          strMessage = "Please Enter The Installment Number You Are Updating"
          .txtYearNo.SetFocus

     Else
     ValidRecord = True
     End If
         
     If Not ValidRecord Then
     MsgBox strMessage, vbCritical + vbOKOnly, "Data Validation"
     End If
            
    End With
Exit Function

err:
    ErrorMessage
End Function

