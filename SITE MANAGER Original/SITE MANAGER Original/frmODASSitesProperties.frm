VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmODASSitesProperties 
   Caption         =   "Print Sites "
   ClientHeight    =   5220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9240
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   9240
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      Begin VB.CommandButton cmdLoad 
         BackColor       =   &H00808000&
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00808000&
         Caption         =   "Print Record"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3960
         Width           =   2175
      End
      Begin VB.TextBox txtSecondPropertyDescription 
         BackColor       =   &H00FFFFFF&
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
         Left            =   6840
         TabIndex        =   10
         Text            =   " "
         Top             =   3480
         Width           =   2175
      End
      Begin VB.TextBox txtSecondPropertyCode 
         BackColor       =   &H00FFFFFF&
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
         Left            =   6840
         TabIndex        =   9
         Top             =   3000
         Width           =   2175
      End
      Begin VB.TextBox txtFirstPropertyDescription 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Left            =   6720
         TabIndex        =   7
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox txtFirstPropertyCde 
         BackColor       =   &H00FFFFFF&
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
         Left            =   6720
         TabIndex        =   6
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Frame Frame3 
         Caption         =   "Second Selection"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   120
         TabIndex        =   3
         Top             =   2640
         Width           =   5295
         Begin MSComctlLib.ListView ListView2 
            Height          =   2055
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   3625
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
      End
      Begin VB.Frame Frame2 
         Caption         =   "First Selection"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   5295
         Begin MSComctlLib.ListView ListView1 
            Height          =   2055
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   3625
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
      End
      Begin VB.Label Label4 
         Caption         =   "Description"
         Height          =   255
         Left            =   5760
         TabIndex        =   12
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Description"
         Height          =   255
         Left            =   5760
         TabIndex        =   11
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Second Propety"
         Height          =   375
         Left            =   5760
         TabIndex        =   8
         Top             =   3120
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "First Property"
         Height          =   255
         Left            =   5760
         TabIndex        =   5
         Top             =   1200
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmODASSitesProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLoad_Click()
With Me
   showALLPROPERTIES
   .cmdPrint.Enabled = True
End With
End Sub

Private Sub cmdPrint_Click()
With Me
  ValidateRECORD
  If bSaveRECORD = True Then
   Load frmRptODASProperties
   frmRptODASProperties.Show vbModal
  End If
End With
End Sub

Private Sub Form_Load()
With Me
  .cmdPrint.Enabled = False
End With
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo err
        Dim i, j As Double
        Dim search As String
        
        If Item.Checked = True Then
            
            j = Screen.ActiveForm.ListView1.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                If Screen.ActiveForm.ListView1.ListItems(i) <> Item Then
                   Screen.ActiveForm.ListView1.ListItems(i).Checked = False
                End If
            Next i
            Me.txtFirstPropertyCde.Text = Item.Text
            Me.txtFirstPropertyDescription.Text = Item.SubItems(1)
            search = Item.Text
            
                showALLProperties2
        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage

End Sub

Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo err
        Dim i, j As Double
        Dim search As String
        
        If Item.Checked = True Then
            
            j = Screen.ActiveForm.ListView1.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                If Screen.ActiveForm.ListView1.ListItems(i) <> Item Then
                   Screen.ActiveForm.ListView1.ListItems(i).Checked = False
                End If
            Next i
            Me.txtSecondPropertyCode.Text = Item.Text
            Me.txtSecondPropertyDescription.Text = Item.SubItems(1)
            search = Item.Text
            
        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub
Private Sub ValidateRECORD()
On Error GoTo err
        With frmODASSitesProperties
                
                If .txtFirstPropertyCde.Text <= " " Then
                    MsgBox " Select One Property From First Property Listing............"
                ElseIf .txtSecondPropertyCode.Text <= " " Then
                    MsgBox "Select One Property From Second Property Listing ..............."
                     
                Else
                        bSaveRECORD = True
                End If
                
        End With
        
Exit Sub
err:
    ErrorMessage
End Sub
