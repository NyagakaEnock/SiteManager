VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmODASPLoadBillBoardPhoto 
   Caption         =   "The BillBoard Photo Loading"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   ScaleHeight     =   5730
   ScaleWidth      =   9255
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Photo To This Site"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   8655
         Begin VB.TextBox txtTransactionNo 
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2040
            TabIndex        =   5
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton cmdSaveImage 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Save"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6120
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   240
            Width           =   2295
         End
         Begin VB.CommandButton cmdInsertImage 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Insert "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3840
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label Label1 
            Caption         =   "Job Brief Number"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   6
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "View Map"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6240
         TabIndex        =   1
         Top             =   5160
         Visible         =   0   'False
         Width           =   2535
      End
      Begin MSComDlg.CommonDialog dlgPHOTO 
         Left            =   4440
         Top             =   1560
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Image imgPHOTO 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   4455
         Left            =   120
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   8655
      End
   End
End
Attribute VB_Name = "frmODASPLoadBillBoardPhoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdInsertImage_Click()
        ShowOpenDialog
End Sub

Private Sub cmdSaveImage_Click()
        saveRecord
End Sub
Private Sub saveRecord()
On Error GoTo err
    With frmODASPLoadBillBoardPhoto
            Set rsCONTROL = New ADODB.Recordset
            
            strSQL = "Select * from ODASPPlotSite Where SiteNo like '" & .txtTransactionNo.Text & "'"
            rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
            If rsCONTROL.BOF Or rsCONTROL.EOF Then
                    rsCONTROL.AddNew
                    ' rsCONTROL!Preparedby = CurrentUserName
                   ' rsCONTROL!dateprepared = Date
            End If
            
            rsCONTROL!photo = .dlgPHOTO.FileName
            bSaveRECORD = False
            
             rsCONTROL.Update
             rsCONTROL.Requery
             MsgBox "Image Saved"
  End With
Exit Sub

err:
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            MsgBox "Update Cancelled! You cannot save this record because some required fields are blank or have invalid data!", vbCritical, "Cancel Update"
            rsCONTROL.CancelUpdate
            rsCONTROL.Requery
    Else
        UpdateErrorMessage
    End If

End Sub

Private Sub cmdView_Click()
  loadIMAGE
End Sub

Private Sub Form_Activate()
        loadIMAGE
        enableALLRECORD
End Sub
Private Sub loadIMAGE()
On Error GoTo err
    
    With frmPQuestionImages
    
        Set rsCONTROL = New ADODB.Recordset
        
        strSQL = "Select * from ODASPPlotSite  Where SiteNo like '" & Me.txtTransactionNo.Text & "'"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        If rsCONTROL.EOF Or rsCONTROL.EOF Then Exit Sub
        
        'If IsNull(rsCONTROL!photo) = False Then
               ' Me.imgPHOTO.Picture = LoadPicture(rsCONTROL!photo)
        'Else
             '   Me.imgPHOTO.Picture = LoadPicture("")
        'End If

    End With

Exit Sub

err:
    UpdateErrorMessage
End Sub

Public Sub ShowOpenDialog()
'If Not NewRecord And Not EditRecord Then Exit Sub
With frmODASPLoadBillBoardPhoto
    
    .imgPHOTO.Picture = LoadPicture(.dlgPHOTO.FileName)
    '.dlgPHOTO.Filter = "JPEG Files (*.jpeg)|*.jpeg|Bitmap Files" & "(*.bmp)|*.bmp" '| All Files (*.*)|*.*"
    .dlgPHOTO.Filter = "Pictures (*.bmp;*.ico;*.jpg)|*.bmp;*.ico;*.jpg"
    .dlgPHOTO.FilterIndex = 2
    .dlgPHOTO.ShowOpen
    If .dlgPHOTO.FileName = "" Then Exit Sub
    .imgPHOTO.Picture = LoadPicture(.dlgPHOTO.FileName)
End With
End Sub

