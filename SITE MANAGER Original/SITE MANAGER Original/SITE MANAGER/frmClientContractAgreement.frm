VERSION 5.00
Begin VB.Form frmClientContractAgreement 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Client contract agreement form"
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3750
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   120
      TabIndex        =   13
      Top             =   5880
      Width           =   11655
      Begin VB.TextBox txtConclusion 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   120
         Width           =   11415
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   11655
      Begin VB.TextBox txtBody1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   1575
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   1800
         Width           =   11295
      End
      Begin VB.TextBox txtBody 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   1575
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   240
         Width           =   11415
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Agreement No"
      Height          =   855
      Left            =   0
      TabIndex        =   2
      Top             =   7320
      Width           =   11895
      Begin VB.CommandButton cmdClose 
         Caption         =   "Clo&se"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9840
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Refresh"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7800
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txtAgreementNo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   120
         X2              =   11520
         Y1              =   240
         Y2              =   240
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   11655
      Begin VB.TextBox txtPreface 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   120
         Width           =   11415
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "CONCLUSION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   5640
      Width           =   11655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "BODY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   11655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "TITLE AND PREFACE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   11775
   End
End
Attribute VB_Name = "frmClientContractAgreement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
On Error GoTo err
With Me
  If MsgBox("This action is going to abort your current procedure ,you will loose any unsaved data  Are you sure you want to continue ?", vbYesNo + vbExclamation, "Cancel") = vbYes Then
    .cmdNew.Caption = "&New"
    .cmdNew.Enabled = True
    .cmdEdit.Caption = "&Edit"
    .cmdEdit.Enabled = True
    .txtPreface.Text = ""
    .txtBody.Text = ""
    .txtBody1.Text = ""
    .txtConclusion.Text = ""
    .txtAgreementNo.Text = ""
    NewRecord = False
    EditRecord = False
  Else

   Exit Sub
End If
End With
Exit Sub
err:
ErrorMessage
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdEdit_Click()
On Error GoTo err
Dim ClauseNo As Variant
With Me
If NewRecord Then Exit Sub
 EditRecord = True
  .cmdNew.Enabled = False
Select Case .cmdEdit.Caption
  Case "&Edit"
  
  ClauseNo = InputBox("Please enter the number of the Agreement you want to edit", "Client Contract Agreement")
  
  If ClauseNo = "" Then
  MsgBox ("No values were entered the system will abort the operation"), vbCritical, "Invalid No"
  Exit Sub
    Else
  Set rsFindRecord = New ADODB.Recordset
     rsFindRecord.Open "SELECT * FROM AdvertClientAgreement WHERE AgreementNo = '" & ClauseNo & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
      
      If rsFindRecord.EOF And rsFindRecord.BOF Then
      MsgBox "The Agreement Number Does not exit or entry is not correct", vbCritical, "Invalid No"
        Set rsFindRecord = Nothing: Exit Sub
        Else
        .txtAgreementNo.Text = rsFindRecord!AgreementNo & ""
        .txtPreface.Text = rsFindRecord!Preface & ""
        .txtBody.Text = rsFindRecord!Body & ""
        .txtBody1.Text = rsFindRecord!Body1 & ""
        .txtConclusion.Text = rsFindRecord!Conclusion & ""
      End If
   End If
        Set rsFindRecord = Nothing
        .cmdEdit.Caption = "&Save"
     Case "&Save"
     If ValidRecord Then
       Set rsLineUpdate = New ADODB.Recordset
       rsLineUpdate.Open "UPDATE AdvertClientAgreement SET Preface = '" & Trim(.txtPreface.Text) & "',Body = '" & Trim(.txtBody.Text) & "',Body1 = '" & Trim(.txtBody1.Text) & "',Conclusion = '" & Trim(.txtConclusion.Text) & "',DateModified = '" & MyCurrentDate & "',ModifiedBy = '" & CurrentUserName & "' WHERE AgreementNo = '" & Trim(.txtAgreementNo.Text) & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
       Set rsLineUpdate = Nothing
     End If
     
     .cmdEdit.Caption = "&Edit"
     .cmdNew.Enabled = True
      EditRecord = False
     .txtPreface.Text = Empty
     .txtAgreementNo.Text = Empty
     .txtBody.Text = Empty
     .txtBody1.Text = Empty
     .txtConclusion.Text = Empty
     MsgBox "Changes Requested Successfully Effected", vbInformation, "Edit Complete"
     
     Case Else
     Exit Sub
End Select
End With
Exit Sub
err:
   ErrorMessage
End Sub

Private Sub cmdNew_Click()
On Error GoTo err
With Me
If EditRecord Then Exit Sub
  NewRecord = True
  .cmdEdit.Enabled = False
Select Case .cmdNew.Caption
   Case "&New"
   .txtAgreementNo.Text = ""
   .txtPreface.Text = ""
   .txtBody.Text = ""
   .txtBody1.Text = ""
   .txtConclusion.Text = ""
   .cmdNew.Caption = "&Save"
   
  Case "&Save"
    If ValidRecord Then
       Set rsNewRecord = New ADODB.Recordset
       rsNewRecord.Open "INSERT INTO AdvertClientAgreement(AgreementNo,Preface,Body,Body1,Conclusion,CreatedBy,DateCreated,AccPeriod) values('" & Trim(.txtAgreementNo.Text) & "','" & Trim(.txtPreface.Text) & "','" & Trim(.txtBody.Text) & "','" & Trim(.txtBody1.Text) & "','" & Trim(.txtConclusion.Text) & "','" & CurrentUserName & "','" & MyCurrentDate & "','" & MyCurrentPeriod & "')", cnCOMMON, adOpenKeyset, adLockOptimistic
    Else
    Exit Sub
    End If
     Set rsNewRecord = Nothing
    
    .cmdNew.Caption = "&New"
    .cmdEdit.Enabled = True
     NewRecord = False
    
  Case Else
  Exit Sub
End Select
End With
Exit Sub
err:
    ErrorMessage
End Sub
Private Function ValidRecord() As Boolean
'On Error Resume Next
With Me
 If .txtAgreementNo.Text = "" Then
    MsgBox "Agreement No is Required"
  .txtAgreementNo.SetFocus
 ElseIf .txtPreface.Text = "" Then
  MsgBox "Preface Required"
  .txtPreface.SetFocus
 ElseIf .txtBody.Text = "" Then
  MsgBox "Body Required"
  .txtBody.SetFocus
 ElseIf .txtBody1.Text = "" Then
  MsgBox "Body Required"
  .txtBody1.SetFocus
 ElseIf .txtConclusion.Text = "" Then
 MsgBox "Conclusion Required"
 .txtConclusion.SetFocus
 Else
 ValidRecord = True
 End If
  
End With
End Function

