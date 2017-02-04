VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmODASPDeliveryMethod 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GENERAL INVENTORY- Delivery Methods"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9780
   Icon            =   "frmODASPDeliveryMethod.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   9780
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView ListView1 
      Height          =   2415
      Left            =   240
      TabIndex        =   8
      ToolTipText     =   "List of current shipping methods"
      Top             =   3600
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   4260
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777152
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox txtDeliveryMethodDescription 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1920
      Width           =   6015
   End
   Begin VB.TextBox txtDeliveryMethod 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   1200
      Width           =   6015
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
      BorderStyle     =   1
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
         Height          =   495
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Width           =   2295
      End
      Begin VB.CommandButton cmdREFRESH 
         BackColor       =   &H80000000&
         Caption         =   "&REFRESH"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4830
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Clear screen"
         Top             =   0
         Width           =   2655
      End
      Begin VB.CommandButton cmdEDIT 
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
         Height          =   495
         Left            =   2415
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Change existing Record"
         Top             =   0
         Width           =   2520
      End
      Begin VB.CommandButton cmdNEW 
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
         Height          =   495
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Add new record"
         Top             =   0
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "List of Current Delivery Methods"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   9
      Top             =   3240
      Width           =   9615
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   240
      X2              =   9120
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label3 
      Caption         =   "Description/Comments"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Delivery Method"
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
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   2055
   End
End
Attribute VB_Name = "frmODASPDeliveryMethod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MyDelivery As clsODASDeliveryMethods

Private Sub cmdEdit_Click()
''On Error GoTo Err
If NewRecord Then Exit Sub
Select Case cmdEdit.Caption
    Case "E&DIT"
        MyDelivery.FindRecord
        'If Not AllowEdit Then Exit Sub
      
    Case "SAVE &CHANGES"
        If EditRecord Then
            If ValidRecord Then
                MyDelivery.EditCurrentRecord
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
    Select Case cmdNEW.Caption
    Case "&NEW"
        MyDelivery.ClearTheScreen
        MyDelivery.AddNewRecord
    Case "&SAVE RECORD"
    If NewRecord Then
        If ValidRecord Then
            MyDelivery.SaveNewRecord
            showDeliveryMethodS
        End If
    End If
    Case Else
    Exit Sub
    End Select
    Exit Sub
err:
    ErrorMessage
End Sub

Private Function ValidRecord() As Boolean
''On Error GoTo Err
With Me
    If .txtDeliveryMethod.Text = Empty Then
        strMessage = "Required Shipping Method!!"
        .txtDeliveryMethod.SetFocus
    ElseIf .txtDeliveryMethodDescription.Text = Empty Then
        strMessage = "Required Shipping Description Method!!"
        .txtDeliveryMethodDescription.SetFocus
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

Private Sub cmdPrint_Click()
Load frmRPTDeliveryMethods
frmRPTDeliveryMethods.Show 1, Me
End Sub

Private Sub cmdRefresh_Click()
If MsgBox(RefreshMessage, vbQuestion + vbYesNo + vbDefaultButton2, "Screen Refresher") = vbNo Then Exit Sub
    NewRecord = False
    EditRecord = False
    MyDelivery.ClearTheScreen
    With Me
        .cmdEdit.Caption = "E&DIT"
        .cmdNEW.Caption = "&NEW"
        .cmdEdit.Enabled = True
    End With
End Sub

Private Sub Form_Activate()
        showDeliveryMethodS
End Sub

Private Sub Form_Initialize()
    Set MyDelivery = New clsODASDeliveryMethods
End Sub


Private Sub Form_Terminate()
    Set MyDelivery = Nothing
End Sub
