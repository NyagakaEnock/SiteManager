VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmODASMFreeSite 
   Caption         =   "Free Sssigned Sites"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11595
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   11595
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Site Attatchment Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   11415
      Begin VB.TextBox txtName 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   2400
         TabIndex        =   22
         Top             =   1320
         Width           =   6975
      End
      Begin VB.Frame Frame2 
         Height          =   3135
         Left            =   120
         TabIndex        =   20
         Top             =   2040
         Width           =   11175
         Begin MSComctlLib.ListView ListView1 
            Height          =   2775
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   4895
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
               Name            =   "Arial"
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
      Begin VB.TextBox txtPlotName 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   960
         TabIndex        =   19
         Top             =   600
         Width           =   4095
      End
      Begin VB.TextBox txtEndDate 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   6120
         TabIndex        =   18
         Top             =   1680
         Width           =   3255
      End
      Begin VB.TextBox txtAccountNo 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   960
         TabIndex        =   17
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtJobBriefNo 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   6120
         TabIndex        =   13
         Top             =   960
         Width           =   3255
      End
      Begin VB.TextBox txtSiteDetails 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   6120
         TabIndex        =   12
         Top             =   600
         Width           =   4815
      End
      Begin VB.TextBox txtSiteNo 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   6120
         TabIndex        =   11
         Top             =   240
         Width           =   3255
      End
      Begin VB.TextBox txtStartDate 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   960
         TabIndex        =   8
         Top             =   1680
         Width           =   3015
      End
      Begin VB.TextBox txtPhysicalLocation 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Top             =   960
         Width           =   4095
      End
      Begin VB.TextBox txtPlotNo 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label12 
         Caption         =   "Client"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "End Date"
         Height          =   255
         Left            =   5160
         TabIndex        =   15
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "Site Detail"
         Height          =   255
         Left            =   5160
         TabIndex        =   14
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Site No"
         Height          =   255
         Left            =   5160
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "JobBriefNo"
         Height          =   255
         Left            =   5160
         TabIndex        =   9
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "StartDate"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Plot Name"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Plot No"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Location"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   1695
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   1164
      ButtonWidth     =   3069
      ButtonHeight    =   1005
      TextAlignment   =   1
      ImageList       =   "ImageList1(1)"
      DisabledImageList=   "ImageList1(1)"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Free Site"
            Key             =   "N"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Edit/Change "
            Key             =   "E"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Search/Find "
            Key             =   "S"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Refresh "
            Key             =   "R"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print &Preview "
            Key             =   "P"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Help System  "
            Key             =   "H"
            ImageIndex      =   1
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Index           =   1
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   8
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMFreeSite.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMFreeSite.frx":067A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMFreeSite.frx":0BBC
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMFreeSite.frx":100E
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMFreeSite.frx":1328
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMFreeSite.frx":19A2
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMFreeSite.frx":201C
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMFreeSite.frx":246E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Index           =   0
         Left            =   10920
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
               Picture         =   "frmODASMFreeSite.frx":2AE8
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMFreeSite.frx":3162
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMFreeSite.frx":35B4
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMFreeSite.frx":38CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMFreeSite.frx":3F48
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMFreeSite.frx":45C2
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMFreeSite.frx":4A14
               Key             =   ""
            EndProperty
         EndProperty
      End
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
End
Attribute VB_Name = "frmODASMFreeSite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LoadAccount()
On Error GoTo err
    With Me
        Set rsFindRecord = New ADODB.Recordset
      
        strSQL = "SELECT * FROM ODASMJobBrief,ODASPAccount where ODASMJobBrief.AccountNo=ODASPAccount.AccountNo AND ODASMJobBrief.JobBriefNo='" & .txtJobBriefNo & "';"
        rsFindRecord.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        If rsFindRecord.BOF And rsFindRecord.EOF Then Exit Sub
           .txtAccountNo = rsFindRecord!AccountNo & ""
           .txtName = rsFindRecord!CompanyName & ""
        
        Set rsFindRecord = Nothing
    End With
Exit Sub
err:
    ErrorMessage
End Sub
Private Sub LoadPLOTDetails()
On Error GoTo err
    With Me
               Set rsFIND = New ADODB.Recordset
               strSQL = "SELECT ODASPPlot.PlotNo,ODASPPlotSite.SiteNo,ODASPPlotSite.JobBriefNo,ODASPPlot.PhysicalLocation,ODASPPlot.CommencementDate,ODASPPlot.expirydate FROM ODASPPlot,ODASPPlotSite where ODASPPlot.PlotNo='" & .txtPlotNo & "' AND ODASPPlotSite.SiteNo='" & .txtSiteNo & "';"
               rsFIND.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
               If rsFIND.BOF And rsFIND.EOF Then Exit Sub
                  .txtPhysicalLocation = rsFIND!PhysicalLocation & ""
                  .txtJobBriefNo = rsFIND!JobBriefNo & ""
                  .txtStartDate = Format(rsFIND!CommencementDate, "dd/mm/yyyy") & ""
                  .txtEndDate = Format(rsFIND!expirydate, "dd/mm/yyyy") & ""
    Set rsFIND = Nothing
    End With
Exit Sub
err:
    ErrorMessage
End Sub
Private Function ValidRecord()
On Error GoTo err
    ValidRecord = False
    With frmODASMContractEditing
      If .txtClause.Text = " " Then
          strMessage = "The Clause Number is Mandatory ..........."
          .txtClause.SetFocus
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


Private Sub cmdSaveClause_Click()
With Me
  saveRecord
End With
End Sub



Private Sub ValidateRECORD()
On Error GoTo err
        With Me
                If .txtSiteNo.Text = "" Then
                    MsgBox "The Site No is Required!!"
                    .txtSiteNo.SetFocus
                ElseIf .txtPlotNo.Text = "" Then
                    MsgBox "The Plot No is Required!!"
                    .txtPlotNo.SetFocus
                Else
                        bSaveRECORD = True
                End If
        End With
Exit Sub
err:
    ErrorMessage
End Sub
Private Sub saveRecord()
On Error GoTo err
    With Me
    
            Set rsSAVE = New ADODB.Recordset
            strSQL = "Select * from ODASPPlotSite Where SiteNo = '" & .txtSiteNo.Text & "'"
            rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            If rsSAVE.BOF Or rsSAVE.EOF Then Exit Sub
                rsSAVE!JobBriefNo = ""
                rsSAVE!Status = "SITE-AVAILABLE"
            bSaveRECORD = False
            rsSAVE.Update
            rsSAVE.Requery
  End With
Exit Sub
err:
        UpdateErrorMessage
End Sub

Private Sub Form_Activate()
LoadPLOTDetails
LoadAccount
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
    With Me
        Select Case Button.Key
            Case "N"
                Select Case Button.Caption
                    Case "&Free Site"
                        baddRECORD = True
                        enableALLRECORD
                        Button.Caption = "&Save Record": Button.Image = 5: .Toolbar1.Buttons(4).Caption = "Cancel": .Toolbar1.Buttons(4).Image = 2
                    Case "&Save Record"
                            ValidateRECORD
                            If bSaveRECORD = True Then
                                saveRecord
                                baddRECORD = False: Button.Caption = "&Free Site": Button.Image = 3: .Toolbar1.Buttons(4).Caption = "&Search/Find ": .Toolbar1.Buttons(4).Image = 7
                        End If
                   End Select
            Case "E"
                Select Case Button.Caption
                    Case "&Edit/Change "
                                Button.Caption = "Save &Changes ": Button.Image = 5
                                
                    Case "Save &Changes "
                    
                      
                        
                        beditRECORD = False: Button.Caption = "&Edit/Change ": Button.Image = 6
                    End Select
                    
            Case "P"
                
            Case "S"
                Select Case Button.Caption
                    Case "&Search/Find "
                        bsearchRECORD = True
                        Button.Caption = "Delete"
                    Case "Cancel"
                        cancelCMD
                    Case "Delete"
                          
                        Button.Caption = "&Search/Find "
                     beditRECORD = False: Button.Caption = "Delete": Button.Image = 6
                End Select
            Case "R"
                    If MsgBox(RefreshMessage, vbExclamation + vbOKCancel + vbDefaultButton2, "Refresh The Screen") = vbCancel Then Exit Sub
                    .Toolbar1.Buttons(2).Caption = "New &Record "
                    .Toolbar1.Buttons(2).Image = 3
                    .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                    .Toolbar1.Buttons(3).Image = 6
                    .Toolbar1.Buttons(4).Caption = "&Search/Find ": .Toolbar1.Buttons(4).Image = 7
                    NewRecord = False: EditRecord = False: MyCommonData.ClearTheScreen

            Case "H"
                Me.HelpCommonDialog.DialogTitle = "Using the Main System"
                Me.HelpCommonDialog.HelpFile = App.HelpFile
                Me.HelpCommonDialog.HelpContext = 19
                Me.HelpCommonDialog.HelpCommand = cdlHelpContext
                Me.HelpCommonDialog.ShowHelp
        End Select
    End With
Exit Sub
err:
ErrorMessage
End Sub



Private Sub txtAccountNo_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub



Private Sub txtEndDate_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtJobBriefNo_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub



Private Sub txtName_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtPhysicalLocation_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtPlotName_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtPlotNo_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtSiteDetails_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtSiteNo_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub


Private Sub txtStartDate_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
