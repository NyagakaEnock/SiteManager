VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmODASMIssueOrder 
   Caption         =   "Issue/Cancel Order"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10260
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   10260
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   10095
      Begin VB.Frame Frame4 
         Caption         =   "Transaction Details"
         Height          =   2415
         Left            =   6720
         TabIndex        =   23
         Top             =   120
         Width           =   3255
         Begin VB.OptionButton optCancel 
            Caption         =   "Cancel?"
            Height          =   255
            Left            =   1800
            TabIndex        =   34
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton optIssue 
            Caption         =   "Issue?"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox txtCancelledBy 
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
            Left            =   1440
            TabIndex        =   29
            Top             =   1680
            Width           =   1575
         End
         Begin VB.TextBox txtDateCancelled 
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
            Left            =   1440
            TabIndex        =   28
            Top             =   2040
            Width           =   1575
         End
         Begin VB.TextBox txtIssueBy 
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
            Left            =   1440
            TabIndex        =   25
            Top             =   960
            Width           =   1575
         End
         Begin VB.TextBox txtDateIssued 
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
            Left            =   1440
            TabIndex        =   24
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label9 
            Caption         =   "DateCancelled"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   2040
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Cancelled By"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   1695
            Width           =   975
         End
         Begin VB.Label Label5 
            Caption         =   "Date Issued"
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Issued By"
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   975
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "LPO Entries"
         Height          =   2655
         Left            =   120
         TabIndex        =   21
         Top             =   2520
         Width           =   9855
         Begin MSComctlLib.ListView ListView2 
            Height          =   2295
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   4048
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
      Begin VB.Frame Frame2 
         Height          =   2415
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   6495
         Begin VB.TextBox txtOrderDescription 
            BackColor       =   &H00FFC0C0&
            Height          =   285
            Left            =   1320
            TabIndex        =   11
            Top             =   600
            Width           =   5055
         End
         Begin VB.TextBox txtDeadlineDate 
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
            Left            =   1320
            TabIndex        =   10
            Top             =   1680
            Width           =   1695
         End
         Begin VB.TextBox txtOrderDate 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   4800
            TabIndex        =   9
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txtOrderNo 
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   1320
            TabIndex        =   8
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtTotalCost 
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
            Left            =   4800
            TabIndex        =   7
            Top             =   1680
            Width           =   1575
         End
         Begin VB.TextBox txtRemarks 
            BackColor       =   &H00FFC0C0&
            Height          =   285
            Left            =   1320
            MultiLine       =   -1  'True
            TabIndex        =   6
            Top             =   1320
            Width           =   5055
         End
         Begin VB.TextBox txtSupplierName 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   2520
            TabIndex        =   5
            Top             =   960
            Width           =   3855
         End
         Begin VB.TextBox txtSupplierCode 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   1320
            TabIndex        =   4
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox txtTotalCostInclusive 
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
            Left            =   4800
            TabIndex        =   3
            Top             =   2040
            Width           =   1575
         End
         Begin VB.TextBox txtTotalVATAmount 
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
            Left            =   1320
            TabIndex        =   2
            Top             =   2040
            Width           =   1695
         End
         Begin VB.Label Label25 
            Caption         =   "Description"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label23 
            Caption         =   "Deadline Date"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   1695
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "Order Date"
            Height          =   255
            Left            =   3720
            TabIndex        =   18
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Order No"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label16 
            Caption         =   "Total Cost"
            Height          =   255
            Left            =   3720
            TabIndex        =   16
            Top             =   1695
            Width           =   1575
         End
         Begin VB.Label Label22 
            Caption         =   "Remarks"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "Supplier "
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "Total Cost Inc"
            Height          =   255
            Left            =   3720
            TabIndex        =   13
            Top             =   2055
            Width           =   1575
         End
         Begin VB.Label Label8 
            Caption         =   "VAT Amount"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   2055
            Width           =   1335
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   10260
      _ExtentX        =   18098
      _ExtentY        =   1164
      ButtonWidth     =   3069
      ButtonHeight    =   1005
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New &Record "
            Key             =   "N"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Edit/Change "
            Key             =   "E"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Refresh "
            Key             =   "R"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print &Preview "
            Key             =   "P"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Help"
            Key             =   "F"
            ImageIndex      =   1
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   9120
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
               Picture         =   "frmODASMIssueOrder.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMIssueOrder.frx":067A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMIssueOrder.frx":0ACC
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMIssueOrder.frx":0DE6
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMIssueOrder.frx":1460
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMIssueOrder.frx":1ADA
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMIssueOrder.frx":1F2C
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
Attribute VB_Name = "frmODASMIssueOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsREQ As New clsODASPurchaseOrder

Private Sub optIssue_Click()
'On Error GoTo err
        Set rsREQ = New clsODASPurchaseOrder
        rsREQ.CheckDetails
        Set rsREQ = Nothing
 
Exit Sub
err:
    ErrorMessage
End Sub
Private Sub optcancel_Click()
'On Error GoTo err
        Set rsREQ = New clsODASPurchaseOrder
        rsREQ.CheckDetails
        Set rsREQ = Nothing

Exit Sub
err:
    ErrorMessage
End Sub

Private Sub Form_Activate()
        Set rsREQ = New clsODASPurchaseOrder
        rsREQ.loadISSUED
        Set rsREQ = Nothing
        loadSUPPLIER
        showORDERITEMS
        disableALLRECORD
End Sub

Private Sub Form_Load()
    OpenODBCConnection
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo err
        
        With Screen.ActiveForm
        
        Select Case Button.Key
                Case "N"
                    Select Case Button.Caption
                    
                    Case "New &Record "
                        If EditRecord Then Exit Sub
                        .ListView2.ListItems.Clear:
                        NewRecord = True: Button.Caption = "&Save Record ": Button.Image = 4
                        enableALLRECORD
                    
                    Case "&Save Record "
                        If NewRecord Then
                            Set rsREQ = New clsODASPurchaseOrder
                            
                            rsREQ.validIssue
                            
                            If bSaveRECORD = True Then
                                    
                                    rsREQ.updateISSUED
                            
                                    If bSaveRECORD = False Then
                                            Button.Caption = "&Next Item"
                                            rsREQ.ClearIssued
                                            frmODASMOrder.ListView1.SetFocus
                                            .Toolbar1.Buttons(3).Caption = "FINISH"
                                    End If
                            End If
                            
                            Set rsREQ = Nothing

                        End If
                    
                    Case "&Next Item"
                            Button.Caption = "&Save Record ": Button.Image = 4
                            enableALLRECORD
                    Case Else
                        Exit Sub
                    End Select
    
        Case "E"
            Select Case Button.Caption
                Case "&Edit/Change "
                        editMYRECORD
                        Button.Caption = "&Save Record ": Button.Image = 4
                        enableALLRECORD

                Case "FINISH"
                        .Toolbar1.Buttons(2).Caption = "New &Record "
                        .Toolbar1.Buttons(2).Image = 2
                        .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                        .Toolbar1.Buttons(3).Image = 5
                        NewRecord = False: EditRecord = False: clearALLRECORD
                Case Else
            End Select
        
        Case "S"
            If MsgBox(RefreshMessage, vbExclamation + vbOKCancel + vbDefaultButton2, "Refresh The Screen") = vbCancel Then Exit Sub
                .Toolbar1.Buttons(2).Caption = "New &Record "
                .Toolbar1.Buttons(2).Image = 2
                .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                .Toolbar1.Buttons(3).Image = 5
                NewRecord = False: EditRecord = False: clearALLRECORD
  
        Case "R"
            If MsgBox(RefreshMessage, vbExclamation + vbOKCancel + vbDefaultButton2, "Refresh The Screen") = vbCancel Then Exit Sub
                .Toolbar1.Buttons(2).Caption = "New &Record "
                .Toolbar1.Buttons(2).Image = 2
                .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                .Toolbar1.Buttons(3).Image = 5
                NewRecord = False: EditRecord = False: clearALLRECORD
        Case "P"
        Case "F"
     
     
        Case Else
            Exit Sub
        End Select
End With
Exit Sub
err:
    ErrorMessage

End Sub

