VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmODASMReceiveinvoice 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RECEIVE INVOICE"
   ClientHeight    =   8100
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11055
   Icon            =   "frmODASMReceiveInvoice.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   11055
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Caption         =   "Invoices Received"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2895
      Left            =   120
      TabIndex        =   48
      Top             =   5040
      Width           =   7095
      Begin MSComctlLib.ListView ListView2 
         Height          =   2535
         Left            =   120
         TabIndex        =   49
         Top             =   240
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   4471
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
   Begin VB.Frame Frame3 
      Caption         =   "Invoice Details"
      Height          =   2895
      Left            =   7320
      TabIndex        =   35
      Top             =   5040
      Width           =   3495
      Begin VB.TextBox txtInvoiceVATAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   1560
         TabIndex        =   51
         Top             =   1680
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DTPickerInvoiceDate 
         Height          =   315
         Left            =   3000
         TabIndex        =   50
         Top             =   600
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Format          =   49545217
         CurrentDate     =   38395
      End
      Begin VB.TextBox txtInvoiceNo 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   1560
         TabIndex        =   41
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtInvoiceDate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   1560
         TabIndex        =   40
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtLPOBalance 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1560
         TabIndex        =   39
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox txtVATRate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   1560
         TabIndex        =   38
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtInvoiceAmountExclusive 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   1560
         TabIndex        =   37
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txtInvoiceAmountInclusive 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   1560
         TabIndex        =   36
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label14 
         Caption         =   "VAT Amount"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   1695
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "InvoiceNo"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   255
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Invoice Date"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   615
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "LPO Balance"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   2415
         Width           =   1215
      End
      Begin VB.Label Label31 
         Caption         =   "VAT RATE"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   975
         Width           =   975
      End
      Begin VB.Label Label33 
         Caption         =   "Price Exclusive"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   1335
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Price Inclusive"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   2055
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Requisitions"
      Height          =   2895
      Left            =   7320
      TabIndex        =   10
      Top             =   2160
      Width           =   3495
      Begin VB.TextBox txtPriceExclusive 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1560
         TabIndex        =   33
         Top             =   1748
         Width           =   1695
      End
      Begin VB.TextBox txtPriceInclusive 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1560
         TabIndex        =   23
         Top             =   2505
         Width           =   1695
      End
      Begin VB.TextBox txtAccountNo 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1560
         TabIndex        =   21
         Top             =   1371
         Width           =   1695
      End
      Begin VB.TextBox txtRequisitionDate 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1560
         TabIndex        =   16
         Top             =   994
         Width           =   1695
      End
      Begin VB.TextBox txtRequisitionNo 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1560
         TabIndex        =   15
         Top             =   617
         Width           =   1695
      End
      Begin VB.TextBox txtVATAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1560
         TabIndex        =   13
         Top             =   2125
         Width           =   1695
      End
      Begin VB.TextBox txtLPONo 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1560
         TabIndex        =   11
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Price Exclusive"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   1763
         Width           =   1095
      End
      Begin VB.Label Label32 
         Caption         =   "%"
         Height          =   255
         Left            =   2640
         TabIndex        =   25
         Top             =   3660
         Width           =   135
      End
      Begin VB.Label Label9 
         Caption         =   "Price Inclusive"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label30 
         Caption         =   "Supplier"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1410
         Width           =   735
      End
      Begin VB.Label Label26 
         Caption         =   "Req Date"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1035
         Width           =   735
      End
      Begin VB.Label Label24 
         Caption         =   "Requsition No"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   635
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "VAT"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2140
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "LPO No"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   10695
      Begin VB.TextBox txtCurrentPeriod 
         Alignment       =   1  'Right Justify
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
         Left            =   8760
         TabIndex        =   53
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtBalance 
         Alignment       =   1  'Right Justify
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
         Left            =   8760
         TabIndex        =   31
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtExpiryDate 
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
         Left            =   3480
         TabIndex        =   28
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtProductCode 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1440
         TabIndex        =   26
         Top             =   600
         Width           =   3135
      End
      Begin VB.TextBox txtTotalCost 
         Alignment       =   1  'Right Justify
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
         Left            =   5880
         TabIndex        =   17
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtDescriptionOfOrder 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   6240
         TabIndex        =   6
         Top             =   600
         Width           =   4215
      End
      Begin VB.TextBox txtCommencementDate 
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
         TabIndex        =   5
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtCompanyName 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   240
         Width           =   7335
      End
      Begin VB.Label Label19 
         Caption         =   "Balance"
         Height          =   255
         Left            =   7800
         TabIndex        =   32
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "Total Cost (Incl)"
         Height          =   255
         Left            =   2880
         TabIndex        =   30
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Expires"
         Height          =   255
         Left            =   2880
         TabIndex        =   29
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Product"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "Total Cost"
         Height          =   255
         Left            =   4920
         TabIndex        =   18
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Desc of Order"
         Height          =   255
         Left            =   4680
         TabIndex        =   9
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label23 
         Caption         =   "Commencement"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label25 
         Caption         =   "Customer"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Related Requistions Not Invoiced"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   7095
      Begin MSComctlLib.ListView ListView1 
         Height          =   2535
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   4471
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6840
      Top             =   240
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
            Picture         =   "frmODASMReceiveInvoice.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMReceiveInvoice.frx":0ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMReceiveInvoice.frx":0F0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMReceiveInvoice.frx":1228
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMReceiveInvoice.frx":18A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMReceiveInvoice.frx":1F1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMReceiveInvoice.frx":236E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   1164
      ButtonWidth     =   3069
      ButtonHeight    =   1005
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
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
            Caption         =   "&Search/Find "
            Key             =   "S"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Refresh "
            Key             =   "R"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print &Preview "
            Key             =   "P"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Help"
            Key             =   "F"
            ImageIndex      =   1
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComDlg.CommonDialog HelpCommonDialog 
         Left            =   10560
         Top             =   -120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         HelpCommand     =   11
         HelpContext     =   72
         HelpFile        =   "REGHELP.HLP"
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuClear 
         Caption         =   "Clear the &Screen"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnumm 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuShow 
      Caption         =   "&Show/View"
      Begin VB.Menu mnuClosedJobs 
         Caption         =   "Closed Jobs"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuKJHGFDGFVHJ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFullInventory 
         Caption         =   "Full Inventory"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help System"
      Begin VB.Menu mnuHow 
         Caption         =   "How to use this System"
         Shortcut        =   ^{F1}
      End
   End
End
Attribute VB_Name = "frmODASMReceiveinvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsINVOICE As clsODASMInvoice

Private Sub DTPickerInvoiceDate_Change()
On Error GoTo err
        With frmODASMReceiveinvoice
            .DTPickerInvoiceDate.Value = .txtInvoiceDate.Text
        End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub Form_Activate()
        Set rsINVOICE = New clsODASMInvoice
        disableALLRECORD
        getREQUISITIONS
        rsINVOICE.loadRECORD
        rsINVOICE.loadREQUISTION
        rsINVOICE.calculateLPOBalance
        showLPOINVOICES
        getVatRate
        Set rsINVOICE = Nothing
End Sub

Private Sub Form_Initialize()
        Set rsINVOICE = New clsODASMInvoice
End Sub

Private Sub Form_Load()
        OpenConnection
End Sub

Private Sub Form_Terminate()
        Set rsINVOICE = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
        Set rsINVOICE = Nothing
        If NewRecord = True Then
            Cancel = True
            MsgBox "Data entry in progress. Click Refresh to Cancel", vbCritical
        Else
            Cancel = False
        End If
End Sub
Private Sub getVatRate()
On Error GoTo err
    With Me
        Set rsFindRecord = New ADODB.Recordset
        rsFindRecord.Open "Select * From ODASPVat Where Ending is null", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        If rsFindRecord.EOF And rsFindRecord.BOF Then Exit Sub
            Me.txtVATRate = rsFindRecord!VATRate
        
    End With
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo err
    ListView1.SortKey = ColumnHeader.Index - 1
    ListView1.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo err
        Dim i, j As Double
        
        'If Item.Checked = True And baddRECORD = True Or beditRECORD = True Or bsearchRECORD = True Then

        If Item.Checked = True Then
            
            j = Screen.ActiveForm.ListView1.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                If Screen.ActiveForm.ListView1.ListItems(i) <> Item Then
                   Screen.ActiveForm.ListView1.ListItems(i).Checked = False
                End If
            Next i
            
            Screen.ActiveForm.txtRequisitionNo.Text = Item.Text
            Set rsINVOICE = New clsODASMInvoice
            rsINVOICE.loadREQUISTION
            Set rsINVOICE = Nothing

        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
        
        With frmODASMReceiveinvoice
        Set rsINVOICE = New clsODASMInvoice
        
        Select Case Button.Key
                Case "N"
                    Select Case Button.Caption
                    
                    Case "New &Record "
                            If editRECORD Then Exit Sub
                            NewRecord = True: Button.Caption = "&Save Record ": Button.Image = 4:
                            rsINVOICE.enableRECORD
                    Case "&Save Record "
                    
                            bsaveRECORD = False
                            rsINVOICE.updateRECORD
                            showALLLPOS
                            If bsaveRECORD = True Then
                                        bsaveRECORD = False
                                        .Toolbar1.Buttons(2).Caption = "New &Record ": Button.Image = 2
'                                        .Toolbar1.Buttons(3).Caption = "&NEXT ITEM "
'                                        .Toolbar1.Buttons(4).Caption = "FINISH"
                                          disableALLRECORD
                            End If
                    
                    Case "&NEXT ITEM "
                            
                            .Toolbar1.Buttons(2).Caption = "&Save Record ": Button.Image = 4
                            rsINVOICE.enableRECORD
                    Case Else
                        Exit Sub
                    End Select
    
        Case "E"
            Select Case Button.Caption
                Case "&Edit/Change "
                
                Case "&Save Record "

                        bsaveRECORD = False
                        rsINVOICE.validateRECORD
                        
                        If bsaveRECORD = True Then
                                rsINVOICE.updateRECORD
                                If bsaveRECORD = False Then
                                          .Toolbar1.Buttons(2).Caption = "New &Record "
                                          .Toolbar1.Buttons(3).Caption = "&NEXT ITEM "
                                          .Toolbar1.Buttons(4).Caption = "FINISH"
                                          disableALLRECORD
                                End If
                        End If
                
                Case "&NEXT ITEM "
                            .Toolbar1.Buttons(3).Caption = "&Save Record "
                            rsINVOICE.enableRECORD
                            rsINVOICE.clearRECORD
                Case Else
            End Select
        
        Case "S"
                .Toolbar1.Buttons(2).Caption = "New &Record "
                .Toolbar1.Buttons(2).Image = 2
                .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                .Toolbar1.Buttons(3).Image = 5
                NewRecord = False: editRECORD = False: clearALLRECORD

        Case "R"
            If MsgBox(RefreshMessage, vbExclamation + vbOKCancel + vbDefaultButton2, "Refresh The Screen") = vbCancel Then Exit Sub
                .Toolbar1.Buttons(2).Caption = "New &Record "
                .Toolbar1.Buttons(2).Image = 2
                .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                .Toolbar1.Buttons(3).Image = 5
                NewRecord = False: editRECORD = False: clearALLRECORD
        Case "P"
        Case "F"
            Me.HelpCommonDialog.DialogTitle = "Using the Main System"
            Me.HelpCommonDialog.HelpFile = App.HelpFile
            Me.HelpCommonDialog.HelpContext = 35
            Me.HelpCommonDialog.HelpCommand = cdlHelpContext
            Me.HelpCommonDialog.ShowHelp
     
        Case Else
            Exit Sub
        End Select
        
        Set rsINVOICE = Nothing
        
End With
Exit Sub
err:
    ErrorMessage

End Sub
Private Sub calculatePRICE()
On Error GoTo err
        
        With frmODASMOpenJobCard
            If .txtExchangeRate.Text <= Empty Or CDbl(.txtUnitPrice.Text) <= 0 Then Exit Sub
            .txtItemQuantity.Text = .UpDownQuantity.Value
            .txtTotalUnitPriceExcl.Text = FormatNumber(CDbl(.txtUnitPrice.Text) * CDbl(.txtItemQuantity.Text) * CDbl(.txtExchangeRate.Text))
            .txtVATAmount.Text = FormatNumber(CDbl(.txtTotalUnitPriceExcl.Text) * (CDbl(.txtVATRate) / 100))
            .txtTotalUnitPriceIncl.Text = FormatNumber(CDbl(.txtTotalUnitPriceExcl) + CDbl(.txtVATAmount.Text))
        End With

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub txtInvoiceAmountExclusive_lostfocus()
On Error GoTo err
        With frmODASMReceiveinvoice
            calculateAMOUNT
        End With
Exit Sub
err:
    ErrorMessage
End Sub


Private Sub calculateAMOUNT()
On Error GoTo err
        With frmODASMReceiveinvoice
            If .txtVATRate.Text = Empty Or .txtInvoiceAmountExclusive.Text = 0 Then Exit Sub
            .txtInvoiceVATAmount.Text = CDbl(.txtInvoiceAmountExclusive) * (CDbl(.txtVATRate.Text) / 100)
            .txtInvoiceAmountInclusive.Text = CDbl(.txtInvoiceAmountExclusive.Text) + CDbl(.txtInvoiceVATAmount)
            
            Set rsINVOICE = New clsODASMInvoice
            rsINVOICE.calculateLPOBalance
            Set rsINVOICE = Nothing

        
        End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub txtInvoiceAmountInclusive_LostFocus()
On Error GoTo err
        Set rsINVOICE = New clsODASMInvoice
        rsINVOICE.calculateLPOBalance
        Set rsINVOICE = Nothing
Exit Sub

err:
    ErrorMessage
End Sub
Private Sub txtVATRate_lostFocus()
        calculateAMOUNT
End Sub

