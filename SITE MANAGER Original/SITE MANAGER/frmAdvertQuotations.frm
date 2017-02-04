VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmAdvertQuotations 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Advertisement Quotations"
   ClientHeight    =   6810
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11160
   Icon            =   "frmAdvertQuotations.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   11160
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtContactTitle 
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   6840
      TabIndex        =   44
      Top             =   3000
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker dtpQuotationDate 
      Height          =   375
      Left            =   9600
      TabIndex        =   10
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   57278465
      CurrentDate     =   38283
   End
   Begin VB.TextBox txtQuotationNo 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Height          =   405
      Left            =   7080
      TabIndex        =   8
      Top             =   720
      Width           =   1815
   End
   Begin VB.Frame Frame4 
      Caption         =   "Item(s) Quoted"
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
      Height          =   3375
      Left            =   5760
      TabIndex        =   6
      Top             =   3480
      Width           =   5295
      Begin VB.CheckBox chkBorder 
         Caption         =   "Bordered"
         Height          =   195
         Left            =   1200
         TabIndex        =   42
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CheckBox chkIlluminate 
         Caption         =   "Illuminated"
         Height          =   255
         Left            =   3720
         TabIndex        =   41
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txtNotes 
         BackColor       =   &H00FFC0C0&
         Height          =   1215
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   39
         ToolTipText     =   "Can allow maximum of 1000 characters"
         Top             =   1920
         Width           =   4095
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Left            =   4080
         Max             =   0
         Min             =   32767
         TabIndex        =   38
         Top             =   960
         Width           =   255
      End
      Begin VB.TextBox txtQuantity 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   4320
         TabIndex        =   37
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtTotalPrice 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   3720
         TabIndex        =   35
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtPrice 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1080
         TabIndex        =   33
         Top             =   1320
         Width           =   1575
      End
      Begin VB.ComboBox cboSiding 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1080
         TabIndex        =   30
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox txtItemName 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1080
         TabIndex        =   28
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox txtItemCode 
         BackColor       =   &H00FFFFC0&
         Height          =   255
         Left            =   4200
         TabIndex        =   26
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtCategoryCode 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4200
         TabIndex        =   24
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cboCategory 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1080
         TabIndex        =   22
         Top             =   240
         Width           =   2415
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   120
         X2              =   5160
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Label Label22 
         Caption         =   "Notes"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label21 
         Caption         =   "Quantity"
         Height          =   255
         Left            =   3360
         TabIndex        =   36
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label20 
         Caption         =   "Total Price"
         Height          =   255
         Left            =   2880
         TabIndex        =   34
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label19 
         Caption         =   "Price"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "Code"
         Height          =   255
         Left            =   3600
         TabIndex        =   31
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label12 
         Caption         =   "Siding"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Item Name"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Code"
         Height          =   255
         Left            =   3600
         TabIndex        =   25
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "Category "
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Client Details"
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
      Height          =   2175
      Left            =   5760
      TabIndex        =   5
      Top             =   1320
      Width           =   5295
      Begin VB.TextBox cboCity 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   3360
         TabIndex        =   47
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtFax 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   3720
         TabIndex        =   46
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txtContactPerson 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   1080
         TabIndex        =   15
         Top             =   1320
         Width           =   4095
      End
      Begin VB.TextBox txtTelephone 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   1080
         TabIndex        =   14
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtMobile 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   3600
         TabIndex        =   13
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtAddress 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   1080
         TabIndex        =   12
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   1080
         TabIndex        =   11
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label Label14 
         Caption         =   "Fax"
         Height          =   255
         Left            =   3240
         TabIndex        =   45
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label13 
         Caption         =   "Cont. Title"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Mobile"
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   3000
         TabIndex        =   21
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Cont.Person"
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Telephone"
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Address"
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   " Name"
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "List Of Items"
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
      Height          =   3135
      Left            =   120
      TabIndex        =   3
      Top             =   3600
      Width           =   5535
      Begin MSComctlLib.ListView ListView2 
         Height          =   2775
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   4895
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
            Name            =   "Times New Roman"
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
      Left            =   7080
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
            Picture         =   "frmAdvertQuotations.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdvertQuotations.frx":0ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdvertQuotations.frx":0F0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdvertQuotations.frx":1228
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdvertQuotations.frx":18A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdvertQuotations.frx":1F1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdvertQuotations.frx":236E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11160
      _ExtentX        =   19685
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
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         HelpCommand     =   11
         HelpContext     =   72
         HelpFile        =   "REGHELP.HLP"
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "List Of Clients"
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
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   5535
      Begin MSComctlLib.ListView ListView1 
         Height          =   2415
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   4260
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
            Name            =   "Times New Roman"
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
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   495
      Left            =   5760
      TabIndex        =   16
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   5760
      X2              =   11040
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label3 
      Caption         =   "Date"
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
      Height          =   255
      Left            =   9000
      TabIndex        =   9
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Quotation No"
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
      Height          =   255
      Left            =   5760
      TabIndex        =   7
      Top             =   720
      Width           =   1215
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
      Begin VB.Menu mnuRegisteredClients 
         Caption         =   "Registered Clients"
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
Attribute VB_Name = "frmAdvertQuotations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim billbno, CustomerCode, PhysicalAddress




Private Sub cboBBNo_GotFocus()

End Sub

Private Sub cboBBNo_LostFocus()

End Sub

Private Sub cboCouncilPeriod_Click()
Screen.ActiveForm.ListView1.SetFocus
End Sub



Private Sub cboCountry_Click()
Screen.ActiveForm.ListView1.SetFocus
End Sub

Private Sub cboElecPeriod_Click()
Screen.ActiveForm.ListView1.SetFocus
End Sub


Private Sub cboLandLord_Click()
Screen.ActiveForm.ListView1.SetFocus
End Sub

Private Sub cboRentPeriod_Click()
Screen.ActiveForm.ListView1.SetFocus
End Sub



Private Sub cboTown_Click()
        Screen.ActiveForm.ListView1.SetFocus
End Sub

Private Sub Combo1_Click()
        Screen.ActiveForm.ListView1.SetFocus
End Sub

Private Sub cboCategory_GotFocus()
If Not NewRecord Then Exit Sub
With Screen.ActiveForm
    AttachSQL = "SELECT CategoryName AS SelectField FROM AdvertCategories ORDER BY CategoryName;"
    .cboCategory.Clear
    MyCommonData.AttachDropDown
End With

End Sub

Private Sub cboCity_GotFocus()
If Not NewRecord Then Exit Sub
With Screen.ActiveForm
    AttachSQL = "SELECT Town AS SelectField FROM ODASPTown ORDER BY Town;"
    .cboCity.Clear
    MyCommonData.AttachDropDown
End With

End Sub

Private Sub cboSiding_GotFocus()
If Not NewRecord Then Exit Sub
With Screen.ActiveForm
    AttachSQL = "SELECT SidingDescription AS SelectField FROM Advertsiding ORDER BY sidingdescription;"
    .cboSiding.Clear
    MyCommonData.AttachDropDown
End With

End Sub

Private Sub cboCategory_Click()
Screen.ActiveForm.ListView2.SetFocus
End Sub
Private Sub cboCategory_LostFocus()
'''On Error GoTo Err
With Screen.ActiveForm

    Set rsFindRecord = cnCOMMON.Execute("SELECT * FROM AdvertCategories WHERE CategoryName='" & Trim(.cboCategory.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        Set rsFindRecord = Nothing: Exit Sub
    Else
        .txtCategoryCode.Text = rsFindRecord!categorycode & ""
        .cboCategory.Text = rsFindRecord!categoryname & ""
        
        .ListView2.SetFocus
        
        If .txtCategoryCode.Text = "AAA" Then
            Call ShowAllBillBoardCategories
        Else
            Call ShowBillBoardsPerCategory
        End If
        
    End If
    
End With
Exit Sub
err:
    ErrorMessage
End Sub
Private Function GetAdvertPrice()
'''On Error GoTo Err
With Screen.ActiveForm
    Set rsFindRecord = cnCOMMON.Execute("SELECT * FROM AdvertPricing WHERE BBNo='" & Trim(.txtItemCode.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        Set rsFindRecord = Nothing
    Else
        GetAdvertPrice = rsFindRecord!BBCharges & ""
     
        
    End If
End With
Exit Function
err:
    ErrorMessage
End Function

Private Sub ShowBillBoardsPerCategory()
'''On Error GoTo Err
With Screen.ActiveForm
.ListView2.ListItems.Clear
.ListView2.ColumnHeaders.Clear

.ListView2.ColumnHeaders.Add , , "Category Code", .ListView2.Width / 6#
.ListView2.ColumnHeaders.Add , , "Category Name", .ListView2.Width / 4
.ListView2.ColumnHeaders.Add , , "Item Code", .ListView2.Width / 6
.ListView2.ColumnHeaders.Add , , "Item Name", .ListView2.Width / 4
.ListView2.ColumnHeaders.Add , , "Length", .ListView2.Width / 6.5
.ListView2.ColumnHeaders.Add , , "Width", .ListView2.Width / 6.5

.ListView2.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM AdvertBBDetails A ,AdvertCategories B WHERE A.CategoryCode = B.CategoryCode AND A.CategoryCode = '" & Trim(.txtCategoryCode.Text) & "' ORDER BY A.Name;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView2.View = lvwList
    Set MyList = .ListView2.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing:  Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!categorycode))


    If Not IsNull(rsLIST!categoryname) Then
        MyList.SubItems(1) = CStr(rsLIST!categoryname)
    End If
     
    If Not IsNull(rsLIST!BillBoardNo) Then
        MyList.SubItems(2) = CStr(rsLIST!BillBoardNo)
    End If
    
    If Not IsNull(rsLIST!Name) Then
        MyList.SubItems(3) = CStr(rsLIST!Name)
    End If
    
    If Not IsNull(rsLIST!Length) Then
        MyList.SubItems(4) = CStr(rsLIST!Length)
    End If
    
    If Not IsNull(rsLIST!Width) Then
        MyList.SubItems(5) = CStr(rsLIST!Width)
    End If
    
    rsLIST.MoveNext
    
Wend


Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
err:
If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub
Private Sub cboCity_Click()
        Screen.ActiveForm.ListView1.SetFocus
End Sub

Private Sub cboSiding_Click()
Screen.ActiveForm.ListView2.SetFocus
End Sub

Private Sub cboSiding_LostFocus()
'''On Error GoTo Err
With Screen.ActiveForm

    Set rsFindRecord = cnCOMMON.Execute("SELECT * FROM AdvertSiding WHERE SidingDescription ='" & Trim(.cboSiding.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        Set rsFindRecord = Nothing: Exit Sub
    Else
        .cboSiding.Text = rsFindRecord!SidingType & ""
        .txtQuantity.SetFocus
    End If
    End With
  Exit Sub
err:
     ErrorMessage
End Sub

Private Sub Check1_Click()

End Sub

Private Sub Form_Activate()
    ShowAllClients
End Sub

Private Sub Form_Initialize()
Set MyCommonData = New clsCommonData
'Set mycommondata New clsCommonData

End Sub

Private Sub Form_Load()
    OpenODBCConnection
End Sub
Public Function AutoPurchaseOrderNo() As String
''''On Error GoTo Err
With Screen.ActiveForm

Dim rsLastID As ADODB.Recordset 'used to retrieve current LastId in the Table
Dim strLastID As String 'SQL statement

Dim strTemp As String 'store current record
Dim iNumPos As Integer 'store position of the first numeral
Dim strPrefix As String 'stores Id Prefix

'Retrieve the last record in the recrdset where order is ascending

'strLastID = "SELECT max(cnCliniclDetails.cnCliniclID) as lastid from cnCliniclDetails"
strLastID = "SELECT MAX(QuotationNo) AS LastID FROM AdvertQuotation;"
Set rsLastID = New ADODB.Recordset

With rsLastID
'open the recordset
    .Open strLastID, cnCOMMON, adOpenKeyset, adLockOptimistic
    If .EOF And .BOF Then 'shows empty recordset
        AutoPurchaseOrderNo = "QT00001" 'format of desired format of the string value
    ElseIf IsNull(!lastid) = True Or !lastid = "" Then
        AutoPurchaseOrderNo = "QT00001"
    Else
       ' If .EOF And .BOF Then .MoveFirst
        '.MoveLast
        strTemp = !lastid
        iNumPos = 1
        Dim sChar As String
        Dim iIDLen As Integer
        iIDLen = Len(strTemp)
        sChar = Mid(strTemp, iNumPos, 1)
        While InStr("1234567890", sChar) = 0
            iNumPos = iNumPos + 1
            sChar = Mid(strTemp, iNumPos, 1)
        Wend
        'store the ID prefix eg AP
        strPrefix = Left(strTemp, iNumPos - 1)
        'store the number portion eg and the length with leading Zeros
        strTemp = Right(strTemp, Len(strTemp) + 1 - iNumPos)
        strTemp = Format(Int(strTemp) + 1, String(iIDLen + 1 - iNumPos, "0"))
        AutoPurchaseOrderNo = strPrefix & strTemp
    End If
End With
End With
    Exit Function
err:
    ErrorMessage
End Function

Private Sub Label18_Click()
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
''On Error GoTo Err
        If Not NewRecord And Not EditRecord Then Item.Checked = False: Exit Sub
        If Screen.ActiveForm.ListView1.ListItems.Count = 0 Or Screen.ActiveForm.ListView1.View <> lvwReport Then Exit Sub
    
        Dim i, j, k
        j = Screen.ActiveForm.ListView1.ListItems.Count
        
        If j = 0 Then Exit Sub
        
        For i = 1 To j
            If Screen.ActiveForm.ListView1.ListItems(i).Text <> Item Then
                Screen.ActiveForm.ListView1.ListItems(i).Checked = False
            End If
        Next i
    
        If Item.Checked = True Then
                loadCUSTOMERDETAILS
        End If
    
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
''On Error GoTo Err
'If Not NewRecord And Not EditRecord Then Item.Checked = False: Exit Sub
If Screen.ActiveForm.ListView2.ListItems.Count = 0 Or Screen.ActiveForm.ListView2.View <> lvwReport Then Exit Sub
    
    Dim i, j, k
    j = Screen.ActiveForm.ListView2.ListItems.Count
    
    If j = 0 Then Exit Sub
    
    For i = 1 To j
        If Screen.ActiveForm.ListView2.ListItems(i).Text <> Item Then
            Screen.ActiveForm.ListView2.ListItems(i).Checked = False
        End If
    Next i
    
      If Item.Checked = True Then
        
'        ClearTextFields
        
        Screen.ActiveForm.txtCategoryCode.Text = Item
        Screen.ActiveForm.cboCategory.Text = Item.SubItems(1)
        Screen.ActiveForm.txtItemCode.Text = Item.SubItems(2)
        Screen.ActiveForm.txtItemName.Text = Item.SubItems(3)
        Screen.ActiveForm.txtPrice.Text = GetAdvertPrice
        
        
        
    ElseIf Item.Checked = False Then
    
'        ClearTextFields
        
    End If
    
Exit Sub
err:
    ErrorMessage
End Sub
Public Sub ClearTextFields()
For Each i In Screen.ActiveForm
    If TypeOf i Is TextBox And i.Name <> "txtTitle" Then
        i.Text = Empty
    End If
    If TypeOf i Is ComboBox Then
        i.Clear
    End If
    If TypeOf i Is Image Then
        i.Picture = LoadPicture("")
    End If
Next i
End Sub
Private Sub mnuClear_Click()
    MyCommonData.ClearTextFields
End Sub

Private Sub mnuClose_Click()
    Unload Screen.ActiveForm
End Sub

Private Sub mnuCurrent_Click()
    Call ShowCurrentSettings

End Sub
Private Sub ShowCurrentSettings()
'''On Error GoTo Err
With Screen.ActiveForm
Screen.MousePointer = vbHourglass

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Brokers Code", .ListView1.Width / 8 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Brokers Name", .ListView1.Width / 2.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Physical Address", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Postal Address", .ListView1.Width / 6
.ListView1.ColumnHeaders.Add , , "Branch", .ListView1.Width / 6
.ListView1.ColumnHeaders.Add , , "Town/City", .ListView1.Width / 10
.ListView1.ColumnHeaders.Add , , "Country", .ListView1.Width / 8
.ListView1.ColumnHeaders.Add , , "Contact Person", .ListView1.Width / 6
.ListView1.ColumnHeaders.Add , , "Contact Title", .ListView1.Width / 3.5
.ListView1.ColumnHeaders.Add , , "Phone Number", .ListView1.Width / 8
.ListView1.ColumnHeaders.Add , , "Email", .ListView1.Width / 5

.ListView1.View = lvwReport: .ListView1.Visible = True

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT ParamInsuranceBrokers.* FROM ParamInsuranceBrokers ORDER BY ParamInsuranceBrokers.BrokersCode;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem, NCount As Double

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Screen.MousePointer = vbDefault: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!BrokersCode))
    
    If Not IsNull(rsLIST!BrokersCode) Then
        MyList.SubItems(1) = CStr(rsLIST!BrokersName)
        MyList.SubItems(2) = CStr(rsLIST!PhysicalAddress)
        MyList.SubItems(3) = CStr(rsLIST!postaladdress)
        MyList.SubItems(4) = CStr(rsLIST!Branch)
        MyList.SubItems(5) = CStr(rsLIST!towncity)
        MyList.SubItems(6) = CStr(rsLIST!country)
        MyList.SubItems(7) = CStr(rsLIST!ContactPerson)
        MyList.SubItems(8) = CStr(rsLIST!ContactTitle)
        MyList.SubItems(9) = CStr(rsLIST!TelephoneNo)
        MyList.SubItems(10) = CStr(rsLIST!Email)

    End If

    rsLIST.MoveNext
    
Wend

    Set MyList = Nothing: Set rsLIST = Nothing: Screen.MousePointer = vbDefault
        
End With
Exit Sub
err:
If err.Number = 3265 Then
    Resume Next
Else
    Screen.MousePointer = vbDefault
    ErrorMessage
End If
End Sub
Private Function ValidRecord() As Boolean
With Screen.ActiveForm
    If .txtQuotationNo.Text = Empty Then
        strMessage = "Quotation Number required...!"
        .txtQuotationNo.SetFocus
    ElseIf .txtName.Text = Empty Then
        strMessage = "Clent name required...!"
        .txtName.SetFocus
    ElseIf .txtItemCode.Text = Empty Then
        strMessage = "BillBoard number required...!"
        .txtItemCode.SetFocus
    ElseIf .cboSiding.Text = Empty Then
        strMessage = "Advertisement siding required...!"
        .cboSiding.SetFocus
    ElseIf .txtQuantity.Text = Empty Then
        strMessage = "Quantity Required...!"
        .txtQuantity.SetFocus
    ElseIf .txtPrice.Text = Empty Then
        strMessage = "Advert Price Required...!"
        .txtPrice.SetFocus
    ElseIf .txtTotalPrice.Text = Empty Then
        strMessage = "Total Price Required...!"
        .txtTotalPrice.SetFocus
    
    Else
        ValidRecord = True
    End If
    If Not ValidRecord Then
        MsgBox strMessage, vbCritical + vbOKOnly, "Data Validation"
    End If
End With
End Function
Private Function ValidMainRecord() As Boolean
With Screen.ActiveForm
    If .txtQuotationNo.Text = Empty Then
        strMessage = "Quotation Number required...!"
        .txtQuotationNo.SetFocus
    ElseIf .txtName.Text = Empty Then
        strMessage = "Clent name required...!"
        .txtName.SetFocus
    ElseIf .dtpQuotationDate.Value = Empty Then
        strMessage = "Quotation date required...!"
        .dtpQuotationDate.SetFocus
     
    Else
        ValidMainRecord = True
    End If
    If Not ValidRecord Then
        MsgBox strMessage, vbCritical + vbOKOnly, "Data Validation"
    End If
End With
End Function
Public Sub RemoveCurrentList2Item()
''On Error GoTo Err
With Screen.ActiveForm
Dim i, j, k
   j = .ListView2.ListItems.Count: i = 1
     If j = 0 Then Exit Sub
     
     For i = 1 To j
      If .ListView2.ListItems(i).Checked = True Then
         .ListView2.ListItems.Remove (i): Exit Sub
      End If
    Next i
End With
Exit Sub
err:
   ErrorMessage
End Sub

Private Sub mnuRegisteredClients_Click()
If Not NewRecord Then Exit Sub
Call ShowAllClients
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
''On Error GoTo Err

        Dim TotalSum, QuoteDate As Variant
        
        With Screen.ActiveForm
        QuoteDate = Format(.dtpQuotationDate.Value, "MMMM dd,yyyy")
        
        Select Case Button.Key
        Case "N"
            Select Case Button.Caption
            Case "New &Record "
                If EditRecord Then Exit Sub
                MyCommonData.ClearTextFields:
                NewRecord = True: Button.Caption = "&Save Record ": Button.Image = 4
                .txtName.SetFocus
                .txtQuotationNo.Text = AutoPurchaseOrderNo
                .dtpQuotationDate.Value = Date
                .cboCity.Text = CurrentRecord
                
            Case "&Save Record "
                If NewRecord Then
                    If ValidRecord Then
                            NewSQL = "INSERT INTO AdvertQuotationItems(QuotationNo,ItemCode,CategoryCode,QuotationDesc,Illuminated,SidingType,BorderType,Quantity,Price,TotalPrice,CreatedBy,DateCreated,AccPeriod)VALUES('" & Trim(.txtQuotationNo.Text) & "','" & Trim(.txtItemCode.Text) & "','" & Trim(.txtCategoryCode.Text) & "','" & Trim(.txtName.Text) & "','" & .chkIlluminate.Value & _
                            "','" & Trim(.cboSiding.Text) & "','" & .chkBorder.Value & "','" & Trim(.txtQuantity.Text) & "'," & CCur(.txtPrice.Text) & "," & CCur(.txtTotalPrice.Text) & ",'" & Trim(CurrentUserName) & "','" & Trim(MyCurrentDate) & "','" & Trim(MyCurrentPeriod) & "');"
                    
                    Set rsNewRecord = New ADODB.Recordset
                    rsNewRecord.Open NewSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                    Set rsNewRecord = Nothing
                     Button.Caption = "NE&XT ITEM"
                    .Toolbar1.Buttons(3).Caption = "FINISH"
                    
                End If
                End If
                
             Case "NE&XT ITEM"
                  .txtItemCode.Text = ""
                  .txtItemName.Text = ""
                  .cboCategory.Text = ""
                  .txtCategoryCode.Text = ""
                  .chkIlluminate.Value = 0
                  .cboSiding.Text = ""
                  .chkBorder.Value = 0
                  .txtQuantity.Text = ""
                  .txtPrice.Text = ""
                  .txtTotalPrice = ""
                  RemoveCurrentList2Item
                  Button.Caption = "&Save Record ": Button.Image = 4
            Case Else
                Exit Sub
            End Select
    
Case "E"
    Select Case Button.Caption
'    Case "&Edit/Change "
'    If NewRecord Then Exit Sub
'        If .txtCode.Text = Empty Then
'            MsgBox "There is NO Current Record to Edit. Please Search and Display a Record First...!", vbCritical + vbOKOnly, ""
'           .txtCode.SetFocus
'        Else
'           .txtCode.Locked = True
'            Button.Caption = "Save &Changes ": Button.Image = 4
'            EditRecord = True
'        End If
'    Case "Save &Changes "
'        If EditRecord Then
'        If ValidRecord Then
'        EditSQL = "Update ParamInsuranceBrokers SET BrokersName = '" & Trim(txtName.Text) & "'" & _
'        " ,Branch = '" & Trim(txtBranch.Text) & "'" & _
'        " ,PhysicalAddress = '" & Trim(txtAddress1.Text) & "'" & _
'        " ,PostalAddress = '" & Trim(txtAddress2.Text) & "'" & _
'        " ,TownCity = '" & Trim(cboTown.Text) & "'" & _
'        " ,Country = '" & Trim(cboCountry.Text) & "'" & _
'        " ,ContactPerson = '" & Trim(txtPerson.Text) & "'" & _
'        " ,ContactTitle = '" & Trim(cboTitle.Text) & "'" & _
'        " ,TelephoneNo = '" & Trim(txtPhone.Text) & "'" & _
'        " ,Email = '" & Trim(txtEmail.Text) & "' WHERE BrokersCode='" & Trim(txtCode.Text) & "';"
'
'            Set rsEditRecord = New ADODB.Recordset
'            rsEditRecord.Open EditSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
'            Set rsEditRecord = Nothing
'            .txtCode.Locked = False: EditRecord = False: Button.Caption = "&Edit/Change ": Button.Image = 5
'        End If
'        End If
     Case "FINISH"
         If ValidMainRecord Then
           
            Set rsNewRecord = New ADODB.Recordset
            rsNewRecord.Open "INSERT INTO AdvertQuotation(QuotationNotes,PhysicalAddress,CustomerCode,ContactTitle,ContactName,DateCreated,CreatedBy,AccPeriod,QuotationNo,Customername,PostalAddress,City,Telephone,MobileNo,FaxNo,QuotationDate)VALUES('" & Trim(.txtNotes.Text) & "','" & PhysicalAddress & "','" & CustomerCode & "','" & Trim(.txtContactTitle.Text) & "','" & Trim(.txtContactPerson.Text) & "','" & MyCurrentDate & "','" & CurrentUserName & "','" & MyCurrentPeriod & "','" & Trim(.txtQuotationNo.Text) & "','" & Trim(.txtName.Text) & "','" & Trim(.txtAddress.Text) & "','" & Trim(.cboCity.Text) & "','" & Trim(.txtTelephone.Text) & "','" & Trim(.txtMobile.Text) & "','" & Trim(.txtFax.Text) & "','" & QuoteDate & "')", cnCOMMON, adOpenKeyset, adLockOptimistic
            Set rsNewRecord = Nothing
            
            Set rsFindRecord = New ADODB.Recordset
               rsFindRecord.Open "SELECT SUM(TotalPrice)as Total FROM AdvertQuotationItems WHERE QuotationNo = '" & Trim(.txtQuotationNo.Text) & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
               If rsFindRecord.EOF And rsFindRecord.BOF Then Set rsFindRecord = Nothing
               Else
               TotalSum = rsFindRecord!Total
               End If
               
            Set rsLineUpdate = New ADODB.Recordset
               rsLineUpdate.Open "UPDATE AdvertQuotation SET TotalCost = " & CCur(TotalSum) & " WHERE QuotationNo = '" & Trim(.txtQuotationNo.Text) & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
            Set rsLineUpdate = Nothing
               
       .Toolbar1.Buttons(2).Caption = "New &Record "
       .Toolbar1.Buttons(2).Image = 2
       .Toolbar1.Buttons(3).Image = 5
       .Toolbar1.Buttons(3).Caption = "&Edit/Change "
             
'     End If
    Case Else
       
        Exit Sub
    End Select
Case "S"
'    If NewRecord Or EditRecord Then Exit Sub
'    INPQRY = InputBox("Please Enter the Brokers Code for the Record to Search and Display...!!!", "Enter Brokers Code...")
'    If Len(INPQRY) = 0 Then
'        MsgBox "Required Search Parameter Missing or the Operation Was Cancelled...! No Work was Done!!!", vbCritical + vbOKOnly, "Missing Parameter"
'        Exit Sub
'    Else
'        Set rsFindRecord = cnCOMMON.Execute("SELECT ParamInsuranceBrokers.* FROM ParamInsuranceBrokers WHERE ParamInsuranceBrokers.BrokersCode='" & Trim(INPQRY) & "';")
'        If rsFindRecord.EOF And rsFindRecord.BOF Then
'            MsgBox "Requested Record Missing or Has Been Deleted. Check your Entries to Ensure they are Accurately Spelt...!", vbOKOnly + vbExclamation, "Record NOT Found...!"
'            Set rsFindRecord = Nothing: Exit Sub
'        Else
'            .txtCode.Text = Trim(rsFindRecord!BrokersCode & "")
'            .txtName.Text = Trim(rsFindRecord!BrokersName & "")
'            .txtBranch.Text = Trim(rsFindRecord!Branch & "")
'            .txtAddress1.Text = Trim(rsFindRecord!PhysicalAddress & "")
'            .txtAddress2.Text = Trim(rsFindRecord!PostalAddress & "")
'            .cboTown.Text = Trim(rsFindRecord!TownCity & "")
'            .cboCountry.Text = Trim(rsFindRecord!Country & "")
'            .txtPerson.Text = Trim(rsFindRecord!ContactPerson & "")
'            .cboTitle.Text = Trim(rsFindRecord!ContactTitle & "")
'            .txtPhone.Text = Trim(rsFindRecord!TelephoneNo & "")
'            .txtEmail.Text = Trim(rsFindRecord!Email & "")
'
'        End If
'        Set rsFindRecord = Nothing
'    End If
Case "R"
    If MsgBox(RefreshMessage, vbExclamation + vbOKCancel + vbDefaultButton2, "Refresh The Screen") = vbCancel Then Exit Sub
        .Toolbar1.Buttons(2).Caption = "New &Record "
        .Toolbar1.Buttons(2).Image = 2
        .Toolbar1.Buttons(3).Caption = "&Edit/Change "
        .Toolbar1.Buttons(3).Image = 5
        NewRecord = False: EditRecord = False: MyCommonData.ClearTheScreen
Case "P"
'    Load frmRptAdvertPrintOut
'    frmRptAdvertPrintOut.Show 1, Screen.activeform
Case "F"
     
     
Case Else
    Exit Sub
End Select
End With
Exit Sub
err:
    ErrorMessage

End Sub




Private Sub txtSidingCost_Change()

End Sub

Private Sub txtQuantity_Change()
'On Error GoTo err
With Screen.ActiveForm
.txtTotalPrice.Text = Val(.txtQuantity) * Val(.txtPrice.Text)
End With
Exit Sub
err:
   ErrorMessage
End Sub

Private Sub VScroll1_Change()
With Screen.ActiveForm
.txtQuantity.Text = .VScroll1.Value
End With
End Sub

Private Sub VScroll1_GotFocus()
With Screen.ActiveForm
.VScroll1.Value = 1
End With
End Sub
