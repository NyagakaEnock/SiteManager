VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmODASMAdministration 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Project Costing (Administration Department)"
   ClientHeight    =   7185
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11115
   Icon            =   "frmODASMAdministration.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   11115
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   0
      TabIndex        =   18
      Top             =   1560
      Width           =   10935
      Begin VB.TextBox txtJobDoneBy 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   1320
         TabIndex        =   22
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtDateOfCommence 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   6600
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtDateOfCompletion 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   9480
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtDeadLineDate 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label27 
         Caption         =   "Warranty"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   255
         Width           =   1335
      End
      Begin VB.Label Label28 
         Caption         =   "DOC"
         Height          =   255
         Left            =   6120
         TabIndex        =   25
         Top             =   255
         Width           =   375
      End
      Begin VB.Label Label30 
         Caption         =   "Date Of Completion"
         Height          =   255
         Left            =   8040
         TabIndex        =   24
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label26 
         Caption         =   "Deadline Date"
         Height          =   255
         Left            =   3600
         TabIndex        =   23
         Top             =   255
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Width           =   11055
      Begin VB.TextBox txtLPONo 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtDate 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   9360
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtJobCardNo 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtDescriptionOfOrder 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   600
         Width           =   3495
      End
      Begin VB.TextBox txtClientName 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   600
         Width           =   4575
      End
      Begin VB.TextBox txtDepartment 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label15 
         Caption         =   "L.P.O No"
         Height          =   255
         Left            =   6240
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Date"
         Height          =   255
         Left            =   8760
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Job  Card No"
         Height          =   255
         Left            =   3600
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Desc Of Order"
         Height          =   255
         Left            =   6240
         TabIndex        =   8
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Name Of Client"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Department"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Job Card Payments Made Under Selected Department"
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
      Height          =   2295
      Left            =   0
      TabIndex        =   3
      Top             =   4800
      Width           =   11175
      Begin MSComctlLib.ListView ListView2 
         Height          =   1935
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   3413
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
            Size            =   9
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
            Picture         =   "frmODASMAdministration.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMAdministration.frx":0ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMAdministration.frx":0F0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMAdministration.frx":1228
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMAdministration.frx":18A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMAdministration.frx":1F1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMAdministration.frx":236E
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
      Width           =   11115
      _ExtentX        =   19606
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
         Left            =   9840
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
      Caption         =   "Job Card Costing Per Department"
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
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   2400
      Width           =   11055
      Begin MSComctlLib.ListView ListView1 
         Height          =   2055
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   3625
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
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
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
Attribute VB_Name = "frmODASMAdministration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim billbno, CustomerCode, PhysicalAddress, JoboNumber, PurchaseOrderNo




Private Sub cboBBNo_GotFocus()

End Sub

Private Sub cboBBNo_LostFocus()

End Sub

Private Sub cboCouncilPeriod_Click()
Me.ListView1.SetFocus
End Sub



Private Sub cboCountry_Click()
Me.ListView1.SetFocus
End Sub

Private Sub cboElecPeriod_Click()
Me.ListView1.SetFocus
End Sub


Private Sub cboLandLord_Click()
Me.ListView1.SetFocus
End Sub

Private Sub cboRentPeriod_Click()
Me.ListView1.SetFocus
End Sub



Private Sub cboTown_Click()
Me.ListView1.SetFocus
End Sub



Private Sub Combo1_Click()
Me.ListView1.SetFocus
End Sub

Private Sub cboCategory_GotFocus()
''On Error GoTo Err
'If Not NewRecord Then Exit Sub
With Me
   
    If .cboCategory.ListCount <> 0 Then Exit Sub
    
     AttachSQL = "SELECT A.CategoryCode AS SelectField FROM ParamDrugCategories A ,GenProductsInventory B WHERE A.CategoryCode = B.CategoryCode ORDER BY A.CategoryCode;"
    .cboCategory.Clear
    MyCommonData.AttachInventDropDown
    
End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub cboCity_GotFocus()
If Not NewRecord Then Exit Sub
With Me
    AttachSQL = "SELECT Town AS SelectField FROM ODASPTown ORDER BY Town;"
    .cboCity.Clear
    MyCommonData.AttachDropDown
End With

End Sub

Private Sub cboSiding_GotFocus()
If Not NewRecord Then Exit Sub
With Me
    AttachSQL = "SELECT SidingDescription AS SelectField FROM Advertsiding ORDER BY sidingdescription;"
    .cboSiding.Clear
    MyCommonData.AttachDropDown
End With

End Sub

Private Sub cboCategory_Click()
Me.ListView3.SetFocus
End Sub
Private Sub cboCategory_LostFocus()
''On Error GoTo Err
With Me

    Set rsFindRecord = cnINVENT.Execute("SELECT * FROM ParamDrugCategories A,GenproductsInventory B WHERE A.CategoryCode = B.CategoryCode AND B.CategoryCODE='" & Trim(.cboCategory.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        Set rsFindRecord = Nothing: Exit Sub
    Else
        .txtCategoryCode.Text = rsFindRecord!categoryname & ""
        .cboCategory.Text = rsFindRecord!categorycode & ""
        
        .ListView2.SetFocus
        
        If .cboCategory.Text = "AAA" Then
            Call ShowAllInventoryItems
        Else
            Call ShowInventoryItemsPerCategory
        End If
        
    End If
    
End With
Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ShowInventoryItemsPerCategory()
''On Error GoTo Err
With Me
.ListView2.ListItems.Clear
.ListView2.ColumnHeaders.Clear

.ListView2.ColumnHeaders.Add , , "Product Code", .ListView2.Width / 5.5
.ListView2.ColumnHeaders.Add , , "Purchase Order No", .ListView2.Width / 5.5
.ListView2.ColumnHeaders.Add , , "Product Name", .ListView2.Width / 4.5
.ListView2.ColumnHeaders.Add , , "Current Quantity", .ListView2.Width / 5.5
.ListView2.ColumnHeaders.Add , , "Quantity Units", .ListView2.Width / 5.5
.ListView2.ColumnHeaders.Add , , "Current Total Pieces", .ListView2.Width / 5.5

.ListView2.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM GenProductsInventory WHERE Discontinued = '" & "N" & "' AND CategoryCode = '" & Trim(.cboCategory.Text) & "' ORDER BY DrugName;", cnINVENT, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView2.View = lvwList
    Set MyList = .ListView2.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!ProductCode))


    If Not IsNull(rsLIST!PurchaseOrderNo) Then
        MyList.SubItems(1) = CStr(rsLIST!PurchaseOrderNo)
    End If
     
    If Not IsNull(rsLIST!drugname) Then
        MyList.SubItems(2) = CStr(rsLIST!drugname)
    End If
    
    If Not IsNull(rsLIST!currentquantity) Then
        MyList.SubItems(3) = CStr(rsLIST!currentquantity)
    End If
    
    If Not IsNull(rsLIST!QuantityUnits) Then
        MyList.SubItems(4) = CStr(rsLIST!QuantityUnits)
    End If
    
    If Not IsNull(rsLIST!TotalPieces) Then
        MyList.SubItems(5) = CStr(rsLIST!TotalPieces)
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

Private Sub ShowAllInventoryProductsStructure()

End Sub

Private Sub ShowAllInventoryItems()
''On Error GoTo Err
With Me
.ListView2.ListItems.Clear
.ListView2.ColumnHeaders.Clear

.ListView2.ColumnHeaders.Add , , "Product Code", .ListView2.Width / 5.5
.ListView2.ColumnHeaders.Add , , "Purchase Order No", .ListView2.Width / 5.5
.ListView2.ColumnHeaders.Add , , "Product Name", .ListView2.Width / 4.5
.ListView2.ColumnHeaders.Add , , "Current Quantity", .ListView2.Width / 5.5
.ListView2.ColumnHeaders.Add , , "Quantity Units", .ListView2.Width / 5.5
.ListView2.ColumnHeaders.Add , , "Current Total Pieces", .ListView2.Width / 5.5

.ListView2.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM GenProductsInventory WHERE Discontinued = '" & "N" & "' ORDER BY DrugName;", cnINVENT, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView2.View = lvwList
    Set MyList = .ListView2.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!ProductCode))


    If Not IsNull(rsLIST!PurchaseOrderNo) Then
        MyList.SubItems(1) = CStr(rsLIST!PurchaseOrderNo)
    End If
     
    If Not IsNull(rsLIST!drugname) Then
        MyList.SubItems(2) = CStr(rsLIST!drugname)
    End If
    
    If Not IsNull(rsLIST!currentquantity) Then
        MyList.SubItems(3) = CStr(rsLIST!currentquantity)
    End If
    
    If Not IsNull(rsLIST!QuantityUnits) Then
        MyList.SubItems(4) = CStr(rsLIST!QuantityUnits)
    End If
    
    If Not IsNull(rsLIST!TotalPieces) Then
        MyList.SubItems(5) = CStr(rsLIST!TotalPieces)
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

Private Function GetAdvertPrice()
''On Error GoTo Err
With Me
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
With Me
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
Private Sub ShowAllClientsStructure()
'''On Error GoTo Err
With Me
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Client Code", .ListView1.Width / 6#
.ListView1.ColumnHeaders.Add , , "Company Name", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Address", .ListView1.Width / 6
.ListView1.ColumnHeaders.Add , , "City", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Phone", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Mobile Phone", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Fax", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Contact Title", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Contact Name", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Physical Adress", .ListView1.Width / 5.5

.ListView1.View = lvwReport
End With
Exit Sub
err:
    ErrorMessage
 End Sub
Private Sub ShowAllClients()
'''On Error GoTo Err
With Me
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Client Code", .ListView1.Width / 6#
.ListView1.ColumnHeaders.Add , , "Company Name", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Address", .ListView1.Width / 5
.ListView1.ColumnHeaders.Add , , "City", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Phone", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Mobile Phone", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Fax", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Contact Title", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Contact Name", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Physical Adress", .ListView1.Width / 1.5
.ListView1.ColumnHeaders.Add , , "Client Code", .ListView1.Width / 6.5

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM AdvertClients ORDER BY CompanyName;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!customerid))


    If Not IsNull(rsLIST!CompanyName) Then
        MyList.SubItems(1) = CStr(rsLIST!CompanyName)
    End If
     
    If Not IsNull(rsLIST!Address) Then
        MyList.SubItems(2) = CStr(rsLIST!Address)
    End If
    
    If Not IsNull(rsLIST!city) Then
        MyList.SubItems(3) = CStr(rsLIST!city)
    End If
    
    If Not IsNull(rsLIST!phone) Then
        MyList.SubItems(4) = CStr(rsLIST!phone)
    End If
    
    If Not IsNull(rsLIST!Mobilephone) Then
        MyList.SubItems(5) = CStr(rsLIST!Mobilephone)
    End If
    
    If Not IsNull(rsLIST!Fax) Then
        MyList.SubItems(6) = CStr(rsLIST!Fax)
    End If
    
    If Not IsNull(rsLIST!ContactTitle) Then
        MyList.SubItems(7) = CStr(rsLIST!ContactTitle)
    End If
    
    If Not IsNull(rsLIST!Contactname) Then
        MyList.SubItems(8) = CStr(rsLIST!Contactname)
    End If
    
    If Not IsNull(rsLIST!PhysicalAddress) Then
        MyList.SubItems(9) = CStr(rsLIST!PhysicalAddress)
    End If
    
    If Not IsNull(rsLIST!customerid) Then
        MyList.SubItems(10) = CStr(rsLIST!customerid)
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
Private Sub ShowAllBillBoardCategoriesStructure()
'''On Error GoTo Err
With Me
.ListView2.ListItems.Clear
.ListView2.ColumnHeaders.Clear

.ListView2.ColumnHeaders.Add , , "Category Code", .ListView2.Width / 6#
.ListView2.ColumnHeaders.Add , , "Category Name", .ListView2.Width / 5
.ListView2.ColumnHeaders.Add , , "Item Code", .ListView2.Width / 6
.ListView2.ColumnHeaders.Add , , "Item Name", .ListView2.Width / 4
.ListView2.ColumnHeaders.Add , , "Length", .ListView2.Width / 6.5
.ListView2.ColumnHeaders.Add , , "Width", .ListView2.Width / 6.5

.ListView2.View = lvwReport
End With
Exit Sub
err:
    ErrorMessage
End Sub


Private Sub ShowAllBillBoardCategories()
''On Error GoTo Err
With Me
.ListView2.ListItems.Clear
.ListView2.ColumnHeaders.Clear

.ListView2.ColumnHeaders.Add , , "Category Code", .ListView2.Width / 6#
.ListView2.ColumnHeaders.Add , , "Category Name", .ListView2.Width / 4
.ListView2.ColumnHeaders.Add , , "Item Code", .ListView2.Width / 5.5
.ListView2.ColumnHeaders.Add , , "Item Name", .ListView2.Width / 4
.ListView2.ColumnHeaders.Add , , "Length", .ListView2.Width / 6.5
.ListView2.ColumnHeaders.Add , , "Width", .ListView2.Width / 6.5

.ListView2.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM AdvertBBDetails A ,AdvertCategories B WHERE A.CategoryCode = B.CategoryCode ORDER BY A.Name;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView2.View = lvwList
    Set MyList = .ListView2.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
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
Me.ListView1.SetFocus
End Sub

Private Sub cboSiding_Click()
Me.ListView2.SetFocus
End Sub

Private Sub cboSiding_LostFocus()
'''On Error GoTo Err
With Me

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

Private Sub cboTax_Click()
Me.ListView2.SetFocus
End Sub

Private Sub cboTax_GotFocus()
''On Error GoTo Err
'If Not NewRecord Then Exit Sub
With Me
   
    If .cboTax.ListCount <> 0 Then Exit Sub
    
     AttachSQL = "SELECT Description AS SelectField FROM ParamTaxes ORDER BY Description;"
    .cboTax.Clear
    MyCommonData.AttachDropDown
    
End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub cboTax_LostFocus()
''On Error GoTo Err
With Me

    Set rsFindRecord = cnCOMMON.Execute("SELECT * FROM ParamTaxes WHERE Description ='" & Trim(.cboTax.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        Set rsFindRecord = Nothing: Exit Sub
    Else
        .cboTax.Text = rsFindRecord!TaxRate & ""
        If .txtPrice.Text = "" And .txtTotalPrice.Text = "" Then Exit Sub
          .txtTaxAmount.Text = Val(CSng(.cboTax.Text) / 100) * Val(CSng(.txtTotalPrice.Text))
          .txtTaxPrice.Text = Val(CSng(.txtTotalPrice.Text)) + Val(.txtTaxAmount.Text)
          .txtTaxAmount.Text = FormatNumber(.txtTaxAmount.Text, 2, vbUseDefault, vbUseDefault, vbTrue)
'          .txtTaxPrice.Text = FormatNumber(.txtTaxAmount.Text, 2, vbUseDefault, vbUseDefault, vbTrue)
         .ListView2.SetFocus
    End If
    
End With
Exit Sub
err:
    ErrorMessage

End Sub

Private Sub Form_Initialize()
Set MyCommonData = New clsCommonData
'Set mycommondata New clsCommonData

End Sub

Private Sub Form_Load()
ShowAllClientsStructure
ShowAllBillBoardCategoriesStructure
ShowAllInventoryProductsStructure
End Sub
Public Function AutoPurchaseOrderNo() As String
''''On Error GoTo Err
With Me

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

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
''On Error GoTo Err
'If Not NewRecord And Not EditRecord Then Item.Checked = False: Exit Sub
If Me.ListView1.ListItems.Count = 0 Or Me.ListView1.View <> lvwReport Then Exit Sub
    
    Dim i, j, k
    j = Me.ListView1.ListItems.Count
    
    If j = 0 Then Exit Sub
    
    For i = 1 To j
        If Me.ListView1.ListItems(i).Text <> Item Then
            Me.ListView1.ListItems(i).Checked = False
        End If
    Next i
    
      If Item.Checked = True Then
        
        Me.txtJobCardNo.Text = Item
        Me.txtDeptCode.Text = Item.SubItems(1)
        Me.txtDepartment.Text = Item.SubItems(2)
        Me.txtSupervisor.Text = Item.SubItems(3)
        Me.txtDateOfCompletion.Text = Item.SubItems(4)
        Me.txtDateOfCommence.Text = Item.SubItems(5)
        Me.txtClientName.Text = Item.SubItems(6)
        Me.txtDeadlineDate.Text = Item.SubItems(7)
        Me.txtLpono.Text = Item.SubItems(8)
        JoboNumber = Item
        
        Call ShowAllItemsUnderSelectedJob
        
             
    ElseIf Item.Checked = False Then

    End If
Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ShowAllItemsUnderSelectedJob()
''On Error GoTo Err
With Me
.ListView3.ListItems.Clear
.ListView3.ColumnHeaders.Clear

.ListView3.ColumnHeaders.Add , , "Job Card No", .ListView3.Width / 5
.ListView3.ColumnHeaders.Add , , "Media Name", .ListView3.Width / 4
.ListView3.ColumnHeaders.Add , , "Site Code", .ListView3.Width / 6.5
.ListView3.ColumnHeaders.Add , , "Site Name", .ListView3.Width / 4.5
.ListView3.ColumnHeaders.Add , , "Quantity", .ListView3.Width / 6.5
.ListView3.ColumnHeaders.Add , , "Item Code", .ListView3.Width / 6.5
.ListView3.ColumnHeaders.Add , , "Length", .ListView3.Width / 6.5
.ListView3.ColumnHeaders.Add , , "Width", .ListView3.Width / 6.5
.ListView3.ColumnHeaders.Add , , "Illuminated", .ListView3.Width / 6.5
.ListView3.ColumnHeaders.Add , , "Siding", .ListView3.Width / 6.5
.ListView3.ColumnHeaders.Add , , "Bordered", .ListView3.Width / 6.5
.ListView3.ColumnHeaders.Add , , "Registered Site", .ListView3.Width / 4.5
.ListView3.ColumnHeaders.Add , , "Other Site", .ListView3.Width / 4.5

.ListView3.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM AdvertJobBrief A ,AdvertJobBriefItems B,AdvertBBDetails C,AdvertSites D WHERE A.JobBriefno = B.JobBriefNo AND C.BillBoardNo = B.ItemCode AND B.SiteCode = D.SiteNo AND B.JoBbriefNo = '" & JoboNumber & "';", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView3.View = lvwList
    Set MyList = .ListView3.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView3.ListItems.Add(, , CStr(rsLIST!JobBriefNo))


    If Not IsNull(rsLIST!categoryname) Then
        MyList.SubItems(1) = CStr(rsLIST!categoryname) + " " + CStr(rsLIST!ItemName)
    End If
     
    If Not IsNull(rsLIST!SiteCode) Then
        MyList.SubItems(2) = CStr(rsLIST!SiteCode)
    End If
    
    If Not IsNull(rsLIST!SiteName) Then
        MyList.SubItems(3) = CStr(rsLIST!SiteName)
    End If
        
    If Not IsNull(rsLIST!quantity) Then
        MyList.SubItems(4) = CStr(rsLIST!quantity)
    End If
    
     If Not IsNull(rsLIST!ItemCode) Then
        MyList.SubItems(5) = CStr(rsLIST!ItemCode)
    End If
    
     If Not IsNull(rsLIST!Length) Then
        MyList.SubItems(6) = CStr(rsLIST!Length)
    End If
    
     If Not IsNull(rsLIST!Width) Then
        MyList.SubItems(7) = CStr(rsLIST!Width)
    End If
    
     If Not IsNull(rsLIST!Illuminate) And (rsLIST!Illuminate) = 0 Then
        MyList.SubItems(8) = CStr("NO")
      ElseIf Not IsNull(rsLIST!Illuminate) And (rsLIST!Illuminate) = 1 Then
        MyList.SubItems(8) = CStr("YES")
    End If
    
    If Not IsNull(rsLIST!Siding) Then
        MyList.SubItems(9) = CStr(rsLIST!Siding)
    End If
    
     If Not IsNull(rsLIST!Border) And (rsLIST!Border) = 0 Then
        MyList.SubItems(10) = CStr("NO")
      ElseIf Not IsNull(rsLIST!Border) And (rsLIST!Border) = 1 Then
        MyList.SubItems(10) = CStr("YES")
    End If
    
    If Not IsNull(rsLIST!SiteName) Then
        MyList.SubItems(11) = CStr(rsLIST!SiteName)
    End If
    
    If Not IsNull(rsLIST!OtherSite) Then
        MyList.SubItems(12) = CStr(rsLIST!OtherSite)
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
Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
''On Error GoTo Err
'If Not NewRecord And Not EditRecord Then Item.Checked = False: Exit Sub
If Me.ListView2.ListItems.Count = 0 Or Me.ListView2.View <> lvwReport Then Exit Sub
    
    Dim i, j, k
    j = Me.ListView2.ListItems.Count
    
    If j = 0 Then Exit Sub
    
    For i = 1 To j
        If Me.ListView2.ListItems(i).Text <> Item Then
            Me.ListView2.ListItems(i).Checked = False
        End If
    Next i
    
      If Item.Checked = True Then

        Me.txtItemCode.Text = Item
        Me.txtItemName.Text = Item.SubItems(2)
'        PurchaseOrderNo = Item.SubItems(1)
        
    ElseIf Item.Checked = False Then
    
       
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

Private Sub ListView3_ItemCheck(ByVal Item As MSComctlLib.ListItem)
''On Error GoTo Err
'If Not NewRecord And Not EditRecord Then Item.Checked = False: Exit Sub
If Me.ListView3.ListItems.Count = 0 Or Me.ListView3.View <> lvwReport Then Exit Sub
    
    Dim i, j, k
    j = Me.ListView3.ListItems.Count
    
    If j = 0 Then Exit Sub
    
    For i = 1 To j
        If Me.ListView3.ListItems(i).Text <> Item Then
            Me.ListView3.ListItems(i).Checked = False
        End If
    Next i
    
      If Item.Checked = True Then
    
        
'        Me.txtItemCode.Text = Item
        Me.txtOrderDesc.Text = Item.SubItems(1)
        Me.txtSideCode.Text = Item.SubItems(2)
        Me.txtLocation.Text = Item.SubItems(3)
        Me.txtOrderQuantity.Text = Item.SubItems(4)
        Me.txtLength.Text = Item.SubItems(6)
        Me.txtWidth.Text = Item.SubItems(7)
        
        Me.txtOtherSite.Text = Item.SubItems(12)
        Me.txtSiding.Text = Item.SubItems(9)
                
        If Item.SubItems(8) = "YES" Then
        Me.chkIlluminated.Value = 1
        Else
        Me.chkIlluminated.Value = 0
        End If
        
        If Item.SubItems(10) = "YES" Then
        Me.chkBorder.Value = 1
        Me.txtBorderLength.Text = Item.SubItems(6)
        Me.txtBorderWidth.Text = Item.SubItems(7)
        Else
        Me.chkBorder.Value = 0
        End If
        
        
    ElseIf Item.Checked = False Then
    
        
    End If
    
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnuClear_Click()
    MyCommonData.ClearTextFields
End Sub

Private Sub mnuClose_Click()
    Unload Me
End Sub

Private Sub mnuCurrent_Click()
    Call ShowCurrentSettings

End Sub
Private Sub ShowCurrentSettings()
'''On Error GoTo Err
With Me
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
With Me
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
With Me
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
With Me
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

Private Sub mnuClosedJobs_Click()
If Not NewRecord Then Exit Sub
Call ShowClosedJobs
End Sub

Private Sub mnuFullInventory_Click()
'On Error GoTo err
'If Not NewRecord Then Exit Sub
Call ShowAllInventoryItems
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnuOpenedJobs_Click()
'If Not NewRecord Then Exit Sub

End Sub
Private Sub ShowClosedJobs()
''On Error GoTo Err
With Me
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Job Card No", .ListView1.Width / 5
.ListView1.ColumnHeaders.Add , , "Dept Code", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Department Name", .ListView1.Width / 5
.ListView1.ColumnHeaders.Add , , "Done By", .ListView1.Width / 5
.ListView1.ColumnHeaders.Add , , "TotalCost", .ListView1.Width / 6.5

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM AdvertJobCard A ,AdvertParamDepartments B WHERE A.DeptCode = B.DepartmentCode AND A.Closed = '" & "Y" & "';", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!JobCardNo))


    If Not IsNull(rsLIST!DeptCode) Then
        MyList.SubItems(1) = CStr(rsLIST!DeptCode)
    End If
     
    If Not IsNull(rsLIST!DepartmentDescription) Then
        MyList.SubItems(2) = CStr(rsLIST!DepartmentDescription)
    End If
    
    If Not IsNull(rsLIST!DoneBy) Then
        MyList.SubItems(3) = CStr(rsLIST!DoneBy)
    End If
    
    If Not IsNull(rsLIST!TotalCost) Then
        MyList.SubItems(4) = CStr(rsLIST!TotalCost)
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

Private Sub optPrices_Click(Index As Integer)
''On Error GoTo Err
With Me
'If Not NewRecord And Not EditRecord Then Exit Sub
Select Case Index
Case 0
    Call GetWholeSaleCost
    Me.optPrices(0).Value = True
    .txtTotalPrice.Text = FormatNumber(CDbl(.txtPrice.Text), 2, vbUseDefault, vbUseDefault, vbTrue)
   
Case 1
    Call GetRetailCost
    Me.optPrices(1).Value = True
   .txtTotalPrice.Text = FormatNumber(CDbl(.txtPrice.Text), 2, vbUseDefault, vbUseDefault, vbTrue)
   
Case Else
    Exit Sub
End Select
Exit Sub
err:
    ErrorMessage
End With
End Sub
Private Sub GetRetailCost()
''On Error GoTo Err
With Me
    Set rsFindRecord = cnINVENT.Execute("SELECT * FROM ProductsCostPriceSetup WHERE DrugCode='" & Trim(.txtItemCode.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        Set rsFindRecord = Nothing
    Else
        .txtPrice.Text = rsFindRecord!RetailCost & ""
        .txtQuantity.Text = 1
        
    End If
End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub GetWholeSaleCost()
''On Error GoTo Err
With Me
    Set rsFindRecord = cnINVENT.Execute("SELECT * FROM ProductsCostPriceSetup WHERE DrugCode='" & Trim(.txtItemCode.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        Set rsFindRecord = Nothing
    Else
        .txtPrice.Text = rsFindRecord!DosageCost & ""
        .txtQuantity.Text = 1
        
    End If
End With
Exit Sub
err:
    ErrorMessage
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
''On Error GoTo Err
Dim TotalSum, QuoteDate As Variant
With Me
QuoteDate = Format(.dtpQuotationDate.Value, "MMMM dd,yyyy")
Select Case Button.Key
Case "N"
    Select Case Button.Caption
    Case "New &Record "
        If EditRecord Then Exit Sub
        MyCommonData.ClearTextFields: .ListView1.ListItems.Clear: .ListView2.ListItems.Clear: .txtName.SetFocus
        NewRecord = True: Button.Caption = "&Save Record ": Button.Image = 4
        .txtName.SetFocus
        .txtQuotationNo.Text = AutoPurchaseOrderNo
        .dtpQuotationDate.Value = MyCurrentDate
        
    Case "&Save Record "
        If NewRecord Then
        If ValidRecord Then
            NewSQL = "INSERT INTO AdvertQuotationItems(ItemName,QuotationNo,ItemCode,CategoryCode,CategoryName,QuotationDesc,Illuminated,SidingType,BorderType,Quantity,Price,TotalPrice,CreatedBy,DateCreated,AccPeriod)VALUES('" & Trim(.txtItemName.Text) & "','" & Trim(.txtQuotationNo.Text) & "','" & Trim(.txtItemCode.Text) & "','" & Trim(.txtCategoryCode.Text) & "','" & Trim(.cboCategory.Text) & "','" & Trim(.txtName.Text) & "','" & .chkIlluminate.Value & _
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
'    frmRptAdvertPrintOut.Show 1, Me
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

Private Sub txtItemName_Change()
'On Error GoTo err
With Me

    If .txtItemName.Text = Empty Then
         .ListView1.ListItems.Clear
    Else
     SearchByProductName
    End If

End With
Exit Sub
err:
   ErrorMessage
End Sub
Private Sub SearchByProductName()
''On Error GoTo Err
With Me
.ListView2.ListItems.Clear
.ListView2.ColumnHeaders.Clear

.ListView2.ColumnHeaders.Add , , "Product Code", .ListView2.Width / 5.5
.ListView2.ColumnHeaders.Add , , "Purchase Order No", .ListView2.Width / 5.5
.ListView2.ColumnHeaders.Add , , "Product Name", .ListView2.Width / 4.5
.ListView2.ColumnHeaders.Add , , "Current Quantity", .ListView2.Width / 5.5
.ListView2.ColumnHeaders.Add , , "Quantity Units", .ListView2.Width / 5.5
.ListView2.ColumnHeaders.Add , , "Current Total Pieces", .ListView2.Width / 5.5

.ListView2.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM GenProductsInventory WHERE Discontinued = '" & "N" & "' AND DrugName LIKE '" & Trim(.txtItemName.Text) & "%' ORDER BY DrugName;", cnINVENT, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView2.View = lvwList
    Set MyList = .ListView2.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!ProductCode))


    If Not IsNull(rsLIST!PurchaseOrderNo) Then
        MyList.SubItems(1) = CStr(rsLIST!PurchaseOrderNo)
    End If
     
    If Not IsNull(rsLIST!drugname) Then
        MyList.SubItems(2) = CStr(rsLIST!drugname)
    End If
    
    If Not IsNull(rsLIST!currentquantity) Then
        MyList.SubItems(3) = CStr(rsLIST!currentquantity)
    End If
    
    If Not IsNull(rsLIST!QuantityUnits) Then
        MyList.SubItems(4) = CStr(rsLIST!QuantityUnits)
    End If
    
    If Not IsNull(rsLIST!TotalPieces) Then
        MyList.SubItems(5) = CStr(rsLIST!TotalPieces)
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
Private Sub txtQuantity_Change()
'On Error GoTo err
With Me
.txtTotalPrice.Text = Val(.txtQuantity) * Val(.txtPrice.Text)
End With
Exit Sub
err:
   ErrorMessage
End Sub

Private Sub txtQuotationNo_Change()

End Sub

Private Sub VScroll1_Change()
With Me
.txtQuantity.Text = .VScroll1.Value
End With
End Sub

Private Sub VScroll1_GotFocus()
With Me
.VScroll1.Value = 1
End With
End Sub
