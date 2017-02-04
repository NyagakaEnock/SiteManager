VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmSysInfoMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SETTINGS AND SYSTEM INFORMATION"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   1140
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSysInfoMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   11910
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin VB.TextBox txtREGS 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   375
      Left            =   11040
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3000
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysInfoMain.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysInfoMain.frx":075C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysInfoMain.frx":0BAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysInfoMain.frx":1000
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysInfoMain.frx":1452
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6375
      Left            =   2400
      TabIndex        =   1
      Top             =   720
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   11245
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
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
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   11245
      _Version        =   393217
      HideSelection   =   0   'False
      Style           =   5
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Visible         =   0   'False
      Begin VB.Menu mnuFileEditSelected 
         Caption         =   "Edit &Selected Record"
      End
      Begin VB.Menu mnu08 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileDeleteSelected 
         Caption         =   "&Delete Selected Record"
      End
   End
   Begin VB.Menu mnuSeacrhCompany 
      Caption         =   "Search&Co"
      Visible         =   0   'False
      Begin VB.Menu mnuSearchCompanyByCode 
         Caption         =   "Search by Company &Code"
      End
      Begin VB.Menu mnu90 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSearchCompanyByName 
         Caption         =   "Search by Company &Name"
      End
   End
   Begin VB.Menu mnuSearchDepartment 
      Caption         =   "Search&Dept"
      Visible         =   0   'False
      Begin VB.Menu mnuSearchDeptByCode 
         Caption         =   "Search by Department &Code"
      End
      Begin VB.Menu mnujjk 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSearchDeptByName 
         Caption         =   "Search by Department &Name"
      End
   End
   Begin VB.Menu mnuSearchStaff 
      Caption         =   "&SearchStaff"
      Visible         =   0   'False
      Begin VB.Menu mnuSearchStaffByID 
         Caption         =   "Search by Staff ID Number"
      End
      Begin VB.Menu mnu8iij 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSearchStaffBySurname 
         Caption         =   "Search by &Surname"
      End
   End
End
Attribute VB_Name = "frmSysInfoMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rsFIND As ADODB.Recordset
Private CSearch As clsSearchForRecord

Private Sub MyTreeSetup()
On Error GoTo Err

Dim NodB As Node
    Set NodB = TreeView1.Nodes.Add(, , "C", "My Company Information", 1)
    Set NodB = TreeView1.Nodes.Add("C", tvwChild, "C1", "New Company Details", 1)
    Set NodB = TreeView1.Nodes.Add("C", tvwChild, "C2", "View My Company Info", 1)
    Set NodB = TreeView1.Nodes.Add("C", tvwChild, "C3", "New Departments/Divisions", 1)
    Set NodB = TreeView1.Nodes.Add("C", tvwChild, "C4", "View List of Departments", 1)
    Set NodB = TreeView1.Nodes.Add("C", tvwChild, "C5", "New Staff Members", 1)
    Set NodB = TreeView1.Nodes.Add("C", tvwChild, "C6", "View Staff Members' Info", 1)

Dim NodD As Node
    Set NodD = TreeView1.Nodes.Add(, , "D", "Other Parameters", 1)
    Set NodD = TreeView1.Nodes.Add("D", tvwChild, "D1", "Employers", 1)
    Set NodD = TreeView1.Nodes.Add("D", tvwChild, "D2", "Countries", 1)
    Set NodD = TreeView1.Nodes.Add("D", tvwChild, "D3", "Cities/Towns", 1)
    Set NodD = TreeView1.Nodes.Add("D", tvwChild, "D4", "Currencies", 1)
    Set NodD = TreeView1.Nodes.Add("D", tvwChild, "D5", "Titles of Courtesy", 1)
    Set NodD = TreeView1.Nodes.Add("D", tvwChild, "D6", "Payment Methods", 1)
    Set NodD = TreeView1.Nodes.Add("D", tvwChild, "D7", "Services and Fees", 1)
    Set NodD = TreeView1.Nodes.Add("D", tvwChild, "D8", "Taxes and Rates", 1)
    Set NodD = TreeView1.Nodes.Add("D", tvwChild, "D9", "Accounting Periods", 1)

Dim NodP As Node
    Set NodP = TreeView1.Nodes.Add(, , "P", "Settings/Options", 1)
    Set NodP = TreeView1.Nodes.Add("P", tvwChild, "P0", "New Options/Settings", 1)
    Set NodP = TreeView1.Nodes.Add("P", tvwChild, "P1", "View Default Currency", 1)
    Set NodP = TreeView1.Nodes.Add("P", tvwChild, "P2", "View Local Currency", 1)
    Set NodP = TreeView1.Nodes.Add("P", tvwChild, "P3", "View V.A.T. Rate", 1)
    Set NodP = TreeView1.Nodes.Add("P", tvwChild, "P4", "View Country Code [Telecom]", 1)
    Set NodP = TreeView1.Nodes.Add("P", tvwChild, "P5", "View Local/Area Code [Telecom]", 1)
    Set NodP = TreeView1.Nodes.Add("P", tvwChild, "P6", "View Defalut Payment Method", 1)
    Set NodP = TreeView1.Nodes.Add("P", tvwChild, "P7", "New MS Office", 1)
    Set NodP = TreeView1.Nodes.Add("P", tvwChild, "P8", "View MS Office", 1)

Dim NodV As Node
    Set NodV = TreeView1.Nodes.Add(, , "V", "Views/Listings", 1)
    Set NodV = TreeView1.Nodes.Add("V", tvwChild, "V1", "My Company Employees", 1)
    Set NodV = TreeView1.Nodes.Add("V", tvwChild, "V2", "List of Employers", 1)
    Set NodV = TreeView1.Nodes.Add("V", tvwChild, "V3", "Authorized Currencies", 1)
    Set NodV = TreeView1.Nodes.Add("V", tvwChild, "V4", "Taxes and Rates", 1)
    Set NodV = TreeView1.Nodes.Add("V", tvwChild, "V5", "Services and Fees", 1)
    Set NodV = TreeView1.Nodes.Add("V", tvwChild, "V6", "Payment Methods", 1)
    
    TreeView1.BorderStyle = vbFixedSingle
    
   Exit Sub
Err:
    MsgBox Err.Description
End Sub

Private Sub cmdSearch_Click()
On Error GoTo Err
With Me
    Select Case .txtREGS.Text
    Case "C1"
        PopupMenu mnuSeacrhCompany
    Case "C2"
        PopupMenu mnuSeacrhCompany
    Case "C3"
        PopupMenu mnuSearchDepartment
    Case "C4"
        PopupMenu mnuSearchDepartment
    Case "C5"
        PopupMenu mnuSearchStaff
    Case "C6"
        PopupMenu mnuSearchStaff
    Case Else
        PopupMenu mnuFile
    End Select
End With
Exit Sub
Err:
ErrorMessage
End Sub


Private Sub GetMainStructure()
On Error GoTo Err

ListView1.ListItems.Clear
ListView1.ColumnHeaders.Clear

ListView1.ColumnHeaders.Add , , "Company Code", ListView1.Width / 7
ListView1.ColumnHeaders.Add , , "Name of Company", ListView1.Width / 3 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Physical Address", ListView1.Width / 2.5 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Postal Address", ListView1.Width / 3.5
ListView1.ColumnHeaders.Add , , "Country", ListView1.Width / 8
ListView1.ColumnHeaders.Add , , "Phone No", ListView1.Width / 5 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Fax/Telex", ListView1.Width / 5 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "E-Mail Address", ListView1.Width / 4.5
ListView1.ColumnHeaders.Add , , "I.T. NO", ListView1.Width / 6 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "PIN NUMBER", ListView1.Width / 5.8 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "NHIF NO", ListView1.Width / 6

ListView1.View = lvwReport

Exit Sub
Err:
    ErrorMessage
End Sub

Private Sub Form_Load()
On Error GoTo Err

    Call OpenConnection
    
    Call MyTreeSetup
    Call FindMyCompany
    
    Set CSearch = New clsSearchForRecord
    
    Exit Sub
Err:
    ErrorMessage
End Sub

Private Sub FindMyCompany()
On Error GoTo Err

ListView1.ListItems.Clear
ListView1.ColumnHeaders.Clear

ListView1.ColumnHeaders.Add , , "Company Code", ListView1.Width / 7
ListView1.ColumnHeaders.Add , , "Name of Company", ListView1.Width / 3 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Physical Address", ListView1.Width / 2.5 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Postal Address", ListView1.Width / 3.5
ListView1.ColumnHeaders.Add , , "Country", ListView1.Width / 8
ListView1.ColumnHeaders.Add , , "Phone No", ListView1.Width / 5 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Fax/Telex", ListView1.Width / 5 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "E-Mail Address", ListView1.Width / 4.5
ListView1.ColumnHeaders.Add , , "I.T. NO", ListView1.Width / 6 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "PIN NUMBER", ListView1.Width / 5.8 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "NHIF NO", ListView1.Width / 6

ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ParamCompanyMaster ORDER BY CompanyCode;", cnHOSPSQL, adOpenKeyset, adLockOptimistic
Dim DF
DF = rsLIST.RecordCount
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = ListView1.ListItems.Add(, , CStr(rsLIST!companycode))

    If Not IsNull(rsLIST!CompanyName) Then
        MyList.SubItems(1) = CStr(rsLIST!CompanyName)
    End If
    
    If Not IsNull(rsLIST!physicaladdress) Then
        MyList.SubItems(2) = CStr(rsLIST!physicaladdress)
    End If
    
    If Not IsNull(rsLIST!postaladdress) And Not IsNull(rsLIST!city) Then
        MyList.SubItems(3) = CStr(Trim(rsLIST!postaladdress)) & " " & CStr(Trim(rsLIST!city))
    End If
    
    If Not IsNull(rsLIST!country) Then
        MyList.SubItems(4) = CStr(rsLIST!country)
    End If
    
    If Not IsNull(rsLIST!telephoneno) Then
        MyList.SubItems(5) = CStr(rsLIST!telephoneno)
    End If
    
    If Not IsNull(rsLIST!telexfax) Then
        MyList.SubItems(6) = CStr(Trim(rsLIST!telexfax))
    End If
    
    If Not IsNull(rsLIST!email) Then
        MyList.SubItems(7) = CStr(rsLIST!email)
    End If
    
    If Not IsNull(rsLIST!coyitno) Then
        MyList.SubItems(8) = CStr(rsLIST!coyitno)
    End If
    
    If Not IsNull(rsLIST!coypinno) Then
        MyList.SubItems(9) = CStr(Trim(rsLIST!coypinno))
    End If
    
    If Not IsNull(rsLIST!coynhifno) Then
        MyList.SubItems(10) = CStr(Trim(rsLIST!coynhifno))
    End If

    rsLIST.MoveNext
    
Wend
Set MyList = Nothing
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Private Sub FindMSOffice()
On Error GoTo Err

ListView1.ListItems.Clear
ListView1.ColumnHeaders.Clear

ListView1.ColumnHeaders.Add , , "Code", ListView1.Width / 4 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Applications", ListView1.Width / 2

ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ParamMSOffice ORDER BY Code;", cnHOSPSQL, adOpenKeyset, adLockOptimistic
Dim DF
DF = rsLIST.RecordCount
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = ListView1.ListItems.Add(, , CStr(rsLIST!code))

    If Not IsNull(rsLIST!appname) Then
        MyList.SubItems(1) = CStr(rsLIST!appname)
    End If
    
    rsLIST.MoveNext
    
Wend
Set MyList = Nothing
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Private Sub FindAreaCode()
On Error GoTo Err

ListView1.ListItems.Clear
ListView1.ColumnHeaders.Clear

ListView1.ColumnHeaders.Add , , "Code", ListView1.Width / 4
ListView1.ColumnHeaders.Add , , "Area Name", ListView1.Width / 3 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Area Code", ListView1.Width / 4 ', lvwColumnCenter

ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM Setareacode ORDER BY Codenumber;", cnHOSPSQL, adOpenKeyset, adLockOptimistic
Dim DF
DF = rsLIST.RecordCount
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = ListView1.ListItems.Add(, , CStr(rsLIST!codenumber))

    If Not IsNull(rsLIST!areaname) Then
        MyList.SubItems(1) = CStr(rsLIST!areaname)
    End If
    
    If Not IsNull(rsLIST!AreaCode) Then
        MyList.SubItems(2) = CStr(rsLIST!AreaCode)
    End If
    
    rsLIST.MoveNext
    
Wend
'Me.ListView1.ColumnHeaders(2).Alignment = lvwColumnRight
Set MyList = Nothing
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Private Sub FindCountryCode()
On Error GoTo Err

ListView1.ListItems.Clear
ListView1.ColumnHeaders.Clear

ListView1.ColumnHeaders.Add , , "Code", ListView1.Width / 4
ListView1.ColumnHeaders.Add , , "Country Name", ListView1.Width / 3 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Country Code", ListView1.Width / 4 ', lvwColumnCenter

ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM Setcountrycode ORDER BY Codenumber;", cnHOSPSQL, adOpenKeyset, adLockOptimistic
Dim DF
DF = rsLIST.RecordCount
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = ListView1.ListItems.Add(, , CStr(rsLIST!codenumber))

    If Not IsNull(rsLIST!countryname) Then
        MyList.SubItems(1) = CStr(rsLIST!countryname)
    End If
    
    If Not IsNull(rsLIST!CountryCode) Then
        MyList.SubItems(2) = CStr(rsLIST!CountryCode)
    End If

    rsLIST.MoveNext
    
Wend
'Me.ListView1.ColumnHeaders(2).Alignment = lvwColumnRight
Set MyList = Nothing
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Private Sub FindVATRate()
On Error GoTo Err

ListView1.ListItems.Clear
ListView1.ColumnHeaders.Clear

ListView1.ColumnHeaders.Add , , "Code", ListView1.Width / 4
ListView1.ColumnHeaders.Add , , "VAT Rate [%]", ListView1.Width / 3 ', lvwColumnCenter

ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ParamVATRate ORDER BY Codenumber;", cnHOSPSQL, adOpenKeyset, adLockOptimistic
Dim DF
DF = rsLIST.RecordCount
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = ListView1.ListItems.Add(, , CStr(rsLIST!codenumber))

    If Not IsNull(rsLIST!VATRate) Then
        MyList.SubItems(1) = CStr(rsLIST!VATRate)
        MyList.SubItems(1) = FormatNumber(rsLIST!VATRate, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    rsLIST.MoveNext
    
Wend
Me.ListView1.ColumnHeaders(2).Alignment = lvwColumnRight
Set MyList = Nothing
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Private Sub FindLocalCurrency()
On Error GoTo Err

ListView1.ListItems.Clear
ListView1.ColumnHeaders.Clear

ListView1.ColumnHeaders.Add , , "Code", ListView1.Width / 5
ListView1.ColumnHeaders.Add , , "Currency Name", ListView1.Width / 3 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Curency Symbol", ListView1.Width / 3 ', lvwColumnCenter

ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM SetLocalCurrency ORDER BY Codenumber;", cnHOSPSQL, adOpenKeyset, adLockOptimistic
Dim DF
DF = rsLIST.RecordCount
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = ListView1.ListItems.Add(, , CStr(rsLIST!codenumber))

    If Not IsNull(rsLIST!currencyname) Then
        MyList.SubItems(1) = CStr(rsLIST!currencyname)
    End If
    
    If Not IsNull(rsLIST!Currency) Then
        MyList.SubItems(2) = CStr(rsLIST!Currency)
    End If
    
    rsLIST.MoveNext
    
Wend
Set MyList = Nothing
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Private Sub FindDefaultPayMethod()
On Error GoTo Err

ListView1.ListItems.Clear
ListView1.ColumnHeaders.Clear

ListView1.ColumnHeaders.Add , , "Serial No", ListView1.Width / 6
ListView1.ColumnHeaders.Add , , "Payment Code", ListView1.Width / 4 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Payment Method", ListView1.Width / 4 ', lvwColumnCenter

ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM SetDefaultPayMethod ORDER BY serialno;", cnHOSPSQL, adOpenKeyset, adLockOptimistic
Dim DF
DF = rsLIST.RecordCount
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = ListView1.ListItems.Add(, , CStr(rsLIST!serialno))

    If Not IsNull(rsLIST!paycode) Then
        MyList.SubItems(1) = CStr(rsLIST!paycode)
    End If
    
    If Not IsNull(rsLIST!PayMethod) Then
        MyList.SubItems(2) = CStr(rsLIST!PayMethod)
    End If
    
    rsLIST.MoveNext
    
Wend

Set MyList = Nothing
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Private Sub FindDefaultCurrency()
On Error GoTo Err

ListView1.ListItems.Clear
ListView1.ColumnHeaders.Clear

ListView1.ColumnHeaders.Add , , "Code", ListView1.Width / 6
ListView1.ColumnHeaders.Add , , "Currency Name", ListView1.Width / 4 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Currency Symbol", ListView1.Width / 4 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Exchange Rate", ListView1.Width / 4 ', lvwColumnCenter

ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM SetDefaultCurrency ORDER BY codenumber;", cnHOSPSQL, adOpenKeyset, adLockOptimistic
Dim DF
DF = rsLIST.RecordCount
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = ListView1.ListItems.Add(, , CStr(rsLIST!codenumber))

    If Not IsNull(rsLIST!currencyname) Then
        MyList.SubItems(1) = CStr(rsLIST!currencyname)
    End If
    
    If Not IsNull(rsLIST!Currency) Then
        MyList.SubItems(2) = CStr(rsLIST!Currency)
    End If
    
    If Not IsNull(rsLIST!exchrate) Then
        MyList.SubItems(3) = CStr(rsLIST!exchrate)
        MyList.SubItems(3) = FormatNumber(rsLIST!exchrate, 5, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    rsLIST.MoveNext
    
Wend
Me.ListView1.ColumnHeaders(4).Alignment = lvwColumnRight
Set MyList = Nothing
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Private Sub FindCompany()
On Error GoTo Err

ListView1.ListItems.Clear
ListView1.ColumnHeaders.Clear

ListView1.ColumnHeaders.Add , , "Dept. Code", ListView1.Width / 8
ListView1.ColumnHeaders.Add , , "Department Name", ListView1.Width / 2.3 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Dept. Head", ListView1.Width / 4.5 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Notes/Descriptions", ListView1.Width / 1.3

ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ParamCompanyDepts ORDER BY DeptCode;", cnHOSPSQL, adOpenKeyset, adLockOptimistic
Dim DF
DF = rsLIST.RecordCount
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = ListView1.ListItems.Add(, , CStr(rsLIST!deptcode))

    If Not IsNull(rsLIST!deptname) Then
        MyList.SubItems(1) = CStr(rsLIST!deptname)
    End If
    
    If Not IsNull(rsLIST!hod) Then
        MyList.SubItems(2) = CStr(rsLIST!hod)
    End If
    
    If Not IsNull(rsLIST!deptnotes) Then
        MyList.SubItems(3) = CStr(rsLIST!deptnotes)
    End If
    
    rsLIST.MoveNext
    
Wend
Set MyList = Nothing
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Private Sub FindDepartments()
On Error GoTo Err

ListView1.ListItems.Clear
ListView1.ColumnHeaders.Clear

ListView1.ColumnHeaders.Add , , "Dept. Code", ListView1.Width / 8
ListView1.ColumnHeaders.Add , , "Department Name", ListView1.Width / 2.3 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Head of Department", ListView1.Width / 4.5 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Staff ID", ListView1.Width / 8 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Official Title", ListView1.Width / 3.1 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Notes/Comments", ListView1.Width / 1.4

ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ParamCompanyDepts ORDER BY DeptCode;", cnHOSPSQL, adOpenKeyset, adLockOptimistic
Dim DF
DF = rsLIST.RecordCount
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = ListView1.ListItems.Add(, , CStr(rsLIST!deptcode))

    If Not IsNull(rsLIST!deptname) Then
        MyList.SubItems(1) = CStr(rsLIST!deptname)
    End If
    
    If Not IsNull(rsLIST!hod) Then
        MyList.SubItems(2) = CStr(rsLIST!hod)
    End If
    
    If Not IsNull(rsLIST!hodstaffid) Then
        MyList.SubItems(3) = CStr(rsLIST!hodstaffid)
    End If
    
    If Not IsNull(rsLIST!officialtitle) Then
        MyList.SubItems(4) = CStr(rsLIST!officialtitle)
    End If
    
    If Not IsNull(rsLIST!deptnotes) Then
        MyList.SubItems(5) = CStr(rsLIST!deptnotes)
    End If
    
    rsLIST.MoveNext
    
Wend
Set MyList = Nothing
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Private Sub FindEmployers()
On Error GoTo Err

ListView1.ListItems.Clear
ListView1.ColumnHeaders.Clear

ListView1.ColumnHeaders.Add , , "Emp. Code", ListView1.Width / 7
ListView1.ColumnHeaders.Add , , "Name of Employer", ListView1.Width / 3 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Physical Address", ListView1.Width / 3 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Postal Address", ListView1.Width / 3
ListView1.ColumnHeaders.Add , , "Country", ListView1.Width / 7
ListView1.ColumnHeaders.Add , , "Phone No", ListView1.Width / 7 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Fax/Telex", ListView1.Width / 7
ListView1.ColumnHeaders.Add , , "Mobile No", ListView1.Width / 7 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Contact Person", ListView1.Width / 3 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Official Title", ListView1.Width / 3

ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ParamEmployer ORDER BY empcode;", cnHOSPSQL, adOpenKeyset, adLockOptimistic
Dim DF
DF = rsLIST.RecordCount
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = ListView1.ListItems.Add(, , CStr(rsLIST!empcode))

    If Not IsNull(rsLIST!empname) Then
        MyList.SubItems(1) = CStr(rsLIST!empname)
    End If
    
    If Not IsNull(rsLIST!physicaladdress) Then
        MyList.SubItems(2) = CStr(rsLIST!physicaladdress)
    End If
    
    If Not IsNull(rsLIST!postaladdress) And Not IsNull(rsLIST!towncity) Then
        MyList.SubItems(3) = CStr(Trim(rsLIST!postaladdress)) & " " & CStr(Trim(rsLIST!towncity))
    End If
    
    If Not IsNull(rsLIST!country) Then
        MyList.SubItems(4) = CStr(rsLIST!country)
    End If
    
    If Not IsNull(rsLIST!telephoneno) Then
        MyList.SubItems(5) = CStr(rsLIST!telephoneno)
    End If
    
    If Not IsNull(rsLIST!mobileno) Then
        MyList.SubItems(6) = CStr(rsLIST!mobileno)
    End If
    
    If Not IsNull(rsLIST!faxtelex) Then
        MyList.SubItems(7) = CStr(rsLIST!faxtelex)
    End If
    
    If Not IsNull(rsLIST!contallnames) Then
        MyList.SubItems(8) = CStr(rsLIST!contallnames)
    End If
    
    If Not IsNull(rsLIST!contofficialtitle) Then
        MyList.SubItems(9) = CStr(rsLIST!contofficialtitle)
    End If
    
    rsLIST.MoveNext
    
Wend
Set MyList = Nothing
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Private Sub FindAccPeriods()
On Error GoTo Err

ListView1.ListItems.Clear
ListView1.ColumnHeaders.Clear

ListView1.ColumnHeaders.Add , , "Period Code", ListView1.Width / 7
ListView1.ColumnHeaders.Add , , "Month Name", ListView1.Width / 7 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Quarter", ListView1.Width / 7 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Sample Period", ListView1.Width / 6 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Start Date", ListView1.Width / 5.2
ListView1.ColumnHeaders.Add , , "End Date", ListView1.Width / 5.2

ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ParamAccPeriod ORDER BY periodcode;", cnHOSPSQL, adOpenKeyset, adLockOptimistic
Dim DF
DF = rsLIST.RecordCount
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = ListView1.ListItems.Add(, , CStr(rsLIST!PeriodCode))

    If Not IsNull(rsLIST!AccPeriod) Then
        MyList.SubItems(1) = CStr(rsLIST!AccPeriod)
    End If
    
    If Not IsNull(rsLIST!quarterofyear) Then
        MyList.SubItems(2) = CStr(rsLIST!quarterofyear)
    End If
    
    If Not IsNull(rsLIST!fullperiod) Then
        MyList.SubItems(3) = CStr(rsLIST!fullperiod)
    End If
    
    If Not IsNull(rsLIST!startdate) Then
        MyList.SubItems(4) = CStr(rsLIST!startdate)
    End If
    
    If Not IsNull(rsLIST!EndDate) Then
        MyList.SubItems(5) = CStr(rsLIST!EndDate)
    End If
    
    rsLIST.MoveNext
    
Wend
Set MyList = Nothing
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Private Sub FindTaxesRates()
On Error GoTo Err

ListView1.ListItems.Clear
ListView1.ColumnHeaders.Clear

ListView1.ColumnHeaders.Add , , "Tax Code", ListView1.Width / 7
ListView1.ColumnHeaders.Add , , "Tax Name", ListView1.Width / 5 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Tax Rate[%]", ListView1.Width / 5 ', lvwColumnCenter

ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ParamTaxes ORDER BY taxcode;", cnHOSPSQL, adOpenKeyset, adLockOptimistic
Dim DF
DF = rsLIST.RecordCount
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = ListView1.ListItems.Add(, , CStr(rsLIST!taxcode))

    If Not IsNull(rsLIST!taxname) Then
        MyList.SubItems(1) = CStr(rsLIST!taxname)
    End If
    
    If Not IsNull(rsLIST!taxrate) Then
        MyList.SubItems(2) = CStr(rsLIST!taxrate)
        MyList.SubItems(2) = FormatNumber(rsLIST!taxrate, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    rsLIST.MoveNext
    
Wend

Me.ListView1.ColumnHeaders(3).Alignment = lvwColumnRight

Set MyList = Nothing

Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Private Sub FindFeeServices()
On Error GoTo Err

ListView1.ListItems.Clear
ListView1.ColumnHeaders.Clear

ListView1.ColumnHeaders.Add , , "Code No", ListView1.Width / 8
ListView1.ColumnHeaders.Add , , "Service / Fee", ListView1.Width / 5 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Description", ListView1.Width / 1.6 ', lvwColumnCenter

ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ParamFeesServices ORDER BY codenumber;", cnHOSPSQL, adOpenKeyset, adLockOptimistic
Dim DF
DF = rsLIST.RecordCount
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = ListView1.ListItems.Add(, , CStr(rsLIST!codenumber))

    If Not IsNull(rsLIST!servicefee) Then
        MyList.SubItems(1) = CStr(rsLIST!servicefee)
    End If
    
    If Not IsNull(rsLIST!descriptions) Then
        MyList.SubItems(2) = CStr(rsLIST!descriptions)
    End If
    
    rsLIST.MoveNext
    
Wend
Set MyList = Nothing
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Private Sub FindPaymentMethods()
On Error GoTo Err

ListView1.ListItems.Clear
ListView1.ColumnHeaders.Clear

ListView1.ColumnHeaders.Add , , "Code No", ListView1.Width / 5
ListView1.ColumnHeaders.Add , , "Payment Method", ListView1.Width / 3 ', lvwColumnCenter

ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ParamPayMethods ORDER BY codenumber;", cnHOSPSQL, adOpenKeyset, adLockOptimistic
Dim DF
DF = rsLIST.RecordCount
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = ListView1.ListItems.Add(, , CStr(rsLIST!codenumber))

    If Not IsNull(rsLIST!PayMethod) Then
        MyList.SubItems(1) = CStr(rsLIST!PayMethod)
    End If
    
    rsLIST.MoveNext
    
Wend
Set MyList = Nothing
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Private Sub FindTitles()
On Error GoTo Err

ListView1.ListItems.Clear
ListView1.ColumnHeaders.Clear

ListView1.ColumnHeaders.Add , , "Title ID", ListView1.Width / 6
ListView1.ColumnHeaders.Add , , "Title of Courtesy", ListView1.Width / 4 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Description", ListView1.Width / 3 ', lvwColumnCenter

ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ParamTitles ORDER BY TitleID;", cnHOSPSQL, adOpenKeyset, adLockOptimistic
Dim DF
DF = rsLIST.RecordCount
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = ListView1.ListItems.Add(, , CStr(rsLIST!titleid))

    If Not IsNull(rsLIST!Title) Then
        MyList.SubItems(1) = CStr(rsLIST!Title)
    End If
    
    If Not IsNull(rsLIST!Description) Then
        MyList.SubItems(2) = CStr(rsLIST!Description)
    End If
    
    rsLIST.MoveNext
    
Wend
Set MyList = Nothing
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Private Sub FindCurrencies()
On Error GoTo Err

ListView1.ListItems.Clear
ListView1.ColumnHeaders.Clear

ListView1.ColumnHeaders.Add , , "Curr. Code", ListView1.Width / 7
ListView1.ColumnHeaders.Add , , "Name of Currency", ListView1.Width / 4 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Symbol", ListView1.Width / 7 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Exchange Rate", ListView1.Width / 6
ListView1.ColumnHeaders.Add , , "Country Code", ListView1.Width / 7

ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ParamCurrencies ORDER BY codenumber;", cnHOSPSQL, adOpenKeyset, adLockOptimistic
Dim DF
DF = rsLIST.RecordCount
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = ListView1.ListItems.Add(, , CStr(rsLIST!codenumber))

    If Not IsNull(rsLIST!desccurrency) Then
        MyList.SubItems(1) = CStr(rsLIST!desccurrency)
    End If
    
    If Not IsNull(rsLIST!Currency) Then
        MyList.SubItems(2) = CStr(rsLIST!Currency)
    End If
    
    If Not IsNull(rsLIST!exchrate) Then
        MyList.SubItems(3) = CStr(rsLIST!exchrate)
        MyList.SubItems(3) = FormatNumber(rsLIST!exchrate, 5, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    If Not IsNull(rsLIST!CountryCode) Then
        MyList.SubItems(4) = CStr(rsLIST!CountryCode)
    End If

    rsLIST.MoveNext
    
Wend

Me.ListView1.ColumnHeaders(4).Alignment = lvwColumnRight

Set MyList = Nothing

Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Private Sub FindCities()
On Error GoTo Err

ListView1.ListItems.Clear
ListView1.ColumnHeaders.Clear

ListView1.ColumnHeaders.Add , , "City Code", ListView1.Width / 8
ListView1.ColumnHeaders.Add , , "Name of City", ListView1.Width / 2.3 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Coyntry Code", ListView1.Width / 4.5 ', lvwColumnCenter

ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ParamCities ORDER BY citycode;", cnHOSPSQL, adOpenKeyset, adLockOptimistic
Dim DF
DF = rsLIST.RecordCount
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = ListView1.ListItems.Add(, , CStr(rsLIST!citycode))

    If Not IsNull(rsLIST!cityname) Then
        MyList.SubItems(1) = CStr(rsLIST!cityname)
    End If
    
    If Not IsNull(rsLIST!CountryCode) Then
        MyList.SubItems(2) = CStr(rsLIST!CountryCode)
    End If
    
    rsLIST.MoveNext
    
Wend
Set MyList = Nothing
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Private Sub FindCountries()
On Error GoTo Err

ListView1.ListItems.Clear
ListView1.ColumnHeaders.Clear

ListView1.ColumnHeaders.Add , , "Country Code", ListView1.Width / 5
ListView1.ColumnHeaders.Add , , "Country Name", ListView1.Width / 5 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Commercial Capital", ListView1.Width / 4 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Continent", ListView1.Width / 4

ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ParamCountries ORDER BY countrycode;", cnHOSPSQL, adOpenKeyset, adLockOptimistic
Dim DF
DF = rsLIST.RecordCount
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = ListView1.ListItems.Add(, , CStr(rsLIST!CountryCode))

    If Not IsNull(rsLIST!country) Then
        MyList.SubItems(1) = CStr(rsLIST!country)
    End If
    
    If Not IsNull(rsLIST!capitalcity) Then
        MyList.SubItems(2) = CStr(rsLIST!capitalcity)
    End If
    
    If Not IsNull(rsLIST!continent) Then
        MyList.SubItems(3) = CStr(rsLIST!continent)
    End If
    
    rsLIST.MoveNext
    
Wend
Set MyList = Nothing
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Private Sub FindEmployees()
On Error GoTo Err

ListView1.ListItems.Clear
ListView1.ColumnHeaders.Clear

ListView1.ColumnHeaders.Add , , "Staff ID", ListView1.Width / 7.5 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Full Names", ListView1.Width / 3 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Birth Date", ListView1.Width / 7.5
ListView1.ColumnHeaders.Add , , "Marital Status", ListView1.Width / 6 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Gender", ListView1.Width / 8
ListView1.ColumnHeaders.Add , , "Nationality", ListView1.Width / 6 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Nat. ID No", ListView1.Width / 7.5 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Passport No", ListView1.Width / 7.5 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "PIN Number", ListView1.Width / 7.5 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Date Hired", ListView1.Width / 7.5
ListView1.ColumnHeaders.Add , , "Official Title", ListView1.Width / 3 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Employ. Type", ListView1.Width / 6.5
ListView1.ColumnHeaders.Add , , "Grade", ListView1.Width / 9 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Dept. Code", ListView1.Width / 8 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Physical Address", ListView1.Width / 2.3 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Postal Address", ListView1.Width / 3.5 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Country", ListView1.Width / 7 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Phone No", ListView1.Width / 6
ListView1.ColumnHeaders.Add , , "Mobile No", ListView1.Width / 6 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "E-Mail Address", ListView1.Width / 4.5 ', lvwColumnCenter

ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ParamEmployeesMaster ORDER BY StaffIDNo;", cnHOSPSQL, adOpenKeyset, adLockOptimistic
Dim DF
DF = rsLIST.RecordCount
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = ListView1.ListItems.Add(, , CStr(rsLIST!staffidno))

    If Not IsNull(rsLIST!allnames) Then
        MyList.SubItems(1) = CStr(rsLIST!allnames)
    End If
    
    If Not IsNull(rsLIST!dateofbirth) Then
        MyList.SubItems(2) = CStr(rsLIST!dateofbirth)
    End If
    
    If Not IsNull(rsLIST!maritalstatus) Then
        MyList.SubItems(3) = CStr(rsLIST!maritalstatus)
    End If
    
    If Not IsNull(rsLIST!gender) Then
        MyList.SubItems(4) = CStr(rsLIST!gender)
    End If
  
    If Not IsNull(rsLIST!nationality) Then
        MyList.SubItems(5) = CStr(rsLIST!nationality)
    End If
    
    If Not IsNull(rsLIST!natidno) Then
        MyList.SubItems(6) = CStr(rsLIST!natidno)
    End If
    
    If Not IsNull(rsLIST!passportno) Then
        MyList.SubItems(7) = CStr(rsLIST!passportno)
    End If
    
    If Not IsNull(rsLIST!pinnumber) Then
        MyList.SubItems(8) = CStr(rsLIST!pinnumber)
    End If
    
    If Not IsNull(rsLIST!datehired) Then
        MyList.SubItems(9) = CStr(rsLIST!datehired)
    End If
  
    If Not IsNull(rsLIST!officialtitle) Then
        MyList.SubItems(10) = CStr(rsLIST!officialtitle)
    End If
    
    If Not IsNull(rsLIST!employtype) Then
        MyList.SubItems(11) = CStr(rsLIST!employtype)
    End If
    
    If Not IsNull(rsLIST!gradecode) Then
        MyList.SubItems(12) = CStr(rsLIST!gradecode)
    End If
    
    If Not IsNull(rsLIST!deptcode) Then
        MyList.SubItems(13) = CStr(rsLIST!deptcode)
    End If
    
    If Not IsNull(rsLIST!physicaladdress) Then
        MyList.SubItems(14) = CStr(rsLIST!physicaladdress)
    End If
    
    If Not IsNull(rsLIST!postaladdress) And Not IsNull(rsLIST!conttowncity) Then
        MyList.SubItems(15) = CStr(Trim(rsLIST!postaladdress)) & " " & CStr(Trim(rsLIST!conttowncity))
    End If
  
    If Not IsNull(rsLIST!contcountry) Then
        MyList.SubItems(16) = CStr(rsLIST!contcountry)
    End If
    
    If Not IsNull(rsLIST!conttelephone) Then
        MyList.SubItems(17) = CStr(rsLIST!conttelephone)
    End If
    
    If Not IsNull(rsLIST!contmobile) Then
        MyList.SubItems(18) = CStr(rsLIST!contmobile)
    End If
  
    If Not IsNull(rsLIST!contemail) Then
        MyList.SubItems(19) = CStr(rsLIST!contemail)
    End If

    rsLIST.MoveNext
Wend
Set MyList = Nothing
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Private Sub FindAllClaimsDrugs()
On Error GoTo Err

ListView1.ListItems.Clear
ListView1.ColumnHeaders.Clear

ListView1.ColumnHeaders.Add , , "Visit No", ListView1.Width / 7.5
ListView1.ColumnHeaders.Add , , "Membership No", ListView1.Width / 6 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Family No", ListView1.Width / 7 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Reg. Date", ListView1.Width / 7 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Trans. Date", ListView1.Width / 7 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Company", ListView1.Width / 7 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Procedures", ListView1.Width / 7
ListView1.ColumnHeaders.Add , , "Visit Date", ListView1.Width / 7 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Visit Time", ListView1.Width / 7
ListView1.ColumnHeaders.Add , , "Clinic", ListView1.Width / 7 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Drug Cost", ListView1.Width / 7 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Proc. Cost", ListView1.Width / 7 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Total Cost", ListView1.Width / 7 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Claim With", ListView1.Width / 7
ListView1.ColumnHeaders.Add , , "Description", ListView1.Width / 7 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Approval", ListView1.Width / 7

ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ProcessClaims;", cnHOSPSQL, adOpenKeyset, adLockOptimistic
Dim DF
DF = rsLIST.RecordCount
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = ListView1.ListItems.Add(, , CStr(rsLIST!visitno))

    If Not IsNull(rsLIST!visitno) Then
        MyList.SubItems(1) = CStr(rsLIST!visitno)
    End If
    
    If Not IsNull(rsLIST!membershipno) Then
        MyList.SubItems(2) = CStr(rsLIST!membershipno)
    End If
    
    If Not IsNull(rsLIST!familyno) Then
        MyList.SubItems(3) = CStr(rsLIST!familyno)
    End If
    
    If Not IsNull(rsLIST!registrationdate) Then
        MyList.SubItems(4) = CStr(rsLIST!registrationdate)
    End If
  
    If Not IsNull(rsLIST!transactiondate) Then
        MyList.SubItems(5) = CStr(rsLIST!transactiondate)
    End If
    
    If Not IsNull(rsLIST!company) Then
        MyList.SubItems(6) = CStr(rsLIST!company)
    End If
    
    If Not IsNull(rsLIST!procedures) Then
        MyList.SubItems(7) = CStr(rsLIST!procedures)
    End If
    
    If Not IsNull(rsLIST!visitdate) Then
        MyList.SubItems(8) = CStr(rsLIST!visitdate)
    End If
    
    If Not IsNull(rsLIST!visittime) Then
        MyList.SubItems(9) = CStr(rsLIST!visittime)
    End If
    
    If Not IsNull(rsLIST!clinic) Then
        MyList.SubItems(10) = CStr(rsLIST!clinic)
    End If
  
    If Not IsNull(rsLIST!drugcost) Then
        MyList.SubItems(11) = CStr(rsLIST!drugcost)
    End If
    
    If Not IsNull(rsLIST!procedurecost) Then
        MyList.SubItems(12) = CStr(rsLIST!procedurecost)
    End If
    
    If Not IsNull(rsLIST!TotalCost) Then
        MyList.SubItems(13) = CStr(rsLIST!TotalCost)
    End If
    
    If Not IsNull(rsLIST!claimwith) Then
        MyList.SubItems(14) = CStr(rsLIST!claimwith)
    End If
    
    If Not IsNull(rsLIST!Description) Then
        MyList.SubItems(15) = CStr(rsLIST!Description)
    End If
    
    If Not IsNull(rsLIST!approved) Then
        MyList.SubItems(16) = CStr(rsLIST!approved)
    End If
  
    rsLIST.MoveNext
Wend
Set MyList = Nothing
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Private Sub FindAllClaimsDiag()
On Error GoTo Err

ListView1.ListItems.Clear
ListView1.ColumnHeaders.Clear

ListView1.ColumnHeaders.Add , , "Code Number", ListView1.Width / 7 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Visit No", ListView1.Width / 7.5
ListView1.ColumnHeaders.Add , , "Membership No", ListView1.Width / 6 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Doctor Code", ListView1.Width / 7 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Details", ListView1.Width / 7 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Disease", ListView1.Width / 7 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Condition", ListView1.Width / 7
ListView1.ColumnHeaders.Add , , "Condition Status", ListView1.Width / 7 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Provider", ListView1.Width / 7
ListView1.ColumnHeaders.Add , , "Disease ID", ListView1.Width / 7 ', lvwColumnCenter

ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ProcessMDiagnosis;", cnHOSPSQL, adOpenKeyset, adLockOptimistic
Dim DF
DF = rsLIST.RecordCount
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = ListView1.ListItems.Add(, , CStr(rsLIST!code1))

    If Not IsNull(rsLIST!visitno) Then
        MyList.SubItems(1) = CStr(rsLIST!visitno)
    End If
    
    If Not IsNull(rsLIST!membershipno) Then
        MyList.SubItems(2) = CStr(rsLIST!membershipno)
    End If
    
    If Not IsNull(rsLIST!doctorcode) Then
        MyList.SubItems(3) = CStr(rsLIST!doctorcode)
    End If
    
    If Not IsNull(rsLIST!Details) Then
        MyList.SubItems(4) = CStr(rsLIST!Details)
    End If
  
    If Not IsNull(rsLIST!disease) Then
        MyList.SubItems(5) = CStr(rsLIST!disease)
    End If
    
    If Not IsNull(rsLIST!Condition) Then
        MyList.SubItems(6) = CStr(rsLIST!Condition)
    End If
    
    If Not IsNull(rsLIST!conditionstatus) Then
        MyList.SubItems(7) = CStr(rsLIST!conditionstatus)
    End If
    
    If Not IsNull(rsLIST!Provider) Then
        MyList.SubItems(8) = CStr(rsLIST!Provider)
    End If
    
    If Not IsNull(rsLIST!diseaseid) Then
        MyList.SubItems(9) = CStr(rsLIST!diseaseid)
    End If
    
    rsLIST.MoveNext
Wend
Set MyList = Nothing
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SysmanMain.Enabled = True
End Sub

Private Sub ListView1_DblClick()
On Error GoTo Err
If Me.ListView1.ListItems.Count = 0 Then Exit Sub

    Set rsFIND = New ADODB.Recordset
    
    If txtREGS.Text = "PBS" Then
        rsFIND.Open "SELECT * FROM PatientsRegister WHERE Regnumber='" & Me.ListView1.SelectedItem.Text & "';", cnHOSPSQL, adOpenKeyset, adLockOptimistic
        With rsFIND
        If .EOF And .BOF Then Exit Sub
            CurrentRecord = !regnumber ': GoTo Outs
        End With
        Load frmEmployeesPersonal
        frmEmployeesPersonal.SSTab1.Tab = 0
        frmEmployeesPersonal.Show 1, Me: GoTo OUTS
    End If
    
    If txtREGS.Text = "MED" Then
        rsFIND.Open "SELECT * FROM PatientsMedical WHERE SerialNumber='" & Me.ListView1.SelectedItem.Text & "';", cnHOSPSQL, adOpenKeyset, adLockOptimistic
        With rsFIND
        If .EOF And .BOF Then Exit Sub
            CurrentRecord = !regnumber ': GoTo Outs
        End With
        Load frmEmployeesPersonal
        frmEmployeesPersonal.SSTab1.Tab = 1
        frmEmployeesPersonal.Show 1, Me: GoTo OUTS
    End If
    
OUTS:
    Set rsFIND = Nothing
    Exit Sub
Err:
    ErrorMessage
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    With Me
        Select Case .txtREGS.Text
        Case "C1"
            PopupMenu mnuSeacrhCompany
        Case "C2"
            PopupMenu mnuSeacrhCompany
        Case "C3"
            PopupMenu mnuSearchDepartment
        Case "C4"
            PopupMenu mnuSearchDepartment
        Case "C5"
            PopupMenu mnuSearchStaff
        Case "C6"
            PopupMenu mnuSearchStaff
        Case Else
            Exit Sub
        End Select
    End With
End If
End Sub

Private Sub mnuSearchCompanyByCode_Click()
On Error GoTo Err
    Me.ListView1.ListItems.Clear
    CSearch.SearchCompanyByName
    Exit Sub
Err:
    ErrorMessage
End Sub

Private Sub mnuSearchCompanyByName_Click()
On Error GoTo Err
    Me.ListView1.ListItems.Clear
    CSearch.SearchCompanyByName
    Exit Sub
Err:
    ErrorMessage
End Sub

Private Sub mnuSearchDeptByCode_Click()
On Error GoTo Err
    Me.ListView1.ListItems.Clear
    CSearch.SearchDepartmentByCode
    Exit Sub
Err:
    ErrorMessage
End Sub

Private Sub mnuSearchDeptByName_Click()
On Error GoTo Err
    Me.ListView1.ListItems.Clear
    CSearch.SearchDepartmentByName
    Exit Sub
Err:
    ErrorMessage
End Sub

Private Sub mnuSearchStaffByID_Click()
On Error GoTo Err
    Me.ListView1.ListItems.Clear
    CSearch.SearchEmployeeByID
    Exit Sub
Err:
    ErrorMessage
End Sub

Private Sub mnuSearchStaffBySurname_Click()
On Error GoTo Err
    Me.ListView1.ListItems.Clear
    CSearch.SearchEmployeeBySurname
    Exit Sub
Err:
    ErrorMessage
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo Err
    Select Case Node.Key
    Case "C1"
        Me.txtREGS.Text = "C1"
        Load frmCompany
        frmCompany.Show 1, Me
    Case "C2"
        FindMyCompany
        Me.txtREGS.Text = "C2"
        Me.ListView1.SetFocus
    Case "C3"
        Me.txtREGS.Text = "C3"
        Load frmCompanyDepartments
        frmCompanyDepartments.Show 1, Me
    Case "C4"
        FindDepartments
        Me.txtREGS.Text = "C4"
        Me.ListView1.SetFocus
    Case "C5"
        Me.txtREGS.Text = "C5"
        Load frmEmployeesPersonal
        frmEmployeesPersonal.Show 1, Me
    Case "C6"
        FindEmployees
        Me.txtREGS.Text = "C6"
        Me.ListView1.SetFocus
    Case "D1"
        Me.txtREGS.Text = "D1"
        Load frmParamEmployers
        frmParamEmployers.Show 1, Me
        FindEmployers
    Case "D2"
        Me.txtREGS.Text = "D2"
        Load frmParamCountries
        frmParamCountries.Show 1, Me
        FindCountries
    Case "D3"
        Me.txtREGS.Text = "D3"
        Load frmParamCities
        frmParamCities.Show 1, Me
        FindCities
    Case "D4"
        Me.txtREGS.Text = "D4"
        Load frmParamCurrencies
        frmParamCurrencies.Show 1, Me
        FindCurrencies
    Case "D5"
        Me.txtREGS.Text = "D5"
        Load frmParamTitles
        frmParamTitles.Show 1, Me
        FindTitles
    Case "D6"
        Me.txtREGS.Text = "D6"
        Load frmParamPayMethods
        frmParamPayMethods.Show 1, Me
        FindPaymentMethods
    Case "D7"
        Me.txtREGS = "D7"
        Load frmParamFeeServices
        frmParamFeeServices.Show 1, Me
        FindFeeServices
    Case "D8"
        Me.txtREGS = "D8"
        Load frmParamTaxes
        frmParamTaxes.Show 1, Me
        FindTaxesRates
    Case "D9"
        Me.txtREGS.Text = "D9"
        Load frmParamAccPeriods
        frmParamAccPeriods.Show 1, Me
        FindAccPeriods
    Case "P0"
        Me.txtREGS.Text = "P0"
        Load frmSettings
        frmSettings.Show 1, Me
    Case "P1"
        FindDefaultCurrency
        Me.txtREGS.Text = "P1"
        Me.ListView1.SetFocus
    Case "P2"
        FindLocalCurrency
        Me.txtREGS = "P2"
        Me.ListView1.SetFocus
    Case "P3"
        FindVATRate
        Me.txtREGS = "P3"
        Me.ListView1.SetFocus
    Case "P4"
        Me.txtREGS.Text = "P4"
        FindCountryCode
        Me.ListView1.SetFocus
    Case "P5"
        FindAreaCode
        Me.txtREGS = "P5"
        Me.ListView1.SetFocus
    Case "P6"
        FindDefaultPayMethod
        Me.txtREGS = "P6"
        Me.ListView1.SetFocus
    Case "P7"
        Me.txtREGS.Text = "P7"
        Load frmOfficeSettings
        frmOfficeSettings.Show 1, Me
    Case "P8"
        FindMSOffice
        Me.txtREGS = "P8"
        Me.ListView1.SetFocus
    Case "V1"
        FindEmployees
        Me.txtREGS.Text = "V1"
        Me.ListView1.SetFocus
    Case "V2"
        FindEmployers
        Me.txtREGS = "V2"
        Me.ListView1.SetFocus
    Case "V3"
        FindCurrencies
        Me.txtREGS = "V3"
        Me.ListView1.SetFocus
    Case "V4"
        FindTaxesRates
        Me.txtREGS = "V4"
        Me.ListView1.SetFocus
    Case "V5"
        FindFeeServices
        Me.txtREGS = "V5"
        Me.ListView1.SetFocus
    Case "V6"
        FindPaymentMethods
        Me.txtREGS = "V6"
        Me.ListView1.SetFocus
    Case Else
        GetMainStructure
        Exit Sub
    End Select
Exit Sub
Err:
    ErrorMessage
End Sub
