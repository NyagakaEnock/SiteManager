VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMainMenu"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Sub GetMainStructure()
On Error GoTo err
    With SchedulingMain
        .ListView1.ListItems.Clear
        .ListView1.ColumnHeaders.Clear
        .ListView1.ColumnHeaders.Add , , "User Name", .ListView1.Width / 6.5
        .ListView1.ColumnHeaders.Add , , "Staff ID", .ListView1.Width / 6.5 ', lvwColumnCenter
        .ListView1.ColumnHeaders.Add , , "Full Names", .ListView1.Width / 3 ', lvwColumnCenter
        .ListView1.ColumnHeaders.Add , , "User Group", .ListView1.Width / 6.5
        .ListView1.ColumnHeaders.Add , , "Account Status", .ListView1.Width / 5
        .ListView1.View = lvwReport
    End With
    Exit Sub
err:
    ErrorMessage
End Sub

Public Sub OpenMain()
On Error GoTo err
    Load frmLogin
    frmLogin.Show 1
    Exit Sub
err:
    ErrorMessage
End Sub

Public Sub ResizeControls()
On Error GoTo err
With SchedulingMain
    .TreeView1.Left = 0: .TreeView1.Top = 720
    .ListView1.Left = 2760: .ListView1.Top = 720
    
    .ListView1.Width = .Width - 2880: .ListView1.Height = .Height - 1860
    .TreeView1.Height = .Height - 1860: .txtTotal.Top = .Height - 1935
End With
Exit Sub
err:
If err.Number = 380 Then Resume Next
    ErrorMessage
End Sub

Public Sub UnLoadMyForm()
    Call UpdateLogoutRecord
    Unload SchedulingMain
End Sub

Public Sub HelpContents()
On Error GoTo err
With SchedulingMain
    Dim nRet As Integer
    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, .Caption
    Else
        'On Error Resume Next
        nRet = WinHelp(.hwnd, App.HelpFile, 3, 0)
        If err Then
            MsgBox err.Description
        End If
    End If
End With
Exit Sub
err:
    ErrorMessage
End Sub

Public Sub HelpIndex()
On Error GoTo err
With SchedulingMain
    Dim nRet As Integer

    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, .Caption
    Else
        'On Error Resume Next
        nRet = WinHelp(.hwnd, App.HelpFile, 261, 0)
        If err Then
            MsgBox err.Description
        End If
    End If
End With
Exit Sub
err:
    ErrorMessage
End Sub

Public Sub HelpAbout()
On Error GoTo err
    Load frmAbout
    frmAbout.Show 1
Exit Sub
err:
    ErrorMessage
End Sub

Public Sub CreateMAINMENU()
On Error GoTo err
Dim Node1 As Node

With SchedulingMain
    
    Dim NodeK As Node

    Set NodeK = .TreeView1.Nodes.Add(, , "K", "Contracts Manager", "C")
    Set NodeK = .TreeView1.Nodes.Add("K", tvwChild, "K1", "Site Acquisition", "F")
    Set NodeK = .TreeView1.Nodes.Add("K", tvwChild, "K101", "Voucher Preparation", "F")
    Set NodeK = .TreeView1.Nodes.Add("K", tvwChild, "K102", "Payment Confirmation", "F")
    Set NodeK = .TreeView1.Nodes.Add("K", tvwChild, "K2", "Assign Properties", "F")
    Set NodeK = .TreeView1.Nodes.Add("K", tvwChild, "K32", "Search Properties Assigned", "F")
    Set NodeK = .TreeView1.Nodes.Add("K", tvwChild, "K34", "All Plots Without Properties", "F")
    Set NodeK = .TreeView1.Nodes.Add("K", tvwChild, "K31", "LandLord Registration", "F")
    Set NodeK = .TreeView1.Nodes.Add("K", tvwChild, "K4", "Prepare Lease", "F")
    Set NodeK = .TreeView1.Nodes.Add("K", tvwChild, "K5", "Contracts", "F")
    Set NodeK = .TreeView1.Nodes.Add("K5", tvwChild, "K52", "ContractsAuthorization", "F")
    Set NodeK = .TreeView1.Nodes.Add("K", tvwChild, "K22", "Edit Lease", "F")
    Set NodeK = .TreeView1.Nodes.Add("K", tvwChild, "K20", "Set Council Rates", "F")
    Set NodeK = .TreeView1.Nodes.Add("K", tvwChild, "K25", "Print Rates Schedule", "F")
    Set NodeK = .TreeView1.Nodes.Add("K", tvwChild, "K21", "Print Rent Installments Sheet", "F")
    Set NodeK = .TreeView1.Nodes.Add("K", tvwChild, "K9", "Print Contract", "F")
    Set NodeK = .TreeView1.Nodes.Add("K9", tvwChild, "K94", "Add Clause", "F")
    Set NodeK = .TreeView1.Nodes.Add("K9", tvwChild, "K93", "Edit Main Contract Clauses", "F")
    Set NodeK = .TreeView1.Nodes.Add("K9", tvwChild, "K95", "Print Edited Contract", "F")
    Set NodeK = .TreeView1.Nodes.Add("K9", tvwChild, "K91", "Print Main Contract", "F")
    Set NodeK = .TreeView1.Nodes.Add("K9", tvwChild, "K92", "Print NACADA Ompliant Contract", "F")
    Set NodeK = .TreeView1.Nodes.Add("K", tvwChild, "K24", "Print Schedule", "F")
    Set NodeK = .TreeView1.Nodes.Add("K", tvwChild, "K150", "Free Assigned Sites", "F")

    
    Dim NodeX As Node
    Set NodeX = .TreeView1.Nodes.Add(, , "X", "Summaries", "C")
    Set NodeX = .TreeView1.Nodes.Add("X", tvwChild, "X2", "Sites on Road Reserve", "F")
    Set NodeX = .TreeView1.Nodes.Add("X", tvwChild, "X3", "Non-Company Structures", "F")
    Set NodeX = .TreeView1.Nodes.Add("X", tvwChild, "X4", "Annual Rent for All Plots", "F")
    Set NodeX = .TreeView1.Nodes.Add("X", tvwChild, "X5", "All Expired Sites", "F")
    Set NodeX = .TreeView1.Nodes.Add("X", tvwChild, "K8", "Almost Expiring", "F")
    Set NodeX = .TreeView1.Nodes.Add("X", tvwChild, "K13", "Leases Due to Expire", "F")
    Set NodeX = .TreeView1.Nodes.Add("X", tvwChild, "K14", "Sites Terminated/Cancelled-By Company", "F")
    Set NodeX = .TreeView1.Nodes.Add("X", tvwChild, "K141", "Sites Terminated/Cancelled-By Owner", "F")
    Set NodeX = .TreeView1.Nodes.Add("X", tvwChild, "K23", "New Sites", "F")
    Set NodeX = .TreeView1.Nodes.Add("X", tvwChild, "K18", "Sites Unpaid", "F")
    Set NodeX = .TreeView1.Nodes.Add("X", tvwChild, "K19", "Sites Paid", "F")
    Set NodeX = .TreeView1.Nodes.Add("X", tvwChild, "K16", "All Un-Leased Plots", "F")
    Set NodeX = .TreeView1.Nodes.Add("X", tvwChild, "K30", "All Notices Issued", "F")
    Set NodeK = .TreeView1.Nodes.Add("X", tvwChild, "K33", "Plots' Details General Search", "F")
    Set NodeK = .TreeView1.Nodes.Add("X", tvwChild, "K17", "All FREE Sites", "F")
    Set NodeK = .TreeView1.Nodes.Add("X", tvwChild, "X6", "Load ", "F")
    Dim NodeN As Node

    Set NodeN = .TreeView1.Nodes.Add(, , "N", "Notices Manager", "C")
    Set NodeN = .TreeView1.Nodes.Add("N", tvwChild, "N1", "Prepare Notice", "F")
    Set NodeN = .TreeView1.Nodes.Add("N", tvwChild, "N3", "Notice Authorization", "F")
    Set NodeN = .TreeView1.Nodes.Add("N", tvwChild, "N4", "Send Notice", "F")
    Set NodeN = .TreeView1.Nodes.Add("N", tvwChild, "N6", "Termination", "F")
    Set NodeN = .TreeView1.Nodes.Add("N", tvwChild, "N12", "Renewal", "F")
    Set NodeN = .TreeView1.Nodes.Add("N", tvwChild, "N13", "Renewal Contracts", "F")
    Set NodeN = .TreeView1.Nodes.Add("N", tvwChild, "N7", "Notices Prepared", "F")
    Set NodeN = .TreeView1.Nodes.Add("N", tvwChild, "N8", "Notices Approved", "F")
    Set NodeN = .TreeView1.Nodes.Add("N", tvwChild, "N9", "Notices Authorized", "F")
    Set NodeN = .TreeView1.Nodes.Add("N", tvwChild, "N10", "Notices Sent", "F")
    Set NodeN = .TreeView1.Nodes.Add("N", tvwChild, "N11", "Notices Acknowledged", "F")

    Dim NodeR As Node

    Set NodeR = .TreeView1.Nodes.Add(, , "R", "Renewals Master", "C")
    Set NodeR = .TreeView1.Nodes.Add("R", tvwChild, "R11", "RewNew Jobs", "F")
    
    Set Node1 = .TreeView1.Nodes.Add(, , "G", "General/Quick Queries", "C")
    Set Node1 = .TreeView1.Nodes.Add("G", tvwChild, "G1", "Show All Pinned Contracts", "F")
    Set Node1 = .TreeView1.Nodes.Add("G", tvwChild, "G2", "Show Contracts Expiring on Specific Date ", "F")
    Set Node1 = .TreeView1.Nodes.Add("G", tvwChild, "G3", "Show CouncilFees Due On A Specific Date", "F")
    Set Node1 = .TreeView1.Nodes.Add("G", tvwChild, "G4", "Show Rent Fees Due On A Specific Date", "F")
    Set Node1 = .TreeView1.Nodes.Add("G", tvwChild, "G5", "Show Empty BillBoards", "F")
    Set Node1 = .TreeView1.Nodes.Add("G", tvwChild, "G7", "Show All LandLords", "F")
    Set Node1 = .TreeView1.Nodes.Add("G", tvwChild, "G13", "Show Sites Un-Allocated", "F")
    Set Node1 = .TreeView1.Nodes.Add("G", tvwChild, "G9", "Show Sites with Adverts", "F")
    Set Node1 = .TreeView1.Nodes.Add("G", tvwChild, "G11", "Show Sites Reserved", "F")
    Set Node1 = .TreeView1.Nodes.Add("G", tvwChild, "G12", "Show Sites To Free", "F")
    Set Node1 = .TreeView1.Nodes.Add("G", tvwChild, "G10", "Show Sites Expiry On Specific Date", "F")
    
    Dim NodeM As Node

    Set NodeM = .TreeView1.Nodes.Add(, , "M", "Site Maintenance", "C")
    Set NodeM = .TreeView1.Nodes.Add("M", tvwChild, "M8", "Site Maintanance Schedule", "F")
    Set NodeM = .TreeView1.Nodes.Add("M", tvwChild, "M2", "Site Due For Maintenance", "F")
    Set NodeM = .TreeView1.Nodes.Add("M2", tvwChild, "M4", "1 week to", "F")
    Set NodeM = .TreeView1.Nodes.Add("M2", tvwChild, "M3", "1 month to", "F")
    Set NodeM = .TreeView1.Nodes.Add("M2", tvwChild, "M6", "Specific Date", "F")
    Set NodeM = .TreeView1.Nodes.Add("M2", tvwChild, "M5", "Specific Period", "F")
    
    Dim NodeZ As Node
    Set NodeZ = .TreeView1.Nodes.Add(, , "Z", "Site Schedule Monitor", "C")
    Set NodeZ = .TreeView1.Nodes.Add("Z", tvwChild, "Z1", "Sites Scheduled", "F")
    Set NodeZ = .TreeView1.Nodes.Add("Z", tvwChild, "Z2", "View Site Schedule", "F")
    Set NodeZ = .TreeView1.Nodes.Add("Z", tvwChild, "Z3", "Seacrh Period with Free Sites ", "F")
End With
Exit Sub
err:
    ErrorMessage
End Sub
