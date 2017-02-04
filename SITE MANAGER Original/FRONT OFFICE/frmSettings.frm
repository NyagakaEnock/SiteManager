VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DEFAULT SETTINGS / OPTIONS"
   ClientHeight    =   3735
   ClientLeft      =   2850
   ClientTop       =   1860
   ClientWidth     =   7215
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   7215
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   6588
      _Version        =   393216
      Tabs            =   6
      Tab             =   5
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Default Currency"
      TabPicture(0)   =   "frmSettings.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Local Currency"
      TabPicture(1)   =   "frmSettings.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "V.A.T. Rate"
      TabPicture(2)   =   "frmSettings.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Country Code [Telecom]"
      TabPicture(3)   =   "frmSettings.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame4"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Area Code [Telecom]"
      TabPicture(4)   =   "frmSettings.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame5"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Payment Method"
      TabPicture(5)   =   "frmSettings.frx":04CE
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).Control(0)=   "Frame6"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).ControlCount=   1
      Begin VB.Frame Frame6 
         Height          =   2895
         Left            =   120
         TabIndex        =   41
         Top             =   720
         Width           =   6855
         Begin VB.TextBox txtPayCode 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   45
            Top             =   1200
            Width           =   1335
         End
         Begin VB.TextBox txtPaySerialNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   44
            Top             =   240
            Width           =   1335
         End
         Begin VB.ComboBox cboPayMEthod 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   1440
            TabIndex        =   43
            Top             =   720
            Width           =   2295
         End
         Begin VB.CommandButton cmdPayChange 
            BackColor       =   &H80000001&
            Caption         =   "&Change"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   5040
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label18 
            Caption         =   "Pay Code"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label17 
            Caption         =   "Payment Method"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label16 
            Caption         =   "Serial No"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame5 
         Height          =   2895
         Left            =   -74880
         TabIndex        =   33
         Top             =   720
         Width           =   6855
         Begin VB.CommandButton cmdAreaChange 
            BackColor       =   &H80000001&
            Caption         =   "&Change"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   5040
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   240
            Width           =   1695
         End
         Begin VB.ComboBox cboAreaName 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   1440
            TabIndex        =   36
            Top             =   720
            Width           =   2295
         End
         Begin VB.TextBox txtAreaCodeNumber 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   35
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtAreaCode 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   1440
            TabIndex        =   34
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label Label15 
            Caption         =   "Code Number"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label14 
            Caption         =   "Area Name"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label13 
            Caption         =   "Local/Area Code"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   1200
            Width           =   1215
         End
      End
      Begin VB.Frame Frame4 
         Height          =   2895
         Left            =   -74880
         TabIndex        =   25
         Top             =   720
         Width           =   6855
         Begin VB.CommandButton cmdCountryChange 
            BackColor       =   &H80000001&
            Caption         =   "&Change"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   5040
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   240
            Width           =   1695
         End
         Begin VB.ComboBox cboCountryName 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   1440
            TabIndex        =   28
            Top             =   720
            Width           =   2295
         End
         Begin VB.TextBox txtCountryCNumber 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtCountryCode 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   1440
            TabIndex        =   26
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label Label12 
            Caption         =   "Code Number"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label11 
            Caption         =   "Country Name"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label10 
            Caption         =   "Country Code"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   1200
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Height          =   2895
         Left            =   -74880
         TabIndex        =   8
         Top             =   720
         Width           =   6855
         Begin VB.TextBox txtVVat 
            Alignment       =   2  'Center
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
            Left            =   1440
            TabIndex        =   22
            Top             =   840
            Width           =   2055
         End
         Begin VB.TextBox txtVCode 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Height          =   375
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   360
            Width           =   2055
         End
         Begin VB.CommandButton cmdVChange 
            BackColor       =   &H80000001&
            Caption         =   "&Change"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   5040
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label9 
            Caption         =   "V.A.T. Rate"
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label8 
            Caption         =   "Code Number"
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2895
         Left            =   -74880
         TabIndex        =   7
         Top             =   720
         Width           =   6855
         Begin VB.TextBox txtLSymbol 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   1200
            Width           =   1335
         End
         Begin VB.TextBox txtLCode 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   240
            Width           =   1335
         End
         Begin VB.ComboBox cboLName 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   1440
            TabIndex        =   15
            Top             =   720
            Width           =   2295
         End
         Begin VB.CommandButton cmdLChange 
            BackColor       =   &H80000001&
            Caption         =   "&Change"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   5040
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label7 
            Caption         =   "Currency Symbol"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "Curency Name"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "Code Number"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2895
         Left            =   -74880
         TabIndex        =   6
         Top             =   720
         Width           =   6855
         Begin VB.TextBox txtRate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   315
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox txtSymbol 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   1200
            Width           =   1335
         End
         Begin VB.ComboBox cboName 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   1320
            TabIndex        =   1
            Top             =   720
            Width           =   2415
         End
         Begin VB.TextBox txtCode 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   0
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdChange 
            BackColor       =   &H80000001&
            Caption         =   "&Change"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   5040
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label4 
            Caption         =   "Exchange Rate"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Symbol"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Currency Name"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Code"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rsDEFAULT As ADODB.Recordset, rsPAYMETHOD As ADODB.Recordset, rsCOUNTRY As ADODB.Recordset, rsAREA As ADODB.Recordset, rsLOCAL As ADODB.Recordset, rsVAT As ADODB.Recordset

Private Sub cboAreaName_Click()
If Not NewRecord And Not EditRecord Then Exit Sub
    Me.cmdAreaChange.SetFocus: Me.cmdAreaChange.Default = True
End Sub

Private Sub cboAreaName_GotFocus()
On Error GoTo err
If Not NewRecord And Not EditRecord Then Exit Sub
    If Me.cboLName.ListCount <> 0 Then Me.cboLName.Refresh: Exit Sub
    
    Dim rsLIST As ADODB.Recordset
    Set rsLIST = cnALIS.Execute("SELECT * FROM paramcurrencies ORDER BY currency;")
    Me.cboLName.Clear
    
    With rsLIST
    If .EOF And .BOF Then Exit Sub
       .MoveFirst
       Do While Not .EOF
            Me.cboLName.AddItem !desccurrency
            .MoveNext
       Loop
    End With
    
    Set rsLIST = Nothing
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub cboCountryName_Click()
If Not NewRecord And Not EditRecord Then Exit Sub
    Me.cmdCountryChange.SetFocus: Me.cmdCountryChange.Default = True
End Sub

Private Sub cboCountryName_GotFocus()
On Error GoTo err
If Not NewRecord And Not EditRecord Then Exit Sub
    If Me.cboCountryName.ListCount <> 0 Then Me.cboCountryName.Refresh: Exit Sub
    
    Dim rsLIST As ADODB.Recordset
    Set rsLIST = cnALIS.Execute("SELECT * FROM Paramcountries ORDER BY country;")
    Me.cboCountryName.Clear
    
    With rsLIST
    If .EOF And .BOF Then Exit Sub
       .MoveFirst
       Do While Not .EOF
            Me.cboCountryName.AddItem !country
            .MoveNext
       Loop
    End With
    
    Set rsLIST = Nothing
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub cboLName_Click()
If Not NewRecord And Not EditRecord Then Exit Sub
    Me.cmdLChange.SetFocus: Me.cmdLChange.Default = True
End Sub

Private Sub cboLName_GotFocus()
On Error GoTo err
If Not NewRecord And Not EditRecord Then Exit Sub
    If Me.cboLName.ListCount <> 0 Then Me.cboLName.Refresh: Exit Sub
    
    Dim rsLIST As ADODB.Recordset
    Set rsLIST = cnALIS.Execute("SELECT * FROM paramcurrencies ORDER BY currency;")
    Me.cboLName.Clear
    
    With rsLIST
    If .EOF And .BOF Then Exit Sub
       .MoveFirst
       Do While Not .EOF
            Me.cboLName.AddItem !desccurrency
            .MoveNext
       Loop
    End With
    
    Set rsLIST = Nothing
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub cboLName_LostFocus()
On Error GoTo err
    Dim rsNAME As ADODB.Recordset
    Set rsNAME = New ADODB.Recordset
    
    rsNAME.Open "SELECT * FROM ParamCurrencies WHERE Desccurrency = '" & Me.cboLName.Text & "';", cnALIS, adOpenKeyset, adLockOptimistic
    
    With rsNAME
    If .EOF And .BOF Then Exit Sub
        Me.txtLSymbol = !Currency & ""
        Me.cmdLChange.SetFocus: Me.cmdLChange.Default = True
    End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub cboName_Click()
If Not NewRecord And Not EditRecord Then Exit Sub
    Me.cmdChange.SetFocus: Me.cmdChange.Default = True
End Sub

Private Sub cboName_GotFocus()
On Error GoTo err
If Not NewRecord And Not EditRecord Then Exit Sub
    If Me.cboName.ListCount <> 0 Then Me.cboName.Refresh: Exit Sub
    
    Dim rsLIST As ADODB.Recordset
    Set rsLIST = cnALIS.Execute("SELECT * FROM paramcurrencies ORDER BY currency;")
    Me.cboName.Clear
    
    With rsLIST
    If .EOF And .BOF Then Exit Sub
       .MoveFirst
       Do While Not .EOF
            Me.cboName.AddItem !desccurrency
            .MoveNext
       Loop
    End With
    
    Set rsLIST = Nothing
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub cboName_LostFocus()
On Error GoTo err
    Dim rsNAME As ADODB.Recordset
    Set rsNAME = New ADODB.Recordset
    
    rsNAME.Open "SELECT * FROM ParamCurrencies WHERE Desccurrency = '" & Me.cboName.Text & "';", cnALIS, adOpenKeyset, adLockOptimistic
    
    With rsNAME
    If .EOF And .BOF Then Exit Sub
        Me.txtSymbol = !Currency & ""
        Me.txtRate = !exchrate & ""
        Me.cmdChange.SetFocus: Me.cmdChange.Default = True
    End With
    Call FormatFigures
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub cboPayMEthod_Click()
If Not NewRecord And Not EditRecord Then Exit Sub
    Me.cmdPayChange.SetFocus: Me.cmdPayChange.Default = True
End Sub

Private Sub cboPayMEthod_GotFocus()
On Error GoTo err
If Not NewRecord And Not EditRecord Then Exit Sub
    If Me.cboPayMEthod.ListCount <> 0 Then Me.cboPayMEthod.Refresh: Exit Sub
    
    Dim rsLIST As ADODB.Recordset
    Set rsLIST = cnALIS.Execute("SELECT * FROM ParamPayMethods ORDER BY CodeNumber;")
    Me.cboPayMEthod.Clear
    
    With rsLIST
    If .EOF And .BOF Then Exit Sub
       .MoveFirst
       Do While Not .EOF
            Me.cboPayMEthod.AddItem !PayMethod
            .MoveNext
       Loop
    End With
    
    Set rsLIST = Nothing
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub cboPayMEthod_LostFocus()
On Error GoTo err
    Dim rsFindRecord As ADODB.Recordset
    Set rsFindRecord = New ADODB.Recordset
    
    rsFindRecord.Open "SELECT * FROM ParamPayMethods WHERE paymethod = '" & Me.cboPayMEthod.Text & "';", cnALIS, adOpenKeyset, adLockOptimistic
    
    With rsFindRecord
    If .EOF And .BOF Then Exit Sub
        Me.txtPayCode = !codenumber & ""
        Me.cmdPayChange.SetFocus: Me.cmdPayChange.Default = True
    End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub cmdAreaChange_Click()
On Error GoTo err
    Select Case cmdAreaChange.Caption
    Case "&Change"
        If MsgBox("Change the Local/Area Telecoms Code Settings?", vbQuestion + vbYesNo + vbDefaultButton2, "Change Defaults") = vbNo Then Exit Sub
        If NewRecord Then EditRecord = False: NewRecord = True: GoTo Proceed
        EditRecord = True
        Editor = True
Proceed:
        Me.cboAreaName.SetFocus: Me.txtAreaCode.Text = Empty
        cmdAreaChange.Caption = "Save &Changes"
    Case "Save &Changes"
    If EditRecord Then
    If ValidLocal Then
        Dim rsCHANGE As ADODB.Recordset
        Set rsCHANGE = New ADODB.Recordset
        
        rsCHANGE.Open "SELECT * FROM setareacode WHERE codenumber='" & Me.txtAreaCodeNumber.Text & "';", cnALIS, adOpenKeyset, adLockOptimistic
        
        With rsCHANGE
        If .EOF And .BOF Then Exit Sub
            !areaname = Me.cboAreaName.Text
            !AreaCode = Me.txtAreaCode.Text
            
            .Update
            .Requery
        End With
        
        EditRecord = False
        Editor = False
        cmdAreaChange.Caption = "&Change"
    End If
    End If
    
    If NewRecord Then
    If ValidLocal Then
        
        With rsLOCAL
        If .EOF And .BOF Then Exit Sub
            .AddNew
            !codenumber = "CN01"
            !areaname = Me.cboAreaName.Text
            !AreaCode = Me.txtAreaCode.Text
            
            .Update
            .Requery
        End With
        
        NewRecord = False
        Editor = False
        cmdAreaChange.Caption = "&Change"
    End If
    End If
    Case Else
        Exit Sub
    End Select
Exit Sub
err:
rsCHANGE.CancelUpdate
rsCHANGE.Requery
rsLOCAL.CancelUpdate
rsLOCAL.Requery

    ErrorMessage
End Sub

Private Sub cmdChange_Click()
On Error GoTo err
    Select Case cmdChange.Caption
    Case "&Change"
        If MsgBox("Change the Default Curency Settings?", vbQuestion + vbYesNo + vbDefaultButton2, "Change Defaults") = vbNo Then Exit Sub
        If NewRecord Then EditRecord = False: NewRecord = True: GoTo Proceed
        EditRecord = True
        Editor = True
Proceed:
        Me.cboName.SetFocus: Me.txtRate.Text = Empty: Me.txtSymbol.Text = Empty
        cmdChange.Caption = "Save &Changes"
    Case "Save &Changes"
    If EditRecord Then
    If ValidDefault Then
        Dim rsCHANGE As ADODB.Recordset
        Set rsCHANGE = New ADODB.Recordset
        
        rsCHANGE.Open "SELECT * FROM SetDefaultCurrency WHERE codenumber='" & Me.txtCode.Text & "';", cnALIS, adOpenKeyset, adLockOptimistic
        
        With rsCHANGE
        If .EOF And .BOF Then Exit Sub
            !currencyname = Me.cboName.Text
            !Currency = Me.txtSymbol.Text
            !exchrate = Me.txtRate.Text
            
            .Update
            .Requery
        End With
        
        EditRecord = False
        Editor = False
        cmdChange.Caption = "&Change"
    End If
    End If
    
    If NewRecord Then
    If ValidDefault Then
        
        With rsDEFAULT
        If .EOF And .BOF Then Exit Sub
            .AddNew
            !codenumber = "CD001"
            !currencyname = Me.cboName.Text
            !Currency = Me.txtSymbol.Text
            !exchrate = Me.txtRate.Text
            
            .Update
            .Requery
        End With
        
        EditRecord = False
        Editor = False
        cmdChange.Caption = "&Change"
    End If
    End If
    Case Else
        Exit Sub
    End Select
Exit Sub
err:
rsCHANGE.CancelUpdate
rsCHANGE.Requery
rsDEFAULT.CancelUpdate
rsDEFAULT.Requery
    ErrorMessage
End Sub

Private Function ValidDefault() As Boolean
On Error GoTo err
    Dim strMessage As String
    With Me
        If .cboName.Text = Empty Then
            strMessage = "Required Currency!"
            Me.cboName.SetFocus
        ElseIf .txtCode.Text = Empty Then
            strMessage = "Required Code!"
            Me.txtCode.SetFocus
            Me.txtCode.Locked = False
        ElseIf .txtRate.Text = Empty Then
            strMessage = "Required Exchange Rate!"
            Me.cboName.SetFocus
        ElseIf .txtSymbol = Empty Then
            strMessage = "Required Currency Symbol!"
            Me.cboName.SetFocus
        Else
            ValidDefault = True
        End If
        If Not ValidDefault Then
            MsgBox strMessage, vbCritical + vbOKOnly, "Data Validation"
        End If
    End With
Exit Function
err:
    ErrorMessage
End Function

Private Sub cmdCountryChange_Click()
On Error GoTo err
    Select Case cmdCountryChange.Caption
    Case "&Change"
        If MsgBox("Change the Country's Telecoms Code Settings?", vbQuestion + vbYesNo + vbDefaultButton2, "Change Defaults") = vbNo Then Exit Sub
        If NewRecord Then EditRecord = False: NewRecord = True: GoTo Proceed
        EditRecord = True
        Editor = True
Proceed:
        Me.cboCountryName.SetFocus: Me.txtCountryCode.Text = Empty
        cmdCountryChange.Caption = "Save &Changes"
    Case "Save &Changes"
    If EditRecord Then
    If ValidCountry Then
        Dim rsCHANGE As ADODB.Recordset
        Set rsCHANGE = New ADODB.Recordset
        
        rsCHANGE.Open "SELECT * FROM Setcountrycode WHERE codenumber='" & Me.txtCountryCNumber.Text & "';", cnALIS, adOpenKeyset, adLockOptimistic
        
        With rsCHANGE
        If .EOF And .BOF Then Exit Sub
            !countryname = Me.cboCountryName.Text
            !CountryCode = Me.txtCountryCode.Text
                  
            .Update
            .Requery
        End With
        
        EditRecord = False
        Editor = False
        cmdCountryChange.Caption = "&Change"
    End If
    End If
    
    If NewRecord Then
    If ValidLocal Then
        
        With rsLOCAL
        If .EOF And .BOF Then Exit Sub
            .AddNew
            !codenumber = "CN01"
            !countryname = Me.cboCountryName.Text
            !CountryCode = Me.txtCountryCode.Text
            
            .Update
            .Requery
        End With
        
        NewRecord = False
        Editor = False
        cmdCountryChange.Caption = "&Change"
    End If
    End If
    Case Else
        Exit Sub
    End Select
Exit Sub
err:
rsCHANGE.CancelUpdate
rsCHANGE.Requery
rsLOCAL.CancelUpdate
rsLOCAL.Requery
    ErrorMessage
End Sub

Private Sub cmdLChange_Click()
On Error GoTo err
    Select Case cmdLChange.Caption
    Case "&Change"
        If MsgBox("Change the Default Curency Settings?", vbQuestion + vbYesNo + vbDefaultButton2, "Change Defaults") = vbNo Then Exit Sub
        If NewRecord Then EditRecord = False: NewRecord = True: GoTo Proceed
        EditRecord = True
        Editor = True
Proceed:
        Me.cboLName.SetFocus: Me.txtLSymbol.Text = Empty
        cmdLChange.Caption = "Save &Changes"
    Case "Save &Changes"
    If EditRecord Then
    If ValidLocal Then
        Dim rsCHANGE As ADODB.Recordset
        Set rsCHANGE = New ADODB.Recordset
        
        rsCHANGE.Open "SELECT * FROM SetLocalCurrency WHERE codenumber='" & Me.txtLCode.Text & "';", cnALIS, adOpenKeyset, adLockOptimistic
        
        With rsCHANGE
        If .EOF And .BOF Then Exit Sub
            !currencyname = Me.cboLName.Text
            !Currency = Me.txtLSymbol.Text
                  
            .Update
            .Requery
        End With
        
        EditRecord = False
        Editor = False
        cmdLChange.Caption = "&Change"
    End If
    End If
    
    If NewRecord Then
    If ValidLocal Then
        
        With rsLOCAL
        If .EOF And .BOF Then Exit Sub
            .AddNew
            !codenumber = "CD001"
            !currencyname = Me.cboLName.Text
            !Currency = Me.txtLSymbol.Text
            
            .Update
            .Requery
        End With
        
        NewRecord = False
        Editor = False
        cmdLChange.Caption = "&Change"
    End If
    End If
    Case Else
        Exit Sub
    End Select
Exit Sub
err:
rsCHANGE.CancelUpdate
rsCHANGE.Requery
rsLOCAL.CancelUpdate
rsLOCAL.Requery
    ErrorMessage
End Sub

Private Function ValidArea() As Boolean
On Error GoTo err
Dim strMessage As String
    With Me
        If .txtAreaCode.Text = Empty Then
            strMessage = "Required Area Code!"
            .txtAreaCode.SetFocus
        ElseIf .txtAreaCodeNumber = Empty Then
            strMessage = "Required Code Number!"
            .txtAreaCodeNumber.SetFocus
            .txtAreaCodeNumber.Locked = False
        ElseIf .cboAreaName.Text = Empty Then
            strMessage = "Required Area Name!"
            .cboAreaName.SetFocus
        Else
            ValidArea = True
        End If
        If Not ValidArea Then
            MsgBox strMessage, vbCritical + vbOKOnly, "Data Validation"
        End If
    End With
Exit Function
err:
    ErrorMessage
End Function


Private Function ValidCountry() As Boolean
On Error GoTo err
Dim strMessage As String
    With Me
        If .txtCountryCNumber.Text = Empty Then
            strMessage = "Required Code Number!"
            .txtCountryCNumber.SetFocus
            .txtCountryCNumber.Locked = False
        ElseIf .txtCountryCode = Empty Then
            strMessage = "Required Country Telecoms Code!"
            .txtCountryCode.SetFocus
        ElseIf .cboCountryName.Text = Empty Then
            strMessage = "Required Country Name!"
            .cboCountryName.SetFocus
        Else
            ValidCountry = True
        End If
        If Not ValidCountry Then
            MsgBox strMessage, vbCritical + vbOKOnly, "Data Validation"
        End If
    End With
Exit Function
err:
    ErrorMessage
End Function

Private Function ValidLocal() As Boolean
On Error GoTo err
Dim strMessage As String
    With Me
        If .txtLCode.Text = Empty Then
            strMessage = "Required Currency Code!"
            .txtLCode.SetFocus
            .txtLCode.Locked = False
        ElseIf .txtLSymbol = Empty Then
            strMessage = "Required Currency Symbol!"
            .cboLName.SetFocus
        ElseIf .cboLName.Text = Empty Then
            strMessage = "Required Currency Name!"
            .cboLName.SetFocus
        Else
            ValidLocal = True
        End If
        If Not ValidLocal Then
            MsgBox strMessage, vbCritical + vbOKOnly, "Data Validation"
        End If
    End With
Exit Function
err:
    ErrorMessage
End Function

Private Sub cmdPayChange_Click()
On Error GoTo err
    Select Case cmdPayChange.Caption
    Case "&Change"
        If MsgBox("Change the Default Payment Method Settings?", vbQuestion + vbYesNo + vbDefaultButton2, "Change Defaults") = vbNo Then Exit Sub
        If NewRecord Then EditRecord = False: NewRecord = True: GoTo Proceed
        EditRecord = True
        Editor = True
Proceed:
        Me.cboPayMEthod.SetFocus: Me.txtPayCode.Text = Empty
        cmdPayChange.Caption = "Save &Changes"
    Case "Save &Changes"
    If EditRecord Then
    If ValidLocal Then
        Dim rsCHANGE As ADODB.Recordset
        Set rsCHANGE = New ADODB.Recordset
        
        rsCHANGE.Open "SELECT * FROM SetDefaultPayMEthod WHERE serialno='" & Me.txtPaySerialNo.Text & "';", cnALIS, adOpenKeyset, adLockOptimistic
        
        With rsCHANGE
        If .EOF And .BOF Then Exit Sub
            !PayMethod = Me.cboPayMEthod.Text
            !paycode = Me.txtPayCode.Text
            
            .Update
            .Requery
        End With
        
        EditRecord = False
        Editor = False
        cmdPayChange.Caption = "&Change"
    End If
    End If
    
    If NewRecord Then
    If ValidLocal Then
        
        With rsLOCAL
        If .EOF And .BOF Then Exit Sub
            .AddNew
            !SerialNo = "CN01"
            !PayMethod = Me.cboPayMEthod.Text
            !paycode = Me.txtPayCode.Text
            
            .Update
            .Requery
        End With
        
        NewRecord = False
        Editor = False
        cmdPayChange.Caption = "&Change"
    End If
    End If
    Case Else
        Exit Sub
    End Select
Exit Sub
err:
rsCHANGE.CancelUpdate
rsCHANGE.Requery
rsLOCAL.CancelUpdate
rsLOCAL.Requery
    ErrorMessage
End Sub

Private Sub cmdVChange_Click()
On Error GoTo err
    Select Case cmdVChange.Caption
    Case "&Change"
        If MsgBox("Change the Default Curency Settings?", vbQuestion + vbYesNo + vbDefaultButton2, "Change Defaults") = vbNo Then Exit Sub
        If NewRecord Then EditRecord = False: NewRecord = True: GoTo Proceed
        EditRecord = True
        Editor = True
Proceed:
        Me.txtVVat.SetFocus: Me.txtVVat.SelStart = 0: Me.txtVVat.SelLength = Len(Me.txtVVat.Text)
        cmdVChange.Caption = "Save &Changes"
    Case "Save &Changes"
    If EditRecord Then
    If ValidVat Then
        Dim rsCHANGE As ADODB.Recordset
        Set rsCHANGE = New ADODB.Recordset
        
        rsCHANGE.Open "SELECT * FROM ParamVATRate WHERE codenumber='" & Me.txtVCode.Text & "';", cnALIS, adOpenKeyset, adLockOptimistic
        
        With rsCHANGE
        If .EOF And .BOF Then Exit Sub
            !codenumber = Me.txtVCode.Text
            !VATRate = Me.txtVVat.Text
                  
            .Update
            .Requery
        End With
        
        EditRecord = False
        Editor = False
        cmdVChange.Caption = "&Change"
    End If
    End If
    
    If NewRecord Then
    If ValidLocal Then
        
        With rsVAT
        If .EOF And .BOF Then Exit Sub
            .AddNew
            !codenumber = "CD001"
            !VATRate = Me.txtVVat.Text
            
            .Update
            .Requery
        End With
        
        NewRecord = False
        Editor = False
        cmdVChange.Caption = "&Change"
    End If
    End If
    Case Else
        Exit Sub
    End Select
Exit Sub
err:
rsCHANGE.CancelUpdate
rsCHANGE.Requery
rsVAT.CancelUpdate
rsVAT.Requery
    
    ErrorMessage
End Sub

Private Function ValidVat() As Boolean
On Error GoTo err
Dim strMessage As String
    With Me
        If .txtVCode.Text = Empty Then
            strMessage = "Required Code!"
            .txtVCode.SetFocus: .txtVCode.Locked = False
        ElseIf .txtVVat.Text = Empty Then
            strMessage = "Required VAT Rate!"
            .txtVVat.SetFocus
        Else
            ValidVat = True
        End If
        If Not ValidVat Then
            MsgBox strMessage, vbCritical + vbOKOnly, "Data Validation"
        End If
    End With
Exit Function
err:
ErrorMessage
End Function

Private Sub Form_Activate()
With ALISSysManager
    Call LoadMainSettings
End With
End Sub

Private Sub Form_Load()
    Call OpenConnection
End Sub

Private Sub LoadMainSettings()
On Error GoTo err
StartOpen:
    Set rsDEFAULT = New ADODB.Recordset
    Set rsLOCAL = New ADODB.Recordset
    Set rsVAT = New ADODB.Recordset
    Set rsCOUNTRY = New ADODB.Recordset
    Set rsAREA = New ADODB.Recordset
    Set rsPAYMETHOD = New ADODB.Recordset
    
loadDEFAULTS:
    rsDEFAULT.Open "SELECT * FROM SetDefaultCurrency;", cnALIS, adOpenKeyset, adLockOptimistic
    If rsDEFAULT.EOF And rsDEFAULT.BOF Then NewRecord = True: EditRecord = False: GoTo LocalCurr
    LoadDefault
    
LocalCurr:
    rsLOCAL.Open "SELECT * FROM SetLocalCurrency;", cnALIS, adOpenKeyset, adLockOptimistic
    If rsLOCAL.EOF And rsLOCAL.BOF Then NewRecord = True: EditRecord = False: GoTo VATRate
    LoadLocal
    
VATRate:
    rsVAT.Open "SELECT * FROM ParamVATRate;", cnALIS, adOpenKeyset, adLockOptimistic
    If rsVAT.EOF And rsVAT.BOF Then NewRecord = True: EditRecord = False: GoTo OUTS
    LoadVAT
    
CountryCode:
    rsCOUNTRY.Open "SELECT * FROM setcountrycode;", cnALIS, adOpenKeyset, adLockOptimistic
    If rsCOUNTRY.EOF And rsVAT.BOF Then NewRecord = True: EditRecord = False: GoTo OUTS
    LoadCountry
    
AreaCode:
    rsAREA.Open "SELECT * FROM setareacode;", cnALIS, adOpenKeyset, adLockOptimistic
    If rsAREA.EOF And rsAREA.BOF Then NewRecord = True: EditRecord = False: GoTo OUTS
    LoadArea
    
PayMethod:
    rsPAYMETHOD.Open "SELECT * FROM SetDefaultPayMethod;", cnALIS, adOpenKeyset, adLockOptimistic
    If rsPAYMETHOD.EOF And rsPAYMETHOD.BOF Then NewRecord = True: EditRecord = False: GoTo OUTS
    LoadMethods

OUTS:
    Me.cboName.SetFocus
    
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub LoadMethods()
With rsPAYMETHOD
If .EOF And .BOF Then Exit Sub
    Me.txtPaySerialNo = !SerialNo & ""
    Me.txtPayCode = !paycode & ""
    Me.cboPayMEthod = !PayMethod & ""
End With
End Sub

Private Sub RecordMethods()
With rsPAYMETHOD
   !SerialNo = Me.txtPaySerialNo
   !paycode = Me.txtPayCode
   !PayMethod = Me.cboPayMEthod
End With
End Sub

Private Sub ClearMethods()
    Me.txtPaySerialNo = ""
    Me.txtPayCode = ""
    Me.cboPayMEthod = ""
End Sub

Private Sub LoadArea()
With rsAREA
If .EOF And .BOF Then Exit Sub
    Me.cboAreaName.Text = !areaname & ""
    Me.txtAreaCode.Text = !AreaCode & ""
    Me.txtAreaCodeNumber.Text = !codenumber & ""
End With
End Sub

Private Sub RecordArea()
With rsAREA
   !areaname = Me.cboAreaName.Text
   !AreaCode = Me.txtAreaCode.Text
   !codenumber = Me.txtAreaCodeNumber.Text
End With
End Sub

Private Sub ClearArea()
    Me.cboAreaName.Text = ""
    Me.txtAreaCode.Text = ""
    Me.txtAreaCodeNumber.Text = ""
End Sub

Private Sub LoadCountry()
With rsCOUNTRY
If .EOF And .BOF Then Exit Sub
    Me.cboCountryName = !countryname & ""
    Me.txtCountryCNumber = !codenumber & ""
    Me.txtCountryCode = !CountryCode & ""
End With
End Sub

Private Sub RecordCountry()
With rsCOUNTRY
   !countryname = Me.cboCountryName
   !codenumber = Me.txtCountryCNumber
   !CountryCode = Me.txtCountryCode
End With
End Sub

Private Sub ClearCountry()
    Me.cboCountryName = ""
    Me.txtCountryCNumber = ""
    Me.txtCountryCode = ""
End Sub

Private Sub LoadDefault()
With rsDEFAULT
If .EOF And .BOF Then Exit Sub
    Me.txtCode = !codenumber & ""
    Me.txtRate = !exchrate & ""
    Me.txtSymbol = !Currency & ""
    Me.cboName = !currencyname & ""
    Call FormatFigures
End With
End Sub

Private Sub FormatFigures()
On Error GoTo err
    Me.txtRate.Text = FormatNumber(Me.txtRate.Text, 5, vbUseDefault, vbUseDefault, vbTrue)
    Me.txtVVat.Text = FormatNumber(Me.txtVVat.Text, 2, vbUseDefault, vbUseDefault, vbTrue)
Exit Sub
err:
If err.Number = 5 Or err.Number = 13 Then Resume Next
    ErrorMessage
End Sub

Private Sub RecordDefault()
With rsDEFAULT
   !codenumber = Me.txtCode
   !exchrate = Me.txtRate
   !Currency = Me.txtSymbol
   !currencyname = Me.cboName
End With
End Sub

Private Sub ClearDefault()
    Me.txtCode = ""
    Me.txtRate = ""
    Me.txtSymbol = ""
    Me.cboName = ""
End Sub

Private Sub LoadLocal()
With rsLOCAL
If .EOF And .BOF Then Exit Sub
    Me.txtLCode.Text = !codenumber & ""
    Me.cboLName.Text = !currencyname & ""
    Me.txtLSymbol.Text = !Currency & ""
End With
End Sub

Private Sub RecordLocal()
With rsLOCAL
   !codenumber = Me.txtLCode.Text
   !currencyname = Me.cboLName.Text
   !Currency = Me.txtLSymbol.Text
End With
End Sub

Private Sub ClearLocal()
    Me.txtLCode.Text = ""
    Me.cboLName.Text = ""
    Me.txtLSymbol.Text = ""
End Sub

Private Sub LoadVAT()
With rsVAT
If .EOF And .BOF Then Exit Sub
    Me.txtVCode.Text = !codenumber & ""
    Me.txtVVat.Text = !VATRate & ""
    Call FormatFigures
End With
End Sub

Private Sub RecordVAT()
With rsVAT
  !codenumber = Me.txtVCode.Text
  !VATRate = Me.txtVVat.Text
End With
End Sub

Private Sub ClearVAT()
    Me.txtVCode.Text = ""
    Me.txtVVat.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If NewRecord Or EditRecord Then MsgBox "Data Entry or Edit in Progress! No Work was Done!", vbInformation + vbOKOnly, "Screen Unload": Cancel = 1
End Sub

Private Sub txtVVat_KeyPress(KeyAscii As Integer)
If Not NewRecord And Not EditRecord Then
    KeyAscii = 0
Else
    Exit Sub
End If
End Sub

Private Sub txtVVat_LostFocus()
    Call FormatFigures
End Sub
