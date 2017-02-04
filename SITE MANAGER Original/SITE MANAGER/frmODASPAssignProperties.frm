VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmODASPAssignProperties 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Assign Properties"
   ClientHeight    =   6705
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   12825
   Icon            =   "frmODASPAssignProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   12825
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   6015
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   12855
      Begin VB.Frame Frame5 
         Caption         =   "List of ALL Properties"
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
         Height          =   3015
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   6975
         Begin MSComctlLib.ListView ListView1 
            Height          =   2655
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   6735
            _ExtentX        =   11880
            _ExtentY        =   4683
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
      Begin VB.Frame Frame4 
         Height          =   3015
         Left            =   7200
         TabIndex        =   4
         Top             =   120
         Width           =   5535
         Begin VB.TextBox txtMedia 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   3960
            TabIndex        =   22
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtPropertyOtherDetails 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   1800
            TabIndex        =   18
            Top             =   2160
            Width           =   3615
         End
         Begin MSComCtl2.DTPicker DTPickerPropertyCommencementDate 
            Height          =   315
            Left            =   5160
            TabIndex        =   17
            Top             =   1800
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   556
            _Version        =   393216
            Format          =   66387969
            CurrentDate     =   38300
         End
         Begin VB.TextBox txtPropertyCommencementDate 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1800
            TabIndex        =   13
            Top             =   1800
            Width           =   3375
         End
         Begin VB.TextBox txtPropertyAmountDue 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   1800
            TabIndex        =   11
            Top             =   1392
            Width           =   3615
         End
         Begin VB.TextBox txtPropertyDateAssigned 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1800
            TabIndex        =   9
            Top             =   1008
            Width           =   3375
         End
         Begin VB.TextBox txtPropertyCode 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1800
            TabIndex        =   6
            Top             =   624
            Width           =   3615
         End
         Begin VB.TextBox txtSiteNo 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1800
            TabIndex        =   5
            Top             =   240
            Width           =   1455
         End
         Begin MSComCtl2.DTPicker DTPickerPropertyDateAssigned 
            Height          =   315
            Left            =   5160
            TabIndex        =   20
            Top             =   1005
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   556
            _Version        =   393216
            Format          =   66387969
            CurrentDate     =   38300
         End
         Begin VB.Label Label8 
            Caption         =   "Media"
            Height          =   255
            Left            =   3360
            TabIndex        =   21
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label7 
            Caption         =   "Other Details"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   2175
            Width           =   1575
         End
         Begin VB.Label Label6 
            Caption         =   "Commencement Date"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   1815
            Width           =   1575
         End
         Begin VB.Label Label4 
            Caption         =   "Amount Due"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   1407
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Date Assigned"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   1023
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Property Code"
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   615
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Site No"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   255
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Properties Assigned"
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
         TabIndex        =   2
         Top             =   3120
         Width           =   12615
         Begin MSComctlLib.ListView ListView2 
            Height          =   2535
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   12375
            _ExtentX        =   21828
            _ExtentY        =   4471
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
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10560
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
            Picture         =   "frmODASPAssignProperties.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASPAssignProperties.frx":0ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASPAssignProperties.frx":0F0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASPAssignProperties.frx":1228
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASPAssignProperties.frx":18A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASPAssignProperties.frx":1F1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASPAssignProperties.frx":236E
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
      Width           =   12825
      _ExtentX        =   22622
      _ExtentY        =   1164
      ButtonWidth     =   3704
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
            Caption         =   "&Preview Properties"
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
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuClear 
         Caption         =   "Clear the &Screen"
         Shortcut        =   ^C
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
Attribute VB_Name = "frmODASPAssignProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsPROPERTIES As clsODASProperties
Dim rsSITE As clsODASSite, MyCommonData As clsCommonData

Private Sub cmdAddNew_Click()

End Sub

Private Sub cmdUpdate_Click()

End Sub

Private Sub cmdSearch_Click()

End Sub

Private Sub cmdEdit_Click()

End Sub

Private Sub cmdDelete_Click()

End Sub

Private Sub cmdPrint_Click()

End Sub

Private Sub DTPickerCommencementDate_CloseUp()
On Error GoTo err
    With frmODASPAssignProperties
        .txtPropertyCommencementDate.Text = .DTPickerPropertyCommencementDate
    End With

Exit Sub
err:
    ErrorMessage
End Sub


Private Sub DTPickerPropertyCommencementDate_CloseUp()
With Me
        .txtPropertyCommencementDate.Text = .DTPickerPropertyCommencementDate.Value
        .txtPropertyDateAssigned.Text = .DTPickerPropertyDateAssigned.Value
End With
End Sub

Private Sub DTPickerPropertyDateAssigned_CloseUp()
On Error GoTo err
    With frmODASPAssignProperties
        .txtPropertyDateAssigned.Text = .DTPickerPropertyDateAssigned
    End With

Exit Sub
err:
    ErrorMessage
End Sub

Private Sub Form_Activate()
        Set MyCommonData = New clsCommonData
        rsPROPERTIES.loadDEFAULTS
        disableALLRECORD
        showALLPROPERTIES
        showACTUALPROPERTIES
        disableFRAME
        
End Sub

Private Sub disableFRAME()
On Error GoTo err
        With frmODASPAssignProperties
            .Frame5.Enabled = False
            .Frame2.Enabled = False
        End With
Exit Sub

err:
    ErrorMessage
End Sub
Private Sub enableFRAME()
On Error GoTo err
        With frmODASPAssignProperties
            .Frame5.Enabled = True
            .Frame2.Enabled = True
        End With
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub Form_Initialize()
            Set rsPROPERTIES = New clsODASProperties
End Sub

Private Sub Form_Terminate()
            Set rsPROPERTIES = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo err

        If NewRecord Then
            Cancel = True
            MsgBox "System cannot close this form when property assignment...", vbCritical + vbOKOnly, "Interruption of Job Brief creation process"
        Else
            Cancel = False
            Set rsSEND = Nothing
        End If
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
Exit Sub
With Me
    frmODASPAssignProperties.txtPropertyCode.Text = Item.Text
    frmODASPAssignProperties.Toolbar1.Buttons(4).Caption = "Delete"
    rsPROPERTIES.loadRECORD
       

End With
err:
    ErrorMessage
End Sub
Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo err
    ListView2.SortKey = ColumnHeader.Index - 1
    ListView2.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo err

    frmODASPAssignProperties.txtPropertyCode.Text = Item.Text
    frmODASPAssignProperties.Toolbar1.Buttons(4).Caption = "Delete"
    rsPROPERTIES.loadRECORD
    
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub ListView3_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo err
    ListView3.SortKey = ColumnHeader.Index - 1
    ListView3.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ListView3_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo err
        Dim i, j As Double

        If Item.Checked = True Then
            
            j = Screen.ActiveForm.ListView3.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                If Screen.ActiveForm.ListView3.ListItems(i) <> Item Then
                   Screen.ActiveForm.ListView3.ListItems(i).Checked = False
                End If
            Next i
            
      Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub
Private Sub clearRECORD()
On Error GoTo err
    With Me
        .txtPropertyAmountDue.Text = 0
        .DTPickerPropertyCommencementDate.Value = Date
        .DTPickerPropertyDateAssigned.Value = Date
    End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
    With Me
        Select Case Button.Key
            Case "N"
                Select Case Button.Caption
                    Case "New &Record "
                            clearRECORD
                            enableALLRECORD
                            enableFRAME
                            NewRecord = True
                            Button.Caption = "&Save Record": Button.Image = 4
                    Case "&Save Record"
                            
                            j = .ListView1.ListItems.Count: k = 0
                            
                            If j = 0 Then Exit Sub
                            
                            For i = 1 To j
                                If .ListView1.ListItems(i).Checked = True Then
                                    k = k + 1
                                End If
                            Next i
                            
                            rsPROPERTIES.updateRECORD
                                strSQL = "select * from ODASPPlotMast Where MAstNo = '" & .txtSiteNo.Text & "';"
                                Set rsSAVE = New ADODB.Recordset
                                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                                    rsSAVE!PropertiesAssigned = "Y"
                                    rsSAVE.Update
                                    rsSAVE.Requery
                            NewRecord = False: Button.Caption = "New &Record ": Button.Image = 2
                End Select
            Case "E"
                Select Case Button.Caption
                    Case "&Edit/Change "
                         If NewRecord Then Exit Sub
                                If .txtSiteNo.Text = Empty Then
                                MsgBox "There is NO Current Record to Edit. Please Search and Display a Record First...!", vbCritical + vbOKOnly, ""
                               .txtSiteNo.SetFocus
                                Else
                                enableALLRECORD
                                enableFRAME
                               .txtSiteNo.Locked = True
                                Button.Caption = "Save &Changes ": Button.Image = 5
                                EditRecord = True
                            End If
                    Case "Save &Changes "
                            rsPROPERTIES.updateRECORD
                       EditRecord = False: Button.Caption = "&Edit/Change ": Button.Image = 6
                End Select
            Case "S"
                Select Case Button.Caption
                Case "&Search/Find "
                    INQUIRY = InputBox("Enter the Site number to search and display;", "Search value request")
                            .txtSiteNo.Text = INQUIRY
                            .txtPropertyDateAssigned.Text = ""
                            .txtPropertyAmountDue.Text = ""
                            .txtPropertyCommencementDate.Text = ""
                            .txtPropertyOtherDetails.Text = ""
                            showALLPROPERTIES
                            showACTUALPROPERTIES
                            disableFRAME
            
                Case "Delete"
                        j = .ListView2.ListItems.Count: k = 0
                        
                        If j = 0 Then Exit Sub
                        
                        For i = 1 To j
                            If .ListView2.ListItems(i).Checked = True Then
                                k = k + 1
                            End If
                        Next i
                    If k = 0 Then Exit Sub
                    
                    If MsgBox("This action will completely delete the " & k & " properties checked. Are you sure you want to continue?", vbYesNo + vbCritical, "Property Deletion") = vbNo Then Exit Sub
                        
                        j = .ListView2.ListItems.Count
                        For i = 1 To j
                            If .ListView2.ListItems(i).Checked = True Then
                            
                                .txtPropertyCode = Trim(.ListView2.ListItems(i).Text)
                                
                                DeleteSQL = "Select * FROM ODASMSiteProperties WHERE BillBoardNo = '" & Trim(.txtSiteNo.Text) & "' and PropertyCode = '" & .txtPropertyCode.Text & "';"
                                Set rsDeleteRecord = New ADODB.Recordset
                                rsDeleteRecord.Open DeleteSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                                
                                If rsDeleteRecord.EOF And rsDeleteRecord.BOF Then GoTo Continue
                                
                                rsDeleteRecord.Delete
                                Set rsDeleteRecord = Nothing
                            End If
Continue:
                        Next i
                           showALLPROPERTIES
                           showACTUALPROPERTIES

                    End Select
            Case "R"
                    If MsgBox(RefreshMessage, vbExclamation + vbOKCancel + vbDefaultButton2, "Refresh The Screen") = vbCancel Then Exit Sub
                    .Toolbar1.Buttons(2).Caption = "New &Record "
                    .Toolbar1.Buttons(2).Image = 2
                    .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                    .Toolbar1.Buttons(3).Image = 5
                    .Toolbar1.Buttons(4).Caption = "&Search/Find "
                    NewRecord = False: EditRecord = False: MyCommonData.ClearTheScreen
            Case "P"
                    CurrentRecord = Me.txtSiteNo
                    frmODASRBBProperties.Show vbModal
        End Select
    End With
Exit Sub
err:
ErrorMessage
End Sub

