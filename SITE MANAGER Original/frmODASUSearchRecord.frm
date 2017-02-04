VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmODASUSearchRecord 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6750
   Icon            =   "frmODASUSearchRecord.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   6750
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "ENTER THE SEARCH VALUE HERE THE PRESS THE ENTER KEY"
         Top             =   3240
         Width           =   6255
      End
      Begin MSComctlLib.ListView listSearch 
         Height          =   2895
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   5106
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Frame Frame2 
         Height          =   1335
         Left            =   3840
         TabIndex        =   2
         Top             =   3540
         Width           =   2535
         Begin VB.CommandButton cmdCancel 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            Cancel          =   -1  'True
            Caption         =   "Cancel"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   1320
            Picture         =   "frmODASUSearchRecord.frx":0ECA
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   200
            Width           =   1095
         End
         Begin VB.CommandButton cmdOk 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            Caption         =   "OK"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   120
            Picture         =   "frmODASUSearchRecord.frx":130C
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   200
            Width           =   1215
         End
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0FFFF&
         BorderWidth     =   3
         Height          =   4815
         Left            =   0
         Top             =   120
         Width           =   6495
      End
   End
End
Attribute VB_Name = "frmODASUSearchRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim validated As Boolean
Dim MultiSelect As Boolean

Private Sub cmdCancel_Click()
CurrentRecord = ""
Me.Hide
End Sub

Private Sub cmdOK_Click()

validated = False
For i = 1 To listSearch.ListItems.Count
        If listSearch.ListItems(i).Checked = True Then
                    CurrentRecord = listSearch.ListItems(i).Text
                    Set selectedListItem = listSearch.ListItems(i)
                    'Debug.Print selectedListItem.SubItems(2)
                    validated = True
                    Exit For
        End If
Next i
If validated = False Then
        MsgBox "Select a record before you proceed", vbExclamation
        Exit Sub
End If

Me.Hide
End Sub

Private Sub Form_Activate()
Me.txtSearch.SetFocus
End Sub

Private Sub Form_Load()
CurrentRecord = Empty
'load_form strSQL
End Sub

Public Function load_form(strSQL As String, Optional windowCaption = "", Optional viewMultiSelect As Boolean = False)
        Screen.MousePointer = vbHourglass
        If Trim(strSQL) = "" Then Exit Function
        If Trim(windowCaption) = "" Then
                    Me.Caption = "Select Record"
        Else
                    Me.Caption = windowCaption
        End If
        MultiSelect = viewMultiSelect
        FillList strSQL, listSearch
        Screen.MousePointer = vbDefault

End Function

Private Sub Form_Unload(Cancel As Integer)
For i = 1 To listSearch.ListItems.Count
        If listSearch.ListItems(i).Checked = True Then
                    validated = True
                    Exit Sub
        End If
Next i
CurrentRecord = ""

End Sub

Private Sub listSearch_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
SortListViewColumn listSearch, ColumnHeader
End Sub

Private Sub listSearch_ItemCheck(ByVal Item As MSComctlLib.ListItem)
If MultiSelect = True Then
Else
        checkOne Item, listSearch
End If
End Sub

Private Sub vmdCancel_Click()

End Sub

Private Sub txtSearch_Change()
'search_in_listview Me.listSearch, "" & Me.txtSearch & ""
'Me.listSearch.FindItem "%" & Me.txtSearch & "%"
End Sub

Private Sub txtSearch_GotFocus()
Me.txtSearch.SelStart = 0
Me.txtSearch.SelLength = Len(Me.txtSearch.Text)
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))

If KeyAscii = vbKeyReturn Then
        FillList strSEARCHSQL, listSearch, Trim(Me.txtSearch.Text)
       
End If
End Sub
