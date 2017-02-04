VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmUSelectDateRange 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Date Range"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   10170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00C0C000&
      Caption         =   "&Print Preview"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   960
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      Begin MSComCtl2.DTPicker DTPickerLastDate 
         Height          =   330
         Left            =   5040
         TabIndex        =   1
         Top             =   360
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   20709379
         CurrentDate     =   39357
      End
      Begin MSComCtl2.DTPicker DTPickerStartDate 
         Height          =   330
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   20709379
         CurrentDate     =   39357
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Start Date"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   600
         TabIndex        =   4
         Top             =   405
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "End Date"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4200
         TabIndex        =   3
         Top             =   405
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmUSelectDateRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Report As String

Private Sub cmdSearch_Click()
If Report = "" Then
        MsgBox "Select Report", vbExclamation
        Exit Sub
ElseIf (Me.DTPickerStartDate.Value > Me.DTPickerLastDate.Value) And Me.DTPickerStartDate.Visible = True Then
        MsgBox "Start Date Cannot be Greater than Last Date", vbExclamation
End If
Select Case Report
    Case "SitesBasedOnDate"
            frmRptODASAllSitesBasedOnDate.strStartDate = Me.DTPickerStartDate.Value
            frmRptODASAllSitesBasedOnDate.strLastDate = Me.DTPickerLastDate.Value
            
            Load frmRptODASAllSitesBasedOnDate
            frmRptODASAllSitesBasedOnDate.Show 1, Me
    Case "ExpiringSitesWithinADateRange"
            
    Case "ExpiringSitesAsAtASingleDate"
            
End Select
End Sub

Private Sub Form_Load()
Me.DTPickerLastDate.Value = Date
Me.DTPickerStartDate.Value = Date
End Sub
