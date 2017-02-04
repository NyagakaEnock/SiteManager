VERSION 5.00
Begin VB.Form frmODASPIssueOrder 
   Caption         =   "Issue/Cancel Order"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   7485
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin VB.Frame Frame3 
         Caption         =   "LPO Entries"
         Height          =   2655
         Left            =   120
         TabIndex        =   21
         Top             =   2520
         Width           =   7095
      End
      Begin VB.Frame Frame2 
         Height          =   2415
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   7095
         Begin VB.TextBox txtOrderDescription 
            BackColor       =   &H00FFC0C0&
            Height          =   285
            Left            =   1320
            TabIndex        =   11
            Top             =   600
            Width           =   5535
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
            Left            =   4560
            TabIndex        =   9
            Top             =   240
            Width           =   2295
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
            Width           =   2055
         End
         Begin VB.TextBox txtRemarks 
            BackColor       =   &H00FFC0C0&
            Height          =   285
            Left            =   1320
            MultiLine       =   -1  'True
            TabIndex        =   6
            Top             =   1320
            Width           =   5535
         End
         Begin VB.TextBox txtSupplierName 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   2520
            TabIndex        =   5
            Top             =   960
            Width           =   4335
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
            Width           =   2055
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
            Left            =   3360
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
            Left            =   3360
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
            Left            =   3360
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
End
Attribute VB_Name = "frmODASPIssueOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
