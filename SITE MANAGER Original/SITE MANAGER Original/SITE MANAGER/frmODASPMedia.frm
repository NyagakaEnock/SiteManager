VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmODASPMedia 
   Caption         =   "Media Codes"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   1170
   ClientWidth     =   9705
   Icon            =   "frmODASPMedia.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmODASPMedia.frx":0442
   ScaleHeight     =   6390
   ScaleWidth      =   9705
   Begin VB.Frame Frame12 
      Height          =   6255
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   9495
      Begin VB.Frame Frame4 
         Height          =   1215
         Left            =   120
         TabIndex        =   35
         Top             =   120
         Width           =   7215
         Begin VB.CheckBox chkLocDepOnSize 
            Caption         =   "Loc Dep On Size"
            Height          =   435
            Left            =   5640
            TabIndex        =   43
            Top             =   660
            Width           =   1455
         End
         Begin VB.TextBox txtMediaDescription 
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2640
            TabIndex        =   40
            Top             =   240
            Width           =   4455
         End
         Begin VB.TextBox txtMediaCode 
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1800
            TabIndex        =   39
            Top             =   240
            Width           =   855
         End
         Begin VB.CheckBox chkInventory 
            Caption         =   "Inventory Item?"
            Height          =   255
            Left            =   2760
            TabIndex        =   38
            Top             =   750
            Width           =   1695
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "Active?"
            Height          =   195
            Left            =   4560
            TabIndex        =   37
            Top             =   780
            Width           =   1215
         End
         Begin VB.TextBox txtFaces 
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1800
            TabIndex        =   36
            Top             =   720
            Width           =   855
         End
         Begin VB.Label lblRelationshipCode 
            Caption         =   "Media Code"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   315
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Maximum No of Faces"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   795
            Width           =   1695
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1215
         Left            =   7440
         TabIndex        =   31
         Top             =   120
         Width           =   1935
         Begin VB.OptionButton OptRequireBillBoard 
            Caption         =   "Require Bill Board?"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   457
            Width           =   1695
         End
         Begin VB.OptionButton optRequireSite 
            Caption         =   "Require Site?"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   120
            Width           =   1695
         End
         Begin VB.OptionButton optRequireNothing 
            Caption         =   "None?"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   795
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Types of Media"
         Height          =   1815
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   9255
         Begin VB.OptionButton optMetalSheet 
            Caption         =   "Metal Sheet?"
            Height          =   255
            Left            =   7440
            TabIndex        =   30
            Top             =   1040
            Width           =   1575
         End
         Begin VB.OptionButton optStreetSign 
            Caption         =   "Street Sign?"
            Height          =   255
            Left            =   3720
            TabIndex        =   29
            Top             =   1040
            Width           =   1575
         End
         Begin VB.OptionButton optBusShelter 
            Caption         =   "Bus Shelter?"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   1440
            Width           =   1575
         End
         Begin VB.OptionButton optBridge 
            Caption         =   "Bridge?"
            Height          =   255
            Left            =   1680
            TabIndex        =   27
            Top             =   1440
            Width           =   1575
         End
         Begin VB.OptionButton optMobileSign 
            Caption         =   "Mobile Sign?"
            Height          =   255
            Left            =   5640
            TabIndex        =   26
            Top             =   1040
            Width           =   1575
         End
         Begin VB.OptionButton optWallPainting 
            Caption         =   "Wall Painting?"
            Height          =   255
            Left            =   5640
            TabIndex        =   25
            Top             =   1440
            Width           =   1575
         End
         Begin VB.OptionButton optPoster 
            Caption         =   "Poster?"
            Height          =   255
            Left            =   7440
            TabIndex        =   24
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton optSignBoard 
            Caption         =   "Sign Board?"
            Height          =   255
            Left            =   7440
            TabIndex        =   23
            Top             =   640
            Width           =   1575
         End
         Begin VB.OptionButton optFleetGraphics 
            Caption         =   "Fleet Graphics?"
            Height          =   255
            Left            =   1680
            TabIndex        =   22
            Top             =   640
            Width           =   1575
         End
         Begin VB.OptionButton optWindowGraphics 
            Caption         =   "Window Graphics?"
            Height          =   255
            Left            =   1680
            TabIndex        =   21
            Top             =   1040
            Width           =   1695
         End
         Begin VB.OptionButton optRailwaySign 
            Caption         =   "Railway Sign?"
            Height          =   255
            Left            =   3720
            TabIndex        =   20
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton optFlexiSign 
            Caption         =   "Flex-Sign?"
            Height          =   255
            Left            =   3720
            TabIndex        =   19
            Top             =   640
            Width           =   1575
         End
         Begin VB.OptionButton optTrolleySign 
            Caption         =   "Trolley Sign?"
            Height          =   255
            Left            =   3720
            TabIndex        =   18
            Top             =   1440
            Width           =   1575
         End
         Begin VB.OptionButton optPrismaSign 
            Caption         =   "Prisma Sign?"
            Height          =   255
            Left            =   5640
            TabIndex        =   17
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton optMileageSign 
            Caption         =   "Mileage Sign?"
            Height          =   255
            Left            =   5640
            TabIndex        =   16
            Top             =   640
            Width           =   1575
         End
         Begin VB.OptionButton optFloorGraphics 
            Caption         =   "Floor Graphics?"
            Height          =   255
            Left            =   1680
            TabIndex        =   15
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton optBanner 
            Caption         =   "Banner?"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   1040
            Width           =   1575
         End
         Begin VB.OptionButton optBackLit 
            Caption         =   "BackLit?"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   640
            Width           =   1575
         End
         Begin VB.OptionButton optBillBoard 
            Caption         =   "BillBoard?"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "List of All Media"
         Height          =   3015
         Left            =   120
         TabIndex        =   9
         Top             =   3120
         Width           =   8055
         Begin MSComctlLib.ListView ListView1 
            Height          =   2655
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   7815
            _ExtentX        =   13785
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
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   8280
         TabIndex        =   3
         Top             =   3120
         Width           =   1095
         Begin VB.CommandButton cmdAddNew 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Picture         =   "frmODASPMedia.frx":0784
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmdUpdate 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Picture         =   "frmODASPMedia.frx":0886
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   990
            Width           =   855
         End
         Begin VB.CommandButton cmdSearch 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Picture         =   "frmODASPMedia.frx":0988
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   1365
            Width           =   855
         End
         Begin VB.CommandButton cmdDelete 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Picture         =   "frmODASPMedia.frx":0A8A
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   1740
            Width           =   855
         End
         Begin VB.CommandButton cmdCancel 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Picture         =   "frmODASPMedia.frx":0B8C
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   2115
            Width           =   855
         End
         Begin VB.CommandButton cmdPrint 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Picture         =   "frmODASPMedia.frx":0C8E
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   2520
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "frmODASPMedia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub loadRECORD()
'On Error GoTo err
    
    Set rsCONTROL = New ADODB.Recordset
    
    strSQL = "Select * from ODASPMedia Where MediaCode = '" & frmODASPMedia.txtMediaCode.Text & "'"
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

    
    With rsCONTROL
        
        If .EOF Or .EOF Then Exit Sub
        
        frmODASPMedia.txtMediaCode = !MediaCode
        frmODASPMedia.txtMediaDescription.Text = !MediaDescription
        frmODASPMedia.txtFaces.Text = !Faces & ""
        
        If !LocDepOnSize = "Y" Then
                frmODASPMedia.chkLocDepOnSize.Value = 1
        Else: frmODASPMedia.chkLocDepOnSize.Value = 0
        End If

        If !InventoryItem = "Y" Then
                frmODASPMedia.chkInventory.Value = 1
        Else: frmODASPMedia.chkInventory.Value = 0
        End If
                    
        If !Status = "A" Then
                frmODASPMedia.chkStatus.Value = 1
        Else: frmODASPMedia.chkStatus.Value = 0
        End If
        
        If !BusShelter = "Y" Then
                frmODASPMedia.optBusShelter.Value = True
        Else: frmODASPMedia.optBusShelter.Value = False
        End If
        
        If !BillBoard = "Y" Then
                frmODASPMedia.optBillBoard.Value = True
        Else: frmODASPMedia.optBillBoard.Value = False
        End If

        If !BackLit = "Y" Then
                frmODASPMedia.optBackLit.Value = True
        Else: frmODASPMedia.optBackLit.Value = False
        End If
        
        If !Banner = "Y" Then
                frmODASPMedia.optBanner.Value = True
        Else: frmODASPMedia.optBanner.Value = False
        End If

        If !FloorGraphics = "Y" Then
                frmODASPMedia.optFloorGraphics.Value = True
        Else: frmODASPMedia.optFloorGraphics.Value = False
        End If
        
        If !WindowGraphics = "Y" Then
                frmODASPMedia.optWindowGraphics.Value = True
        Else: frmODASPMedia.optWindowGraphics.Value = False
        End If

        If !Bridge = "Y" Then
                frmODASPMedia.optBridge.Value = True
        Else: frmODASPMedia.optBridge.Value = False
        End If
        
        If !RailwaySign = "Y" Then
                frmODASPMedia.optRailwaySign.Value = True
        Else: frmODASPMedia.optRailwaySign.Value = False
        End If

        If !FlexiSign = "Y" Then
                frmODASPMedia.optFlexiSign.Value = True
        Else: frmODASPMedia.optFlexiSign.Value = False
        End If
        
        If !StreetSign = "Y" Then
                frmODASPMedia.optStreetSign.Value = True
        Else: frmODASPMedia.optStreetSign.Value = False
        End If
        
        If !TrolleySign = "Y" Then
                frmODASPMedia.optTrolleySign.Value = True
        Else: frmODASPMedia.optTrolleySign.Value = False
        End If
        
        If !PrismaSign = "Y" Then
                frmODASPMedia.optPrismaSign.Value = True
        Else: frmODASPMedia.optPrismaSign.Value = False
        End If
        
        If !MileageSign = "Y" Then
                frmODASPMedia.optMileageSign.Value = True
        Else: frmODASPMedia.optMileageSign.Value = False
        End If

        If !MobileSign = "Y" Then
                frmODASPMedia.optMobileSign.Value = True
        Else: frmODASPMedia.optMobileSign.Value = False
        End If
        
        If !WallPainting = "Y" Then
                frmODASPMedia.optWallPainting.Value = True
        Else: frmODASPMedia.optWallPainting.Value = False
        End If

        If !Poster = "Y" Then
                frmODASPMedia.optPoster.Value = True
        Else: frmODASPMedia.optPoster.Value = False
        End If
        
        If !SignBoard = "Y" Then
                frmODASPMedia.optSignBoard.Value = True
        Else: frmODASPMedia.optSignBoard.Value = False
        End If
        
        If !MetalSheet = "Y" Then
                frmODASPMedia.optMetalSheet.Value = True
        Else: frmODASPMedia.optMetalSheet.Value = False
        End If


    End With

Exit Sub

err:
    ErrorMessage
End Sub


Private Sub cmdAddNew_Click()
        baddRECORD = True
        clearALLRECORD
        enableALLRECORD
        disableButtons
End Sub


Private Sub cmdCancel_Click()
        enableButtons
        clearALLRECORD
        disableALLRECORD
        baddRECORD = False
End Sub


Private Sub cmdDelete_Click()
'On Error GoTo err

If txtMediaCode.Text = "" Then
            MsgBox "There is no current record to delete", vbInformation, "Delete Information"
Else
        If MsgBox("Are you sure you want to completely delete the current record?", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
            Set rsCONTROL = New ADODB.Recordset
    
            strSQL = "Select * from ODASPMediaCode Where MediaCode = '" & frmODASPMedia.txtMediaCode.Text & "'"
            rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

            
            With rsCONTROL
                
                If .EOF And .BOF Then Exit Sub
                .Delete
                .Requery
                clearALLRECORD
                getALLOPERATIONS
            End With
    End If
        '/* End if Msg Box
        
End If
        '/* If txt = ""
        
Exit Sub

err:
    ErrorMessage

End Sub

Private Sub cmdEdit_Click()
        editMYRECORD
End Sub

Private Sub ValidateRECORD()
'On Error GoTo err
        With frmODASPMedia
        
                bSaveRECORD = False
                
                If .txtMediaCode.Text = Empty Then
                        MsgBox "The Operation Type MUST be Entered"
                        .txtMediaCode.SetFocus
                
                ElseIf .txtMediaDescription.Text <= Empty Then
                        MsgBox "The Description of the Media cannot be Left Blank"
                        .txtMediaDescription.SetFocus
                
                ElseIf .txtFaces.Text <= Empty Then
                        MsgBox "The Maximum Number of Faces"
                        .txtFaces.SetFocus
                        
                ElseIf .OptRequireBillBoard.Value = False And .optRequireSite.Value = False And .optRequireNothing.Value = False Then
                        MsgBox "You Must Select One of the options"
                        .OptRequireBillBoard.SetFocus
                Else
                        bSaveRECORD = True
                End If

        End With
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub SaveRECORD()
''On Error GoTo err
    Set rsCONTROL = New ADODB.Recordset
    
    strSQL = "Select * from ODASPMedia Where MediaCode = '" & frmODASPMedia.txtMediaCode.Text & "'"
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
   With rsCONTROL
        If .BOF Or .EOF Then
                .AddNew
                !MediaCode = frmODASPMedia.txtMediaCode
                !PreparedBy = CurrentUserName
                !DatePrepared = Date
        End If
        
        !MediaDescription = frmODASPMedia.txtMediaDescription
        !Faces = CDbl(frmODASPMedia.txtFaces.Text)
        
        If frmODASPMedia.chkLocDepOnSize = 1 Then
                !LocDepOnSize = "Y"
            Else: !LocDepOnSize = "N"
        End If

        If frmODASPMedia.chkStatus = 1 Then
                !Status = "A"
            Else: !Status = "I"
        End If
        
        If frmODASPMedia.chkInventory = 1 Then
                !InventoryItem = "Y"
            Else: !InventoryItem = "N"
        End If
        
        If frmODASPMedia.OptRequireBillBoard = True Then
                !RequireBillBoard = "Y"
            Else: !RequireBillBoard = "N"
        End If
        
        If frmODASPMedia.optRequireNothing = True Then
                !RequireNothing = "Y"
            Else: !RequireNothing = "N"
        End If
        
        If frmODASPMedia.optRequireSite = True Then
                !RequireSite = "Y"
            Else: !RequireSite = "N"
        End If

        If frmODASPMedia.optBackLit = True Then
                !BackLit = "Y"
            Else: !BackLit = "N"
        End If
        
        If frmODASPMedia.optBanner = True Then
                !Banner = "Y"
            Else: !Banner = "N"
        End If
        
        If frmODASPMedia.optBillBoard = True Then
                !BillBoard = "Y"
            Else: !BillBoard = "N"
        End If
        
        If frmODASPMedia.optBridge = True Then
                !Bridge = "Y"
            Else: !Bridge = "N"
        End If
        
        If frmODASPMedia.optFleetGraphics = True Then
                !FleetGraphics = "Y"
            Else: !FleetGraphics = "N"
        End If
        
        If frmODASPMedia.optFlexiSign = True Then
                !FlexiSign = "Y"
            Else: !FlexiSign = "N"
        End If
        
        If frmODASPMedia.optFloorGraphics = True Then
                !FloorGraphics = "Y"
            Else: !FloorGraphics = "N"
        End If
        
        If frmODASPMedia.optWindowGraphics = True Then
                !WindowGraphics = "Y"
            Else: !WindowGraphics = "N"
        End If

        If frmODASPMedia.optMetalSheet = True Then
                !MetalSheet = "Y"
            Else: !MetalSheet = "N"
        End If
        
        If frmODASPMedia.optMileageSign = True Then
                !MileageSign = "Y"
            Else: !MileageSign = "N"
        End If
        
        If frmODASPMedia.optMobileSign = True Then
                !MobileSign = "Y"
            Else: !MobileSign = "N"
        End If
        
        If frmODASPMedia.optPoster = True Then
                !Poster = "Y"
            Else: !Poster = "N"
        End If
        
        If frmODASPMedia.optPrismaSign = True Then
                !PrismaSign = "Y"
            Else: !PrismaSign = "N"
        End If

        If frmODASPMedia.optRailwaySign = True Then
                !RailwaySign = "Y"
            Else: !RailwaySign = "N"
        End If
        
        If frmODASPMedia.optSignBoard = True Then
                !SignBoard = "Y"
            Else: !SignBoard = "N"
        End If
        
        If frmODASPMedia.optStreetSign = True Then
                !StreetSign = "Y"
            Else: !StreetSign = "N"
        End If
        
        If frmODASPMedia.optTrolleySign = True Then
                !TrolleySign = "Y"
            Else: !TrolleySign = "N"
        End If
        
        If frmODASPMedia.optWallPainting = True Then
                !WallPainting = "Y"
            Else: !WallPainting = "N"
        End If
        
        If frmODASPMedia.optBridge = True Then
                !Bridge = "Y"
            Else: !Bridge = "N"
        End If
        
        If frmODASPMedia.optBusShelter = True Then
                !BusShelter = "Y"
            Else: !BusShelter = "N"
        End If


        bSaveRECORD = False
        
         .Update
         .Requery
  End With
Exit Sub

err:
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            MsgBox "Update Cancelled! You cannot save this record because some required fields are blank or have invalid data!", vbCritical, "Cancel Update"
            rsCONTROL.CancelUpdate
            rsCONTROL.Requery
    Else
        UpdateErrorMessage
    End If

End Sub


Private Sub cmdUpdate_Click()
        bSaveRECORD = False
        ValidateRECORD
        If bSaveRECORD = True Then
            SaveRECORD
                If bSaveRECORD = False Then
                    enableButtons
                    disableALLRECORD
                    baddRECORD = False
                End If
        End If
        showALLMEDIA

        
End Sub

Private Sub cmdSearch_Click()
        searchMyRecord
End Sub

Private Sub Form_Activate()
    disableALLRECORD
    enableButtons
    showALLMEDIA
End Sub

Private Sub Form_Load()

    OpenODBCConnection
      
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'On Error GoTo err
    ListView1.SortKey = ColumnHeader.Index - 1
    ListView1.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'On Error GoTo err
        Dim i, j As Double
        
        If Item.Checked = True And baddRECORD = True Or beditRECORD = True Or bsearchRECORD = True Then
            
            j = Screen.ActiveForm.ListView1.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                If Screen.ActiveForm.ListView1.ListItems(i) <> Item Then
                   Screen.ActiveForm.ListView1.ListItems(i).Checked = False
                End If
            Next i
                        
            frmODASPMedia.txtMediaCode.Text = Item.Text
            loadRECORD
        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub



