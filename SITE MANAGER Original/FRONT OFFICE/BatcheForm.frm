VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ListBox List1 
      Height          =   8055
      Left            =   9480
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&START"
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdStart_Click()

Dim rschange As ADODB.Recordset
Set rschange = New ADODB.Recordset
Dim rschange1 As ADODB.Recordset
Set rschange1 = New ADODB.Recordset
Dim rschange2 As ADODB.Recordset
Set rschange2 = New ADODB.Recordset
Dim rschange3 As ADODB.Recordset
Set rschange3 = New ADODB.Recordset

Dim agnum As String, agnum2 As String
Dim num


 


rschange.Open "Select * from ALISPAgent", cnALIS, adOpenKeyset, adLockOptimistic

num = rschange.RecordCount
rschange1.Open " Update ALISPAgent SET ag1= SUBSTRING(AgentNo,1,3) ;", cnALIS, adOpenKeyset, adLockOptimistic
rschange2.Open " Update ALISPAgent SET ag2= SUBSTRING(AgentNo,5,3)  ;", cnALIS, adOpenKeyset, adLockOptimistic
rschange3.Open " Update ALISPAgent SET ag3= SUBSTRING(AgentNo,9,13) ;", cnALIS, adOpenKeyset, adLockOptimistic
rschange.MoveFirst

Dim i As Integer

For i = 1 To num

agnum = rschange!Ag1 & rschange!ag2 & rschange!ag3
rschange!AgentNo2 = agnum
rschange.MoveNext
Form1.List1.AddItem agnum
Next i

End Sub

Private Sub Form_Load()
Call OpenConnection
End Sub
