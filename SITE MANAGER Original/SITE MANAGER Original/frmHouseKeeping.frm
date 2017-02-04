VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmHouseKeeping 
   Caption         =   "MonthlyStockTake"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkSelectAll 
      Caption         =   "Check1"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   8160
      Width           =   255
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   8055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   14208
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777152
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmHouseKeeping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ActiveSearchEdit, NextMonthTotalStockCost, NextMonthRetailStockCost, NextMonthWholesaleStockCost, GrossProfit, LastMonthTotalSales, LastMonthTotalRetailSales, LastMonthTotalWholesalesales, LastMonthTotalStockCost, LastMonthSoldRetailStockCost, LastMonthSoldStockCost, LastMonthSoldRetailStock, LastMonthSoldStock, OpeningStockForLastMonth, ReceiptNo, PorderNo, TransNo, ProductCode, TotalAmountBought, TotalMonthCost, ProductName, ProductPrice, CurrentQuantity, TotalPrice, SaleType, QuantityUnits, TotalPieces, PorderDate
 
Private Function ItemsSelected() As Boolean
'On Error GoTo Err
With Me
Dim i, j, k
j = .ListView1.ListItems.Count: k = 0

If j = 0 Or .ListView1.View <> lvwReport Then

    ItemsSelected = False
    
    strMessage = "There are no Items on the Datasheet to Perform this Operation!!!"
    
Else

    For i = 1 To j
        If .ListView1.ListItems(i).Checked = True Then
            k = k + 1
        End If
    Next i
    
    If k = 0 Then
        ItemsSelected = False
        strMessage = "Please Select at Least ONE item from the datasheet!!"
    ElseIf k >= 1 Then
        ItemsSelected = True
    End If
    
End If
If Not ItemsSelected Then
    MsgBox strMessage, vbCritical + vbOKOnly, "Data Validation"
End If
End With
Exit Function
Err:
    ErrorMessage
End Function

Private Function GetWholesaleAmount() As Variant
'On Error GoTo Err
With Me
    Set rsFindRecord = cnCOMMON.Execute("SELECT WholesaleQuantity FROM GeneralProducts WHERE DrugCode='" & Trim(ProductCode) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        Set rsFindRecord = Nothing
    Else
        GetWholesaleAmount = rsFindRecord!WholeSaleQuantity & ""
        
        
    End If
End With
Exit Function
Err:
    ErrorMessage
End Function

Private Function GetCurrentStockQuantity() As Variant
'On Error GoTo Err
With Me
    Set rsFindRecord = cnCOMMON.Execute("SELECT CurrentQuantity FROM GenProductsInventory WHERE DrugCode ='" & ProductCode & "'")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        GetCurrentStockQuantity = 0: Set rsFindRecord = Nothing
    ElseIf IsNull(rsFindRecord!CurrentQuantity) = True Or rsFindRecord!CurrentQuantity = "" Then
        GetCurrentStockQuantity = 0: Set rsFindRecord = Nothing
    Else
        GetCurrentStockQuantity = CDbl(rsFindRecord!CurrentQuantity)
    End If
        
    Set rsFindRecord = Nothing
Exit Function
Err:
    ErrorMessage
    End With
End Function
Private Function GetLastMonthWholeSaleStockSold() As Variant

'On Error GoTo Err
With Me
    Set rsFindRecord = cnCOMMON.Execute("SELECT SUM(Quantity) as Total FROM PharmPointOfSale WHERE DrugCode ='" & ProductCode & "' AND AccPeriod = '" & LastAccPeriod & "' AND SaleType = '" & "WholeSale" & "'")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        GetLastMonthWholeSaleStockSold = 0: Set rsFindRecord = Nothing
    ElseIf IsNull(rsFindRecord!Total) = True Or rsFindRecord!Total = "" Then
       GetLastMonthWholeSaleStockSold = 0: Set rsFindRecord = Nothing
    Else
       GetLastMonthWholeSaleStockSold = CDbl(rsFindRecord!Total)
    End If
        
    Set rsFindRecord = Nothing
Exit Function
Err:
    ErrorMessage
    End With
End Function

Private Function GetLastMonthWholeSaleSales() As Variant
'On Error GoTo Err
With Me
    Set rsFindRecord = cnCOMMON.Execute("SELECT SUM(TotalPrice) as Total FROM PharmPointOfSale WHERE DrugCode ='" & ProductCode & "' AND AccPeriod = '" & LastAccPeriod & "' AND SaleType = '" & "WholeSale" & "'")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        GetLastMonthWholeSaleSales = 0: Set rsFindRecord = Nothing
    ElseIf IsNull(rsFindRecord!Total) = True Or rsFindRecord!Total = "" Then
       GetLastMonthWholeSaleSales = 0: Set rsFindRecord = Nothing
    Else
       GetLastMonthWholeSaleSales = CDbl(rsFindRecord!Total)
    End If
        
    Set rsFindRecord = Nothing
Exit Function
Err:
    ErrorMessage
    End With
End Function

Private Function GetLastMonthRetailSales() As Variant
'On Error GoTo Err
With Me
    Set rsFindRecord = cnCOMMON.Execute("SELECT SUM(TotalPrice) as Total FROM PharmPointOfSale WHERE DrugCode ='" & ProductCode & "' AND AccPeriod = '" & LastAccPeriod & "' AND SaleType = '" & "Retail" & "'")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        GetLastMonthRetailSales = 0: Set rsFindRecord = Nothing
    ElseIf IsNull(rsFindRecord!Total) = True Or rsFindRecord!Total = "" Then
       GetLastMonthRetailSales = 0: Set rsFindRecord = Nothing
    Else
       GetLastMonthRetailSales = CDbl(rsFindRecord!Total)
    End If
        
    Set rsFindRecord = Nothing
Exit Function
Err:
    ErrorMessage
    End With
End Function


Private Function GetLastMonthRetailStockSold() As Variant
'On Error GoTo Err
With Me
    Set rsFindRecord = cnCOMMON.Execute("SELECT SUM(Quantity) as Total FROM PharmPointOfSale WHERE DrugCode ='" & ProductCode & "' AND AccPeriod = '" & LastAccPeriod & "' AND SaleType = '" & "Retail" & "'")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        GetLastMonthRetailStockSold = 0: Set rsFindRecord = Nothing
    ElseIf IsNull(rsFindRecord!Total) = True Or rsFindRecord!Total = "" Then
       GetLastMonthRetailStockSold = 0: Set rsFindRecord = Nothing
    Else
       GetLastMonthRetailStockSold = CDbl(rsFindRecord!Total)
    End If
        
    Set rsFindRecord = Nothing
Exit Function
Err:
    ErrorMessage
    End With
End Function

Private Function GetOpeningStockForLastMonth() As Variant
'On Error GoTo Err
With Me
    Set rsFindRecord = cnCOMMON.Execute("SELECT QuantityReceived FROM AccPeriodStock WHERE ProductCode ='" & ProductCode & "'")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        GetOpeningStockForLastMonth = 0: Set rsFindRecord = Nothing
    ElseIf IsNull(rsFindRecord!QuantityReceived) = True Or rsFindRecord!QuantityReceived = "" Then
        GetOpeningStockForLastMonth = 0: Set rsFindRecord = Nothing
    Else
        GetOpeningStockForLastMonth = CDbl(rsFindRecord!QuantityReceived)
    End If
        
    Set rsFindRecord = Nothing
Exit Function
Err:
    ErrorMessage
    End With
End Function

Private Function GetCurrentStockTotalPieces() As Variant
'On Error GoTo Err
With Me
    Set rsFindRecord = cnCOMMON.Execute("SELECT TotalPieces FROM GenProductsInventory  WHERE DrugCode='" & ProductCode & "'")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        GetCurrentStockTotalPieces = 0: Set rsFindRecord = Nothing
    ElseIf IsNull(rsFindRecord!TotalPieces) = True Or rsFindRecord!TotalPieces = "" Then
        GetCurrentStockTotalPieces = 0: Set rsFindRecord = Nothing
    Else
        GetCurrentStockTotalPieces = CDbl(rsFindRecord!TotalPieces) Mod GetWholesaleAmount
    End If
        
    Set rsFindRecord = Nothing
Exit Function
Err:
    ErrorMessage
    End With
End Function

Private Function GetWholesaleCost() As Variant
'On Error GoTo Err
With Me
    Set rsFindRecord = cnCOMMON.Execute("SELECT DosageCost FROM ProductsCostPriceSetup  WHERE DrugCode = '" & ProductCode & "' ;")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        GetWholesaleCost = 0: Set rsFindRecord = Nothing
    ElseIf IsNull(rsFindRecord!DosageCost) = True Or rsFindRecord!DosageCost = "" Then
        GetWholesaleCost = 0: Set rsFindRecord = Nothing
    Else
        GetWholesaleCost = CDbl(rsFindRecord!DosageCost)
    End If
        
    Set rsFindRecord = Nothing
Exit Function
Err:
    ErrorMessage
    End With
End Function

Private Function GetRetailCost() As Variant
'On Error GoTo Err
With Me
    Set rsFindRecord = cnCOMMON.Execute("SELECT RetailCost FROM ProductsCostPriceSetup  WHERE DrugCode = '" & ProductCode & "' ;")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        GetRetailCost = 0: Set rsFindRecord = Nothing
    ElseIf IsNull(rsFindRecord!RetailCost) = True Or rsFindRecord!RetailCost = "" Then
        GetRetailCost = 0: Set rsFindRecord = Nothing
    Else
        GetRetailCost = CDbl(rsFindRecord!RetailCost)
    End If
        
    Set rsFindRecord = Nothing
Exit Function
Err:
    ErrorMessage
    End With
End Function

Private Sub Form_Load()
'On Error GoTo Err
Dim WholesaleStock, RetailStock, TotalMonthStock, WholesaleAmount, RetailAmount As Variant
 Dim i, j, k
With Me
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Product Code ", .ListView1.Width / 6.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Product Name", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Current Quantity", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Quantity Units", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Total Pieces", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "P.Order Date", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "P.Order Number", .ListView1.Width / 5 ', lvwColumnCenter

.ListView1.View = lvwReport


Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM GenProductsInventory A,AccPeriodStock B  WHERE A.DrugCode= B.ProductCode;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!Drugcode))

    If Not IsNull(rsLIST!DrugName) Then
        MyList.SubItems(1) = CStr(rsLIST!DrugName)
    End If
    
    If Not IsNull(rsLIST!CurrentQuantity) Then
        MyList.SubItems(2) = CStr(rsLIST!CurrentQuantity)
    End If
        
    If Not IsNull(rsLIST!QuantityUnits) Then
        MyList.SubItems(3) = CStr(rsLIST!QuantityUnits)
    End If
    
    If Not IsNull(rsLIST!TotalPieces) Then
        MyList.SubItems(4) = CStr(rsLIST!TotalPieces)
    End If
    
    If Not IsNull(rsLIST!PorderDate) Then
        MyList.SubItems(5) = CStr(rsLIST!PorderDate)
    End If
    
    If Not IsNull(rsLIST!PorderNo) Then
        MyList.SubItems(6) = CStr(rsLIST!PorderNo)
    End If
    
    
        
    rsLIST.MoveNext
    
Wend
 .chkSelectAll.Value = 1
 Call chkSelectAll_Click

Set MyList = Nothing: Set rsLIST = Nothing

    j = .ListView1.ListItems.Count
    If j = 0 Or .ListView1.View <> lvwReport Then Exit Sub
 
    Set rsNewRecord = New ADODB.Recordset
     rsNewRecord.Open "INSERT INTO PreviousStockRecords(PorderNo,ProductCode,ProductName,AccPeriod,TransDate,CreatedBy)(SELECT PorderNo,ProductCode,ProductName,AccPeriod,DateCreated,CreatedBy FROM AccPeriodStock)", cnCOMMON, adOpenKeyset, adLockOptimistic
    Set rsNewRecord = Nothing
    

For i = 1 To j
If .ListView1.ListItems(i).Checked = True Then
    
        ProductCode = .ListView1.ListItems(i).Text
        ProductName = .ListView1.ListItems(i).SubItems(1)
        CurrentQuantity = .ListView1.ListItems(i).SubItems(2)
        QuantityUnits = .ListView1.ListItems(i).SubItems(3)
        TotalPieces = .ListView1.ListItems(i).SubItems(4)
        PorderDate = .ListView1.ListItems(i).SubItems(5)
        PorderNo = .ListView1.ListItems(i).SubItems(6)
    
    'previous month's sale details
    OpeningStockForLastMonth = GetOpeningStockForLastMonth
    LastMonthSoldStock = GetLastMonthWholeSaleStockSold
    LastMonthSoldRetailStock = GetLastMonthRetailStockSold
    LastMonthSoldStockCost = GetLastMonthWholeSaleStockSold * GetWholesaleCost
    LastMonthSoldRetailStockCost = GetLastMonthRetailStockSold * GetRetailCost
    LastMonthTotalStockCost = LastMonthSoldStockCost + LastMonthSoldRetailStockCost
    LastMonthTotalWholesalesales = GetLastMonthWholeSaleSales
    LastMonthTotalRetailSales = GetLastMonthRetailSales
    LastMonthTotalSales = LastMonthTotalWholesalesales + LastMonthTotalRetailSales
    GrossProfit = LastMonthTotalSales - LastMonthTotalStockCost
    TotalMonthCost = GetTotalMonthCost
    TotalAmountBought = GetTotalAmountBought
    
    
    'next month's stock details
    NextMonthWholesaleStockCost = GetCurrentStockQuantity * GetWholesaleCost
    NextMonthRetailStockCost = GetCurrentStockTotalPieces * GetRetailCost
    NextMonthTotalStockCost = NextMonthWholesaleStockCost + NextMonthRetailStockCost
    WholesaleAmount = GetCurrentStockQuantity
    RetailAmount = GetCurrentStockTotalPieces
   
    
    
    
    Set rsLineUpdate = New ADODB.Recordset
      rsLineUpdate.Open "UPDATE previousStockRecords SET WholesaleSold = '" & LastMonthSoldStock & "',Retailsold = '" & LastMonthSoldRetailStock & "',  GrossProfit = '" & GrossProfit & "',QuantityReceived = '" & TotalAmountBought & "',ClosingStockCost = " & TotalMonthCost & ",ClosingWholesaleAmount = '" & WholesaleAmount & "',ClosingRetailAmount = '" & RetailAmount & "',DateCreated = '" & MyCurrentDate & "' WHERE ProductCode = '" & ProductCode & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
    Set rsLineUpdate = Nothing
    
    Set rsDeleteRecord = New ADODB.Recordset
       rsDeleteRecord.Open "DELETE  FROM AccPeriodStock WHERE ProductCode = '" & ProductCode & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
    Set rsDeleteRecord = Nothing
    
    Set rsNewRecord = New ADODB.Recordset
      rsNewRecord.Open "INSERT INTO AccPeriodStock (PorderNo,ProductCode,ProductName,Cost,AccPeriod,CurrentAccPeriod,OpeningStockCost,OpeningWholesaleAmount,OpeningRetailAmount,DateCreated,CreatedBy,DateUpDated) VALUES('" & PorderNo & "','" & ProductCode & "','" & ProductName & "'," & NextMonthTotalStockCost & ",'" & MyCurrentPeriod & "','" & MyCurrentPeriod & "'," & NextMonthTotalStockCost & ",'" & WholesaleAmount & "','" & RetailAmount & "','" & MyCurrentDate & "','" & CurrentUserName & "','" & MyCurrentDate & "') ", cnCOMMON, adOpenKeyset, adLockOptimistic
    Set rsNewRecord = Nothing
    End If
  Next i
   
   
End With
Exit Sub

Err:
If Err.Number = 3265 Or 364 Then Resume Next
    ErrorMessage
End Sub

Private Sub chkSelectAll_Click()
'On Error GoTo Err
With Me
Dim i, j, k
j = .ListView1.ListItems.Count

If j = 0 Or .ListView1.View <> lvwReport Then .chkSelectAll.Value = 0: Exit Sub

Select Case .chkSelectAll.Value
Case 0
    For i = 1 To j
        .ListView1.ListItems(i).Checked = False
    Next i
Case 1
    For i = 1 To j
        .ListView1.ListItems(i).Checked = True
    Next i
Case Else
    Exit Sub
End Select
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Private Function GetTotalMonthCost() As Variant
Dim LastMonth As Variant
LastMonth = LastAccPeriod
'On Error GoTo Err
With Me
    Set rsFindRecord = cnCOMMON.Execute("SELECT Sum(Cost) as Total FROM AccPeriodStock WHERE ProductCode = '" & ProductCode & "' and AccPeriod = '" & LastMonth & "' ;")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        GetTotalMonthCost = 0: Set rsFindRecord = Nothing
    ElseIf IsNull(rsFindRecord!Total) = True Or rsFindRecord!Total = "" Then
        GetTotalMonthCost = 0: Set rsFindRecord = Nothing
    Else
        GetTotalMonthCost = CDbl(rsFindRecord!Total)
    End If
        
    Set rsFindRecord = Nothing
Exit Function
Err:
    ErrorMessage
    End With
End Function

Private Function GetTotalAmountBought() As Variant
'On Error GoTo Err
With Me
    Set rsFindRecord = cnCOMMON.Execute("SELECT Sum(QuantityReceived) as Total FROM AccPeriodStock WHERE ProductCode = '" & ProductCode & "' and AccPeriod = '" & LastAccPeriod & "' ;")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        GetTotalAmountBought = 0: Set rsFindRecord = Nothing
    ElseIf IsNull(rsFindRecord!Total) = True Or rsFindRecord!Total = "" Then
        GetTotalAmountBought = 0: Set rsFindRecord = Nothing
    Else
        GetTotalAmountBought = CDbl(rsFindRecord!Total)
    End If
        
    Set rsFindRecord = Nothing
Exit Function
Err:
    ErrorMessage
    End With
End Function

