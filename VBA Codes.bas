Attribute VB_Name = "Module1"
Public conn As Object
Public rs As Object

Sub ConnectToDatabase()
    On Error GoTo ErrHandler

    ' Create a connection object
    Set conn = CreateObject("ADODB.Connection")

    ' Specify the database path
    Dim dbPath As String
    dbPath = "C:\Users\K KIRAN KUMAR\OneDrive\Documents\South-Indian-Restaurant.accdb" ' Adjust if necessary

    ' Open the connection
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath

    MsgBox "Database connection successful!", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error connecting to database: " & Err.Description
End Sub

Sub LoadOrders()
    Dim ws As Worksheet
    Dim query As String
    Dim row As Integer

    On Error GoTo ErrHandler

    ' Ensure the database connection is active
    If conn Is Nothing Then
        Call ConnectToDatabase
    End If

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("OrderDetailsOverview")
    ws.Cells.Clear ' Clear existing data

    ' SQL query to fetch order details
    query = "SELECT [Orders].[OrderID], [Orders].[TableNumber], [Orders].[OrderDate], [Menu].[ItemName], [OrderDetails].[Quantity], " & _
            "[OrderDetails].[UnitPrice], ([OrderDetails].[Quantity] * [OrderDetails].[UnitPrice]) AS TotalPrice, [Orders].[PaymentStatus] " & _
            "FROM ([Orders] INNER JOIN [OrderDetails] ON [Orders].[OrderID] = [OrderDetails].[OrderID]) " & _
            "INNER JOIN [Menu] ON [OrderDetails].[ItemID] = [Menu].[ItemID]"

    ' Debug: Print the query to the Immediate Window
    Debug.Print "Executing query: " & query

    ' Execute the query
    Set rs = conn.Execute(query) ' <-- Error likely occurs here

    ' Add headers to the worksheet
    ws.Range("A1:H1").Value = Array("OrderID", "TableNumber", "OrderDate", "ItemName", "Quantity", "UnitPrice", "TotalPrice", "PaymentStatus")
    ws.Rows(1).Font.Bold = True

    ' Populate data from the database
    row = 2
    Do Until rs.EOF
        ws.Cells(row, 1).Value = rs.Fields("OrderID").Value
        ws.Cells(row, 2).Value = rs.Fields("TableNumber").Value
        ws.Cells(row, 3).Value = rs.Fields("OrderDate").Value
        ws.Cells(row, 4).Value = rs.Fields("ItemName").Value
        ws.Cells(row, 5).Value = rs.Fields("Quantity").Value
        ws.Cells(row, 6).Value = rs.Fields("UnitPrice").Value
        ws.Cells(row, 7).Value = rs.Fields("TotalPrice").Value
        ws.Cells(row, 8).Value = rs.Fields("PaymentStatus").Value

        rs.MoveNext
        row = row + 1
    Loop

    rs.Close
    MsgBox "Order details loaded successfully!", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error loading orders: " & Err.Description, vbCritical
    Debug.Print "Error: " & Err.Description
    Debug.Print "Query: " & query
End Sub

Sub CreatePivotTable()
    Dim ws As Worksheet
    Dim pivotWs As Worksheet
    Dim pivotCache As pivotCache
    Dim pivotTable As pivotTable
    Dim dataRange As Range
    
    ' Set worksheets
    Set ws = ThisWorkbook.Sheets("OrderDetailsOverview")
    Set pivotWs = ThisWorkbook.Sheets("SalesSummary")
    
    ' Clear existing pivot table
    pivotWs.Cells.Clear
    
    ' Define data range
    Set dataRange = ws.Range("A1:H" & ws.Cells(ws.Rows.Count, "A").End(xlUp).row)
    
    ' Create pivot cache
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)
    
    ' Create pivot table
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotWs.Range("A3"), TableName:="SalesPivot")
    
    ' Add fields
    With pivotTable
        .PivotFields("ItemName").Orientation = xlRowField
        .PivotFields("TotalPrice").Orientation = xlDataField
    End With
    
    MsgBox "Pivot Table Created!"
End Sub

Private Sub NewOrder_Click()
    ' Clear input fields for a new order
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("NewOrdersForm")
    
    ws.Range("B3").Value = "" ' Clear Table Number
    ws.Range("B4").Value = Date ' Set today's date for Order Date
    ws.Range("B5").Value = "" ' Clear Menu Item
    ws.Range("B6").Value = "" ' Clear Quantity
    ws.Range("B7").Value = "" ' Clear Payment Status
    MsgBox "New order form is ready.", vbInformation
End Sub

Private Sub SubmitOrder_Click()
    ' Submit the order to the database
    Dim conn As Object
    Dim rs As Object
    Dim sql As String
    Dim ws As Worksheet
    Dim orderID As Long
    Dim tableNo As String, orderDate As String, paymentStatus As String
    Dim menuItem As String, quantity As Double, unitPrice As Double, totalAmount As Double
    
    On Error GoTo ErrorHandler
    
    ' Set worksheet
    Set ws = ThisWorkbook.Worksheets("NewOrdersForm")
    
    ' Fetch input values
    tableNo = ws.Range("B3").Value
    orderDate = ws.Range("B4").Value
    paymentStatus = ws.Range("B7").Value
    menuItem = ws.Range("B5").Value
    quantity = ws.Range("B6").Value
    
    If tableNo = "" Or paymentStatus = "" Or menuItem = "" Or quantity <= 0 Then
        MsgBox "Please fill in all fields correctly.", vbExclamation
        Exit Sub
    End If
    
    ' Calculate unit price and total amount
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\K KIRAN KUMAR\OneDrive\Documents\South-Indian-Restaurant.accdb;"
    
    ' Get Unit Price from Menu table
    sql = "SELECT Price FROM Menu WHERE ItemName = '" & menuItem & "'"
    Set rs = conn.Execute(sql)
    
    If Not rs.EOF Then
        unitPrice = rs.Fields("Price").Value
    Else
        MsgBox "Item not found in the database.", vbExclamation
        rs.Close
        conn.Close
        Exit Sub
    End If
    rs.Close
    
    ' Calculate total amount
    totalAmount = quantity * unitPrice
    
    ' Insert order into Orders table
    sql = "INSERT INTO Orders (TableNo, OrderDate, TotalAmount, PaymentStatus) " & _
          "VALUES ('" & tableNo & "', #" & orderDate & "#, " & totalAmount & ", '" & paymentStatus & "')"
    conn.Execute sql
    
    ' Get the last inserted OrderID
    Set rs = conn.Execute("SELECT @@IDENTITY AS LastID")
    orderID = rs.Fields("LastID").Value
    rs.Close
    
    ' Insert order details into Order Details table
    sql = "INSERT INTO [Order Details] (OrderID, ItemID, Quantity, UnitPrice) " & _
          "VALUES (" & orderID & ", (SELECT ItemID FROM Menu WHERE ItemName = '" & menuItem & "'), " & _
          quantity & ", " & unitPrice & ")"
    conn.Execute sql
    
    MsgBox "Order submitted successfully!", vbInformation

    ' Clear form for next order
    Call NewOrder_Click

    ' Close connection
    conn.Close
    Set conn = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    If Not conn Is Nothing Then conn.Close
    Set conn = Nothing
End Sub











