# VBA_Analysis_for_Automobile_firm
Develop Macro to automate Annual report generation
Sub AssgnQ01()

' Create Excel File Annual overview

Workbooks.Add
ActiveWorkbook.SaveAs "Annual Overview.xlsx"
Sheets(1).Select
Sheets(1).Name = "Dashboard"
Range("A2").Select
ActiveCell.FormulaR1C1 = "Please Enter Report Year"
Selection.Font.Bold = True
Columns("A:A").EntireColumn.AutoFit
Range("B2").Select
Dim input1 As String
input1 = InputBox("Please Enter Report Year", "Enter Year", Int(4))
Range("B2").Value = input1
Range("B2").Select
Selection.Font.Bold = True

' Create Copy of Source File into Destination File


Dim SourceFile01 As Workbook
Dim SourceFile02 As Workbook
Dim SourceFile03 As Workbook
Dim SourceFile04 As Workbook
Dim SourceFile05 As Workbook
Dim SourceFile06 As Workbook
Dim SourceFile07 As Workbook
Dim SourceFile08 As Workbook
Dim DestinationFile As Workbook
Dim SourceSheet01 As Worksheet
Dim SourceSheet02 As Worksheet
Dim SourceSheet03 As Worksheet
Dim SourceSheet04 As Worksheet
Dim SourceSheet05 As Worksheet
Dim SourceSheet06 As Worksheet
Dim SourceSheet07 As Worksheet
Dim SourceSheet08 As Worksheet
Dim DestinationSheet01 As Worksheet
Dim DestinationSheet02 As Worksheet
Dim DestinationSheet03 As Worksheet
Dim DestinationSheet04 As Worksheet
Dim DestinationSheet05 As Worksheet
Dim DestinationSheet06 As Worksheet
Dim DestinationSheet07 As Worksheet
Dim DestinationSheet08 As Worksheet

'Setting all the source files and sheets
Set SourceFile01 = Workbooks("customers.csv")
Set SourceFile02 = Workbooks("employees.csv")
Set SourceFile03 = Workbooks("offices.csv")
Set SourceFile04 = Workbooks("orderdetails.csv")
Set SourceFile05 = Workbooks("orders.csv")
Set SourceFile06 = Workbooks("payments.csv")
Set SourceFile07 = Workbooks("productlines.csv")
Set SourceFile08 = Workbooks("products.csv")
Set DestinationFile = Workbooks("Annual Overview.xlsx")

'Setting Source Sheets
Set SourceSheet01 = SourceFile01.Sheets("customers")
Set SourceSheet02 = SourceFile02.Sheets("employees")
Set SourceSheet03 = SourceFile03.Sheets("offices")
Set SourceSheet04 = SourceFile04.Sheets("orderdetails")
Set SourceSheet05 = SourceFile05.Sheets("orders")
Set SourceSheet06 = SourceFile06.Sheets("payments")
Set SourceSheet07 = SourceFile07.Sheets("productlines")
Set SourceSheet08 = SourceFile08.Sheets("products")

'Copy source data into annual overview workbook
SourceSheet01.Copy After:=DestinationFile.Sheets(1)
Set DestinationSheet01 = DestinationFile.Sheets(2)
DestinationSheet01.Name = "CustomerCopy"

SourceSheet02.Copy After:=DestinationFile.Sheets(2)
Set DestinationSheet02 = DestinationFile.Sheets(3)
DestinationSheet02.Name = "employeesCopy"

SourceSheet03.Copy After:=DestinationFile.Sheets(3)
Set DestinationSheet03 = DestinationFile.Sheets(4)
DestinationSheet03.Name = "officesCopy"

SourceSheet04.Copy After:=DestinationFile.Sheets(4)
Set DestinationSheet04 = DestinationFile.Sheets(5)
DestinationSheet04.Name = "orderdetailsCopy"

SourceSheet05.Copy After:=DestinationFile.Sheets(5)
Set DestinationSheet05 = DestinationFile.Sheets(6)
DestinationSheet05.Name = "ordersCopy"

SourceSheet06.Copy After:=DestinationFile.Sheets(6)
Set DestinationSheet06 = DestinationFile.Sheets(7)
DestinationSheet06.Name = "paymentsCopy"

SourceSheet07.Copy After:=DestinationFile.Sheets(7)
Set DestinationSheet07 = DestinationFile.Sheets(8)
DestinationSheet07.Name = "productlinesCopy"

SourceSheet08.Copy After:=DestinationFile.Sheets(8)
Set DestinationSheet08 = DestinationFile.Sheets(9)
DestinationSheet08.Name = "productsCopy"

'Preparation of Sheet Q01

Sheets.Add After:=DestinationFile.Sheets(9)
Sheets("employeescopy").Select
Range("A1:C240").Select
Selection.Copy
Sheets(10).Select
ActiveSheet.Paste
Sheets("employeescopy").Select
Range("F1").Select
Range(Selection, Selection.End(xlDown)).Select
Application.CutCopyMode = False
Selection.Copy
Sheets(10).Select
Range("D1").Select
ActiveSheet.Paste
Range("I15").Select
Sheets(10).Select
Range("E1").Select
Application.CutCopyMode = False
ActiveCell.FormulaR1C1 = "Sales Revenue"
Range("F1").Select
Sheets(10).Select
ActiveCell.FormulaR1C1 = "Rank of employees"
Range("E2").Select
Columns("E:E").EntireColumn.AutoFit
Columns("F:F").EntireColumn.AutoFit
Range("E3").Select
Sheets(10).Name = "Q01"


'Preparation of Sheet YearlyOrder

Sheets("ordersCopy").Select
Range("H1").Select
ActiveCell.FormulaR1C1 = "Year"
Range("H2").Select
ActiveCell.FormulaR1C1 = "=YEAR(RC[-6])"
Range("H2").Select
Selection.AutoFill Destination:=Range("H2:H327")
Range("H2:H327").Select
Range("I3").Select


Sheets.Add After:=DestinationFile.Sheets(10)
Sheets(11).Name = "YearlyOrder"
Sheets(11).Select
Range("A1").Select
ActiveCell.FormulaR1C1 = "orderNumber"
Range("B1").Select
ActiveCell.FormulaR1C1 = "orderDate"
Range("C1").Select
ActiveCell.FormulaR1C1 = "requiredDate"
Range("D1").Select
ActiveCell.FormulaR1C1 = "shippedDate"
Range("E1").Select
ActiveCell.FormulaR1C1 = "status"
Range("F1").Select
ActiveCell.FormulaR1C1 = "comments"
Range("G1").Select
ActiveCell.FormulaR1C1 = "customerNumber"
Range("H1").Select
ActiveCell.FormulaR1C1 = "Total Price Per Order"
Range("I1").Select
ActiveCell.FormulaR1C1 = "SalesRepEmployeeNo"
Range("M2").Select
ActiveCell.FormulaR1C1 = "Selected Year"
Range("N2").Select
ActiveCell.FormulaR1C1 = "=Dashboard!RC[-12]"
Range("L5").Select
ActiveCell.FormulaR1C1 = "Date Value"
Range("L7").Select
ActiveCell.FormulaR1C1 = "2003"
Range("L8").Select
ActiveCell.FormulaR1C1 = "2004"
Range("L9").Select
ActiveCell.FormulaR1C1 = "2005"
Range("M7").Select
ActiveCell.FormulaR1C1 = "=DATEVALUE(""01/01/2003"")"
Range("M8").Select
ActiveCell.FormulaR1C1 = "=DATEVALUE(""01/01/2004"")"
Range("M9").Select
ActiveCell.FormulaR1C1 = "=DATEVALUE(""01/01/2005"")"
Range("M5").Select
ActiveCell.FormulaR1C1 = _
    "=IF(R[-3]C[1]=2003,R[2]C,IF(R[-3]C[1]=2004,R[3]C,IF(R[-3]C[1]=2005,R[4]C,""NA"")))"
Range("M6").Select
Range("A2").Select
ActiveCell.Formula2R1C1 = _
    "=FILTER(ordersCopy!RC:R[325]C[6],ordersCopy!RC[7]:R[325]C[7]=RC[13])"
Range("A3").Select
Range("B2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.NumberFormat = "m/d/yyyy"
Range("C2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.NumberFormat = "m/d/yyyy"
Range("D2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.NumberFormat = "m/d/yyyy"
Range("K10").Select


Columns("A:A").EntireColumn.AutoFit
Columns("B:B").EntireColumn.AutoFit
Columns("C:C").EntireColumn.AutoFit
Columns("D:D").EntireColumn.AutoFit
Columns("E:E").EntireColumn.AutoFit
Columns("G:G").EntireColumn.AutoFit
Columns("H:H").EntireColumn.AutoFit
Columns("I:I").EntireColumn.AutoFit
Columns("L:L").EntireColumn.AutoFit
Columns("M:M").EntireColumn.AutoFit


Sheets("orderdetailsCopy").Select
Columns("E:E").EntireColumn.AutoFit
Range("F1").Select
ActiveCell.FormulaR1C1 = "Total Price"
Range("F2").Select
Application.CutCopyMode = False
ActiveCell.FormulaR1C1 = "=RC[-3]*RC[-2]"
Range("F2").Select
Selection.AutoFill Destination:=Range("F2:F2997")
Range("F2:F2997").Select
Range("H13").Select



Range("G1").Select
Sheets("YearlyOrder").Select
Columns("G:G").EntireColumn.AutoFit
Range("H1").Select
ActiveCell.FormulaR1C1 = "Total Price per order"
Range("H2").Select
Columns("H:H").EntireColumn.AutoFit
ActiveCell.FormulaR1C1 = _
   "=SUMIF(orderdetailsCopy!R2C1:R2997C1,YearlyOrder!RC[-7],orderdetailsCopy!R2C6:R2997C6)"
Range("H2").Select
Selection.AutoFill Destination:=Range("H2:H327")
Range("H2:H327").Select
Range("J10").Select





Sheets("YearlyOrder").Select
Range("I1").Select
ActiveCell.FormulaR1C1 = "SalesRep Emp ID"
Range("I2").Select
Columns("I:I").EntireColumn.AutoFit
Range("I2").Select
ActiveCell.FormulaR1C1 = _
    "=INDEX(CustomerCopy!R2C12:R123C12,MATCH(YearlyOrder!RC[-2],CustomerCopy!R2C1:R123C1,0),1)"
Range("I2").Select
Selection.AutoFill Destination:=Range("I2:I327")
Range("I2:I327").Select
Range("J9").Select


Sheets("Q01").Select
Range("E2").Select
ActiveCell.FormulaR1C1 = _
    "=SUMIF(YearlyOrder!R2C9:R327C9,'Q01'!RC[-4],YearlyOrder!R2C8:R327C8)"
Range("E2").Select
Range("E2").Select
Selection.AutoFill Destination:=Range("E2:E24")
Sheets("Q01").Select
Range("F15").Select

Range("F2").Select
ActiveCell.FormulaR1C1 = "=RANK.EQ(RC[-1],R2C5:R24C5)"
Range("F2").Select
Selection.AutoFill Destination:=Range("F2:F24")
Range("F2:F24").Select
Range("H2").Select

Sheets("officesCopy").Select
Range("A1:B8").Select
Selection.Copy
Sheets("Q01").Select
Range("K1").Select
ActiveSheet.Paste
Range("M15").Select
Columns("L:L").EntireColumn.AutoFit
Columns("K:K").EntireColumn.AutoFit
Sheets("officesCopy").Select
Range("G1:G8").Select
Application.CutCopyMode = False
Selection.Copy
Sheets("Q01").Select
Range("M1").Select
ActiveSheet.Paste
Range("N1").Select
Application.CutCopyMode = False
ActiveCell.FormulaR1C1 = "Sales Revenue"
Range("O1").Select
ActiveCell.FormulaR1C1 = "Rank"
Range("O2").Select
Columns("N:N").EntireColumn.AutoFit

Range("N2").Select
ActiveCell.FormulaR1C1 = "=SUMIF(R2C4:R24C4,RC[-3],R2C5:R24C5)"
Range("N2").Select
Selection.AutoFill Destination:=Range("N2:N8")
Range("N2:N8").Select
Range("O2").Select
ActiveCell.FormulaR1C1 = "=RANK.EQ(RC[-1],R2C14:R8C14)"
Range("O2").Select
Selection.AutoFill Destination:=Range("O2:O8")
Range("O2:O8").Select
Range("N4").Select

Sheets("Q01").Select
Range("A1:P1").Select
Selection.Font.Bold = True

Worksheets("Q01").Move After:=Worksheets("YearlyOrder")

' Dashboard Preparation for Q01

Sheets("Dashboard").Select
Range("B4").Select
ActiveCell.FormulaR1C1 = "Annual Sales Gross Revenue"
Range("C5").Select
ActiveCell.FormulaR1C1 = "Employee ID"
Range("D5").Select
ActiveCell.FormulaR1C1 = "First Name"
Range("E5").Select
ActiveCell.FormulaR1C1 = "Last Name"
Range("F5").Select
ActiveCell.FormulaR1C1 = "Office ID"
Range("G5").Select
ActiveCell.FormulaR1C1 = "Total Sales"
Range("H5").Select
ActiveCell.FormulaR1C1 = "Rank"

Range("K5").Select
ActiveCell.FormulaR1C1 = "Office Code"
Range("L5").Select
ActiveCell.FormulaR1C1 = "City"
Range("M5").Select
ActiveCell.FormulaR1C1 = "Country"
Range("N5").Select
ActiveCell.FormulaR1C1 = "Total Sales"
Range("O5").Select
ActiveCell.FormulaR1C1 = "Rank"

Range("K15").Select
ActiveCell.FormulaR1C1 = "Top Three Employees"
Selection.Font.Bold = True

Range("K22").Select
ActiveCell.FormulaR1C1 = "Top Three Offices"
Selection.Font.Bold = True

Range("C6").Select
ActiveCell.Formula2R1C1 = "=SORT('Q01'!R[-4]C[-2]:R[18]C[3],6,1)"
Range("K6").Select
ActiveCell.Formula2R1C1 = "=SORT('Q01'!R[-4]C:R[2]C[4],5,1)"
Range("K16").Select
Application.CutCopyMode = False
ActiveCell.Formula2R1C1 = "=R[-11]C[-8]:R[-8]C[-3]"
Range("K23").Select
Application.CutCopyMode = False
ActiveCell.Formula2R1C1 = "=R[-18]C:R[-15]C[4]"
Range("B4").Select
Selection.Font.Bold = True
Range("C5:H5").Select
Selection.Font.Bold = True
Range("K5:O5").Select
Selection.Font.Bold = True
Columns("A:A").EntireColumn.AutoFit
Columns("B:B").EntireColumn.AutoFit
Columns("C:C").EntireColumn.AutoFit
Columns("D:D").EntireColumn.AutoFit
Columns("E:E").EntireColumn.AutoFit
Columns("F:F").EntireColumn.AutoFit
Columns("G:G").EntireColumn.AutoFit
Columns("H:H").EntireColumn.AutoFit
Columns("K:K").EntireColumn.AutoFit
Columns("L:L").EntireColumn.AutoFit
Columns("M:M").EntireColumn.AutoFit
Columns("N:N").EntireColumn.AutoFit
Columns("O:O").EntireColumn.AutoFit

'Q2 Section

Sheets("YearlyOrder").Select
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Late Shipment"
    Range("J2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNUMBER(RC[-6]),IF(RC[-6]-RC[-8]>5,RC[-9],0),RC[-9])"
    Range("J2").Select
    Selection.AutoFill Destination:=Range("J2:J327")
    Range("J2:J327").Select
    Range("K17").Select
    Sheets("Q01").Select
    Sheets.Add After:=ActiveSheet
    Sheets(12).Select
    Sheets(12).Name = "Q02A"
    Range("A2").Select
    ActiveCell.Formula2R1C1 = _
        "=FILTER(YearlyOrder!RC:R[98]C[9],YearlyOrder!RC[9]:R[98]C[9]>10)"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "OrderNum"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "OrderDate"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "requiredDate"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "ShippedDate"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Status"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "comments"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "customerNo"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "TotalPrice"
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "SalesRep EmpID"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Late Order"
    Range("L2").Select

    Sheets("orderdetailsCopy").Select
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Late Shipped"
    Range("G2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNUMBER(MATCH(RC[-6],Q02A!R2C1:R17C1,0)),""YES"",""NO"")"
    Range("G2").Select
    Selection.AutoFill Destination:=Range("G2:G2997")
    Range("G2:G2997").Select
    Range("H3").Select
    Sheets("Q02A").Select
    Range("L2").Select
    ActiveCell.Formula2R1C1 = _
        "=FILTER(orderdetailsCopy!RC[-10]:R[4998]C[-5],orderdetailsCopy!RC[-5]:R[4998]C[-5]=""YES"")"
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "Product"
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "QuantityOrdered"
    Range("N1").Select
    ActiveCell.FormulaR1C1 = "PriceEach"
    Range("O1").Select
    ActiveCell.FormulaR1C1 = "OrderLineNo"
    Range("P1").Select
    ActiveCell.FormulaR1C1 = "Total Price"
    Range("Q1").Select
    ActiveCell.FormulaR1C1 = "Shipped Late"
    Range("J23").Select
    Sheets.Add After:=ActiveSheet
    
    Sheets(13).Select
    Sheets(13).Name = "Q02B"
    Range("A2").Select
    ActiveCell.Formula2R1C1 = "=UNIQUE(Q02A!RC[11]:R[133]C[11])"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(Q02A!R2C12:R500C12,Q02B!RC[-1])"
    Range("B2").Select
    Selection.AutoFill Destination:=Range("B2:B92")
    Range("B2:B92").Select
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Products shipped late"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Late Shipment Count"
    Range("D3").Select
    Columns("A:A").EntireColumn.AutoFit
    Columns("B:B").EntireColumn.AutoFit
    Range("D2").Select
    ActiveCell.Formula2R1C1 = "=UNIQUE(Q02A!RC[5]:R[15]C[5])"
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(Q02A!R2C9:R17C9,Q02B!RC[-1])"
    Range("E2").Select
    Selection.AutoFill Destination:=Range("E2:E13")
    Range("E2:E13").Select
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "EmployeeID"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "LateShipment Count"
    Range("E2").Select
    Columns("D:D").EntireColumn.AutoFit
    Columns("E:E").EntireColumn.AutoFit
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Max Product Late Count"
    Range("G6").Select
    Columns("G:G").EntireColumn.AutoFit
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "=MAX(R[1]C[-6]:R[90]C[-6])"
    Range("H2").Select
    ActiveWindow.SmallScroll Down:=-1
    Range("G3").Select
    ActiveCell.FormulaR1C1 = "Products that were shipped late most of the time"
    Range("G6").Select
    ActiveCell.Formula2R1C1 = _
        "=FILTER(R[-4]C[-6]:R[194]C[-5],R[-4]C[-5]:R[194]C[-5]=R[-5]C[1])"
    Range("I6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISBLANK(RC[-2]),"""",INDEX(productsCopy!R[-4]C[-7]:R[105]C[-7],MATCH(Q02B!RC[-2],productsCopy!R[-4]C[-8]:R[105]C[-8],0),1))"
    Range("I6").Select
    Selection.AutoFill Destination:=Range("I6:I46")
    Range("I6:I46").Select
    Range("G48").Select
    ActiveCell.FormulaR1C1 = "Max Emp Late Count"
    Range("H48").Select
    ActiveCell.FormulaR1C1 = "=MAX(R[-46]C[-3]:R[-39]C[-3])"
    Range("H49").Select
    Range("G51").Select
    ActiveCell.Formula2R1C1 = _
        "=FILTER(R[-49]C[-3]:R[-42]C[-2],R[-49]C[-2]:R[-42]C[-2]=R[-3]C[1])"
    Sheets("Q02B").Select
    
    
    
    
        Range("I51").Select
    ActiveCell.FormulaR1C1 = _
        "=INDEX(employeesCopy!R2C3:R24C3,MATCH(Q02B!RC[-2],employeesCopy!R2C1:R24C1,0),1)"
    Range("I51").Select
    Selection.AutoFill Destination:=Range("I51:I56"), Type:=xlFillDefault
    Range("I51:I56").Select
    Range("J51").Select
    

    Range("G5").Select
    ActiveCell.FormulaR1C1 = "Product ID"
    Range("H5").Select
    ActiveCell.FormulaR1C1 = "Late Shipped"
    Range("I5").Select
    ActiveCell.FormulaR1C1 = "Product Name"
    Range("G9").Select
    Range("G50").Select
    ActiveCell.FormulaR1C1 = "Employee ID"
    Range("H50").Select
    ActiveCell.FormulaR1C1 = "Late Shipped Count"
    Range("I50").Select
    ActiveCell.FormulaR1C1 = "First Name"
    Range("I51").Select
    
    'Dashboard For Question 2
    
    Sheets("Dashboard").Select
    Range("B35").Select
    ActiveCell.FormulaR1C1 = "Late Shipment Details"
    Range("B37").Select
    ActiveCell.FormulaR1C1 = "Product that was frequently shipped late"
    Range("C39").Select
    ActiveCell.Formula2R1C1 = "=Q02B!R[-34]C[4]:R[-31]C[6]"
    Range("K37").Select
    ActiveCell.FormulaR1C1 = "Employee whose orders were frequently shipped late"
    Range("K39").Select
    ActiveCell.Formula2R1C1 = "=Q02B!R[11]C[-4]:R[17]C[-2]"
    Range("B35").Select
    Selection.Font.Bold = True
    Range("B37").Select
    Selection.Font.Bold = True
    Range("K37").Select
    Selection.Font.Bold = True
    Range("C39:E39").Select
    Selection.Font.Bold = True
    Range("K39:M39").Select
    Selection.Font.Bold = True
    Range("O40").Select
    
    'Question 03 Preparation
    
    Sheets("YearlyOrder").Select
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Shipped"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNUMBER(RC[-7]),""Yes"",IF(ISBLANK(RC[-7]),""NA"",""No""))"
    Range("K2").Select
    Selection.AutoFill Destination:=Range("K2:K327")
    Range("K2:K327").Select
    Range("M16").Select
    Sheets.Add After:=ActiveSheet
    Sheets(11).Select
    Sheets(11).Name = "YearlyOrder02"
    Range("A1").Select
    ActiveCell.Formula2R1C1 = "=YearlyOrder!RC:RC[10]"
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "Issues"
    Range("A2").Select
    ActiveCell.Formula2R1C1 = _
        "=FILTER(YearlyOrder!RC:R[325]C[10],YearlyOrder!RC:R[325]C>1)"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-6]=""NULL"",""No"",IF(ISBLANK(RC[-8]),""NA"",""Yes""))"
    Range("L2").Select
    Selection.AutoFill Destination:=Range("L2:L250")
    Range("L2:L65").Select
    Sheets("Q02B").Select
    Sheets.Add After:=ActiveSheet
    Sheets(15).Select
    Sheets(15).Name = "Q03"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "All Orders that are not shipped yet"
    Range("A2").Select
    ActiveCell.Formula2R1C1 = "=YearlyOrder02!R[-1]C:R[-1]C[11]"
    Range("A3").Select
    ActiveCell.Formula2R1C1 = _
        "=FILTER(YearlyOrder02!R[-1]C:R[62]C[11],YearlyOrder02!R[-1]C[10]:R[62]C[10]=""No"")"
    Range("N1").Select
    ActiveCell.FormulaR1C1 = "All Orders having issues or disputes"
    Range("N2").Select
    ActiveCell.Formula2R1C1 = "=YearlyOrder02!R[-1]C[-13]:R[-1]C[-2]"
    Range("N3").Select
    ActiveCell.Formula2R1C1 = _
        "=FILTER(YearlyOrder02!R[-1]C[-13]:R[247]C[-2],YearlyOrder02!R[-1]C[-2]:R[247]C[-2]=""Yes"")"
    Range("N4").Select
    ActiveWindow.ScrollColumn = 2
    Range("L16").Select
    
    
    
    
    'Dashboard for Question 3

    Sheets("Dashboard").Select
    Range("B49").Select
    ActiveCell.FormulaR1C1 = "All Orders that are not shipped yet"
    Range("B49").Select
    Selection.Font.Bold = True
    Range("B51").Select
    ActiveCell.Formula2R1C1 = "='Q03'!R[-49]C:R[-40]C[10]"
    Range("B51:L51").Select
    Selection.Font.Bold = True
    Columns("H:H").ColumnWidth = 8.01
    Columns("G:G").ColumnWidth = 20.01
    Range("B72").Select
    ActiveCell.FormulaR1C1 = "All Orders having Issues or disputes"
    Range("B72").Select
    Selection.Font.Bold = True
    Range("B74").Select
    ActiveCell.Formula2R1C1 = "='Q03'!R[-72]C[12]:R[-17]C[23]"
    Range("B75").Select
    Range("B74:M74").Select
    Selection.Font.Bold = True
    Range("K71").Select
    
    'Preparation for Question 4
    
        Sheets("YearlyOrder02").Select
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "Month"
    Range("M2").Select
    ActiveCell.FormulaR1C1 = "=TEXT(RC[-11],""MMM"")"
    Range("M2").Select
    Selection.AutoFill Destination:=Range("M2:M250"), Type:=xlFillDefault
    Range("M2:M250").Select
    Range("J254").Select
    Sheets("Q03").Select
    Sheets.Add After:=ActiveSheet
    Sheets(16).Select
    Sheets(16).Name = "Q04"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Sales Month"
    Range("A2").Select
    ActiveCell.Formula2R1C1 = "=UNIQUE(YearlyOrder02!RC[12]:R[248]C[12])"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Total Revenue"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = _
        "=SUMIF(YearlyOrder02!R2C13:R250C13,'Q04'!RC[-1],YearlyOrder02!R2C8:R250C8)"
    Range("B2").Select
    Selection.AutoFill Destination:=Range("B2:B13")
    Range("B2:B13").Select
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Total Revenue difference with mean"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Average"
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(R2C2:R13C2)"
    Range("D2").Select
    Selection.AutoFill Destination:=Range("D2:D13")
    Range("D2:D13").Select
    Range("C2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-1]-RC[1]"
    Range("C2").Select
    Selection.AutoFill Destination:=Range("C2:C13")
    Range("C2:C13").Select
    Range("C1").Select
    Columns("C:C").EntireColumn.AutoFit
    
    'Chart for Q04
    
    Range("A1:D13").Select
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range("'Q04'!$A$1:$D$13")
    ActiveChart.FullSeriesCollection(1).ChartType = xlColumnClustered
    ActiveChart.FullSeriesCollection(1).AxisGroup = 1
    ActiveChart.FullSeriesCollection(2).ChartType = xlColumnClustered
    ActiveChart.FullSeriesCollection(2).AxisGroup = 1
    ActiveChart.FullSeriesCollection(3).ChartType = xlLine
    ActiveChart.FullSeriesCollection(3).AxisGroup = 1
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "Monthly Revenue Chart"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "Monthly Revenue Chart"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 21).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 15).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 14
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Spacing = 0
        .Strike = msoNoStrike
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(16, 6).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 14
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Spacing = 0
        .Strike = msoNoStrike
    End With
    Range("F2").Select
    
        
    
    'Dashboard For Question 4
    
    Sheets("Dashboard").Select
    Sheets.Add After:=ActiveSheet
    Sheets(2).Select
    Sheets(2).Name = "Dashboard2"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "Monthly Revenue Figures"
    Range("B4").Select
    ActiveCell.Formula2R1C1 = "='Q04'!R[-3]C[-1]:R[9]C[2]"
    Range("H4").Select
    
    'Chart for Dashboard Question 4
    
    Range("B4:E17").Select
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range("'Dashboard2'!$B$4:$E$17")
    ActiveChart.FullSeriesCollection(1).ChartType = xlColumnClustered
    ActiveChart.FullSeriesCollection(1).AxisGroup = 1
    ActiveChart.FullSeriesCollection(2).ChartType = xlColumnClustered
    ActiveChart.FullSeriesCollection(2).AxisGroup = 1
    ActiveChart.FullSeriesCollection(3).ChartType = xlLine
    ActiveChart.FullSeriesCollection(3).AxisGroup = 1
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "Monthly Revenue Chart"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "Monthly Revenue Chart"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 21).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 15).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 14
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Spacing = 0
        .Strike = msoNoStrike
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(16, 6).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 14
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Spacing = 0
        .Strike = msoNoStrike
    End With
    Range("F2").Select
    
    'Preparation for Question 5
    
    Sheets("paymentsCopy").Select
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=IF(YEAR(RC[-2])=YearlyOrder!R2C14,""Yes"",""No"")"
    Range("E2").Select
    Selection.AutoFill Destination:=Range("E2:E500")
    Range("E2:E274").Select
    ActiveWindow.SmallScroll Down:=0
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Current Year"
    Range("I6").Select
    
    Sheets("paymentsCopy").Select
    Range("G1").Select
    Application.CutCopyMode = False
    ActiveCell.Formula2R1C1 = "=RC[-6]:RC[-2]"
    Range("G2").Select
    ActiveCell.Formula2R1C1 = _
        "=FILTER(RC[-6]:R[498]C[-2],RC[-2]:R[498]C[-2]=""Yes"")"
    Range("G3").Select
    Range("N11").Select
    
    
    Sheets("Q04").Select
    Sheets.Add After:=ActiveSheet
    Sheets(18).Select
    Sheets(18).Name = "Q05"
    Range("A3").Select
    ActiveCell.Formula2R1C1 = "=UNIQUE(paymentsCopy!R[-1]C[6]:R[97]C[6])"
    Range("A4").Select
    Range("B3").Select
    Range("C3").Select
    ActiveCell.FormulaR1C1 = "=RANK.EQ(RC[-1],R3C2:R34C2)"
    Range("C3").Select
    Selection.AutoFill Destination:=Range("C3:C34")
    Range("C3:C34").Select
    Range("B3").Select
    ActiveCell.FormulaR1C1 = _
        "=SUMIF(paymentsCopy!R2C7:R100C7,'Q05'!RC[-1],paymentsCopy!R2C10:R100C10)"
    Range("B3").Select
    Selection.AutoFill Destination:=Range("B3:B34")
    Range("B3:B34").Select
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "Cust ID"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "Payment"
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "Rank"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Customer Pyment for current Year"
    Range("A2").Select
    
    Sheets("Q05").Select
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Sort Customers according to the payments"
    Range("E2").Select
    Application.CutCopyMode = False
    ActiveCell.Formula2R1C1 = "=RC[-4]:RC[-2]"
    Range("E3").Select
    ActiveCell.Formula2R1C1 = "=SORT(RC[-4]:R[47]C[-2],3,1)"
    Range("E4").Select
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "Customer Name"
    Range("H3").Select
    Columns("H:H").EntireColumn.AutoFit
    ActiveCell.FormulaR1C1 = _
        "=INDEX(CustomerCopy!R2C2:R123C2,MATCH('Q05'!RC[-3],CustomerCopy!R2C2:R123C2,0),1)"
    Range("H3").Select
    ActiveCell.FormulaR1C1 = _
        "=INDEX(CustomerCopy!R2C2:R123C2,MATCH('Q05'!RC[-3],CustomerCopy!R2C1:R123C1,0),1)"
    Range("H3").Select
    Selection.AutoFill Destination:=Range("H3:H50")
    Range("H3:H50").Select
    Range("K6").Select
    Columns("H:H").EntireColumn.AutoFit
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Highest Valued Customers"
    Range("J2").Select
    Application.CutCopyMode = False
    ActiveCell.Formula2R1C1 = "=RC[-5]:R[10]C[-2]"
    Range("J3").Select
    Columns("M:M").EntireColumn.AutoFit
    Range("J2:M2").Select
    Selection.Font.Bold = True
    Range("J1").Select
    Selection.Font.Bold = True
    Range("E1").Select
    Selection.Font.Bold = True
    Range("E2:H2").Select
    Selection.Font.Bold = True
    Range("A2:C2").Select
    Selection.Font.Bold = True
    Range("A1:C1").Select
    Selection.Font.Bold = True
    Range("M19").Select
    
    
    Sheets("Q05").Select
    Range("O1").Select
    ActiveCell.FormulaR1C1 = "Largest Order Value"
    Range("Q1").Select
    ActiveCell.FormulaR1C1 = "=MAX(YearlyOrder!R[1]C[-9]:R[326]C[-9])"
    Range("O3").Select
    ActiveCell.FormulaR1C1 = "Order number of Max order Value"
    Range("S3").Select
    ActiveCell.FormulaR1C1 = _
        "=INDEX(YearlyOrder!R[-1]C[-18]:R[324]C[-18],MATCH('Q05'!R[-2]C[-2],YearlyOrder!R[-1]C[-11]:R[324]C[-11],0),1)"
    Range("O5").Select
    ActiveCell.FormulaR1C1 = "Order Details of Largest Order Value"
    Range("O7").Select
    ActiveCell.Formula2R1C1 = _
        "=FILTER(orderdetailsCopy!R[-5]C[-14]:R[4993]C[-8],orderdetailsCopy!R[-5]C[-14]:R[4993]C[-14]='Q05'!R[-4]C[4])"
    Range("O6").Select
    ActiveCell.FormulaR1C1 = "Order No"
    Range("P6").Select
    ActiveCell.FormulaR1C1 = "Product No"
    Range("Q6").Select
    ActiveCell.FormulaR1C1 = "Qty"
    Range("R6").Select
    ActiveCell.FormulaR1C1 = "Price each"
    Range("S6").Select
    ActiveCell.FormulaR1C1 = "Order Line"
    Range("T6").Select
    ActiveCell.FormulaR1C1 = "Total Amount"
    Range("U6").Select
    ActiveCell.FormulaR1C1 = "Late Shipped"
    Range("V6").Select
    Columns("P:P").EntireColumn.AutoFit
    Columns("T:T").EntireColumn.AutoFit
    Columns("U:U").EntireColumn.AutoFit
    Range("O1").Select
    Selection.Font.Bold = True
    Range("O3").Select
    Selection.Font.Bold = True
    Range("O5").Select
    Selection.Font.Bold = True
    Range("O6:U6").Select
    Selection.Font.Bold = True
    Range("M22").Select

    'Question 5 Dashboard
    
    Sheets("Dashboard2").Select
    Range("B25").Select
    ActiveCell.FormulaR1C1 = "Details of Largest Value Order"
    Range("B29").Select
    ActiveCell.FormulaR1C1 = "='Q05'!R[-26]C[13]"
    Range("E29").Select
    ActiveCell.FormulaR1C1 = "='Q05'!R[-26]C[14]"
    Range("B31").Select
    ActiveCell.Formula2R1C1 = "='Q05'!R[-25]C[13]:R[-9]C[19]"
    Range("B25").Select
    Selection.Font.Bold = True
    Range("B27").Select
    ActiveCell.FormulaR1C1 = "='Q05'!R[-26]C[13]"
    Range("D27").Select
    ActiveCell.FormulaR1C1 = "='Q05'!R[-26]C[13]"
    Range("B27").Select
    Selection.Font.Bold = True
    Range("B29").Select
    Selection.Font.Bold = True
    Range("B31:H31").Select
    Selection.Font.Bold = True
    Range("J32").Select
    
    Range("K25").Select
    ActiveCell.FormulaR1C1 = "Highest Value Customer"
    Range("P27").Select
    ActiveCell.FormulaR1C1 = "='Q05'!R[-24]C[-3]"
    Range("P27").Select
    Selection.ClearContents
    Range("M29").Select
    ActiveCell.Formula2R1C1 = "='Q05'!R[-27]C[-3]#"
    Range("M27").Select
    Selection.Font.Bold = True
    Range("M29:P29").Select
    Selection.Font.Bold = True
    Range("K25").Select
    Selection.Font.Bold = True
    Range("K26").Select
    
    
    'jerlkdvmd
    
    
    Sheets("orderdetailsCopy").Select
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Years"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Current Years"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNUMBER(MATCH(RC[-7],YearlyOrder!RC[-7]:R[150]C[-7],0)),""Yes"",""No"")"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNUMBER(MATCH(RC[-7],YearlyOrder!R2C1:R152C1,0)),""Yes"",""No"")"
    Range("H2").Select
    Selection.AutoFill Destination:=Range("H2:H2997")
    Range("H2:H2997").Select
    Range("I2").Select
    Sheets("YearlyOrder02").Select
    Sheets("Q05").Select
    Sheets.Add After:=ActiveSheet
    Sheets(19).Select
    Sheets(19).Name = "Q06"
    Range("A2").Select
    ActiveCell.Formula2R1C1 = _
        "=FILTER(orderdetailsCopy!RC:R[2995]C[7],orderdetailsCopy!RC[7]:R[2995]C[7]=""Yes"")"
    Range("A1").Select
    ActiveCell.Formula2R1C1 = "=orderdetailsCopy!RC:RC[7]"
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Product ID"
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "Total Sales Revenue"
    Range("K2").Select
    ActiveCell.Formula2R1C1 = "=UNIQUE(RC[-9]:R[1420]C[-9])"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(R2C2:R1422C2,RC[-1],R2C6:R1422C6)"
    Range("L2").Select
    Selection.AutoFill Destination:=Range("L2:L110")
    Range("L2:L110").Select
    Columns("L:L").EntireColumn.AutoFit
    Range("O2").Select
    ActiveCell.FormulaR1C1 = "Product with highest sales revenue"
    Range("S2").Select
    ActiveCell.FormulaR1C1 = ""
    Range("O4").Select
    ActiveCell.FormulaR1C1 = "Max Amount"
    Range("Q4").Select
    ActiveCell.FormulaR1C1 = "=MAX(R[-2]C[-5]:R[106]C[-5])"
    Range("S2").Select
    ActiveCell.FormulaR1C1 = _
        "=INDEX(RC[-8]:R[108]C[-7],MATCH(R[2]C[-2],RC[-7]:R[108]C[-7],0),1)"
    Range("O6").Select
    ActiveCell.FormulaR1C1 = "Product Details"
    Range("O8").Select
    ActiveCell.FormulaR1C1 = "Product Name"
    Range("P8").Select
    ActiveCell.FormulaR1C1 = "Product ID"
    Range("Q8").Select
    ActiveCell.FormulaR1C1 = "Total Qty Sold"
    Range("R8").Select
    ActiveCell.FormulaR1C1 = "Total Sales Revenue"
    Range("S8").Select
    Columns("O:O").EntireColumn.AutoFit
    Columns("P:P").EntireColumn.AutoFit
    Columns("Q:Q").EntireColumn.AutoFit
    Columns("R:R").EntireColumn.AutoFit
    Range("O9").Select
    ActiveCell.FormulaR1C1 = _
        "=INDEX(productsCopy!R[-8]C[-13]:R[102]C[-13],MATCH('Q06'!R[-7]C[4],productsCopy!R[-7]C[-14]:R[102]C[-14],0),1)"
    Range("O10").Select
    ActiveWindow.Zoom = 89
    Range("O9").Select
    ActiveCell.FormulaR1C1 = _
        "=INDEX(productsCopy!R[-7]C[-13]:R[102]C[-13],MATCH('Q06'!R[-7]C[4],productsCopy!R[-7]C[-14]:R[102]C[-14],0),1)"
    Range("P9").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[-7]C[3]"
    Range("Q9").Select
    ActiveCell.FormulaR1C1 = _
        "=SUMIF(R[-7]C[-15]:R[1413]C[-15],RC[-1],R[-7]C[-14]:R[1413]C[-14])"
    Range("R9").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[-5]C[-1]"
    Range("A1:H1").Select
    Selection.Font.Bold = True
    Range("K1:L1").Select
    Selection.Font.Bold = True
    Range("O2").Select
    Selection.Font.Bold = True
    Range("O4").Select
    Selection.Font.Bold = True
    Range("O6").Select
    Selection.Font.Bold = True
    Range("O8:R8").Select
    Selection.Font.Bold = True
    
    'Dashboard for Question 6
    
    Sheets("Dashboard2").Select
    Range("B51").Select
    ActiveCell.FormulaR1C1 = "='Q06'!R[-49]C[13]"
    Range("D51").Select
    ActiveCell.FormulaR1C1 = "='Q06'!R[-49]C[13]"
    Range("B53").Select
    ActiveCell.FormulaR1C1 = "='Q06'!R[-49]C[13]"
    Range("D53").Select
    ActiveCell.FormulaR1C1 = "='Q06'!R[-49]C[13]"
    Range("B55").Select
    ActiveCell.Formula2R1C1 = "='Q06'!R[-47]C[13]:R[-46]C[16]"
    Range("B49").Select
    ActiveCell.FormulaR1C1 = "Product Details with highest Sales Revenue"
    Range("B49").Select
    Selection.Font.Bold = True
    Range("B51").Select
    Selection.Font.Bold = True
    Range("B53").Select
    Selection.Font.Bold = True
    Range("B55:E55").Select
    Selection.Font.Bold = True
    Range("G50").Select
    Columns("B:B").EntireColumn.AutoFit
    Columns("C:C").EntireColumn.AutoFit
    Columns("D:D").EntireColumn.AutoFit
    Columns("E:E").EntireColumn.AutoFit
    
    'Preparation of question 7 Summary Page
    
    Sheets("ordersCopy").Select
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "Total Revenue"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = _
        "=SUMIF(orderdetailsCopy!R2C1:R2997C1,RC[-8],orderdetailsCopy!R2C6:R2997C6)"
    Range("I2").Select
    Selection.AutoFill Destination:=Range("I2:I327")
    Range("I2:I327").Select
    Range("K5").Select
    Sheets("Q06").Select
    Sheets.Add After:=ActiveSheet
    Sheets(20).Select
    Sheets(20).Name = "Summary_Sheet"
    Range("J21").Select
    
    Sheets("Summary_Sheet").Select
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "Total Revenue generated per year"
    Range("B2").Select
    Selection.Font.Bold = True
    Range("B4").Select
    ActiveCell.FormulaR1C1 = "Year"
    Range("C4").Select
    ActiveCell.FormulaR1C1 = "Revenue"
    Range("B4:C4").Select
    Selection.Font.Bold = True
    Range("B5").Select
    ActiveCell.FormulaR1C1 = "2003"
    Range("B6").Select
    ActiveCell.FormulaR1C1 = "2004"
    Range("B7").Select
    ActiveCell.FormulaR1C1 = "2005"
    Range("C5").Select
    ActiveCell.FormulaR1C1 = _
        "=SUMIF(ordersCopy!R2C8:R400C8,'Summary_Sheet'!RC[-1],ordersCopy!R2C9:R400C9)"
    Range("C5").Select
    Selection.AutoFill Destination:=Range("C5:C7")
    Range("C5:C7").Select
    Range("B13").Select
    ActiveCell.FormulaR1C1 = "Product Sold on High Quantity"
    Range("B13").Select
    Selection.Font.Bold = True
    Range("B15").Select
    ActiveCell.FormulaR1C1 = "Sr No"
    Range("C15").Select
    ActiveCell.FormulaR1C1 = "Product Id"
    Range("D15").Select
    ActiveCell.FormulaR1C1 = "Qty"
    Range("B30").Select
    ActiveCell.FormulaR1C1 = "Product with High Revenues"
    Range("B30").Select
    Selection.Font.Bold = True
    Range("B32").Select
    ActiveCell.FormulaR1C1 = "Sr No"
    Range("C32").Select
    ActiveCell.FormulaR1C1 = "Product ID"
    Range("D32").Select
    ActiveCell.FormulaR1C1 = "Total Revenue"
    Range("P2").Select
    ActiveCell.FormulaR1C1 = "Revenue Generated each month for selected year"
    Range("P4").Select
    ActiveCell.FormulaR1C1 = "Year"
    Range("Q4").Select
    ActiveCell.FormulaR1C1 = "=YearlyOrder!R[-2]C[-3]"
    Range("P2").Select
    Selection.Font.Bold = True
    Range("P4").Select
    Selection.Font.Bold = True
    Range("Q4").Select
    Selection.Font.Bold = True
    Range("P6").Select
    ActiveCell.FormulaR1C1 = "Month"
    Range("Q6").Select
    ActiveCell.FormulaR1C1 = "Revenue"
    Range("P7").Select
    ActiveCell.Formula2R1C1 = "='Q04'!R[-5]C[-15]:R[6]C[-14]"
    Range("P6:Q6").Select
    Selection.Font.Bold = True
    Range("P22").Select
    ActiveCell.FormulaR1C1 = "Employees with high Sales Values"
    Range("P22").Select
    Selection.Font.Bold = True
    Range("P24").Select
    ActiveCell.FormulaR1C1 = "Emp ID"
    Range("P24").Select
    Selection.Font.Bold = True
    Range("F16").Select

    Sheets("Q06").Select
    Range("O15").Select
    ActiveCell.FormulaR1C1 = "Product Sold on high volumes"
    Range("P15").Select
    ActiveCell.FormulaR1C1 = "=MAX(R[-13]C[-13]:R[1407]C[-13])-1"
    Range("O17").Select
    ActiveCell.FormulaR1C1 = "Product ID"
    Range("P17").Select
    ActiveCell.FormulaR1C1 = "Qty"
    Range("O18").Select
    ActiveCell.Formula2R1C1 = _
        "=FILTER(R[-16]C[-13]:R[1404]C[-12],R[-16]C[-12]:R[1404]C[-12]>R[-3]C[1])"
    Range("O19").Select
    
    Range("R17").Select
    ActiveCell.FormulaR1C1 = "Product ID"
    Range("S17").Select
    ActiveCell.FormulaR1C1 = "Total Sales Revenue"
    Range("R18").Select
    ActiveCell.Formula2R1C1 = "=SORT(R[-16]C[-7]:R[92]C[-6],2,-1)"
    Range("R17").Select
    
    Sheets("Summary_Sheet").Select
    Range("C16").Select
    ActiveCell.Formula2R1C1 = "='Q06'!R[2]C[12]:R[10]C[13]"
    Range("B16").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("B17").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("B18").Select
    ActiveCell.FormulaR1C1 = "3"
    Range("B19").Select
    ActiveCell.FormulaR1C1 = "4"
    Range("B20").Select
    ActiveCell.FormulaR1C1 = "5"
    Range("B21").Select

    Range("C33").Select
    ActiveCell.Formula2R1C1 = "='Q06'!R[-15]C[15]:R[-11]C[16]"
    Range("B33").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("B34").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("B35").Select
    ActiveCell.FormulaR1C1 = "3"
    Range("B36").Select
    ActiveCell.FormulaR1C1 = "4"
    Range("B37").Select
    ActiveCell.FormulaR1C1 = "5"
    Range("B38").Select
    
    
    Range("P25").Select
    ActiveCell.Formula2R1C1 = "=Dashboard!R[-9]C[-5]#"
    Range("P25:U25").Select
    Selection.Font.Bold = True
    Range("R23").Select
    
    
    Range("B5:C7").Select
    ' ActiveSheet.Shapes.AddDashboardChart2(201, xlColumnClustered).Select
    ' ActiveChart.SetSourceData Source:=Range("Summary_Sheet!$B$5:$C$7")
    Range("F4").Select
    
    ' ActiveSheet.ChartObjects(1).Activate
    ' ActiveSheet.ChartObjects(1).Activate
    ' ActiveSheet.Shapes(1).IncrementLeft -136.8
    ' ActiveSheet.Shapes(1).IncrementTop -60
    ' Range("G1").Select
    
    Range("B32:D37").Select
    ActiveSheet.Shapes.AddChart2(251, xlPie).Select
    ActiveChart.SetSourceData Source:=Range("Summary_Sheet!$B$32:$D$37")
    ' ActiveSheet.Shapes(2).IncrementLeft -138.6
    ' ActiveSheet.Shapes(2).IncrementTop 67.8
    
    Sheets("Summary_Sheet").Move Before:=Sheets("Dashboard")
    
    
    
    

    


End Sub
