'05-01-17 written by Ben Currer-Burgess
'Macro for stripping out shopify orders that should not be shipped via royal mail and distinguishing between different royal mail shipping options

'This Macro is designed to work with Shopify and will work on an orders_export.csv file from shopify
'The resulting file can be saved and uploaded to Royal mail DMO.

'on initial setup you will need to:
'	- set up royal mail DMO mapping to match the resulting file
'		this will be easier to do if you run this macro on an order.csv first so you have a visual reference to work from
'	- edit this macro, wherever you find a # you need to replace that string with the correct text for your situation
'   


Sub CleanShopifyCsv()

Application.ScreenUpdating = False
'=======================================
' Section 1 - for checking orders for large products and setting the shipping option as blank
'=======================================
    ' Specific products to check for and remove because they do not get sent with royal mail you can add as many as you need, just follow the sequence
    [BO2].Value = "#Exact product name 1"
    [BO3].Value = "#Exact product name 2"
    [BO4].Value = "#Exact product name 3"

        
    Dim pRange As Range
    Dim nShip As Range
    'look at list of products not sent by normal courier
    Set pRange = Range("BO2")
    ' active cell = column BO
    pRange.Select
    ' when selected cell is not blank
    While ActiveCell.Value <> ""
    ' Select Line item... ACTIVE CELL = column R
    Range("R2").Select
    ' while selected cell is not blank
    While ActiveCell.Value <> ""
    ' look at order numbers (column A) to determine where an order starts and ends
    If ActiveCell.Offset(0, -17).Value <> ActiveCell.Offset(-1, -17).Value Then
    ' set nship as shipping (Column O) in correct cell
    Set nShip = ActiveCell.Offset(0, -3)

    End If
    ' If current lineitem is in the list of products not sent by normal courier
    If ActiveCell.Value = pRange.Value Then
    'set the shipping method as whatever is in column BP
    nShip.Value = pRange.Offset(0, 1).Value
    ' then Select next product in column R
    ActiveCell.Offset(1, 0).Select

    Else
    'otherwise select next product in column R
    ActiveCell.Offset(1, 0).Select

    End If

    Wend
    'Look at the next item in the list of products not sent by normal courier
    Set pRange = pRange.Offset(1, 0)
    pRange.Select

    Wend
'=======================================
' End of section 1
' Section 2 - clean up sheet
'=======================================
    ' Delete the information that does not relate to shipping
    Range("BO:BO").Delete Shift:=xlToLeft
    Range("B:N").Delete Shift:=xlToLeft
    Range("C:U").Delete Shift:=xlToLeft
    Range("E:E").Delete Shift:=xlToLeft
    Range("K:AG").Delete Shift:=xlToLeft
    ' Sort the sheet based on shipping method (column B)
    Dim oneRange As Range
    Dim aCell As Range
    Set oneRange = Range("A:J")
    Set aCell = Range("B:B")
    oneRange.Sort Key1:=aCell, Order1:=xlAscending, Header:=xlYes
    ' Add in new column headers 
    [K1].Value = "Service Ref"
    [L1].Value = "Service"
    [M1].Value = "Items"
    [N1].Value = "Weight"
    [O1].Value = "Format"
    [P1].Value = "Enhancement"
'------------------------------
' Delete empty rows section 2.2
'------------------------------
    Dim Firstrow As Long
    Dim Lastrow As Long
    Dim Lrow As Long
    Dim CalcMode As Long

    With Application
        CalcMode = .Calculation
        .Calculation = xlCalculationManual
    End With

    'use the ActiveSheet
    With ActiveSheet
    .Select

        'Set the first and last row to loop through
    Firstrow = 2
    Lastrow = .Cells(.Rows.Count, "A").End(xlUp).Row
    
   'loop from Lastrow to Firstrow (bottom to top)
    For Lrow = Lastrow To Firstrow Step -1
        
   'check the values in the B column
    With .Cells(Lrow, "B")
    If Not IsError(.Value) Then
            
            ' This will delete each row where column B is blank or has a non royal mail shipping method
            If  .Value = "" Or _
                .Value = "#exact shipping method name that is not royal mail 1" Or _
                .Value = "#exact shipping method name that is not royal mail 2" _
                Then .EntireRow.Delete
            End If
    End With
    Next Lrow
    End With
    
    With Application
        .Calculation = CalcMode
    End With
'------------------------------
' End of section 2.2
'------------------------------
    ' Adjust width of shipping column
    Columns(2).AutoFit
'=======================================
' End section 2
' Section 3 - Add in royal mail content
'=======================================
                '----------------------------
                'All standard options
                '----------------------------
    Lrow = Range("A" & Rows.Count).End(xlUp).Row
    Set MR = Range("A2:A" & Lrow)
    For Each Cell In MR
    If Not Cell.Value = "" Then
    'service ref (from royal mail)
    Cell.Offset(0, 10).Value = "#Default service reference code here"
    'items 
    Cell.Offset(0, 12).Value = "#Default quantity"
    'weight
    Cell.Offset(0, 13).Value = "#Default weight in grams"
    End If
    Next

    'Auto fill royal mail shipping options based on shopify shipping options/ countries
    Lrow = Range("B" & Rows.Count).End(xlUp).Row
    Set MR = Range("B2:B" & Lrow)
    For Each Cell In MR
            '----------------------------
            'For UK postage
            '----------------------------
    'Standard shipping options from shopify
    If Cell.Value = "#Exact royal mail shipping method from shopify 1" Or _
        Cell.Value = "#Exact royal mail shipping method from shopify 2" Or _
        Cell.Value = "#Exact royal mail shipping method from shopify 3" Then
        'service
        Cell.Offset(0, 10).Value = "#Exact royal mail service reference for the above shipping methods"
        'Format
        Cell.Offset(0, 13).Value = "#Exact royal mail service format for the above shipping methods"
        'Enhancement (if required)
        cell.Offset(0, 14).Value = "#Exact royal mail service enhancement (if required)"
            '----------------------------
            'For standard international postage
            '----------------------------
    'Standard international shipping options from shopify
    ElseIf Cell.Value = "#Exact royal mail shipping method from shopify 4" Or _
        Cell.Value = "#Exact royal mail shipping method from shopify 5" Or _
        Cell.Value = "#Exact royal mail shipping method from shopify 6" Then
        'service
        Cell.Offset(0, 10).Value = "#Exact royal mail service reference for the above shipping methods"
        'Format
        Cell.Offset(0, 13).Value = "#Exact royal mail service format for the above shipping methods"
        'Enhancement (if required)
        cell.Offset(0, 14).Value = "#Exact royal mail service enhancement (if required)"
    End If
    Next
            '----------------------------
            'For international exceptions
            '----------------------------
    Lrow = Range("J" & Rows.Count).End(xlUp).Row
    Set MR = Range("J2:J" & Lrow)
    For Each Cell In MR

    ' List of countries that don't take signed for but do take tracked
    If Cell.Value = "AD" Or Cell.Value = "AT" Or _
        Cell.Value = "BE" Or Cell.Value = "CA" Or _
        Cell.Value = "IC" Or Cell.Value = "HR" Or _
        Cell.Value = "DK" Or Cell.Value = "FI" Or _
        Cell.Value = "FR" Or Cell.Value = "DE" Or _
        Cell.Value = "HK" Or Cell.Value = "HU" Or _
        Cell.Value = "IS" Or Cell.Value = "IE" Or _
        Cell.Value = "IT" Or Cell.Value = "LV" Or _
        Cell.Value = "LI" Or Cell.Value = "LT" Or _
        Cell.Value = "LU" Or Cell.Value = "MY" Or _
        Cell.Value = "MT" Or Cell.Value = "NL" Or _
        Cell.Value = "NZ" Or Cell.Value = "PL" Or _
        Cell.Value = "PT" Or Cell.Value = "SM" Or _
        Cell.Value = "SG" Or Cell.Value = "SK" Or _
        Cell.Value = "SI" Or Cell.Value = "KR" Or _
        Cell.Value = "ES" Or Cell.Value = "SE" Or _
        Cell.Value = "CH" Or Cell.Value = "TR" Or _
        Cell.Value = "US" Or Cell.Value = "VA" Or _
        Cell.Value = "AU" Then
        'Service to tracked, if you do not have a tracked service then set this to blank: ""
        Cell.Offset(0, 2).Value = "#Exact code for international tracked"

    'List of countries that don't take signed for OR tracked but do take signed and tracked
    ElseIf Cell.Value = "AR" Or _
        Cell.Value = "BY" Or _
        Cell.Value = "BG" Or _
        Cell.Value = "KH" Or _
        Cell.Value = "KY" Or _
        Cell.Value = "CY" Or _
        Cell.Value = "CZ" Or _
        Cell.Value = "EC" Or _
        Cell.Value = "GI" Or _
        Cell.Value = "GR" Or _
        Cell.Value = "ID" Or _
        Cell.Value = "JP" Or _
        Cell.Value = "MD" Or _
        Cell.Value = "RO" Or _
        Cell.Value = "RS" Or _
        Cell.Value = "TH" Or _
        Cell.Value = "TT" Or _
        Cell.Value = "AE" Then
        'Service to Signed and tracked, if you do not have a signed and tracked service then set this to blank: ""
        Cell.Offset(0, 2).Value = "#Exact code for international signed and tracked"
    End If
    Next
'=======================================
'End of section 3
'Section 4 - file compliance
'=======================================
    'Remove erroneous apostrophes
        With ActiveSheet.UsedRange
            .Value = .Value
        End With

    'Remove all commas to stop csv format breaking
        ExecuteExcel4Macro _
            "FORMULA.REPLACE("","","""",2,1,FALSE,FALSE,,FALSE,FALSE,FALSE,FALSE)"

    'Sort the sheet based on Order Number
    Set oneRange = Range("A:P")
    Set aCell = Range("A:A")
    oneRange.Sort Key1:=aCell, Order1:=xlAscending, Header:=xlYes
'=======================================
'End of section 4
'=======================================
Application.ScreenUpdating = True
MsgBox "Macro complete"
End Sub