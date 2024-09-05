Attribute VB_Name = "M41Platt"
Sub Platt(fname, fpath, XLSpath, XLSname, path, emailmessage)
    ' Reset Variables
    emailmessage = ""
    Fail = 0
    PossibleError = 0
    tempsheetoffset = 0
    docno = ""

    ' Prep Temp Sheet
    Call FormatTempSheet

    ' Check if PDF is Order
    If fname Like "*[0-9][0-9][0-9][0-9]_ORDER_[0-9][0-9][0-9][0-9]*" Or _
       fname Like "ORDER_[0-9][A-Z][0-9][0-9][0-9][0-9][0-9]*" Or _
       fname Like "ORDER_[A-Z][0-9][0-9][0-9][0-9][0-9][0-9]*" Or _
       (fname Like "* [0-9][A-Z][0-9][0-9][0-9][0-9][0-9] *" And Not fname Like "* PAC *") Then
        
        Call Convert_PDF_to_Excel(fname, fpath, XLSpath, XLSname, emailmessage)
        PDFtype = "Order"
        Call ProcessPlattOrderPDF(fname, fpath, XLSpath, XLSname, PDFtype, path, emailmessage)
        Exit Sub
    End If

    ' Check if PDF is an Invoice
    If fname Like "*[0-9][0-9][0-9][0-9]_INVOICE_[0-9][0-9][0-9][0-9]*" Then
        PDFtype = "Invoice"
        Call Convert_PDF_to_Excel(fname, fpath, XLSpath, XLSname, emailmessage)
        Call ProcessPlattInvoicePDF(fname, fpath, XLSpath, XLSname, PDFtype, path, emailmessage)
        Exit Sub
    End If
End Sub
Sub ProcessPlattOrderPDF(fname, fpath, XLSpath, XLSname, PDFtype, path, emailmessage)
    'fname -> Original PDF file name
    'fpath -> Original PDF file path
    'XLSname - > Converted to excel file name
    'XLSpath -> Converted to excel file path
    'PDFtype -> Invoice or Order

    ' Safeguard against getting sent here when doing Submittal webscrape
    If Not UCase(path) Like "*ATTACHMENT*" Then
        'MsgBox "Inside ProcessPlatt Order Ack module"
        Call ProcessPlattInvoicePDF(fname, fpath, XLSpath, XLSname, PDFtype, path, emailmessage)
        Exit Sub
    End If

    ' If order acknowledgement, no need to scrape convert. Immediately route to Webscrape
    If fname Like "ORDER_[0-9][A-Z][0-9][0-9][0-9][0-9]*" Or _
        fname Like "ORDER_[A-Z][0-9][0-9][0-9][0-9][0-9]*" Then
        docno = Replace(fname, "ORDER_", "")
        docno = Replace(docno, ".PDF", "")
        Workbooks(XLSname).Close SaveChanges:=False
        Kill XLSpath
        GoTo docno:
    End If

    ' Platt PDF to Excel Order scrape Macro Data Algorithm
    For xoffset = 1 To 200
        'Assign variable 'line'
        Line = Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset, 0)
        
        'Platt Document number
        If Line Like "*[0-9][A-Z][0-9][0-9][0-9][0-9][0-9]*" Or _
            Line Like "* [A-Z][0-9][0-9][0-9][0-9][0-9][0-9] *" Then
            If docno = "" Then
                docno = Line
                
                ' Trim Platt Document number down to 7 characters
                For Repeat = 1 To Len(docno)
                    If Left(docno, 7) Like "[0-9][A-Z][0-9][0-9][0-9][0-9][0-9]" Or _
                    Left(docno, 7) Like "[A-Z][0-9][0-9][0-9][0-9][0-9][0-9]" Then
                        docno = Left(docno, 7)
                        Exit For
                    Else
                        docno = Right(docno, (Len(docno) - 1))
                    End If
                Next Repeat
                
                ' Error Check Platt document number is 7 characters of correct format
                If docno Like "*[0-9][A-Z][0-9][0-9][0-9][0-9][0-9]*" Or _
                   docno Like "*[A-Z][0-9][0-9][0-9][0-9][0-9][0-9]" Then
                    'do nothing
                Else
                    ' Message User that found Platt Document number doesn't conform
                    MsgBox "Error parsing Platt DocNo, Doc no is" & Chr(13) & docno
                    MsgBox "Freeze"
                    MsgBox "Freeze"
                    MsgBox "Freeze"
                End If
            End If
        End If
        
        ' Order Date
        If Line Like "DATE *" Then
            OrderDate = Mid(Line, 6, 9)
            OrderDate = Replace(OrderDate, "T", "")
            OrderDate = Replace(OrderDate, " ", "")
            'MsgBox "OrderDate=:" & OrderDate & ":"
        End If
        
        ' DECO PO#
        If Line Like "Phone*" And Not Workbooks(XLSname).Sheets(1).Range("A2").Offset(xoffset, 0) _
            Like "Dutton*" Then
            line2 = Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset + 1, 0)
            DecoPO = Left(line2, 24)
            DecoPO = Replace(DecoPO, " ", "")
            'MsgBox "decoPO=:" & decoPO & ":"
        End If
        
        ' Platt Invoice Number (Or, document number since this is an order acknowledgement)
        If Line Like "PO BOX 418759*" Then
            vendorInvoice = Mid(Line, (Len(Line) - 16), 9)
            vendorInvoice = Replace(vendorInvoice, " ", "")
            'MsgBox "VendorInvoice=:" & VendorInvoice & ":"
        End If
        
        ' Shipping and Handling
        If UCase(Line) Like "*SHIP*HANDLING*" Then
            line2 = Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset, 0)
            line2 = Replace(line2, "SHIP/HANDLING", "")
            line2 = Replace(line2, " ", "")
            If Not line2 Like "*.*" Then line2 = line2 & ".00"
            If line2 Like "*.[0-9]" Then line2 = line2 & "0"
            shipping = line2
            'If Line2 <> "" Then MsgBox "Detected Shipping charges of " & shipping
        End If
        
        ' Invoice Total
        If Line Like "END OF ORDER*" Then
            line2 = Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset + 1, 0)
            If Not line2 Like "*.*" Then line2 = line2 & ".00"
            If line2 Like "*.[0-9]" Then line2 = line2 & "0"
            TotalInvoice = line2
            'MsgBox "TotalInvoice=:" & Totalinvoice & ":"
        End If
        
        ' Tax
        If Line Like "END OF ORDER*" Then
            line2 = Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset - 1, 0)
            Tax = Right(line2, 13)
            Tax = Replace(Tax, " ", "")
            'MsgBox "Tax=:" & Tax & ":"
        End If
    Next xoffset
    
    ' End of Macro Data Scrape
    
    ' Begin Line items Scrape
    For xoffset = 3 To 200
        ' Assign variable 'Line'
        Line = Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset, 0)
        
        ' Find excel Line to Parse
        If Line Like "0[0-9][0-9] *" Then
            
            ' Line item description
            itemDesc = Mid(Line, 24, 29)
            'ItemDesc = Replace(ItemDesc, " ", "")
            'MsgBox "ItemDesc=:" & ItemDesc & ":"
            
            ' Check for line item Description error
            If itemDesc = "" Then PossibleError = PossibleError + 1
            
            ' Unit
            Unit = Mid(Line, 53, 2)
            Unit = Replace(Unit, " ", "")
            If Unit = "FT" Then Unit = "EA"
            'MsgBox "Unit=:" & Unit & ":"
            
            ' Quantity
            Quantity = Mid(Line, 58, 7)
            Quantity = Replace(Quantity, " ", "")
            'MsgBox "Quantity=:" & Quantity & ":"
            
            ' Shipped Quantity
            SHIP = Mid(Line, 65, 7)
            SHIP = Replace(SHIP, " ", "")
            'MsgBox "Shipped=:" & Ship & ":"
            
            ' Per Unit Price
            unitprice = Mid(Line, 79, 11)
            unitprice = Replace(unitprice, " ", "")
            If Unit = "C" Then
                If unitprice = "" Then unitprice = 0
                unitprice = unitprice / 100
                Unit = "EA"
            End If
            If Unit = "M" Then
                unitprice = unitprice / 1000
                Unit = "Ea"
            End If
            ' MsgBox "UnitPrice=:" & UnitPrice & ":"
            
            ' Line price total
            lineprice = Right(Line, 10)
            lineprice = Replace(lineprice, " ", "")
            If SHIP = 0 Then lineprice = unitprice * Quantity
            
            ' Write data to thisworkbook Temp Sheet
            'MsgBox "LinePrice=:" & LinePrice & ":"
            ThisWorkbook.Sheets("Temp").Range("P2").Offset(tempsheetoffset, 0) = itemDesc
            ThisWorkbook.Sheets("Temp").Range("Q2").Offset(tempsheetoffset, 0) = Unit
            ThisWorkbook.Sheets("Temp").Range("R2").Offset(tempsheetoffset, 0) = Quantity
            ThisWorkbook.Sheets("Temp").Range("S2").Offset(tempsheetoffset, 0) = unitprice
            ThisWorkbook.Sheets("Temp").Range("T2").Offset(tempsheetoffset, 0) = lineprice
            ThisWorkbook.Sheets("Temp").Range("A2").Offset(tempsheetoffset, 0) = DecoPO
            ThisWorkbook.Sheets("Temp").Range("B2").Offset(tempsheetoffset, 0) = OrderDate
            ThisWorkbook.Sheets("Temp").Range("C2").Offset(tempsheetoffset, 0) = "234"
            ThisWorkbook.Sheets("Temp").Range("AH2").Offset(tempsheetoffset, 0) = Tax
            ThisWorkbook.Sheets("Temp").Range("N2").Offset(tempsheetoffset, 0) = TotalInvoice
            ThisWorkbook.Sheets("Temp").Range("H2").Offset(tempsheetoffset, 0) = vendorInvoice
            ThisWorkbook.Sheets("Temp").Range("B2").Offset(tempsheetoffset, 0) = OrderDate
            tempsheetoffset = tempsheetoffset + 1
        End If
    Next xoffset

    'MsgBox "Done scraping PDF sheet data"

    ' Close and kill XLS
    Workbooks(XLSname).Close SaveChanges:=False
    Kill XLSpath

    ' Error check and common fixes
    Call SelfHealTempPage

    'Check if PO conforms before bothering to Enter
    TargetPO = ThisWorkbook.Sheets("Temp").Range("A2")
             
docno:
' If successfully fetched Platt DocNo, then Webscrape
    If docno <> "" Then
        'MsgBox "Successful read Document Number " & DocNo & Chr(13) & "for " & TargetPO & ", try to acquire the data through webscrape"
        PDFtype = "Order"
        
        Call PlattWebscrape(Fail, fpath, TargetPO, docno, PDFtype, shipping, path, fname, emailmessage)
        
        'Check if failed to webscrape
        If Not UCase(path) Like "*ATTACHEMENT*" Then Exit Sub 'just scraping submittlas
            ' failed to webscrape
            If Fail > 2 Then
                'do something about repeated failed attemps
                If Dir("\\server2\Faxes\ & " & TargetPO & " Re-Process.pdf") = "" Then Name fpath As "\\server2\Faxes\ & " & TargetPO & " Re-Process.pdf"
                Fail = 0
        End If
        
        'set renaming variables and call to send PDF
        Invoice_or_Order = "ORDACK"
        Target_folder = "PLATT - 234"
        Call Rename_and_move_pdf(fname, fpath, NewName, NewPath, Target_folder, Invoice_or_Order)
        
        'Webscrape and move pdf complete so exit this sub
        Exit Sub
        
    End If
            
' If trying to scrape submittals and did not fetch DocNo then exit here
    If Not UCase(path) Like "*ATTACHEMENT*" Then Exit Sub 'just scraping submittlas
            
' Message user that webscrape was not done and failed to pick up DocNo
    If docno = "" Then MsgBox "didn't pick up a Doc number on a platt order!"
            
' Enter Data into Sage with PDF scrape only
    Call CheckPONumber(TargetPO, Found)

'Found = 3 TargetPO is SHOP, Send PDF to Fax File
    If Found = 3 Then
        MsgBox "File was Shop PO, returned to Platt module"
    End If

    If Found = 1 And PossibleError < 1 Then 'Good TargetPO and no errors
        Call ClickOnSage
        xoffset = 0
        Call SageEnterPOfromTEMP(xoffset, emailmessage)
        If emailmessage = "Job entered was not valid in sage" Then
            sourcePath = fpath
            TargetPath = "\\server2\Dropbox\Attachments\_Re Run\" & fname
            Call PDF_MoveToFolder(sourcePath, TargetPath, specialmessage)
            updatelog = "Job entered was not valid in sage " & fname
            Call logupdate(updatelog)
            Exit Sub
        End If
        'rename and move file
        TotalInvoiceAmount = ThisWorkbook.Sheets("Temp").Range("N2").Offset(xoffset, 0)
        TotalInvoiceAmount = Replace(TotalInvoiceAmount, "$", "")
        If Not TotalInvoiceAmount Like "*.*" Then TotalInvoiceAmount = TotalInvoiceAmount & ".00"
        If TotalInvoiceAmount Like "*.[0-9]" Then TotalInvoiceAmount = TotalInvoiceAmount & "0"
        pdfoption1 = ThisWorkbook.Sheets("Temp").Range("A2") & " " _
        & "ORDACK " & ThisWorkbook.Sheets("Temp").Range("H2") & " (" _
        & TotalInvoiceAmount & ").pdf"
        pdfoption1 = Replace(pdfoption1, "$", "")
        
        If Dir("\\server2\Faxes\PLATT" & "\" & pdfoption1) = "" Then _
        Name fpath As "\\server2\Faxes\PLATT\" & pdfoption1
        'Sage Minimize
        SetCursorPos 1083, 11
        Call Mouse_left_button_press
        Call Mouse_left_button_Letgo
        Application.Wait (Now + TimeValue("00:00:01"))
        Else
        'MsgBox "Did not enter " & TargetPO & Chr(13) & "PossibleErrors =" & Possibleerror _
        & Chr(13) & "Found =" & Found & Chr(13) & "Document#" & DocNo
    End If

    If Found = 2 Then 'TargetPO matches a subcontract number
        'rename and move file
        TotalInvoiceAmount = ThisWorkbook.Sheets("Temp").Range("N2").Offset(xoffset, 0)
        TotalInvoiceAmount = Replace(TotalInvoiceAmount, "$", "")
        If Not TotalInvoiceAmount Like "*.*" Then TotalInvoiceAmount = TotalInvoiceAmount & ".00"
        If TotalInvoiceAmount Like "*.[0-9]" Then TotalInvoiceAmount = TotalInvoiceAmount & "0"
        pdfoption1 = ThisWorkbook.Sheets("Temp").Range("A2") & " " _
        & "ORDACK " & ThisWorkbook.Sheets("Temp").Range("H2") & " (" _
        & TotalInvoiceAmount & ").pdf"
        pdfoption1 = "Contract " & Replace(pdfoption1, "$", "")
        
        If Dir("\\server2\Faxes\" & "\" & pdfoption1) = "" Then _
        Name fpath As "\\server2\Faxes\" & pdfoption1 & " Subcontract"
        
        SetCursorPos 1083, 11 '--------------------------'Sage Minimize
        Call Mouse_left_button_press
        Call Mouse_left_button_Letgo
        Application.Wait (Now + TimeValue("00:00:01"))
        
        'MsgBox "Sent TargetPO that matches Subcontract number to Fax Folder " & PDFoption1
        
        Exit Sub
    
    End If


'MsgBox fpath
    If Dir(fpath) <> "" And UCase(path) Like "*ATTACH*" Or emailmessage Like "*already been entered*" And Dir(fpath) <> "" Then Kill (fpath)


End Sub
Sub ProcessPlattInvoicePDF(fname, fpath, XLSpath, XLSname, PDFtype, path, emailmessage)
'fname -> Original PDF file name
'fpath -> Original PDF file path
'XLSname - > Converted to excel file name
'XLSpath -> Converted to excel file path
'PDFtype -> Invoice or Order
'path -> Path of the parent folder being scanned, to discern whether this is a submittal scrape or processing invoices
'MsgBox "at platt ProcessPlattInvoicePDF(fname, fpath, XLSpath, XLSname, PDFtype, path)"


For xoffset = 1 To 200 ' Gather Macro Information
    Line = Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset, 0)
    
    ' Platt Document Number
    If Line Like "*[0-9][A-Z][0-9][0-9][0-9][0-9][0-9]*" Or _
       Line Like "[A-Z][0-9][0-9][0-9][0-9][0-9][0-9]" Then
        If docno = "" Then
            docno = Replace(Line, " ", "")
            vendorInvoice = docno
            
            ' Error check
            If docno Like "*ORIGINAL*" Then
                MsgBox "Platt Doc no is an original document number. Please investigate."
            End If
        End If
    End If
    
    ' Remove lines with problematic values
    If Line = "#VALUE!" Then
        Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset, 0) = ""
        Line = "nothing"
    End If
    
    ' Order and Invoice Dates
    If Line Like "*/*/*/*/*" Then
        line2 = Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset - 2, 0)
        InvoiceDate = line2
        OrderDate = Replace(Replace(Left(Line, 9), "T", ""), " ", "")
        
        If InvoiceDate = "" Or Not InvoiceDate Like "*/*/*" Then
            InvoiceDate = OrderDate
        End If
    End If
    
    ' DECO PO#
    If Line Like "*/*/*/*/*" And DecoPO = "" Then
        DecoPO = Replace(Right(Line, 20), " ", "")
        If DecoPO = "" Then
            MsgBox "Error, decoPO is empty."
            PossibleError = PossibleError + 1
        End If
    End If
    
    ' Shipping and Handling
    If UCase(Line) Like "*SHIP*HANDLING*" Then
        line2 = Replace(Replace(Line, "SHIP/HANDLING", ""), " ", "")
        If Not line2 Like "*.*" Then line2 = line2 & ".00"
        If line2 Like "*.[0-9]" Then line2 = line2 & "0"
        shipping = line2
    End If
    
    ' Total Invoice Amount
    If Line Like "Pay Online /*" Then
        For Repeat = 1 To 5
            line2 = Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset + Repeat, 0)
            If line2 Like "*[0-9]*" Then
                If Not line2 Like "*.*" Then line2 = line2 & ".00"
                If line2 Like "*.[0-9]" Then line2 = line2 & "0"
                TotalInvoice = line2
            End If
        Next Repeat
        
        ' Refine / filter Total Invoice
        If Len(TotalInvoice) > 8 Or TotalInvoice Like "*[a-zA-Z]*" Then
            TotalInvoice = "0.00"
        End If
    End If
    
    ' Tax
    If Line Like "Pay Online /*" Then
        line2 = Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset + 1, 0)
        Tax = Replace(Right(line2, 8), " ", "")
    End If

Next xoffset

' Error check
'If docno = "" Then MsgBox "Failed to get Docno so will not webscrape, hit break to investigate"

'Now Scrape Line Items

For xoffset = 3 To 200 'Gather Macro Information 'Now input Line Items
    If Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset, 0) Like "*[0-9]*" And _
        Workbooks(XLSname).Sheets(1).Range("B1").Offset(xoffset, 0) Like "*[0-9]*" And _
        Workbooks(XLSname).Sheets(1).Range("C1").Offset(xoffset, 0) Like "*[0-9]*" Then
        found1 = 1
        
        'ITEM DESC
        itemDesc = Workbooks(XLSname).Sheets(1).Range("D1").Offset(xoffset, 0) '
        'ItemDesc = Replace(ItemDesc, " ", "")
        'MsgBox "ItemDesc=:" & ItemDesc & ":"
        
        
        'UNIT
        Unit = Workbooks(XLSname).Sheets(1).Range("F1").Offset(xoffset, 0) '
        Unit = Replace(Unit, " ", "")
        If Unit = "FT" Then Unit = "EA"
        'MsgBox "Unit=:" & Unit & ":"
        
        
        'QUANTITY
        Quantity = Workbooks(XLSname).Sheets(1).Range("B1").Offset(xoffset, 0) '
        Quantity = Replace(Quantity, " ", "")
        'MsgBox "Quantity=:" & Quantity & ":"
        
        
        'BACK ORDER
        BackOrder = Workbooks(XLSname).Sheets(1).Range("C1").Offset(xoffset, 0) '
        BackOrder = Replace(BackOrder, " ", "")
        'MsgBox "Ship=:" & Ship & ":"
        
        'Unit Pricing
        unitprice = Workbooks(XLSname).Sheets(1).Range("E1").Offset(xoffset, 0)
        unitprice = Replace(unitprice, " ", "")
        'MsgBox "UnitPrice=:" & UnitPrice & ":"
        
        'Line total price
        lineprice = Workbooks(XLSname).Sheets(1).Range("G1").Offset(xoffset, 0)
        lineprice = Replace(lineprice, " ", "")
        'MsgBox "LinePrice=:" & LinePrice & ":"
        lineprice = Workbooks(XLSname).Sheets(1).Range("H1").Offset(xoffset, 0)
        lineprice = Replace(lineprice, " ", "")
        
        'Write data to thisworkbook temp sheet
        'MsgBox "LinePrice=:" & LinePrice & ":"
        ThisWorkbook.Sheets("Temp").Range("P2").Offset(tempsheetoffset, 0) = itemDesc
        ThisWorkbook.Sheets("Temp").Range("Q2").Offset(tempsheetoffset, 0) = Unit
        ThisWorkbook.Sheets("Temp").Range("R2").Offset(tempsheetoffset, 0) = Quantity
        ThisWorkbook.Sheets("Temp").Range("S2").Offset(tempsheetoffset, 0) = unitprice
        ThisWorkbook.Sheets("Temp").Range("T2").Offset(tempsheetoffset, 0) = lineprice
        ThisWorkbook.Sheets("Temp").Range("A2").Offset(tempsheetoffset, 0) = DecoPO
        ThisWorkbook.Sheets("Temp").Range("B2").Offset(tempsheetoffset, 0) = OrderDate
        ThisWorkbook.Sheets("Temp").Range("C2").Offset(tempsheetoffset, 0) = "234 - PLATT ELECTRIC SUPPLY"
        ThisWorkbook.Sheets("Temp").Range("AH2").Offset(tempsheetoffset, 0) = Tax
        ThisWorkbook.Sheets("Temp").Range("N2").Offset(tempsheetoffset, 0) = TotalInvoice
        ThisWorkbook.Sheets("Temp").Range("H2").Offset(tempsheetoffset, 0) = docno
        ThisWorkbook.Sheets("Temp").Range("B2").Offset(tempsheetoffset, 0) = OrderDate
        ThisWorkbook.Sheets("Temp").Range("J2").Offset(tempsheetoffset, 0) = InvoiceDate
        ThisWorkbook.Sheets("Temp").Range("AG2").Offset(tempsheetoffset, 0) = BackOrder
        tempsheetoffset = tempsheetoffset + 1
    End If
Next xoffset

'Total Invoice Backup plan // retrieve last number on sheet
    For xoffset = 3 To 200
        If Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset, 0) = "" Then
            If Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset + 1, 0) = "" Then
                ThisWorkbook.Sheets("Temp").Range("N2") = Workbooks(XLSname).Sheets(1).Range("A1").Offset(xoffset - 1, 0)
                'MsgBox ThisWorkbook.Sheets("Temp").Range("N2")
            End If
        End If
    If ThisWorkbook.Sheets("Temp").Range("N2") <> "" Then Exit For
    Next xoffset

' Close Excel Workbook and kill
    Workbooks(XLSname).Close SaveChanges:=False
    Kill XLSpath

' Find common errors and fix
    Call SelfHealTempPage
    
'Check if PO conforms before bothering to Enter
    TargetPO = DecoPO
    Call CheckPONumber(TargetPO, Found)

'special circumstance
    If TargetPO = "24AY-MC-0050824" Then docno = "5E07040"


'If DOC NO not found, but PO is found, try to acquire DOC NO
    If Found > 0 And docno = "" Or Found > 0 And docno Like "*ORIGINAL*" Then
        If TargetPO <> "" Then
            TargetPO = DecoPO
            'MsgBox "Could not acquire doc number from PDF, will attempt to webscrape it"
                If TotalInvoice = 0 Then
                    MsgBox "UPDATE: cannot find docno and Will not be able to identify correct invoice becuase did not get totalInvoice"
                End If
            Call PlattFindInvoiceNumber(TargetPO, Found, docno, emailmessage, Fail, TotalInvoice)
        End If
    End If
    
           
   
' if Fail = 3 then
' Webscrape if got document number
    If docno <> "" And Fail < 2 Then
    
        TargetPO = DecoPO
        'MsgBox "Successful read Document Number " & DocNo & Chr(13) & "for " & TargetPO & ", try to acquire the data through webscrape"
        'MsgBox DocNo
        PDFtype = "Invoice"
        
        Call PlattWebscrape(Fail, fpath, TargetPO, docno, PDFtype, shipping, path, fname, emailmessage):
        If UCase(emailmessage) Like "*TEMP*" Then Exit Sub
        
        If Not UCase(path) Like "*ATTACHMENT*" Then Exit Sub 'only here to webscrape submittals
        
        If Fail > 2 Then
            'do something about repeated failed attemps
            For Repeat = 1 To 20
            If Dir("\\server2\Faxes\" & TargetPO & " Re Process " & Repeat & "1234_INVOICE_1234.pdf") = "" Then
                'Name fpath As "\\server2\Faxes\" & TargetPO & " Re Process " & Repeat & "1234_INVOICE_1234.pdf"
                Exit For
            End If
            Next Repeat
            'MsgBox "renamed file"
            Fail = 0
        End If
        
       'set renaming variables and call to send PDF
        Invoice_or_Order = "INVOICE"
        Target_folder = "PLATT - 234"
        Call Rename_and_move_pdf(fname, fpath, NewName, NewPath, Target_folder, Invoice_or_Order)
        
        'Webscrape and move pdf complete so exit this sub
        Exit Sub
    Else
        For Repeat = 1 To 100
            If emailmessage <> "Busy" Then MsgBox "Did not get doc no. or something, couldnt do webscrape. Hit break to investigate"
        Next Repeat
    End If
                
    If Not UCase(path) Like "*ATTACHEMENT*" Then Exit Sub 'just scraping submittlas

'Enter Invoice Manually -> got all the data from reading the PDF
    If Found = 1 And possiblerror < 1 Then
        If ThisWorkbook.Sheets("Temp").Range("A2") = "" And docno = "" Then
            MsgBox "About to click on Sage to enter Invoice," & DecoPO _
                & "but data didn't populate to the Temp sheet, BREAK and investigate" & Chr(13) _
                & "TempSheetOffset=" & tempsheetoffset & Chr(13) & "Lrow=" & lrow & Chr(13) _
                & "Found line itmes on XLS Sheet = " & found1 & Chr(13) & "Assume line item did not parse correctly"
      
        End If
        'MsgBox "Freeze"
        Call ClickOnSage
        xoffset = 0
        Call SageEnterINVOICEfromTEMP(xoffset, emailmessage, fpath)
        If emailmessage = "Temp Sheet Total Error" Then Exit Sub
        TotalInvoiceAmount = ThisWorkbook.Sheets("Temp").Range("N2").Offset(xoffset, 0)
        TotalInvoiceAmount = Replace(TotalInvoiceAmount, "$", "")
        If Not TotalInvoiceAmount Like "*.*" Then TotalInvoiceAmount = TotalInvoiceAmount & ".00"
        If TotalInvoiceAmount Like "*.[0-9]" Then TotalInvoiceAmount = TotalInvoiceAmount & "0"
        pdfoption1 = ThisWorkbook.Sheets("Temp").Range("A2") & " " _
        & "ORDACK " & ThisWorkbook.Sheets("Temp").Range("H2") & " (" _
        & TotalInvoiceAmount & ").pdf"
        pdfoption1 = Replace(pdfoption1, "$", "")
        'MsgBox fpath
        'MsgBox PDFOption1
        If Dir("\\server2\Faxes\PLATT" & "\" & pdfoption1) = "" And UCase(path) Like "*ATTACH*" Then _
        Name fpath As "\\server2\Faxes\PLATT\" & pdfoption1
        
        SetCursorPos 20, 234 '--------------------------Sage, Click 4 Accounts Payable
        Application.Wait (Now + TimeValue("00:00:01"))
        Call Mouse_left_button_press
        Call Mouse_left_button_Letgo
    
        SetCursorPos 1083, 11 '--------------------------'Sage, Minimize
        Call Mouse_left_button_press
        Call Mouse_left_button_Letgo
        Application.Wait (Now + TimeValue("00:00:01"))
        
        If Dir(fpath) <> "" And UCase(fpath) Like "*ATTACH*" Or emailmessage Like "*already been entered*" And Dir(fpath) <> "" Then Kill (fpath)
    End If
    
    

    If PossibleError > 0 Or Fail > 2 Then MsgBox "Did not attempt to enter " & fname & Chr(13) & "Due to detected error in scraping data" & _
    Chr(13) & "Variables:" & Chr(13) & "PossibleError->" & PossibleError & Chr(13) & "Fail->" & Fail

End Sub
Sub PlattWebscrape(Fail, fpath, TargetPO, docno, PDFtype, shipping, path, fname, emailmessage):
    Dim ws As Worksheet
    Dim vendoritemno As String
    Set ws = ThisWorkbook.Sheets("Temp")


' If fpath = "DocmentRun" then sent here from User Module looking for a specific invoice
    If fpath = "DocumentRun" Then
        If docno = "" Then docno = InputBox("Enter the Platt Document Number")
    End If

' Prepare this workbooks temp sheet
    Call FormatTempSheet

'Set Target URL, Reference format: https://www.platt.com/Order.aspx?itemid=2D22679&CustNum=36850
    docno = Replace(docno, "ORIGINALORDER:", "")
    docno = Replace(docno, ".PDF", "")

'error check
    If UCase(docno) Like "*PDF*" Then
        MsgBox "Halt, Platt DocNo has the term PDF in it->" & Chr(13) & docno
    End If

    TargetURL = "https://www.platt.com/Orders/" & docno & "     "

' Open Chrome
    Call OpenChrome(TargetURL)
    
' COPY->PASTE webpage Data
    Application.CutCopyMode = False
    Application.SendKeys ("^a"), True
    Sleep (250)
    For Repeat = 1 To 3
        Application.SendKeys ("^c"), True
        Sleep 500
    Next Repeat
    ws.Paste Destination:=ws.Range("BA2")
    Sleep 500
    ws.DrawingObjects.Delete
    ws.Range("BA2:CA300").UnMerge
    Sleep 500

'Read Paste Data and check if at login page
    Found = 0
    For x = 0 To 100
        For y = 0 To 26
            If UCase(ws.Range("BA2").Offset(x, y)) Like "*PASSWORD?*" Then Found = 1
        Next y
        If Found = 1 Then Exit For
    Next x
    
    If Found = 1 Then
        ' MsgBox "We're at the login page"
        ' tab 16 times to hit "Login" with pre-populated data
        MsgBox "Need to login!"
    End If
    
    
'Check if at ORDER PAGE
    Found = 0
    For x = 0 To 100
        For y = 0 To 26
            If UCase(ws.Range("BA2").Offset(x, y)) Like "*" & docno & "*" Then Found = 1
        Next y
        If Found = 1 Then Exit For
    Next x
        
'If webpage copy determines we landed at the correct page, then the copy->paste data is good and we can close chrome
    If Found = 1 Then
        If UCase(path) Like "*ATTACHMENT*" Then
            'MsgBox "We're at the ORDER page"
            'close chrome unless we're going to go on to download submittal sheets
            Application.SendKeys ("^w")
            Sleep 500
            Application.SendKeys ("^w")
        End If
    Else
        Fail = Fail + 1
        If Fail > 2 Then
            Application.SendKeys ("^w")
            Sleep 500
            MsgBox "Failed to get to Platt Document Page"
            Exit Sub
        End If
    End If
    
' Reset variables
    tempsheetoffset = 0
    InvoiceDate = ""
    PossibleError = 0
    TotalInvoice = ""
    yoffset = 0

' Scrape copied wepage data
    For xoffset = 1 To 1000 'Gather Macro Information
           'MsgBox "Freeze"
           Line = ws.Range("BA2").Offset(xoffset, 0)
           Line = Replace(Line, vbLf, "")
    
        If Line = "CREDIT MEMO" And PDFtpye = "" Then
            PDFtype = "Invoice"
            MsgBox "PDFtype_>" & PDFtype
        End If
    
           'INVOICE Date
            If Line Like "*Placed on*" And PDFtype = "Invoice" And InvoiceDate = "" Then  '
               InvoiceDate = Line
               'If Not InvoiceDate Like "*/*/*" Then InvoiceDate = ""
               InvoiceDate = Replace(InvoiceDate, "Placed on", "")
                'InvoiceDate = Replace(InvoiceDate, " ", "")
               InvoiceDate = Format(InvoiceDate, "MM/DD/YYYY")
               InvoiceDate = Format(Date, "MM/DD/YYYY")
               'MsgBox "InvoiceDate=:" & InvoiceDate & ":"
           End If
           
           'ORDER DATE
           If Line Like "*Placed on*" And PDFtype = "Order" Then '
               OrderDate = Line
               'If Not OrderDate Like "*/*/*" Then OrderDate = ""
               OrderDate = Replace(OrderDate, "Placed on", "")
               OrderDate = Format(OrderDate, "MM/DD/YYYY")
               'MsgBox "OrderDate=:" & OrderDate & ":"
           End If
    
           'DELIVERY METHOD
           'ws.Range("AE2") = ""
           
            'TOTAL $
            If UCase(Line) Like "TOTAL" Then 'ws.Range("BA2").Offset(xoffset, 1) Like "*[$]*[0-9].[0-9]*" Then '
               'MsgBox "Freeze"
                If ws.Range("BA2").Offset(xoffset, 1) <> "" And _
                    ws.Range("BA2").Offset(xoffset, 1) Like "*[0-9]*" Then _
                    VendorTotalInvoice = CDbl(ws.Range("BA2").Offset(xoffset, yoffset + 1))
               'MsgBox "Platt VendorTotalInvoice=:" & VendorTotalInvoice & ":"
           End If
            
           'DECO PO
           If Line = "PO:" Then '
            DecoPO = ws.Range("BA2").Offset(xoffset + 1, yoffset)
            DecoPO = Replace(DecoPO, vbLf, "")
               DecoPO = Replace(DecoPO, "PO", "")
               DecoPO = Replace(DecoPO, "#", "")
               DecoPO = Replace(DecoPO, " ", "")
               'MsgBox "decoPO=:" & DecoPO & ":"
           End If
           
            'Vendor Invoice #
            vendorInvoice = docno
            
           'TAX
           If ws.Range("BA2").Offset(xoffset, 0) = "Tax" And ws.Range("BA2").Offset(xoffset, 1) <> "" Then
                Tax = ws.Range("BA2").Offset(xoffset, 1)
                'If tax Like "*[0-9]*" Then MsgBox "Platt Webscrape Found Tax Amount->" & tax
           End If
           
           'SHIPPING and HANDLING
            If UCase(Line) Like "*SHIP*HANDLING*" And Not UCase(Line) Like "*PLATT*" Then '
                For Repeat = 1 To 10
                    If ws.Range("BA2").Offset(xoffset, Repeat) <> "" Then
                        handling = ws.Range("BA2").Offset(xoffset, Repeat)
                        'If Handling Like "*[0-9]*" Then MsgBox "Platt Webscrape Found ship/handle Amount->" & Handling
                        Exit For
                    End If
                Next Repeat
           End If
           
    Next xoffset

    'For Repeat = 1 To 10
    '    MsgBox Freeze
    'Next Repeat
            
 ' Message user if failed to fetch Invoice Total
    If VendorTotalInvoice = "" And PDFtype = "Invoice" Or VendorTotalInvoice = 0 And PDFtype = "Invoice" Then
       If emailmessage <> "busy" Then
           MsgBox "Failed to scrape invoice total"
       End If
    End If
            
 ' Get Line Items
     For xoffset = 0 To 1000 'Now get Line Items
         yoffset = 0
         'For Yoffset = 0 To 20
             Line = ws.Range("BA2").Offset(xoffset, 0)
             Line = Replace(Line, vbLf, "")
             
              'ITEM DESCRIPTION
             If Line Like "*Item*#*" Then
     
                 itemDesc = ws.Range("BA2").Offset(xoffset - 1, 0)
                 itemDesc = Replace(itemDesc, "Item # ", "")
                 itemDesc = Replace(itemDesc, vbLf, "")
                 itemDesc = Replace(itemDesc, "+", "")
                 If Left(itemDesc, 1) = " " Then itemDesc = Right(itemDesc, Len(itemDesc - 1))
                 If Left(itemDesc, 1) = "," Then itemDesc = Right(itemDesc, Len(itemDesc - 1))
                 'remove problematic symbols
                 itemDesc = Replace(itemDesc, "°", "")
                 itemDesc = Replace(itemDesc, Chr(173), "")
                 If Len(itemDesc) > 60 Then ItemDec = Left(itemDesc, 60)
                 If itemDesc Like "1000*" Then MsgBox "PLATT WEBSCRAPER ERROR, description seems to be quantity" & Chr(13) & "Description->" & itemDesc
                 If itemDesc = "" Then PossibleError = PossibleError + 1
                 'MsgBox "ItemDesc=:" & ItemDesc
    
                ' Vendor Item No. (vendoritemno)
                vendoritemno = Replace(Line, "Item #", "")
                vendoritemno = Replace(vendoritemno, " ", "")
                'MsgBox ":" & vendoritemno & ":"
                
                    
                 'UNIT Type
                 For y = 0 To 6
                     If ws.Range("BA2").Offset(xoffset + y, yoffset) Like "*Price:*" Then Exit For
                 Next y
                 If y > 6 Then MsgBox "Found line item but couldn't find the unit!" & Chr(13) & "excel row ->" & xoffset + 2 & _
                      "might be one of those cases where the units are on the line below the item. fixed by resetting and going again"
                 
                 Unit = Right(ws.Range("BA2").Offset(xoffset + y, yoffset), 2)
                 'MsgBox unit
                 Unit = Replace(Unit, "Price:", "")
                 Unit = Replace(Unit, vbLf, "")
                 Unit = Replace(Unit, vbLf, "")
                 Unit = Replace(Unit, "(100 EA)", "")
                 Unit = Replace(Unit, "(100 FT)", "")
                 Unit = Replace(Unit, " ", "")
                 
                 If Unit Like "*E*E*" Then PossibleError = PossibleError + 1 'idicates that rows are combined in excel conversion
                 For re = 0 To 15
                     If Left(Unit, 1) Like "[0-9]" Or Left(Unit, 1) Like "/" Or Left(Unit, 1) Like "." _
                     Or Left(Unit, 1) Like "$" Then Unit = Right(Unit, Len(Unit) - 1)
                 Next re
                 If Unit = "FT" Then Unit = "EA"
                 If Unit Like "*C*" Then Unit = "C"
                 If Unit Like "*M*" Then Unit = "M"
                  
                 'QUANTITY ORDERED OR SHIPPED
                 For y = 0 To 6
                     If ws.Range("BA2").Offset(xoffset + y, 0) Like "*Order Qty:*" And _
                         PDFtype = "Order" Then Exit For
                     If ws.Range("BA2").Offset(xoffset + y, 0) Like "*Ship Qty:*" And _
                         PDFtype = "Invoice" Then Exit For
                 Next y
                 Quantity = ws.Range("BA2").Offset(xoffset + y, 0)
                 Quantity = Replace(Quantity, "Order Qty:", "")
                 Quantity = Replace(Quantity, "Ship Qty:", "")
                 Quantity = Replace(Quantity, vbLf, "")
                 Quantity = Replace(Quantity, " ", "")
                 If y = 7 Then Quantity = 0 'Ship Qty not found aka -> none shipped
                 Negative = 0
                 For re = 0 To 15
                     If UCase(Right(Quantity, 1)) = "-" Then Negative = 1
                     If UCase(Right(Quantity, 1)) Like "[A-Z]" Then Quantity = Left(Quantity, Len(Quantity) - 1)
                 Next re
                 If Negative = 1 Then Quantity = Quantity * -1
                 Quantity = Replace(Quantity, " ", "")
                 'MsgBox "Quantity=:" & Quantity & ":"
                 
    
                 'UNIT PRICE
                 For y = 0 To 6
                     If ws.Range("BA2").Offset(xoffset + y, 0) Like "*Price:*" Then Exit For
                 Next y
                 unitprice = ws.Range("BA2").Offset(xoffset + y, 0)
                 'MsgBox unitprice
                 unitprice = Replace(unitprice, vbLf, "")
                 unitprice = Replace(unitprice, "Price:", "")
                 unitprice = Replace(unitprice, "(100 FT)", "")
                 unitprice = Replace(unitprice, " ", "")
                 unitprice = Replace(unitprice, "$", "")
                 For re = 0 To Len(unitprice)
                     If Not Right(unitprice, 1) Like "[0-9]" Then unitprice = Left(unitprice, Len(unitprice) - 1)
                 Next re
                 
                 ' Unit price conversion
                 If Unit = "C" Then
                     If unitprice = "" Then unitprice = 0
                     unitprice = unitprice / 100
                     Unit = "EA"
                 End If
                 If Unit = "M" Then
                     unitprice = unitprice / 1000
                     Unit = "Ea"
                 End If
                 If unitprice = "" Then unitprice = "0"
                 If unitprice Like "*[A-Z]*" Then unitprice = "0"
                 'MsgBox "UnitPrice=:" & unitprice & ":"
                 
                'LINE TOTAL ref: Pre-tax Total: $27.35
                For y = 0 To 6
                     If ws.Range("BA2").Offset(xoffset + y, 0) Like "*Pre-tax Total:*" Then
                        'MsgBox "found line total"
                        lineprice = ws.Range("BA2").Offset(xoffset + y, 0)
                        lineprice = Replace(lineprice, "Pre-tax Total:", "")
                        lineprice = Replace(lineprice, "$", "")
                        lineprice = Replace(lineprice, " ", "")
                        If UCase(Quantity) Like "*[A-Z]*" Then Quantity = 1
                        ws.Range("S2").Offset(tempsheetoffset, 0) = unitprice
                        If Quantity <> "" And Quantity <> "0" And lineprice <> "" And _
                            lineprice <> "0" And unitprice <> "" Then
                            If Quantity * unitprice <> lineprice Then unitprice = lineprice / Quantity
                        End If
                        
                    End If
                 Next y
                 
                 'Write Data to thisworkbok temp sheet
                 ws.Range("P2").Offset(tempsheetoffset, 0) = itemDesc
                 ws.Range("Q2").Offset(tempsheetoffset, 0) = Unit
                 ws.Range("R2").Offset(tempsheetoffset, 0) = Quantity
                 ws.Range("S2").Offset(tempsheetoffset, 0).NumberFormat = "0.0000"
                 ws.Range("S2").Offset(tempsheetoffset, 0) = unitprice
                 ws.Range("T2").Offset(tempsheetoffset, 0) = lineprice
                 ws.Range("A2").Offset(tempsheetoffset, 0) = DecoPO
                 ws.Range("B2").Offset(tempsheetoffset, 0) = OrderDate
                 ws.Range("C2").Offset(tempsheetoffset, 0) = "234"
                 ws.Range("AH2").Offset(tempsheetoffset, 0) = Tax
                 ws.Range("H2").Offset(tempsheetoffset, 0) = vendorInvoice
                 ws.Range("N2").Offset(tempsheetoffset, 0) = VendorTotalInvoice
                 ws.Range("O2").Offset(tempsheetoffset, 0).NumberFormat = "@"
                 ws.Range("O2").Offset(tempsheetoffset, 0) = vendoritemno
                 ws.Range("J2").Offset(tempsheetoffset, 0) = InvoiceDate
                 tempsheetoffset = tempsheetoffset + 1
                 'Cross Check
                 If itemDesc = Quantity Then _
                     MsgBox "Platt Webscrape Transcription Error, Description " & itemDesc & " is Equal to Quantity " & Quantity & " on line " & xoffset _
                     & Chr(13) & "temp sheet line " & tempsheetoffset
             End If
         'Next yOffset
     Next xoffset
 
' If there was sales tax, write it in at the end now
    If Tax <> "" Then
            ws.Range("P2").Offset(tempsheetoffset, 0) = "Sales Tax on " & docno
            ws.Range("Q2").Offset(tempsheetoffset, 0) = ""
            ws.Range("R2").Offset(tempsheetoffset, 0) = 1
            ws.Range("S2").Offset(tempsheetoffset, 0).NumberFormat = "0.0000"
            ws.Range("S2").Offset(tempsheetoffset, 0) = Tax
            ws.Range("T2").Offset(tempsheetoffset, 0) = lineprice
            ws.Range("A2").Offset(tempsheetoffset, 0) = DecoPO
            ws.Range("B2").Offset(tempsheetoffset, 0) = OrderDate
            ws.Range("C2").Offset(tempsheetoffset, 0) = "234"
            ws.Range("AH2").Offset(tempsheetoffset, 0) = Tax
            ws.Range("H2").Offset(tempsheetoffset, 0) = vendorInvoice
            ws.Range("N2").Offset(tempsheetoffset, 0) = VendorTotalInvoice
            ws.Range("O2").Offset(tempsheetoffset, 0).NumberFormat = "@"
            ws.Range("O2").Offset(tempsheetoffset, 0) = vendoritemno
            ws.Range("J2").Offset(tempsheetoffset, 0) = InvoiceDate
            tempsheetoffset = tempsheetoffset + 1
    End If
     
 
'If there was shipping costs, insert it now at the end of the line items
    If handling <> "" Then
            ws.Range("S2").Offset(tempsheetoffset, 0) = handling 'actual cost
            ws.Range("A2").Offset(tempsheetoffset, 0) = DecoPO
            ws.Range("B2").Offset(tempsheetoffset, 0) = OrderDate
            ws.Range("C2").Offset(tempsheetoffset, 0) = "234"
            ws.Range("AH2").Offset(tempsheetoffset, 0) = Tax
            ws.Range("H2").Offset(tempsheetoffset, 0) = vendorInvoice
            ws.Range("N2").Offset(tempsheetoffset, 0) = VendorTotalInvoice
            ws.Range("P2").Offset(tempsheetoffset, 0) = "Shipping and Handling" 'line description
            ws.Range("R2").Offset(tempsheetoffset, 0) = 1
            ws.Range("J2").Offset(tempsheetoffset, 0) = InvoiceDate
    End If
                
    Application.SendKeys ("^w")
                
'MsgBox "Done scraping Web data"
    Call SelfHealTempPage

'Check if PO confors before bothering to Enter
    TargetPO = ws.Range("A2")
    'MsgBox TargetPO
    Call CheckPONumber(TargetPO, Found)


            
' If CheckPONumber determined the PO was a subcontract,
    If Found = 3 Then
        'MsgBox "Returned to Platt Module, TargetPO was SHOP->" & TargetPO & ", Move to Fax File" & Chr(13) & fpath & Chr(13) & docno
        Set fso = CreateObject("Scripting.filesystemobject")
        fso.MoveFile fpath, "\\server2\Faxes\" & docno & " Shop Expense.pdf"
        Exit Sub
    End If


' Exit now if we're only here to scrape submittlas
    If Not UCase(path) Like "*ATTACHMENTS*" Then Call Platt_download_product_sheet(path, fname)
    If Not UCase(path) Like "*ATTACHMENTS*" Then Exit Sub
            

'If TargetPO is acceptable, ENTER in SAGE, Move File
    If Found = 1 And PossibleError < 1 Then
        Application.Wait (Now + TimeValue("00:00:01"))
        Application.SendKeys ("^w")
        Application.Wait (Now + TimeValue("00:00:01"))
        Call ClickOnSage
        xoffset = 0
        emailmessage = "Platt"
        If fname Like "*INV*" Then PDFtype = "Invoice"
        If PDFtype = "Invoice" Then
            Call SageEnterINVOICEfromTEMP(xoffset, emailmessage, fpath)
            If emailmessage = "Temp Sheet Total Error" Then Exit Sub
            Else
            Call SageEnterPOfromTEMP(xoffset, emailmessage)
            If emailmessage = "Job entered was not valid in sage" Then
                sourcePath = fpath
                TargetPath = "\\server2\Dropbox\Attachments\_Re Run\" & fname
                Call PDF_MoveToFolder(sourcePath, TargetPath, specialmessage)
                updatelog = "Job entered was not valid in sage " & fname
                Call logupdate(updatelog)
            Exit Sub
        End If
        End If
        
        'If did not successfully entered info into Sage
        If emailmessage <> "Saved" Then Exit Sub
        If fpath = "DocumentRun" Then Exit Sub
        
       'rename and move file
        TotalInvoiceAmount = ws.Range("N2").Offset(xoffset, 0)
        TotalInvoiceAmount = Replace(TotalInvoiceAmount, "$", "")
        If Not TotalInvoiceAmount Like "*.*" Then TotalInvoiceAmount = TotalInvoiceAmount & ".00"
        If TotalInvoiceAmount Like "*.[0-9]" Then TotalInvoiceAmount = TotalInvoiceAmount & "0"
        
        'if invoice build name with INV
        If PDFtype = "Invoice" Then
            pdfoption1 = ws.Range("A2") & " " _
            & "INV " & ws.Range("H2") & " (" _
            & TotalInvoiceAmount & ").pdf"
            pdfoption1 = Replace(pdfoption1, "$", "")
        Else
        'if is order, build name with ORDACK
            pdfoption1 = ws.Range("A2") & " " _
            & "ORDACK " & ws.Range("H2") & " (" _
            & TotalInvoiceAmount & ").pdf"
            pdfoption1 = Replace(pdfoption1, "$", "")
        End If
        
        If emailmessage = "Saved" Then
            'MsgBox "\\server2\Faxes\PLATT\" & PDFOption1
            'MsgBox fpath
            If Dir("\\server2\Faxes\PLATT - 234" & "\" & pdfoption1) = "" And Dir(fpath) <> "" Then _
            Name fpath As "\\server2\Faxes\PLATT - 234\" & pdfoption1
            updatelog = pdfoption1
            Call logupdate(updatelog)
            Application.Wait (Now + TimeValue("00:00:06"))
        End If
        'MsgBox fpath
        
        'Sage Minimize
        SetCursorPos 1083, 11
        Call Mouse_left_button_press
        Call Mouse_left_button_Letgo
        Application.Wait (Now + TimeValue("00:00:01"))
        
        Else
    
    
        MsgBox "Did not enter-> " & fname & Chr(13) & "TargetPO->" & TargetPO & Chr(13) & "PossibleErrors =" & PossibleError _
        & Chr(13) & "Found =" & Found
    
    End If
    
End Sub
Sub PlattFindInvoiceNumber(TargetPO, Found, docno, emailmessage, Fail, TotalInvoice)


    file = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    If Dir("C:\Program Files\Google\Chrome\Application\chrome.exe") <> "" Then file = "C:\Program Files\Google\Chrome\Application\chrome.exe"
    
    
    Fail = 0
    Shell (file)
'<<Maximize
    Application.Wait (Now + TimeValue("00:00:01"))
    Application.SendKeys "%{ }" '
    Application.Wait (Now + TimeValue("00:00:01"))
    Application.SendKeys "x"
    Application.Wait (Now + TimeValue("00:00:01"))

'Filter TargetPO if needed
'MsgBox TargetPO
    Call CheckPONumber(TargetPO, Found)
    If Found <> 1 Then
        MsgBox "Freeze"
        MsgBox "Freeze"
        MsgBox "Freeze"
        MsgBox "Freeze"
    MsgBox "Freeze"

    End If


start:
    ThisWorkbook.Sheets("Temp").Range("BA2:CA300") = ""
    ThisWorkbook.Sheets("Temp").Range("BA2:CA300").UnMerge
    'Click into navigation field // Navigate to Platt
    Application.Wait (Now + TimeValue("00:00:04"))
    Application.SendKeys ("^l"), True
    Sleep (250)

    TargetURL = "https://www.platt.com/orders"
    Sleep 250
    Application.SendKeys (TargetURL), True
    Sleep 1000
    Application.SendKeys ("~"), True
    Sleep 1000
    Application.SendKeys ("~"), True
    Sleep 5000
    

'COPY->PASTE Page Data
    Application.CutCopyMode = False
    Application.SendKeys ("^a"), True
    Sleep (250)
    For Repeat = 1 To 3
        Application.SendKeys ("^c"), True
        Sleep 500
    Next Repeat
    ThisWorkbook.Sheets("Temp").Paste Destination:=ThisWorkbook.Sheets("Temp").Range("BA2")
    Sleep 500
    ThisWorkbook.Sheets("Temp").DrawingObjects.Delete
    ThisWorkbook.Sheets("Temp").Range("BA2:CA300").UnMerge
    Sleep 500
'Check if at login page
    'tab 16 times to hit "Login" with pre-populated data
    Found = 0
    For x = 0 To 100
        For y = 0 To 26
            If UCase(ThisWorkbook.Sheets("Temp").Range("BA2").Offset(x, y)) Like "*PASSWORD?*" Then Found = 1
        Next y
        If Found = 1 Then Exit For
    Next x
    If Found = 1 Then
       MsgBox "We're at the login page"
    End If



'SCAN FOR DATA
    Found = 0
    'MsgBox TargetPO
    'MsgBox TotalInvoice
    For x = 0 To 200
        For y = 0 To 5
        Line = ThisWorkbook.Sheets("Temp").Range("BA2").Offset(x, y)
            
            If Line Like "*" & TargetPO & "*" Then
               ' MsgBox "FOund PO -> Stage 1 ->" & TargetPO
                Excel_cost = ThisWorkbook.Sheets("Temp").Range("BA2").Offset(x, y + 3)
                Excel_cost = Replace(Excel_cost, "$", "")
                Excel_cost = Replace(Excel_cost, ",", "")
                Excel_cost = Replace(Excel_cost, ".00", "")
                'MsgBox "Excel_cost->" & Excel_cost & Chr(13) & "TotalInvoice->" & TotalInvoice
                If TotalInvoice Like "*" & Excel_cost & "*" Or Excel_cost = TotalInvoice Then
                    Found = Found + 1
                    docno = ThisWorkbook.Sheets("Temp").Range("BA2").Offset(x, y - 2)
                    ';MsgBox "found invoice number->" & docno
                End If
            End If
            If Found > 0 Then Exit For
        Next y
        If Found > 0 Then Exit For
    Next x
'If there is only one invoice then it must be the correct one......
    If Found > 0 Then
        'MsgBox "Found DocNO to be->" & DocNo
        Exit Sub
    End If

MsgBox "Didn't find PO & total on orders page" & Chr(13) & TargetPO & Chr(13) & TotalInvoice


'Else need to determine which invoice is the one on the PDF
'MsgBox "Freeze"
    Found = 0
    PDFInvoiceTotal = ThisWorkbook.Sheets("Temp").Range("N2")
    If Right(PDFInvoiceTotal, 1) = "-" Then
        PDFInvoiceTotal = Left(PDFInvoiceTotal, Len(PDFInvoiceTotal) - 1)
        PDFInvoiceTotal = PDFInvoiceTotal * -1
    End If

'Error check
    If PDFInvoiceTotal = "" Then MsgBox "Freeze / no total to compare"


    For x = 0 To 500
        docno = ThisWorkbook.Sheets("Temp").Range("BA2").Offset(x, 0)
        If docno Like "[A-Z][0-9][0-9][0-9][0-9][0-9][0-9]" Or docno Like "[0-9][A-Z][0-9][0-9][0-9][0-9][0-9]" Then
            'MsgBox DocNo & "->" & ThisWorkbook.Sheets("Temp").Range("BA2").Offset(x, 6) & Chr(13) & _
                            "trying to match ->" & PDFInvoiceTotal
            If ThisWorkbook.Sheets("Temp").Range("BA2").Offset(x, 6) = PDFInvoiceTotal Then Exit Sub
        End If
    Next x

'MsgBox "didn't find any matching invoice numbers"
'emailmessage = "Didn't Save"
'MsgBox "Searched for Document number->" & docno & Chr(13) & "But didn't find it"

Fail = 3

End Sub

Sub Platt_download_product_sheet(path, fname)
'MsgBox "in Platt_download_product_sheet ()"


'identify submittal folder
'\\server2\Dropbox\Acct\100 Jobs\2022 JOBS\2202 Marysville Stormwater Treatment (McClure)\300 Accounting\Backup
'\\server2\Dropbox\Acct\100 Jobs\2022 JOBS\2202 Marysville Stormwater Treatment (McClure)\Submittals
'\\server2\Dropbox\Acct\100 Jobs\2021 JOBS\2105 Arlington HS (Kassel)\300 Accounting\Backup
'\\server2\Dropbox\Acct\100 Jobs\2021 JOBS\2105 Arlington HS (Kassel)\Submittals
'
Submittal_Folder_Location = Replace(path, "300 Accounting\Backup", "Submittals")

'identify current order page
    Application.Wait (Now + TimeValue("00:00:01"))
    Application.SendKeys ("%d")
    Application.Wait (Now + TimeValue("00:00:01"))
    Set Clipboard = New MSForms.DataObject
    Application.CutCopyMode = False
    Clipboard.Clear
    Order_URL = ""
    Application.SendKeys ("^c")
    Sleep 150
    Clipboard.GetFromClipboard
    Order_URL = Clipboard.GetText
    If Order_URL = "" Then
        MsgBox "URL copy went wrong!"
    End If
    
'Run loop to download cut sheets
For x = 0 To 100
Get_item_description:
    Call Get_item_description(x, item_description, file_description, page_data)
    If item_description = "" Then GoTo close_webpage
    answer = 0 'whether or not this item has already been downlaoded
    Call Check_if_already_downloaded_this_submittal(Submittal_Folder_Location, file_description, answer)
    If answer = 2 Then 'already downloaded this submittal
        x = x + 1
        GoTo Get_item_description:
    End If
try = 0
enter_product_data:
    Application.Wait (Now + TimeValue("00:00:01"))
    Application.SendKeys ("^f")
    Application.Wait (Now + TimeValue("00:00:01"))
    Application.SendKeys item_description, True
    Application.Wait (Now + TimeValue("00:00:01"))
    'advance cursor to selected item
    Application.SendKeys ("+^~")
    Application.Wait (Now + TimeValue("00:00:01"))
    'click item to advance to specific page
    Application.SendKeys ("^~")
    'wait for page to load
wait_for_product_page_to_load:
    Application.Wait (Now + TimeValue("00:00:04"))
    'check if we have made it to another page
    Application.SendKeys ("%d")
    Application.Wait (Now + TimeValue("00:00:01"))
    Set Clipboard = New MSForms.DataObject
    Application.CutCopyMode = False
    Clipboard.Clear
    product_URL = ""
    Application.SendKeys ("^c")
    Sleep 150
    Clipboard.GetFromClipboard
    product_URL = Clipboard.GetText
    If Order_URL = "" Then
        MsgBox "URL copy went wrong!"
    End If
    If product_URL = Order_URL Then
        try = try + 1
        If try > 2 Then
            'goto enter the data again
            try = 0
            x = x + 1
            GoTo Get_item_description:
        End If
        GoTo wait_for_product_page_to_load:
    End If
    
    
    
try = 1
click_on_product_data:

If try = 1 Then findtext = "Catalog"
If try > 1 Then findtext = "Cut Sheet"
    'search and click on catalog page
    Application.Wait (Now + TimeValue("00:00:01"))
    Application.SendKeys ("^f")
    Application.Wait (Now + TimeValue("00:00:01"))
    Application.SendKeys (findtext)
    Application.Wait (Now + TimeValue("00:00:01"))
    Application.SendKeys ("{esc}")
    For Repeat = 1 To 20
        Sleep 5
    Next Repeat
    'advance cursor to selected item
    Application.SendKeys ("~")

wait_for_catalog_page_to_load:
    Application.Wait (Now + TimeValue("00:00:04"))
    Application.SendKeys ("%d")
    Application.Wait (Now + TimeValue("00:00:01"))
    Set Clipboard = New MSForms.DataObject
    Application.CutCopyMode = False
    Clipboard.Clear
    Check_URL = ""
    Application.SendKeys ("^c")
    Sleep 150
    Clipboard.GetFromClipboard
    Check_URL = Clipboard.GetText
    If Check_URL = product_URL Then
        try = try + 1
        If try > 3 Then
            'x = x + 1
            Call Chrome_close_secondary_tabs
            GoTo restore_primary_tab:
        End If
        GoTo click_on_product_data
    End If
    
Call Submittals_click_download_PDF(file_description)
   
restore_primary_tab:
    'Restore ORDER_URL in primary tab
    Application.Wait (Now + TimeValue("00:00:01"))
    Application.SendKeys ("%d")
    Application.Wait (Now + TimeValue("00:00:01"))
    Application.SendKeys Order_URL
    Application.SendKeys "~"
wait_for_order_page_to_Load:
    Application.Wait (Now + TimeValue("00:00:05"))

    
Next x

close_webpage:

'Update submittals page with download data
    If fname = "" Then MsgBox "Freeze"
    For x = 0 To 10000
        If ThisWorkbook.Sheets("Submittals").Range("A1").Offset(x, 0) = fname Or ThisWorkbook.Sheets("Submittals").Range("A1").Offset(x, 0) = "" Then Exit For
    Next x
    ThisWorkbook.Sheets("Submittals").Range("A1").Offset(x, 0) = fname

For Repeat = 1 To 5
    Application.SendKeys ("^w")
    Application.Wait (Now + TimeValue("00:00:01"))
Next Repeat


End Sub

Sub PlattWebscrapeCall()

' This module is if a webscrape of an invoice number is called directly from the forms module
    fpath = "DocumentRun"

' Execute the webscrape
    Call PlattWebscrape(Fail, fpath, TargetPO, docno, PDFtype, fname, emailmessage):

End Sub
