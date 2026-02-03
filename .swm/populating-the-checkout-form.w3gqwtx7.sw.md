---
title: Populating the Checkout Form
---
This document describes how the checkout form is populated with guest, rate, and financial information. The process retrieves the current transaction, fills in guest details, and calculates charges and balances to ensure the form is ready for checkout.

# Populating Checkout Form with Guest and Rate Data

<SwmSnippet path="/HotelManagementSystem/Forms/frmCheckOut.frm" line="658">

---

In <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="658:4:4" line-data="Private Sub Form_Load()">`Form_Load`</SwmToken>, we kick things off by starting a database transaction and loading the transaction record for the current room where the status is 'Check In'. We bind the rate type data to the rate type dropdown, then start populating the form controls with the guest's check-in details. We need to call into <SwmPath>[HotelManagementSystem/Modules/modProcedure.bas](HotelManagementSystem/Modules/modProcedure.bas)</SwmPath> next to handle the actual data binding for the rate type dropdown, since that's not handled inline here.

```visual basic
Private Sub Form_Load()
On Error GoTo err

    CN.BeginTrans

    RS.CursorLocation = adUseClient
    RS.Open "SELECT * FROM Transactions WHERE RoomNumber = " & RoomNumber & " AND Status = 'Check In'", CN, adOpenStatic, adLockOptimistic

    bind_dc "SELECT * FROM [Rate Type]", "RateType", dcRateType, "RateTypeID", True

    txtRoomNumber.Text = RoomNumber
    
    With RS
        txtGuestName.Tag = .Fields("FolioNumber")
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Modules/modProcedure.bas" line="180">

---

<SwmToken path="HotelManagementSystem/Modules/modProcedure.bas" pos="180:4:4" line-data="Public Sub bind_dc(ByVal srcSQL As String, ByVal srcBindField As String, ByRef srcDC As DataCombo, Optional srcColBound As String, Optional ShowFirstRec As Boolean)">`bind_dc`</SwmToken> handles binding the results of a SQL query to a <SwmToken path="HotelManagementSystem/Modules/modProcedure.bas" pos="180:30:30" line-data="Public Sub bind_dc(ByVal srcSQL As String, ByVal srcBindField As String, ByRef srcDC As DataCombo, Optional srcColBound As String, Optional ShowFirstRec As Boolean)">`DataCombo`</SwmToken> control. It sets up which field to display, which to bind, and fills the dropdown with the query results. If <SwmToken path="HotelManagementSystem/Modules/modProcedure.bas" pos="180:44:44" line-data="Public Sub bind_dc(ByVal srcSQL As String, ByVal srcBindField As String, ByRef srcDC As DataCombo, Optional srcColBound As String, Optional ShowFirstRec As Boolean)">`ShowFirstRec`</SwmToken> is true, it also sets the dropdown to the first record and stores some metadata in the Tag property. This is how the rate type dropdown gets its data and initial state.

```visual basic
Public Sub bind_dc(ByVal srcSQL As String, ByVal srcBindField As String, ByRef srcDC As DataCombo, Optional srcColBound As String, Optional ShowFirstRec As Boolean)
    Dim RS As New Recordset
    
    RS.CursorLocation = adUseClient
    RS.Open srcSQL, CN, adOpenStatic, adLockOptimistic
    
    With srcDC
        .ListField = srcBindField
        .BoundColumn = srcColBound
        Set .RowSource = RS
        'Display the first record
        If ShowFirstRec = True Then
            If Not RS.RecordCount < 1 Then
                .BoundText = RS.Fields(srcColBound)
                .Tag = RS.RecordCount & "*~~~~~*" & RS.Fields(srcColBound)
            Else
                .Tag = "0*~~~~~*0"
            End If
        End If
    End With
    Set RS = Nothing
End Sub
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Forms/frmCheckOut.frm" line="672">

---

Back in <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="658:4:4" line-data="Private Sub Form_Load()">`Form_Load`</SwmToken>, after binding the rate type dropdown, we fill in the rest of the form fields using the transaction record. We fetch the guest name with a separate query, set the check-out date based on <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="674:3:3" line-data="        If AutoCheckOut = True Then">`AutoCheckOut`</SwmToken> logic, and fill in other details like days, adults, children, and rates. We need to call into <SwmPath>[HotelManagementSystem/Modules/modFunction.bas](HotelManagementSystem/Modules/modFunction.bas)</SwmPath> next to actually run the SQL and get the guest name.

```visual basic
        txtGuestName.Text = getValueAt("SELECT [Name] FROM qry_CheckIn WHERE FolioNumber = '" & .Fields("FolioNumber") & " '", "Name")
        txtDateIn.Text = .Fields("DateIn")
        If AutoCheckOut = True Then
            If .Fields("DateOut") >= Date Then
                dtpDateOut.Value = .Fields("DateOut")
            Else
                dtpDateOut.Value = Date
            End If
        Else
            dtpDateOut.Value = .Fields("DateOut")
        End If
        dcRateType.BoundText = .Fields("RateType")
        txtDays.Text = dtpDateOut.Value - CDate(txtDateIn.Text)
        txtAdults.Text = .Fields("Adults")
        txtChildrens.Text = .Fields("Childrens")
        txtRate.Text = toMoney(.Fields("Rate"))
        txtOtherCharges.Text = toMoney(.Fields("OtherCharges"))
        txtDiscount.Text = toMoney(.Fields("Discount"))
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Modules/modFunction.bas" line="168">

---

<SwmToken path="HotelManagementSystem/Modules/modFunction.bas" pos="168:4:4" line-data="Public Function getValueAt(ByVal srcSQL As String, ByVal whichField As String) As String">`getValueAt`</SwmToken> runs a SQL query and grabs the value of a specific field from the first record. It's used here to fetch the guest name for display in the checkout form.

```visual basic
Public Function getValueAt(ByVal srcSQL As String, ByVal whichField As String) As String
    Dim RS As New Recordset
    
    RS.CursorLocation = adUseClient
    RS.Open srcSQL, CN, adOpenStatic, adLockReadOnly
    If RS.RecordCount > 0 Then getValueAt = RS.Fields(whichField)
    
    Set RS = Nothing
End Function
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Forms/frmCheckOut.frm" line="690">

---

Back in <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="658:4:4" line-data="Private Sub Form_Load()">`Form_Load`</SwmToken>, after getting the guest name, we fill in the amount paid and other monetary fields, formatting them for display. We need to call into <SwmPath>[HotelManagementSystem/Modules/modFunction.bas](HotelManagementSystem/Modules/modFunction.bas)</SwmPath> next to handle the currency formatting.

```visual basic
        txtAmountPaid.Text = toMoney(.Fields("AmountPaid"))
    End With
    
    dcRateType.Enabled = False
    
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Modules/modFunction.bas" line="216">

---

<SwmToken path="HotelManagementSystem/Modules/modFunction.bas" pos="216:4:4" line-data="Public Function toMoney(ByVal srcCurr As String) As String">`toMoney`</SwmToken> takes a string, treats blanks as zero, and formats the value as currency with two decimals and thousands separators. It's used to make sure all monetary fields look consistent in the UI.

```visual basic
Public Function toMoney(ByVal srcCurr As String) As String
   toMoney = Format$(IIf(Trim(srcCurr) = "", 0, srcCurr), "#,##0.00")
End Function
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Forms/frmCheckOut.frm" line="695">

---

Back in <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="658:4:4" line-data="Private Sub Form_Load()">`Form_Load`</SwmToken>, after formatting the monetary fields, we call <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="695:3:3" line-data="    Call ComputeAddRate">`ComputeAddRate`</SwmToken> to fetch and update the room rate and extra charges for the current room and rate type.

```visual basic
    Call ComputeAddRate
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Forms/frmCheckOut.frm" line="726">

---

<SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="726:4:4" line-data="Private Sub ComputeAddRate()">`ComputeAddRate`</SwmToken> queries the database for the current room's rates and extra charges based on the selected rate type. If found, it updates the UI fields with these values, formatting the main rate for display. We need to call <SwmPath>[HotelManagementSystem/Modules/modFunction.bas](HotelManagementSystem/Modules/modFunction.bas)</SwmPath> next to handle the formatting.

```visual basic
Private Sub ComputeAddRate()
    Dim rsRoomRates As New ADODB.Recordset
    
    With rsRoomRates
        .Open "SELECT * FROM [Room Rates] WHERE RoomNumber = " & RoomNumber & " AND RateTypeID = " & dcRateType.BoundText, CN, adOpenStatic, adLockOptimistic
    
        If .RecordCount > 0 Then
            txtRate.Text = toMoney(!RoomRate)
            txtAdults.Tag = !ExtraAdultRate
            txtChildrens.Tag = !ExtraChildRate
        End If
    End With
    
    rsRoomRates.Close
End Sub
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Forms/frmCheckOut.frm" line="696">

---

Back in <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="658:4:4" line-data="Private Sub Form_Load()">`Form_Load`</SwmToken>, after updating the rates, we call <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="696:3:3" line-data="    Call ComputeRate">`ComputeRate`</SwmToken> to recalculate all the totals. Then we clear and reload the temporary rate per period data for the current folio, making sure the calculations use the latest info.

```visual basic
    Call ComputeRate

    
    CN.Execute "DELETE FolioNumber " & _
                "From [Rate Per Period Temp] " & _
                "WHERE FolioNumber='" & txtGuestName.Tag & "'"

    CN.Execute "INSERT INTO [Rate Per Period Temp] " & _
                "SELECT [Rate Per Period].* " & _
                "From [Rate Per Period] " & _
                "WHERE FolioNumber='" & txtGuestName.Tag & "'"
                
    CN.CommitTrans
    
    Exit Sub

err:
    CN.RollbackTrans
```

---

</SwmSnippet>

## Calculating Charges and Balances

<SwmSnippet path="/HotelManagementSystem/Forms/frmCheckOut.frm" line="718">

---

In <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="718:4:4" line-data="Private Sub ComputeRate()">`ComputeRate`</SwmToken>, we calculate total charges using <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="719:9:9" line-data="    txtTotalCharges.Text = toMoney(ComputeRatePerPeriod)">`ComputeRatePerPeriod`</SwmToken>, then add other charges for the subtotal. We need to call <SwmPath>[HotelManagementSystem/Modules/modFunction.bas](HotelManagementSystem/Modules/modFunction.bas)</SwmPath> next to convert the string values to numbers for further calculations.

```visual basic
Private Sub ComputeRate()
    txtTotalCharges.Text = toMoney(ComputeRatePerPeriod)
    txtSubTotal.Text = toMoney(toNumber(txtTotalCharges.Text) + toNumber(txtOtherCharges.Text))
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Forms/frmCheckOut.frm" line="720">

---

Back in <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="696:3:3" line-data="    Call ComputeRate">`ComputeRate`</SwmToken>, after converting and adding the charges, we format the subtotal for display. We need to call <SwmPath>[HotelManagementSystem/Modules/modFunction.bas](HotelManagementSystem/Modules/modFunction.bas)</SwmPath> again to handle the next calculation steps.

```visual basic
    txtSubTotal.Text = toMoney(toNumber(txtTotalCharges.Text) + toNumber(txtOtherCharges.Text))
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Modules/modFunction.bas" line="182">

---

<SwmToken path="HotelManagementSystem/Modules/modFunction.bas" pos="182:4:4" line-data="Public Function toNumber(ByVal srcCurrency As String, Optional RetZeroIfNegative As Boolean) As Double">`toNumber`</SwmToken> converts a currency string (with or without commas) to a double. If <SwmToken path="HotelManagementSystem/Modules/modFunction.bas" pos="182:17:17" line-data="Public Function toNumber(ByVal srcCurrency As String, Optional RetZeroIfNegative As Boolean) As Double">`RetZeroIfNegative`</SwmToken> is set, it forces values less than 1 to zero. This is used to prep values for calculations in <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="696:3:3" line-data="    Call ComputeRate">`ComputeRate`</SwmToken>.

```visual basic
Public Function toNumber(ByVal srcCurrency As String, Optional RetZeroIfNegative As Boolean) As Double
    If srcCurrency = "" Then
        toNumber = 0
    Else
        Dim retValue As Double
        If InStr(1, srcCurrency, ",") > 0 Then
            retValue = Val(Replace(srcCurrency, ",", "", , , vbTextCompare))
        Else
            retValue = Val(srcCurrency)
        End If
        If RetZeroIfNegative = True Then
            If retValue < 1 Then retValue = 0
        End If
        toNumber = retValue
        retValue = 0
    End If
End Function
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Forms/frmCheckOut.frm" line="721">

---

Back in <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="696:3:3" line-data="    Call ComputeRate">`ComputeRate`</SwmToken>, after getting the subtotal, we apply the discount as a percentage and format the result for display. We need to call <SwmPath>[HotelManagementSystem/Modules/modFunction.bas](HotelManagementSystem/Modules/modFunction.bas)</SwmPath> again to convert and format the values.

```visual basic
    txtTotal.Text = toMoney(toNumber(txtSubTotal.Text) - (toNumber(txtSubTotal.Text) * toNumber(txtDiscount.Text) / 100))
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Forms/frmCheckOut.frm" line="721">

---

Back in <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="696:3:3" line-data="    Call ComputeRate">`ComputeRate`</SwmToken>, after calculating the total, we subtract the amount paid to get the balance, formatting it for display. We need to call <SwmPath>[HotelManagementSystem/Modules/modFunction.bas](HotelManagementSystem/Modules/modFunction.bas)</SwmPath> again for the conversion and formatting.

```visual basic
    txtTotal.Text = toMoney(toNumber(txtSubTotal.Text) - (toNumber(txtSubTotal.Text) * toNumber(txtDiscount.Text) / 100))
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Forms/frmCheckOut.frm" line="722">

---

Back in <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="696:3:3" line-data="    Call ComputeRate">`ComputeRate`</SwmToken>, we format the balance for display, calling <SwmPath>[HotelManagementSystem/Modules/modFunction.bas](HotelManagementSystem/Modules/modFunction.bas)</SwmPath> for consistency.

```visual basic
    txtBalance.Text = toMoney(toNumber(txtTotal.Text) - toNumber(txtAmountPaid.Text))
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Forms/frmCheckOut.frm" line="722">

---

Back in <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="696:3:3" line-data="    Call ComputeRate">`ComputeRate`</SwmToken>, after all calculations, we format and display the final balance. Each calculation step uses <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="722:7:7" line-data="    txtBalance.Text = toMoney(toNumber(txtTotal.Text) - toNumber(txtAmountPaid.Text))">`toMoney`</SwmToken> to keep the UI consistent.

```visual basic
    txtBalance.Text = toMoney(toNumber(txtTotal.Text) - toNumber(txtAmountPaid.Text))
End Sub
```

---

</SwmSnippet>

## Finalizing Checkout Form and Error Handling

<SwmSnippet path="/HotelManagementSystem/Forms/frmCheckOut.frm" line="714">

---

Back in <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="658:4:4" line-data="Private Sub Form_Load()">`Form_Load`</SwmToken>, after all calculations, we handle any errors by showing a message and logging the details. We call <SwmPath>[HotelManagementSystem/Modules/modProcedure.bas](HotelManagementSystem/Modules/modProcedure.bas)</SwmPath> to handle the error prompt and logging.

```visual basic
    prompt_err err, Name, "txtDays_Change"
    Screen.MousePointer = vbDefault
End Sub
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Modules/modProcedure.bas" line="87">

---

<SwmToken path="HotelManagementSystem/Modules/modProcedure.bas" pos="87:4:4" line-data="Public Sub prompt_err(ByVal sError As ErrObject, ByVal ModuleName As String, ByVal OccurIn As String)">`prompt_err`</SwmToken> shows an error message with details and logs the error to a file with all the relevant info for debugging.

```visual basic
Public Sub prompt_err(ByVal sError As ErrObject, ByVal ModuleName As String, ByVal OccurIn As String)
    MsgBox "Error From: " & ModuleName & vbNewLine & _
           "Occur In: " & OccurIn & vbNewLine & _
           "Error Number: " & sError.Number & vbNewLine & _
           "Description: " & sError.Description, vbCritical, "Application Error"
    'Save the error log (The save error log will be display later on in the program)
    Open App.Path & "\Error.log" For Append As #1
        Print #1, Format(Date, "MMM-dd-yyyy") & "~~~~~" & Time & "~~~~~" & sError.Number & "~~~~~" & sError.Description & "~~~~~" & ModuleName & "~~~~~" & OccurIn
    Close #1
End Sub
```

---

</SwmSnippet>

&nbsp;

*This is an auto-generated document by Swimm 🌊 and has not yet been verified by a human*

<SwmMeta version="3.0.0" repo-id="Z2l0aHViJTNBJTNBY3RzLVZCNi1Qcm9qZWN0cyUzQSUzQVN3aW1tLURlbW8=" repo-name="cts-VB6-Projects"><sup>Powered by [Swimm](https://app.swimm.io/)</sup></SwmMeta>
