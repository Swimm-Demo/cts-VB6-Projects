---
title: Updating stay duration and charges after check-out date change
---
When a user selects a new check-out date, the hotel management system updates the stay duration and recalculates all charges, discounts, and balances. The updated billing information is immediately reflected in the UI.

# Handling Check-Out Date Changes

<SwmSnippet path="/HotelManagementSystem/Forms/frmCheckOut.frm" line="638">

---

In <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="638:4:4" line-data="Private Sub dtpDateOut_Change()">`dtpDateOut_Change`</SwmToken>, we update the days stayed by subtracting the check-in date from the new check-out date, then immediately call <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="641:3:3" line-data="    Call ComputeRate">`ComputeRate`</SwmToken> to refresh all rate and charge calculations based on the new duration.

```visual basic
Private Sub dtpDateOut_Change()
    txtDays.Text = dtpDateOut.Value - CDate(txtDateIn.Text)
    
    Call ComputeRate
```

---

</SwmSnippet>

## Calculating and Formatting Charges

<SwmSnippet path="/HotelManagementSystem/Forms/frmCheckOut.frm" line="718">

---

In <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="718:4:4" line-data="Private Sub ComputeRate()">`ComputeRate`</SwmToken>, we calculate the total charges using <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="719:9:9" line-data="    txtTotalCharges.Text = toMoney(ComputeRatePerPeriod)">`ComputeRatePerPeriod`</SwmToken>, then format the result as currency with <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="719:7:7" line-data="    txtTotalCharges.Text = toMoney(ComputeRatePerPeriod)">`toMoney`</SwmToken>. We also prepare the subtotal by adding other charges, again formatting for display. The next step needs <SwmPath>[HotelManagementSystem/Modules/modFunction.bas](HotelManagementSystem/Modules/modFunction.bas)</SwmPath> for currency formatting and conversion.

```visual basic
Private Sub ComputeRate()
    txtTotalCharges.Text = toMoney(ComputeRatePerPeriod)
    txtSubTotal.Text = toMoney(toNumber(txtTotalCharges.Text) + toNumber(txtOtherCharges.Text))
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Modules/modFunction.bas" line="216">

---

ToMoney handles formatting any string as a currency value, defaulting empty inputs to zero and ensuring consistent display with two decimals and separators.

```visual basic
Public Function toMoney(ByVal srcCurr As String) As String
   toMoney = Format$(IIf(Trim(srcCurr) = "", 0, srcCurr), "#,##0.00")
End Function
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Forms/frmCheckOut.frm" line="720">

---

Back in <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="641:3:3" line-data="    Call ComputeRate">`ComputeRate`</SwmToken>, after formatting total charges, we sum the charges and other charges (converting both to numbers for accuracy), then format the subtotal for display. We call <SwmPath>[HotelManagementSystem/Modules/modFunction.bas](HotelManagementSystem/Modules/modFunction.bas)</SwmPath> again to handle the parsing and formatting.

```visual basic
    txtSubTotal.Text = toMoney(toNumber(txtTotalCharges.Text) + toNumber(txtOtherCharges.Text))
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Modules/modFunction.bas" line="182">

---

ToNumber parses currency strings to numeric values, stripping commas for correct conversion and optionally zeroing out small/negative results if requested.

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

Back in <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="641:3:3" line-data="    Call ComputeRate">`ComputeRate`</SwmToken>, after getting the subtotal, we calculate the total by applying any discount percentage, again parsing and formatting values using <SwmPath>[HotelManagementSystem/Modules/modFunction.bas](HotelManagementSystem/Modules/modFunction.bas)</SwmPath> for accuracy and display.

```visual basic
    txtTotal.Text = toMoney(toNumber(txtSubTotal.Text) - (toNumber(txtSubTotal.Text) * toNumber(txtDiscount.Text) / 100))
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Forms/frmCheckOut.frm" line="721">

---

Back in <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="641:3:3" line-data="    Call ComputeRate">`ComputeRate`</SwmToken>, after calculating the discounted total, we format it for display using <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="721:7:7" line-data="    txtTotal.Text = toMoney(toNumber(txtSubTotal.Text) - (toNumber(txtSubTotal.Text) * toNumber(txtDiscount.Text) / 100))">`toMoney`</SwmToken>

```visual basic
    txtTotal.Text = toMoney(toNumber(txtSubTotal.Text) - (toNumber(txtSubTotal.Text) * toNumber(txtDiscount.Text) / 100))
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Forms/frmCheckOut.frm" line="722">

---

Back in <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="641:3:3" line-data="    Call ComputeRate">`ComputeRate`</SwmToken>, we determine the remaining balance by subtracting the amount paid from the total, then format the result for display using <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="722:7:7" line-data="    txtBalance.Text = toMoney(toNumber(txtTotal.Text) - toNumber(txtAmountPaid.Text))">`toMoney`</SwmToken>.

```visual basic
    txtBalance.Text = toMoney(toNumber(txtTotal.Text) - toNumber(txtAmountPaid.Text))
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Forms/frmCheckOut.frm" line="722">

---

Back in <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="641:3:3" line-data="    Call ComputeRate">`ComputeRate`</SwmToken>, after calculating the balance, we format it for display, keeping all monetary values consistent in the UI.

```visual basic
    txtBalance.Text = toMoney(toNumber(txtTotal.Text) - toNumber(txtAmountPaid.Text))
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Forms/frmCheckOut.frm" line="723">

---

Finally, <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="641:3:3" line-data="    Call ComputeRate">`ComputeRate`</SwmToken> ends after updating all charge-related fields

```visual basic
End Sub
```

---

</SwmSnippet>

## Completing the Date Change Update

<SwmSnippet path="/HotelManagementSystem/Forms/frmCheckOut.frm" line="642">

---

Finally, after returning from <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="641:3:3" line-data="    Call ComputeRate">`ComputeRate`</SwmToken>, <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="638:4:4" line-data="Private Sub dtpDateOut_Change()">`dtpDateOut_Change`</SwmToken> ends, leaving all UI fields updated to reflect the new check-out date.

```visual basic
End Sub
```

---

</SwmSnippet>

&nbsp;

*This is an auto-generated document by Swimm 🌊 and has not yet been verified by a human*

<SwmMeta version="3.0.0" repo-id="Z2l0aHViJTNBJTNBY3RzLVZCNi1Qcm9qZWN0cyUzQSUzQVN3aW1tLURlbW8=" repo-name="cts-VB6-Projects"><sup>Powered by [Swimm](https://app.swimm.io/)</sup></SwmMeta>
