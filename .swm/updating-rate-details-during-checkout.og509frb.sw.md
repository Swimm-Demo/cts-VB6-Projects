---
title: Updating rate details during checkout
---
This document describes the flow for updating rate details during checkout. When the rate per period label is clicked, a dialog is presented for user input. After changes, charges and balances are recalculated and the UI is updated to show the new values.

# Triggering Rate Per Period Dialog and Recalculation

<SwmSnippet path="/HotelManagementSystem/Forms/frmCheckOut.frm" line="810">

---

In <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="810:4:4" line-data="Private Sub lblRatePerPeriod_Click()">`lblRatePerPeriod_Click`</SwmToken>, we set up <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="811:3:3" line-data="    With frmRatePerPeriod">`frmRatePerPeriod`</SwmToken> with the current folio number, show it modally for user input, and then immediately call <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="816:3:3" line-data="        Call ComputeRate">`ComputeRate`</SwmToken> to refresh all rate-related fields based on any changes made in the dialog.

```visual basic
Private Sub lblRatePerPeriod_Click()
    With frmRatePerPeriod
        .FolioNumber = txtGuestName.Tag
        
        .Show vbModal
        
        Call ComputeRate
```

---

</SwmSnippet>

## Calculating and Formatting Checkout Values

<SwmSnippet path="/HotelManagementSystem/Forms/frmCheckOut.frm" line="718">

---

In <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="718:4:4" line-data="Private Sub ComputeRate()">`ComputeRate`</SwmToken>, we recalculate all checkout-related amounts, converting results to formatted currency strings for display. The next step is to call utility functions from <SwmPath>[HotelManagementSystem/Modules/modFunction.bas](HotelManagementSystem/Modules/modFunction.bas)</SwmPath> to handle the formatting and parsing cleanly.

```visual basic
Private Sub ComputeRate()
    txtTotalCharges.Text = toMoney(ComputeRatePerPeriod)
    txtSubTotal.Text = toMoney(toNumber(txtTotalCharges.Text) + toNumber(txtOtherCharges.Text))
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Modules/modFunction.bas" line="216">

---

<SwmToken path="HotelManagementSystem/Modules/modFunction.bas" pos="216:4:4" line-data="Public Function toMoney(ByVal srcCurr As String) As String">`toMoney`</SwmToken> handles converting any string (even empty) into a standardized currency format for display

```visual basic
Public Function toMoney(ByVal srcCurr As String) As String
   toMoney = Format$(IIf(Trim(srcCurr) = "", 0, srcCurr), "#,##0.00")
End Function
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Forms/frmCheckOut.frm" line="720">

---

Back in <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="718:4:4" line-data="Private Sub ComputeRate()">`ComputeRate`</SwmToken>, after formatting total charges, we sum them with other charges (using <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="720:9:9" line-data="    txtSubTotal.Text = toMoney(toNumber(txtTotalCharges.Text) + toNumber(txtOtherCharges.Text))">`toNumber`</SwmToken> for parsing), then format the subtotal for display with <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="720:7:7" line-data="    txtSubTotal.Text = toMoney(toNumber(txtTotalCharges.Text) + toNumber(txtOtherCharges.Text))">`toMoney`</SwmToken>. This keeps calculations accurate and the UI clean.

```visual basic
    txtSubTotal.Text = toMoney(toNumber(txtTotalCharges.Text) + toNumber(txtOtherCharges.Text))
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Modules/modFunction.bas" line="182">

---

<SwmToken path="HotelManagementSystem/Modules/modFunction.bas" pos="182:4:4" line-data="Public Function toNumber(ByVal srcCurrency As String, Optional RetZeroIfNegative As Boolean) As Double">`toNumber`</SwmToken> parses currency strings into numbers, stripping out commas so calculations don't break on formatted input. It also optionally forces negatives and zero to become zero if needed.

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

Back in <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="718:4:4" line-data="Private Sub ComputeRate()">`ComputeRate`</SwmToken>, after parsing and summing values, we apply the discount as a percentage, then format the result for display. This keeps the UI in sync with user-entered discounts.

```visual basic
    txtTotal.Text = toMoney(toNumber(txtSubTotal.Text) - (toNumber(txtSubTotal.Text) * toNumber(txtDiscount.Text) / 100))
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Forms/frmCheckOut.frm" line="721">

---

Back in <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="718:4:4" line-data="Private Sub ComputeRate()">`ComputeRate`</SwmToken>, after calculating the total with discount, we use <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="721:7:7" line-data="    txtTotal.Text = toMoney(toNumber(txtSubTotal.Text) - (toNumber(txtSubTotal.Text) * toNumber(txtDiscount.Text) / 100))">`toMoney`</SwmToken> again to ensure the total is displayed as a formatted currency string.

```visual basic
    txtTotal.Text = toMoney(toNumber(txtSubTotal.Text) - (toNumber(txtSubTotal.Text) * toNumber(txtDiscount.Text) / 100))
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Forms/frmCheckOut.frm" line="722">

---

Back in <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="718:4:4" line-data="Private Sub ComputeRate()">`ComputeRate`</SwmToken>, we subtract the amount paid from the total to get the balance, then format it for display. This step ensures the UI always shows the up-to-date balance.

```visual basic
    txtBalance.Text = toMoney(toNumber(txtTotal.Text) - toNumber(txtAmountPaid.Text))
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Forms/frmCheckOut.frm" line="722">

---

Back in <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="718:4:4" line-data="Private Sub ComputeRate()">`ComputeRate`</SwmToken>, after calculating the balance, we use <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="722:7:7" line-data="    txtBalance.Text = toMoney(toNumber(txtTotal.Text) - toNumber(txtAmountPaid.Text))">`toMoney`</SwmToken> to keep the display consistent and readable for the user.

```visual basic
    txtBalance.Text = toMoney(toNumber(txtTotal.Text) - toNumber(txtAmountPaid.Text))
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Forms/frmCheckOut.frm" line="723">

---

Finally, <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="718:4:4" line-data="Private Sub ComputeRate()">`ComputeRate`</SwmToken> wraps up after updating all fields

```visual basic
End Sub
```

---

</SwmSnippet>

## Completing the Rate Update Interaction

<SwmSnippet path="/HotelManagementSystem/Forms/frmCheckOut.frm" line="817">

---

Back in <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="810:4:4" line-data="Private Sub lblRatePerPeriod_Click()">`lblRatePerPeriod_Click`</SwmToken>, after <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="718:4:4" line-data="Private Sub ComputeRate()">`ComputeRate`</SwmToken> runs, the function ends, leaving the UI updated with any changes from the modal dialog.

```visual basic
    End With
End Sub
```

---

</SwmSnippet>

&nbsp;

*This is an auto-generated document by Swimm 🌊 and has not yet been verified by a human*

<SwmMeta version="3.0.0" repo-id="Z2l0aHViJTNBJTNBY3RzLVZCNi1Qcm9qZWN0cyUzQSUzQVN3aW1tLURlbW8=" repo-name="cts-VB6-Projects"><sup>Powered by [Swimm](https://app.swimm.io/)</sup></SwmMeta>
