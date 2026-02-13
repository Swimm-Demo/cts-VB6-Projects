---
title: Updating Balance on Payment Change
---
This document explains how the checkout form recalculates and displays the remaining balance when the user updates the amount paid. The system processes the input, computes the new balance, formats it as currency, and updates the display to provide immediate feedback.

# Calculating and Formatting Balance on Payment Change

<SwmSnippet path="/HotelManagementSystem/Forms/frmCheckOut.frm" line="825">

---

In <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="825:4:4" line-data="Private Sub txtAmountPaid_Change()">`txtAmountPaid_Change`</SwmToken>, we trigger the recalculation of the balance whenever the user updates the amount paid. This means we need to convert the total and paid amounts from text to numbers, subtract them, and then format the result as currency for display. To do this, we call utility functions in <SwmPath>[HotelManagementSystem/Modules/modFunction.bas](HotelManagementSystem/Modules/modFunction.bas)</SwmPath> to handle the conversions and formatting.

```visual basic
Private Sub txtAmountPaid_Change()
    txtBalance.Text = toMoney(toNumber(txtTotal.Text) - toNumber(txtAmountPaid.Text))
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Modules/modFunction.bas" line="216">

---

<SwmToken path="HotelManagementSystem/Modules/modFunction.bas" pos="216:4:4" line-data="Public Function toMoney(ByVal srcCurr As String) As String">`toMoney`</SwmToken> handles converting a string (possibly empty) into a properly formatted currency string. It ensures that even if the input is blank, the output is always a valid currency value, using '#,##<SwmToken path="HotelManagementSystem/Modules/modFunction.bas" pos="217:30:32" line-data="   toMoney = Format$(IIf(Trim(srcCurr) = &quot;&quot;, 0, srcCurr), &quot;#,##0.00&quot;)">`0.00`</SwmToken>' formatting.

```visual basic
Public Function toMoney(ByVal srcCurr As String) As String
   toMoney = Format$(IIf(Trim(srcCurr) = "", 0, srcCurr), "#,##0.00")
End Function
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Forms/frmCheckOut.frm" line="826">

---

Back in <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="825:4:4" line-data="Private Sub txtAmountPaid_Change()">`txtAmountPaid_Change`</SwmToken>, after converting the text inputs to numbers and subtracting, we immediately format the result as currency before displaying it. This keeps the UI consistent and user-friendly.

```visual basic
    txtBalance.Text = toMoney(toNumber(txtTotal.Text) - toNumber(txtAmountPaid.Text))
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Modules/modFunction.bas" line="182">

---

<SwmToken path="HotelManagementSystem/Modules/modFunction.bas" pos="182:4:4" line-data="Public Function toNumber(ByVal srcCurrency As String, Optional RetZeroIfNegative As Boolean) As Double">`toNumber`</SwmToken> converts a currency string (with or without commas) into a double for calculation. It also optionally forces negative or zero values to become zero if requested.

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

<SwmSnippet path="/HotelManagementSystem/Forms/frmCheckOut.frm" line="827">

---

Finally, in <SwmToken path="HotelManagementSystem/Forms/frmCheckOut.frm" pos="825:4:4" line-data="Private Sub txtAmountPaid_Change()">`txtAmountPaid_Change`</SwmToken>, after all conversions and formatting, the balance textbox is updated instantly to reflect the latest calculation, keeping the UI in sync with user input.

```visual basic
End Sub
```

---

</SwmSnippet>

&nbsp;

*This is an auto-generated document by Swimm 🌊 and has not yet been verified by a human*

<SwmMeta version="3.0.0" repo-id="Z2l0aHViJTNBJTNBY3RzLVZCNi1Qcm9qZWN0cyUzQSUzQVN3aW1tLURlbW8=" repo-name="cts-VB6-Projects"><sup>Powered by [Swimm](https://app.swimm.io/)</sup></SwmMeta>
