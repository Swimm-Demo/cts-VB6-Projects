---
title: Saving and Updating Client and Bank Details
---
This document describes how users can save or update client and bank details using a form. The system ensures data integrity by wrapping the operation in a transaction. After saving, users receive feedback and can choose to add another record, which resets the form for new input.

# Saving Client and Bank Details with Transaction Handling

<SwmSnippet path="/HotelManagementSystem/Forms/frmAccounts.frm" line="268">

---

In <SwmToken path="HotelManagementSystem/Forms/frmAccounts.frm" pos="268:4:4" line-data="Private Sub cmdSave_Click()">`cmdSave_Click`</SwmToken>, we handle adding or editing client and bank records, using <SwmToken path="HotelManagementSystem/Forms/frmAccounts.frm" pos="296:11:11" line-data="      .Fields(&quot;CreditTerm&quot;) = toNumber(txtEntry(14).Text)">`toNumber`</SwmToken> to clean up numeric input before saving. We clear and repopulate bank details, and wrap everything in a transaction for safety.

```visual basic
Private Sub cmdSave_Click()
    On Error GoTo err

    If Trim(txtEntry(1).Text) = "" Then Exit Sub
    
    CN.BeginTrans

    If State = adStateAddMode Or State = adStatePopupMode Then
        RS.AddNew
        
        RS.Fields("ClientID") = PK
        RS.Fields("AddedByFK") = CurrUser.USER_PK
    Else
        RS.Fields("DateModified") = Now
        RS.Fields("LastUserFK") = CurrUser.USER_PK
    End If
    
    With RS
      .Fields("Company") = txtEntry(1).Text
      .Fields("CategoryID") = dcCategory.BoundText
      .Fields("Tin") = txtEntry(2).Text
      .Fields("OwnersName") = txtEntry(3).Text
      .Fields("Address") = txtEntry(4).Text
      .Fields("CityID") = dcCity.BoundText
      .Fields("PurchaserName") = txtEntry(6).Text
      .Fields("Mobile") = txtEntry(7).Text
      .Fields("Landline") = txtEntry(8).Text
      .Fields("Fax") = txtEntry(9).Text
      .Fields("CreditTerm") = toNumber(txtEntry(14).Text)
      .Fields("CreditLimit") = toNumber(txtEntry(15).Text)
      .Fields("BlackListed") = IIf(chkBlackListed.Value = 1, True, False)
      .Fields("Remarks") = txtEntry(16).Text
       
      .Update
    End With

    Dim rsClientBank As New Recordset

    rsClientBank.CursorLocation = adUseClient
    rsClientBank.Open "SELECT * FROM Clients_Bank WHERE ClientID=" & PK, CN, adOpenStatic, adLockOptimistic
    
    DeleteItems
    
    Dim c As Integer
    
    With Grid
        'Save the details of the records
        For c = 1 To cIRowCount
            .Row = c
            If State = adStateAddMode Or State = adStatePopupMode Then
AddNew:
                'Add qty received in Local Purchase Details
                rsClientBank.AddNew

                rsClientBank![ClientID] = PK
                rsClientBank![BankID] = toNumber(.TextMatrix(c, 5))
                rsClientBank![AccountNo] = .TextMatrix(c, 3)
                rsClientBank![AccountName] = .TextMatrix(c, 4)

                rsClientBank.Update
            ElseIf State = adStateEditMode Then
                rsClientBank.Filter = "BankID = " & toNumber(.TextMatrix(c, 5))
            
                If rsClientBank.RecordCount = 0 Then GoTo AddNew

                rsClientBank![ClientID] = PK
                rsClientBank![BankID] = toNumber(.TextMatrix(c, 5))
                rsClientBank![AccountNo] = .TextMatrix(c, 3)
                rsClientBank![AccountName] = .TextMatrix(c, 4)

                rsClientBank.Update
            End If

        Next c
    End With

```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Modules/modFunction.bas" line="182">

---

<SwmToken path="HotelManagementSystem/Modules/modFunction.bas" pos="182:4:4" line-data="Public Function toNumber(ByVal srcCurrency As String, Optional RetZeroIfNegative As Boolean) As Double">`toNumber`</SwmToken> takes a string (possibly with commas) and converts it to a double. If the optional <SwmToken path="HotelManagementSystem/Modules/modFunction.bas" pos="182:17:17" line-data="Public Function toNumber(ByVal srcCurrency As String, Optional RetZeroIfNegative As Boolean) As Double">`RetZeroIfNegative`</SwmToken> is set, it returns 0 for anything less than 1. This is mainly to sanitize user input for numeric fields, so we don't end up with weird values in the DB.

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

<SwmSnippet path="/HotelManagementSystem/Forms/frmAccounts.frm" line="344">

---

After coming back from <SwmToken path="HotelManagementSystem/Forms/frmAccounts.frm" pos="296:11:11" line-data="      .Fields(&quot;CreditTerm&quot;) = toNumber(txtEntry(14).Text)">`toNumber`</SwmToken>, we commit the transaction, set a flag, and show the user a message box. If we're adding a new record, we ask if they want to add another—if yes, we call <SwmToken path="HotelManagementSystem/Forms/frmAccounts.frm" pos="355:1:1" line-data="            ResetFields">`ResetFields`</SwmToken> to clear the form for the next entry. Otherwise, or in other modes, we just unload the form. <SwmToken path="HotelManagementSystem/Forms/frmAccounts.frm" pos="355:1:1" line-data="            ResetFields">`ResetFields`</SwmToken> is used here to keep the clearing logic in one place and avoid repeating code.

```visual basic
    'Clear variables
    c = 0
    Set rsClientBank = Nothing
    
    CN.CommitTrans

    HaveAction = True
    
    If State = adStateAddMode Then
        MsgBox "New record has been successfully saved.", vbInformation
        If MsgBox("Do you want to add another new record?", vbQuestion + vbYesNo) = vbYes Then
            ResetFields
         Else
            Unload Me
        End If
    ElseIf State = adStatePopupMode Then
        MsgBox "New record has been successfully saved.", vbInformation
        Unload Me
    Else
        MsgBox "Changes in  record has been successfully saved.", vbInformation
        Unload Me
    End If

    Exit Sub

err:
    CN.RollbackTrans
```

---

</SwmSnippet>

## Resetting Form Fields for New Entry

<SwmSnippet path="/HotelManagementSystem/Forms/frmAccounts.frm" line="261">

---

<SwmToken path="HotelManagementSystem/Forms/frmAccounts.frm" pos="261:4:4" line-data="Private Sub ResetFields()">`ResetFields`</SwmToken> calls <SwmToken path="HotelManagementSystem/Forms/frmAccounts.frm" pos="262:1:1" line-data="  clearText Me">`clearText`</SwmToken> to wipe all text fields, sets <SwmToken path="HotelManagementSystem/Forms/frmAccounts.frm" pos="264:1:1" line-data="  txtEntry(15).Text = &quot;0.00&quot;">`txtEntry`</SwmToken>(15) to <SwmToken path="HotelManagementSystem/Forms/frmAccounts.frm" pos="264:11:13" line-data="  txtEntry(15).Text = &quot;0.00&quot;">`0.00`</SwmToken> (probably a money field), and puts the cursor back in <SwmToken path="HotelManagementSystem/Forms/frmAccounts.frm" pos="264:1:1" line-data="  txtEntry(15).Text = &quot;0.00&quot;">`txtEntry`</SwmToken>(1) so the user can start typing right away. <SwmToken path="HotelManagementSystem/Forms/frmAccounts.frm" pos="262:1:1" line-data="  clearText Me">`clearText`</SwmToken> (from <SwmPath>[HotelManagementSystem/Modules/modProcedure.bas](HotelManagementSystem/Modules/modProcedure.bas)</SwmPath>) is called to handle the bulk clearing logic.

```visual basic
Private Sub ResetFields()
  clearText Me
  
  txtEntry(15).Text = "0.00"
  txtEntry(1).SetFocus
End Sub
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Modules/modProcedure.bas" line="228">

---

<SwmToken path="HotelManagementSystem/Modules/modProcedure.bas" pos="228:4:4" line-data="Public Sub clearText(ByRef sForm As Form)">`clearText`</SwmToken> loops through all controls on the form and tries to clear TextBoxes by setting the control to <SwmToken path="HotelManagementSystem/Modules/modProcedure.bas" pos="231:19:19" line-data="        If (TypeOf Control Is TextBox) Then Control = vbNullString">`vbNullString`</SwmToken>. But in VB6, this doesn't actually clear the text—you'd need to set Control.Text instead. So, this function doesn't really do what it's supposed to.

```visual basic
Public Sub clearText(ByRef sForm As Form)
    Dim Control As Control
    For Each Control In sForm.Controls
        If (TypeOf Control Is TextBox) Then Control = vbNullString
    Next Control
    Set Control = Nothing
End Sub
```

---

</SwmSnippet>

## Error Handling and Logging After Save

<SwmSnippet path="/HotelManagementSystem/Forms/frmAccounts.frm" line="371">

---

Back in <SwmToken path="HotelManagementSystem/Forms/frmAccounts.frm" pos="371:10:10" line-data="    prompt_err err, Name, &quot;cmdSave_Click&quot;">`cmdSave_Click`</SwmToken>, after <SwmToken path="HotelManagementSystem/Forms/frmAccounts.frm" pos="261:4:4" line-data="Private Sub ResetFields()">`ResetFields`</SwmToken> (or closing the form), if an error was caught, we call <SwmToken path="HotelManagementSystem/Forms/frmAccounts.frm" pos="371:1:1" line-data="    prompt_err err, Name, &quot;cmdSave_Click&quot;">`prompt_err`</SwmToken> from <SwmPath>[HotelManagementSystem/Modules/modProcedure.bas](HotelManagementSystem/Modules/modProcedure.bas)</SwmPath> to show the error and log it. This helps with debugging and makes sure we don't miss silent failures.

```visual basic
    prompt_err err, Name, "cmdSave_Click"
    Screen.MousePointer = vbDefault
End Sub
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Modules/modProcedure.bas" line="87">

---

<SwmToken path="HotelManagementSystem/Modules/modProcedure.bas" pos="87:4:4" line-data="Public Sub prompt_err(ByVal sError As ErrObject, ByVal ModuleName As String, ByVal OccurIn As String)">`prompt_err`</SwmToken> pops up a message box with error details (module, location, number, description) and then logs the same info to <SwmPath>[HotelManagementSystem/Error.log](HotelManagementSystem/Error.log)</SwmPath> in the app folder. This way, you get immediate feedback and a persistent error history for debugging.

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
