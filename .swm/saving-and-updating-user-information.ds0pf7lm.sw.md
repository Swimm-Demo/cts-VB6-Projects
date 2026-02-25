---
title: Saving and Updating User Information
---
This document describes the process of saving or updating user information and permissions. User input is validated for completeness. For new users, permissions are initialized for all forms. The user record is saved or updated in the database, and the system offers to reset the form for another entry.

# Saving User Data and Permissions

<SwmSnippet path="/HotelManagementSystem/Forms/frmUsers.frm" line="162">

---

In <SwmToken path="HotelManagementSystem/Forms/frmUsers.frm" pos="162:4:4" line-data="Private Sub cmdSave_Click()">`cmdSave_Click`</SwmToken>, we check the first three input fields for empty values using <SwmToken path="HotelManagementSystem/Forms/frmUsers.frm" pos="163:3:3" line-data="    If is_empty(txtEntry(0), True) = True Then Exit Sub">`is_empty`</SwmToken>. If any are empty, we bail out immediately. This prevents incomplete user data from being processed. Next, we call the validation logic in <SwmPath>[HotelManagementSystem/Modules/modFunction.bas](HotelManagementSystem/Modules/modFunction.bas)</SwmPath> to handle the actual check and user notification.

```visual basic
Private Sub cmdSave_Click()
    If is_empty(txtEntry(0), True) = True Then Exit Sub
    If is_empty(txtEntry(1), True) = True Then Exit Sub
    If is_empty(txtEntry(2), True) = True Then Exit Sub
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Modules/modFunction.bas" line="122">

---

<SwmToken path="HotelManagementSystem/Modules/modFunction.bas" pos="122:4:4" line-data="Public Function is_empty(ByRef sText As Variant, Optional UseTagValue As Boolean) As Boolean">`is_empty`</SwmToken> checks if a field is blank, pops up a message if needed, and puts the cursor back in the empty field. This stops the flow until the user fixes the input.

```visual basic
Public Function is_empty(ByRef sText As Variant, Optional UseTagValue As Boolean) As Boolean
    On Error Resume Next
    If sText.Text = "" Then
        is_empty = True
        If UseTagValue = True Then
            MsgBox "The field '" & sText.Tag & "' is required.Please check it!", vbExclamation
        Else
            MsgBox "The field is required.Please check it!", vbExclamation
        End If
        sText.SetFocus
    Else
        is_empty = False
    End If
End Function
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Forms/frmUsers.frm" line="166">

---

Back in <SwmToken path="HotelManagementSystem/Forms/frmUsers.frm" pos="162:4:4" line-data="Private Sub cmdSave_Click()">`cmdSave_Click`</SwmToken>, after validation, we check if we're adding or editing. For new users, we set up the record and call <SwmToken path="HotelManagementSystem/Forms/frmUsers.frm" pos="173:3:3" line-data="        Call AddPermission">`AddPermission`</SwmToken> to initialize their permissions. For edits, we just update the modified fields.

```visual basic
    
    If State = adStateAddMode Then
        RS.AddNew
        RS.Fields("PK") = PK
        RS.Fields("DateAdded") = Now
        RS.Fields("AddedByFK") = CurrUser.USER_PK
        
        Call AddPermission
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Forms/frmUsers.frm" line="257">

---

<SwmToken path="HotelManagementSystem/Forms/frmUsers.frm" pos="257:4:4" line-data="Public Sub AddPermission()">`AddPermission`</SwmToken> creates a permission entry for the new user for every form in the system by running a single SQL insert. This sets up the baseline permissions structure for the user.

```visual basic
Public Sub AddPermission()
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "INSERT INTO [User Permission] ( UserID, FormID ) " _
            & "SELECT '" & Me.txtEntry(0).Text & "', Form.FormID " _
            & "FROM Form"

    CN.Execute sSQL
    
    Exit Sub
    
RAE:
    Set vRS = Nothing
End Sub
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Forms/frmUsers.frm" line="174">

---

After <SwmToken path="HotelManagementSystem/Forms/frmUsers.frm" pos="173:3:3" line-data="        Call AddPermission">`AddPermission`</SwmToken>, still in <SwmToken path="HotelManagementSystem/Forms/frmUsers.frm" pos="162:4:4" line-data="Private Sub cmdSave_Click()">`cmdSave_Click`</SwmToken>, we fill in the user fields from the form, encrypt the password, and convert the admin checkbox value before saving everything to the database.

```visual basic
    Else
        RS.Fields("DateModified") = Now
        RS.Fields("LastUserFK") = CurrUser.USER_PK
    End If
    'Phill 2:12
    With RS
        .Fields("UserID") = txtEntry(0).Text
        .Fields("Password") = Enc.EncryptString(txtEntry(1).Text)
        .Fields("CompleteName") = txtEntry(2).Text
        .Fields("Admin") = changeYNValue(Check1.Value)
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Modules/modFunction.bas" line="138">

---

<SwmToken path="HotelManagementSystem/Modules/modFunction.bas" pos="138:4:4" line-data="Public Function changeYNValue(ByVal srcStr As String) As String">`changeYNValue`</SwmToken> just flips between 'Y'/'N' and '1'/'0' so the admin status matches what the database expects.

```visual basic
Public Function changeYNValue(ByVal srcStr As String) As String
    Select Case srcStr
        Case "Y": changeYNValue = "1"
        Case "N": changeYNValue = "0"
        Case "1": changeYNValue = "Y"
        Case "0": changeYNValue = "N"
    End Select
End Function
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Forms/frmUsers.frm" line="184">

---

After updating and saving, <SwmToken path="HotelManagementSystem/Forms/frmUsers.frm" pos="162:4:4" line-data="Private Sub cmdSave_Click()">`cmdSave_Click`</SwmToken> sets a flag, shows a confirmation, and if the user wants to add another, calls <SwmToken path="HotelManagementSystem/Forms/frmUsers.frm" pos="192:1:1" line-data="            ResetFields">`ResetFields`</SwmToken> to prep the form for the next entry.

```visual basic
        .Update
    End With
    
    HaveAction = True
    
    If State = adStateAddMode Then
        MsgBox "New record has been successfully saved.", vbInformation
        If MsgBox("Do you want to add another new record?", vbQuestion + vbYesNo) = vbYes Then
            ResetFields
```

---

</SwmSnippet>

## Resetting the User Entry Form

<SwmSnippet path="/HotelManagementSystem/Forms/frmUsers.frm" line="156">

---

In <SwmToken path="HotelManagementSystem/Forms/frmUsers.frm" pos="156:4:4" line-data="Private Sub ResetFields()">`ResetFields`</SwmToken>, we call <SwmToken path="HotelManagementSystem/Forms/frmUsers.frm" pos="157:1:1" line-data="    clearText Me">`clearText`</SwmToken> to wipe all text fields on the form, making sure everything is blank for the next user entry.

```visual basic
Private Sub ResetFields()
    clearText Me
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Modules/modProcedure.bas" line="228">

---

<SwmToken path="HotelManagementSystem/Modules/modProcedure.bas" pos="228:4:4" line-data="Public Sub clearText(ByRef sForm As Form)">`clearText`</SwmToken> loops through the form's controls and blanks out any <SwmToken path="HotelManagementSystem/Modules/modProcedure.bas" pos="231:10:10" line-data="        If (TypeOf Control Is TextBox) Then Control = vbNullString">`TextBox`</SwmToken>, leaving other controls as-is.

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

<SwmSnippet path="/HotelManagementSystem/Forms/frmUsers.frm" line="158">

---

After <SwmToken path="HotelManagementSystem/Forms/frmUsers.frm" pos="157:1:1" line-data="    clearText Me">`clearText`</SwmToken>, <SwmToken path="HotelManagementSystem/Forms/frmUsers.frm" pos="156:4:4" line-data="Private Sub ResetFields()">`ResetFields`</SwmToken> unchecks the admin box and puts the cursor back in the first entry field so the user can start typing right away.

```visual basic
    Check1.Value = 0
    txtEntry(0).SetFocus
End Sub
```

---

</SwmSnippet>

## Preparing for the Next User Entry

<SwmSnippet path="/HotelManagementSystem/Forms/frmUsers.frm" line="193">

---

After <SwmToken path="HotelManagementSystem/Forms/frmUsers.frm" pos="156:4:4" line-data="Private Sub ResetFields()">`ResetFields`</SwmToken> in <SwmToken path="HotelManagementSystem/Forms/frmUsers.frm" pos="162:4:4" line-data="Private Sub cmdSave_Click()">`cmdSave_Click`</SwmToken>, we grab a new primary key using <SwmToken path="HotelManagementSystem/Forms/frmUsers.frm" pos="193:5:5" line-data="            PK = getIndex(&quot;tbl_SM_Users&quot;)">`getIndex`</SwmToken> so the next user entry doesn't reuse the old one.

```visual basic
            PK = getIndex("tbl_SM_Users")
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Modules/modADO.bas" line="35">

---

<SwmToken path="HotelManagementSystem/Modules/modADO.bas" pos="35:4:4" line-data="Public Function getIndex(ByVal srcTable As String) As Long">`getIndex`</SwmToken> fetches and increments the next available key for the user table, making sure every new user gets a unique ID.

```visual basic
Public Function getIndex(ByVal srcTable As String) As Long
    On Error GoTo err
    Dim RS As New Recordset
    Dim RI As Long
    
    RS.CursorLocation = adUseClient
    RS.Open "SELECT * FROM [KEY GENERATOR] WHERE TableName = '" & srcTable & "'", CN, adOpenStatic, adLockOptimistic
    
    RI = RS.Fields("NextNo")
    CN.BeginTrans
    RS.Fields("NextNo") = RI + 1
    RS.Update
    CN.CommitTrans
    getIndex = RI
    
    srcTable = ""
    RI = 0
    Set RS = Nothing
    Exit Function
err:
        ''Error when incounter a null value
        If err.Number = 94 Then
            getIndex = 1
            Resume Next
        Else
            MsgBox err.Description
        End If
        CN.RollbackTrans
End Function
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Forms/frmUsers.frm" line="194">

---

After getting the new PK, <SwmToken path="HotelManagementSystem/Forms/frmUsers.frm" pos="162:4:4" line-data="Private Sub cmdSave_Click()">`cmdSave_Click`</SwmToken> either preps for another entry or closes the form, depending on the user's choice.

```visual basic
         Else
            Unload Me
        End If
    Else
        MsgBox "Changes in  record has been successfully saved.", vbInformation
        Unload Me
    End If
End Sub
```

---

</SwmSnippet>

&nbsp;

*This is an auto-generated document by Swimm 🌊 and has not yet been verified by a human*

<SwmMeta version="3.0.0" repo-id="Z2l0aHViJTNBJTNBY3RzLVZCNi1Qcm9qZWN0cyUzQSUzQVN3aW1tLURlbW8=" repo-name="cts-VB6-Projects"><sup>Powered by [Swimm](https://app.swimm.io/)</sup></SwmMeta>
