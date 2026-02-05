---
title: Viewing Customer Record Modification History
---
This document describes how users can view the modification history of a customer record within the customer management feature of the Hotel Management System. When a user requests this information, the system retrieves the dates when the record was added and last modified, looks up the names of the users who performed these actions, and presents all this information to the user.

# Showing Record Modification History

<SwmSnippet path="/HotelManagementSystem/Forms/frmCustomers.frm" line="274">

---

CmdUsrHistory_Click kicks off the flow by grabbing the <SwmToken path="HotelManagementSystem/Forms/frmCustomers.frm" pos="281:13:13" line-data="    tDate1 = Format$(RS.Fields(&quot;DateAdded&quot;), &quot;MMM-dd-yyyy HH:MM AMPM&quot;)">`DateAdded`</SwmToken> and <SwmToken path="HotelManagementSystem/Forms/frmCustomers.frm" pos="282:13:13" line-data="    tDate2 = Format$(RS.Fields(&quot;DateModified&quot;), &quot;MMM-dd-yyyy HH:MM AMPM&quot;)">`DateModified`</SwmToken> fields from the current record, formatting them for readability, and then pulling the names of the users who performed those actions using <SwmToken path="HotelManagementSystem/Forms/frmCustomers.frm" pos="284:5:5" line-data="    tUser1 = getValueAt(&quot;SELECT PK,CompleteName FROM tbl_SM_Users WHERE PK = &quot; &amp; RS.Fields(&quot;AddedByFK&quot;), &quot;CompleteName&quot;)">`getValueAt`</SwmToken>. We call <SwmToken path="HotelManagementSystem/Forms/frmCustomers.frm" pos="284:5:5" line-data="    tUser1 = getValueAt(&quot;SELECT PK,CompleteName FROM tbl_SM_Users WHERE PK = &quot; &amp; RS.Fields(&quot;AddedByFK&quot;), &quot;CompleteName&quot;)">`getValueAt`</SwmToken> next because we only have the user IDs in the main record, so we need to look up the actual names from the users table. Finally, it puts all this info together and pops it up in a message box for the user.

```visual basic
Private Sub cmdUsrHistory_Click()
    On Error Resume Next
    Dim tDate1 As String
    Dim tDate2 As String
    Dim tUser1 As String
    Dim tUser2 As String
    
    tDate1 = Format$(RS.Fields("DateAdded"), "MMM-dd-yyyy HH:MM AMPM")
    tDate2 = Format$(RS.Fields("DateModified"), "MMM-dd-yyyy HH:MM AMPM")
    
    tUser1 = getValueAt("SELECT PK,CompleteName FROM tbl_SM_Users WHERE PK = " & RS.Fields("AddedByFK"), "CompleteName")
    tUser2 = getValueAt("SELECT PK,CompleteName FROM tbl_SM_Users WHERE PK = " & RS.Fields("LastUserFK"), "CompleteName")
    
    MsgBox "Date Added: " & tDate1 & vbCrLf & _
           "Added By: " & tUser1 & vbCrLf & _
           "" & vbCrLf & _
           "Last Modified: " & tDate2 & vbCrLf & _
           "Modified By: " & tUser2, vbInformation, "Modification History"
           
    tDate1 = vbNullString
    tDate2 = vbNullString
    tUser1 = vbNullString
    tUser2 = vbNullString
End Sub
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Modules/modFunction.bas" line="168">

---

GetValueAt handles the actual DB lookup for a single field value. It runs the SQL, checks if any records came back, and if so, returns the requested field as a string. This is how we turn user IDs into display names in the main flow.

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

&nbsp;

*This is an auto-generated document by Swimm 🌊 and has not yet been verified by a human*

<SwmMeta version="3.0.0" repo-id="Z2l0aHViJTNBJTNBY3RzLVZCNi1Qcm9qZWN0cyUzQSUzQVN3aW1tLURlbW8=" repo-name="cts-VB6-Projects"><sup>Powered by [Swimm](https://app.swimm.io/)</sup></SwmMeta>
