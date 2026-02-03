---
title: Preparing the Client Account Form
---
This document describes how the client account form is prepared for user interaction. When the form loads, it retrieves client and bank details, fills dropdowns for category and city, and sets up the form for either adding a new client or editing an existing one. The form is populated with all necessary data to allow users to efficiently add or update client information.

# Loading Client Data and Populating Dropdowns

<SwmSnippet path="/HotelManagementSystem/Forms/frmAccounts.frm" line="406">

---

In <SwmToken path="HotelManagementSystem/Forms/frmAccounts.frm" pos="406:4:4" line-data="Private Sub Form_Load()">`Form_Load`</SwmToken>, we start by opening recordsets for the main client and their bank details using the current PK. This pulls in all the relevant data for the form. Right after, we call <SwmToken path="HotelManagementSystem/Forms/frmAccounts.frm" pos="413:1:1" line-data="    bind_dc &quot;SELECT * FROM Clients_Category&quot;, &quot;Category&quot;, dcCategory, &quot;CategoryID&quot;, True">`bind_dc`</SwmToken> twice to fill the Category and City dropdowns with up-to-date lists from the database. We need to call into <SwmPath>[HotelManagementSystem/Modules/modProcedure.bas](HotelManagementSystem/Modules/modProcedure.bas)</SwmPath> here because that's where the actual dropdown binding logic lives, so the controls get populated with the right data for user interaction.

```visual basic
Private Sub Form_Load()
    RS.CursorLocation = adUseClient
    RS.Open "SELECT * FROM Clients WHERE ClientID = " & PK, CN, adOpenStatic, adLockOptimistic
        
    rsClientBank.CursorLocation = adUseClient
    rsClientBank.Open "SELECT * FROM qry_Clients_Bank WHERE ClientID = " & PK, CN, adOpenStatic, adLockOptimistic
    
    bind_dc "SELECT * FROM Clients_Category", "Category", dcCategory, "CategoryID", True
    bind_dc "SELECT * FROM Cities", "City", dcCity, "CityID", True
   
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Modules/modProcedure.bas" line="180">

---

<SwmToken path="HotelManagementSystem/Modules/modProcedure.bas" pos="180:4:4" line-data="Public Sub bind_dc(ByVal srcSQL As String, ByVal srcBindField As String, ByRef srcDC As DataCombo, Optional srcColBound As String, Optional ShowFirstRec As Boolean)">`bind_dc`</SwmToken> takes care of wiring up a <SwmToken path="HotelManagementSystem/Modules/modProcedure.bas" pos="180:30:30" line-data="Public Sub bind_dc(ByVal srcSQL As String, ByVal srcBindField As String, ByRef srcDC As DataCombo, Optional srcColBound As String, Optional ShowFirstRec As Boolean)">`DataCombo`</SwmToken> to a SQL query result. It sets which field to show, which field to bind, and connects the recordset as the data source. If <SwmToken path="HotelManagementSystem/Modules/modProcedure.bas" pos="180:44:44" line-data="Public Sub bind_dc(ByVal srcSQL As String, ByVal srcBindField As String, ByRef srcDC As DataCombo, Optional srcColBound As String, Optional ShowFirstRec As Boolean)">`ShowFirstRec`</SwmToken> is true, it pre-selects the first record and tags the control with the record count and value, or marks it as empty if there are no records.

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

<SwmSnippet path="/HotelManagementSystem/Forms/frmAccounts.frm" line="416">

---

After coming back from <SwmPath>[HotelManagementSystem/Modules/modProcedure.bas](HotelManagementSystem/Modules/modProcedure.bas)</SwmPath>, <SwmToken path="HotelManagementSystem/Forms/frmAccounts.frm" pos="406:4:4" line-data="Private Sub Form_Load()">`Form_Load`</SwmToken> checks if we're adding or editing. If we're adding or in popup mode, it updates the caption, disables user history, and calls <SwmToken path="HotelManagementSystem/Forms/frmAccounts.frm" pos="421:1:1" line-data="        GeneratePK">`GeneratePK`</SwmToken> to get a new unique primary key for the new client. If not, it preps the form for editing instead.

```visual basic
    'Check the form state
    If State = adStateAddMode Or State = adStatePopupMode Then
        Caption = "Create New Entry"
        cmdUsrHistory.Enabled = False
        
        GeneratePK
    Else
        Caption = "Edit Entry"
```

---

</SwmSnippet>

## Generating a Unique Client Identifier

<SwmSnippet path="/HotelManagementSystem/Forms/frmAccounts.frm" line="431">

---

<SwmToken path="HotelManagementSystem/Forms/frmAccounts.frm" pos="431:4:4" line-data="Private Sub GeneratePK()">`GeneratePK`</SwmToken> just calls <SwmToken path="HotelManagementSystem/Forms/frmAccounts.frm" pos="432:5:10" line-data="    PK = getIndex(&quot;Clients&quot;)">`getIndex("Clients")`</SwmToken> to fetch the next available unique PK for a new client. The actual logic for generating and updating the PK is handled in <SwmPath>[HotelManagementSystem/Modules/modADO.bas](HotelManagementSystem/Modules/modADO.bas)</SwmPath>.

```visual basic
Private Sub GeneratePK()
    PK = getIndex("Clients")
End Sub
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Modules/modADO.bas" line="35">

---

<SwmToken path="HotelManagementSystem/Modules/modADO.bas" pos="35:4:4" line-data="Public Function getIndex(ByVal srcTable As String) As Long">`getIndex`</SwmToken> looks up the next PK for a table from the KEY GENERATOR table, increments it in a transaction, and returns the old value as the new PK. If the value is missing, it defaults to 1. This keeps PKs unique and avoids collisions.

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

## Finalizing Form State and Data Loading

<SwmSnippet path="/HotelManagementSystem/Forms/frmAccounts.frm" line="424">

---

After coming back from <SwmToken path="HotelManagementSystem/Forms/frmAccounts.frm" pos="421:1:1" line-data="        GeneratePK">`GeneratePK`</SwmToken>, <SwmToken path="HotelManagementSystem/Forms/frmAccounts.frm" pos="406:4:4" line-data="Private Sub Form_Load()">`Form_Load`</SwmToken> finishes up by either prepping the form for a new entry or, if we're editing, calling <SwmToken path="HotelManagementSystem/Forms/frmAccounts.frm" pos="424:1:1" line-data="        DisplayForEditing">`DisplayForEditing`</SwmToken> to load the existing client data into the UI and enabling the <SwmToken path="HotelManagementSystem/Forms/frmAccounts.frm" pos="425:1:1" line-data="        cmdPH.Enabled = True">`cmdPH`</SwmToken> button.

```visual basic
        DisplayForEditing
        cmdPH.Enabled = True
    End If

End Sub
```

---

</SwmSnippet>

# Loading and Displaying Existing Client Details

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
    node1["Display client details in form fields"]
    click node1 openCode "HotelManagementSystem/Forms/frmAccounts.frm:188:203"
    node1 --> node2{"Are there any bank accounts?"}
    click node2 openCode "HotelManagementSystem/Forms/frmAccounts.frm:213:214"
    node2 -->|"Yes"| loop1
    node2 -->|"No"| node4{"Is form in edit mode?"}
    
    subgraph loop1["For each bank account"]
        node3["Display bank account in grid"]
        click node3 openCode "HotelManagementSystem/Forms/frmAccounts.frm:215:233"
    end
    loop1 --> node4{"Is form in edit mode?"}
    click node4 openCode "HotelManagementSystem/Forms/frmAccounts.frm:238:241"
    node4 -->|"Yes"| node5["Adjust grid for editing"]
    click node5 openCode "HotelManagementSystem/Forms/frmAccounts.frm:239:240"
    node5 --> node6["Finish"]
    click node6 openCode "HotelManagementSystem/Forms/frmAccounts.frm:244:249"
    node4 -->|"No"| node6

classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;

%% Swimm:
%% %%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
%% flowchart TD
%%     node1["Display client details in form fields"]
%%     click node1 openCode "<SwmPath>[HotelManagementSystem/Forms/frmAccounts.frm](HotelManagementSystem/Forms/frmAccounts.frm)</SwmPath>:188:203"
%%     node1 --> node2{"Are there any bank accounts?"}
%%     click node2 openCode "<SwmPath>[HotelManagementSystem/Forms/frmAccounts.frm](HotelManagementSystem/Forms/frmAccounts.frm)</SwmPath>:213:214"
%%     node2 -->|"Yes"| loop1
%%     node2 -->|"No"| node4{"Is form in edit mode?"}
%%     
%%     subgraph loop1["For each bank account"]
%%         node3["Display bank account in grid"]
%%         click node3 openCode "<SwmPath>[HotelManagementSystem/Forms/frmAccounts.frm](HotelManagementSystem/Forms/frmAccounts.frm)</SwmPath>:215:233"
%%     end
%%     loop1 --> node4{"Is form in edit mode?"}
%%     click node4 openCode "<SwmPath>[HotelManagementSystem/Forms/frmAccounts.frm](HotelManagementSystem/Forms/frmAccounts.frm)</SwmPath>:238:241"
%%     node4 -->|"Yes"| node5["Adjust grid for editing"]
%%     click node5 openCode "<SwmPath>[HotelManagementSystem/Forms/frmAccounts.frm](HotelManagementSystem/Forms/frmAccounts.frm)</SwmPath>:239:240"
%%     node5 --> node6["Finish"]
%%     click node6 openCode "<SwmPath>[HotelManagementSystem/Forms/frmAccounts.frm](HotelManagementSystem/Forms/frmAccounts.frm)</SwmPath>:244:249"
%%     node4 -->|"No"| node6
%% 
%% classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/HotelManagementSystem/Forms/frmAccounts.frm" line="181">

---

<SwmToken path="HotelManagementSystem/Forms/frmAccounts.frm" pos="181:4:4" line-data="Private Sub DisplayForEditing()">`DisplayForEditing`</SwmToken> loads client info into the form fields and pulls related bank records into the grid. It handles grid row logic to avoid blanks and sets grid properties for edit mode. If something goes wrong, it calls <SwmToken path="HotelManagementSystem/Forms/frmAccounts.frm" pos="253:1:1" line-data="    prompt_err err, Name, &quot;DisplayForEditing&quot;">`prompt_err`</SwmToken> for error handling.

```visual basic
Private Sub DisplayForEditing()
    On Error GoTo err
    Dim rsClients As New Recordset
    
    rsClients.CursorLocation = adUseClient
    rsClients.Open "SELECT * FROM qry_Clients WHERE ClientID = " & PK, CN, adOpenStatic, adLockOptimistic
    
    With rsClients
        txtEntry(1).Text = .Fields("Company")
        dcCategory.BoundText = .Fields![CategoryID]
        txtEntry(2).Text = .Fields("Tin")
        txtEntry(3).Text = .Fields("OwnersName")
        txtEntry(4).Text = .Fields("Address")
        dcCity.BoundText = .Fields![CityID]
        txtEntry(6).Text = .Fields("PurchaserName")
        txtEntry(7).Text = .Fields("Mobile")
        txtEntry(8).Text = .Fields("Landline")
        txtEntry(9).Text = .Fields("Fax")
        txtEntry(14).Text = .Fields("CreditTerm")
        txtEntry(15).Text = .Fields("CreditLimit")
        chkBlackListed.Value = IIf(.Fields("BlackListed") = True, 1, 0)
        txtEntry(16).Text = .Fields("Remarks")
    End With
    
    'Display the details
    Dim rsClientBank As New Recordset

    cIRowCount = 0
    
    rsClientBank.CursorLocation = adUseClient
    rsClientBank.Open "SELECT * FROM qry_Clients_Bank WHERE ClientID=" & PK, CN, adOpenStatic, adLockOptimistic
    
    If rsClientBank.RecordCount > 0 Then
        rsClientBank.MoveFirst
        While Not rsClientBank.EOF
          cIRowCount = cIRowCount + 1     'increment
            With Grid
                If .Rows = 2 And .TextMatrix(1, 5) = "" Then
                    .TextMatrix(1, 1) = rsClientBank![Bank]
                    .TextMatrix(1, 2) = rsClientBank![Branch]
                    .TextMatrix(1, 3) = rsClientBank![AccountNo]
                    .TextMatrix(1, 4) = rsClientBank![AccountName]
                    .TextMatrix(1, 5) = rsClientBank![BankID]
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = rsClientBank![Bank]
                    .TextMatrix(.Rows - 1, 2) = rsClientBank![Branch]
                    .TextMatrix(.Rows - 1, 3) = rsClientBank![AccountNo]
                    .TextMatrix(.Rows - 1, 4) = rsClientBank![AccountName]
                    .TextMatrix(.Rows - 1, 5) = rsClientBank![BankID]
                End If
            End With
            rsClientBank.MoveNext
        Wend
        Grid.Row = 1
        Grid.ColSel = 5
        'Set fixed cols
        If State = adStateEditMode Then
            Grid.FixedRows = Grid.Row: 'Grid.SelectionMode = flexSelectionFree
            Grid.FixedCols = 1
        End If
    End If

    rsClientBank.Close
    'Clear variables
    Set rsClientBank = Nothing
        
    'txtEntry(1).SetFocus
    Exit Sub
err:
    If err.Number = 94 Then Resume Next
    
    prompt_err err, Name, "DisplayForEditing"
    Screen.MousePointer = vbDefault
End Sub
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Modules/modProcedure.bas" line="87">

---

<SwmToken path="HotelManagementSystem/Modules/modProcedure.bas" pos="87:4:4" line-data="Public Sub prompt_err(ByVal sError As ErrObject, ByVal ModuleName As String, ByVal OccurIn As String)">`prompt_err`</SwmToken> pops up a message box with error details and appends the info to <SwmPath>[HotelManagementSystem/Error.log](HotelManagementSystem/Error.log)</SwmPath> for later review. It logs the date, time, error number, description, module, and location.

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
