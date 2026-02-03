---
title: Displaying formatted records in a ListView
---
This document describes how records are displayed in a <SwmToken path="HotelManagementSystem/Modules/modProcedure.bas" pos="47:12:12" line-data="Public Sub FillListView(ByRef sListView As ListView, ByRef sRecordSource As Recordset, ByVal sNumOfFields As Byte, ByVal sNumIco As Byte, ByVal with_num As Boolean, ByVal show_first_rec As Boolean, Optional srcHiddenField As String)">`ListView`</SwmToken> for users, with each item formatted for clarity and usability. The flow receives data records and formatting options, prepares the <SwmToken path="HotelManagementSystem/Modules/modProcedure.bas" pos="47:12:12" line-data="Public Sub FillListView(ByRef sListView As ListView, ByRef sRecordSource As Recordset, ByVal sNumOfFields As Byte, ByVal sNumIco As Byte, ByVal with_num As Boolean, ByVal show_first_rec As Boolean, Optional srcHiddenField As String)">`ListView`</SwmToken>, checks for available data, and adds each record as a formatted item. Formatting ensures that dates, currency, and other fields are presented in a consistent, user-friendly way.

# Populating and Formatting <SwmToken path="HotelManagementSystem/Modules/modProcedure.bas" pos="47:12:12" line-data="Public Sub FillListView(ByRef sListView As ListView, ByRef sRecordSource As Recordset, ByVal sNumOfFields As Byte, ByVal sNumIco As Byte, ByVal with_num As Boolean, ByVal show_first_rec As Boolean, Optional srcHiddenField As String)">`ListView`</SwmToken> Data

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
    node1["Start: Prepare to fill list"] --> node2{"Are there records to display?"}
    click node1 openCode "HotelManagementSystem/Modules/modProcedure.bas:47:51"
    click node2 openCode "HotelManagementSystem/Modules/modProcedure.bas:52:53"
    node2 -->|"No"| node3["Exit (nothing to display)"]
    click node3 openCode "HotelManagementSystem/Modules/modProcedure.bas:52:53"
    node2 -->|"Yes"| loop1
    subgraph loop1["For each record in data source"]
        node4{"Show item numbers? (with_num)"}
        click node4 openCode "HotelManagementSystem/Modules/modProcedure.bas:55:59"
        node4 -->|"Yes"| node5["Add item with number"]
        click node5 openCode "HotelManagementSystem/Modules/modProcedure.bas:56:56"
        node4 -->|"No"| node6["Add item with first field"]
        click node6 openCode "HotelManagementSystem/Modules/modProcedure.bas:58:58"
        node5 --> node7{"Attach hidden field? (srcHiddenField)"}
        node6 --> node7
        click node7 openCode "HotelManagementSystem/Modules/modProcedure.bas:60:60"
        node7 -->|"Yes"| node8["Attach hidden field to item"]
        click node8 openCode "HotelManagementSystem/Modules/modProcedure.bas:60:60"
        node7 -->|"No"| node9["Continue"]
        node8 --> node10["Format fields for display (FormatRS)"]
        node9 --> node10
        click node10 openCode "HotelManagementSystem/Modules/modProcedure.bas:61:79"
        click node14 openCode "HotelManagementSystem/Modules/modProcedure.bas:84:84"
    end

classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;

%% Swimm:
%% %%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
%% flowchart TD
%%     node1["Start: Prepare to fill list"] --> node2{"Are there records to display?"}
%%     click node1 openCode "<SwmPath>[HotelManagementSystem/Modules/modProcedure.bas](HotelManagementSystem/Modules/modProcedure.bas)</SwmPath>:47:51"
%%     click node2 openCode "<SwmPath>[HotelManagementSystem/Modules/modProcedure.bas](HotelManagementSystem/Modules/modProcedure.bas)</SwmPath>:52:53"
%%     node2 -->|"No"| node3["Exit (nothing to display)"]
%%     click node3 openCode "<SwmPath>[HotelManagementSystem/Modules/modProcedure.bas](HotelManagementSystem/Modules/modProcedure.bas)</SwmPath>:52:53"
%%     node2 -->|"Yes"| loop1
%%     subgraph loop1["For each record in data source"]
%%         node4{"Show item numbers? (<SwmToken path="HotelManagementSystem/Modules/modProcedure.bas" pos="47:44:44" line-data="Public Sub FillListView(ByRef sListView As ListView, ByRef sRecordSource As Recordset, ByVal sNumOfFields As Byte, ByVal sNumIco As Byte, ByVal with_num As Boolean, ByVal show_first_rec As Boolean, Optional srcHiddenField As String)">`with_num`</SwmToken>)"}
%%         click node4 openCode "<SwmPath>[HotelManagementSystem/Modules/modProcedure.bas](HotelManagementSystem/Modules/modProcedure.bas)</SwmPath>:55:59"
%%         node4 -->|"Yes"| node5["Add item with number"]
%%         click node5 openCode "<SwmPath>[HotelManagementSystem/Modules/modProcedure.bas](HotelManagementSystem/Modules/modProcedure.bas)</SwmPath>:56:56"
%%         node4 -->|"No"| node6["Add item with first field"]
%%         click node6 openCode "<SwmPath>[HotelManagementSystem/Modules/modProcedure.bas](HotelManagementSystem/Modules/modProcedure.bas)</SwmPath>:58:58"
%%         node5 --> node7{"Attach hidden field? (<SwmToken path="HotelManagementSystem/Modules/modProcedure.bas" pos="47:62:62" line-data="Public Sub FillListView(ByRef sListView As ListView, ByRef sRecordSource As Recordset, ByVal sNumOfFields As Byte, ByVal sNumIco As Byte, ByVal with_num As Boolean, ByVal show_first_rec As Boolean, Optional srcHiddenField As String)">`srcHiddenField`</SwmToken>)"}
%%         node6 --> node7
%%         click node7 openCode "<SwmPath>[HotelManagementSystem/Modules/modProcedure.bas](HotelManagementSystem/Modules/modProcedure.bas)</SwmPath>:60:60"
%%         node7 -->|"Yes"| node8["Attach hidden field to item"]
%%         click node8 openCode "<SwmPath>[HotelManagementSystem/Modules/modProcedure.bas](HotelManagementSystem/Modules/modProcedure.bas)</SwmPath>:60:60"
%%         node7 -->|"No"| node9["Continue"]
%%         node8 --> node10["Format fields for display (<SwmToken path="HotelManagementSystem/Modules/modProcedure.bas" pos="65:10:10" line-data="                            X.SubItems(i) = FormatRS(sRecordSource.Fields(CInt(i) - 1))">`FormatRS`</SwmToken>)"]
%%         node9 --> node10
%%         click node10 openCode "<SwmPath>[HotelManagementSystem/Modules/modProcedure.bas](HotelManagementSystem/Modules/modProcedure.bas)</SwmPath>:61:79"
%%         click node14 openCode "<SwmPath>[HotelManagementSystem/Modules/modProcedure.bas](HotelManagementSystem/Modules/modProcedure.bas)</SwmPath>:84:84"
%%     end
%% 
%% classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/HotelManagementSystem/Modules/modProcedure.bas" line="47">

---

<SwmToken path="HotelManagementSystem/Modules/modProcedure.bas" pos="47:4:4" line-data="Public Sub FillListView(ByRef sListView As ListView, ByRef sRecordSource As Recordset, ByVal sNumOfFields As Byte, ByVal sNumIco As Byte, ByVal with_num As Boolean, ByVal show_first_rec As Boolean, Optional srcHiddenField As String)">`FillListView`</SwmToken> kicks off the process by clearing the <SwmToken path="HotelManagementSystem/Modules/modProcedure.bas" pos="47:12:12" line-data="Public Sub FillListView(ByRef sListView As ListView, ByRef sRecordSource As Recordset, ByVal sNumOfFields As Byte, ByVal sNumIco As Byte, ByVal with_num As Boolean, ByVal show_first_rec As Boolean, Optional srcHiddenField As String)">`ListView`</SwmToken> and checking if there's any data to show. It loops through the Recordset, adding each record as a <SwmToken path="HotelManagementSystem/Modules/modProcedure.bas" pos="47:12:12" line-data="Public Sub FillListView(ByRef sListView As ListView, ByRef sRecordSource As Recordset, ByVal sNumOfFields As Byte, ByVal sNumIco As Byte, ByVal with_num As Boolean, ByVal show_first_rec As Boolean, Optional srcHiddenField As String)">`ListView`</SwmToken> item, optionally using a numeric index or the first field as the key, and can attach a hidden field as metadata. For each field, it decides which value to display based on the flags, then calls <SwmToken path="HotelManagementSystem/Modules/modProcedure.bas" pos="65:10:10" line-data="                            X.SubItems(i) = FormatRS(sRecordSource.Fields(CInt(i) - 1))">`FormatRS`</SwmToken> to handle display formatting before assigning the value to the <SwmToken path="HotelManagementSystem/Modules/modProcedure.bas" pos="47:12:12" line-data="Public Sub FillListView(ByRef sListView As ListView, ByRef sRecordSource As Recordset, ByVal sNumOfFields As Byte, ByVal sNumIco As Byte, ByVal with_num As Boolean, ByVal show_first_rec As Boolean, Optional srcHiddenField As String)">`ListView`</SwmToken> subitems. We need to call <SwmToken path="HotelManagementSystem/Modules/modProcedure.bas" pos="65:10:10" line-data="                            X.SubItems(i) = FormatRS(sRecordSource.Fields(CInt(i) - 1))">`FormatRS`</SwmToken> next to make sure things like dates and currency look right in the UI.

```visual basic
Public Sub FillListView(ByRef sListView As ListView, ByRef sRecordSource As Recordset, ByVal sNumOfFields As Byte, ByVal sNumIco As Byte, ByVal with_num As Boolean, ByVal show_first_rec As Boolean, Optional srcHiddenField As String)
    Dim X As Variant
    Dim i As Byte
    On Error Resume Next
    sListView.ListItems.Clear
    If sRecordSource.RecordCount < 1 Then Exit Sub
    sRecordSource.MoveFirst
    Do While Not sRecordSource.EOF
        If with_num = True Then
            Set X = sListView.ListItems.Add(, , sRecordSource.AbsolutePosition, sNumIco, sNumIco)
        Else
            Set X = sListView.ListItems.Add(, , "" & sRecordSource.Fields(0), sNumIco, sNumIco)
        End If
            If srcHiddenField <> "" Then X.Tag = sRecordSource.Fields(srcHiddenField)
            For i = 1 To sNumOfFields - 1
                If show_first_rec = True Then
                    If with_num = True Then
                        If sRecordSource.Fields(CInt(i) - 1).Type = adDouble Then
                            X.SubItems(i) = FormatRS(sRecordSource.Fields(CInt(i) - 1))
                        Else
                            X.SubItems(i) = "" & FormatRS(sRecordSource.Fields(CInt(i) - 1))
                        End If
                    Else
                        If sRecordSource.Fields(CInt(i)).Type = adDouble Then
                            X.SubItems(i) = FormatRS(sRecordSource.Fields(CInt(i)))
                        Else
                            X.SubItems(i) = "" & FormatRS(sRecordSource.Fields(CInt(i)))
                        End If
                    End If
                Else
                    X.SubItems(i) = "" & FormatRS(sRecordSource.Fields(CInt(i) + 1))
                End If
            Next i
        sRecordSource.MoveNext
    Loop
    i = 0
    Set X = Nothing
End Sub
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Modules/modFunction.bas" line="7">

---

<SwmToken path="HotelManagementSystem/Modules/modFunction.bas" pos="7:4:4" line-data="Public Function FormatRS(ByVal srcField As Field, Optional AllowNewLine As Boolean) As String">`FormatRS`</SwmToken> handles the display formatting for each field. If <SwmToken path="HotelManagementSystem/Modules/modFunction.bas" pos="7:17:17" line-data="Public Function FormatRS(ByVal srcField As Field, Optional AllowNewLine As Boolean) As String">`AllowNewLine`</SwmToken> is False, it strips out newlines to keep the <SwmToken path="HotelManagementSystem/Modules/modProcedure.bas" pos="47:12:12" line-data="Public Sub FillListView(ByRef sListView As ListView, ByRef sRecordSource As Recordset, ByVal sNumOfFields As Byte, ByVal sNumIco As Byte, ByVal with_num As Boolean, ByVal show_first_rec As Boolean, Optional srcHiddenField As String)">`ListView`</SwmToken> clean. It checks the field type: currency values get formatted with two decimals, dates get a readable string, and everything else is just shown as-is. This keeps the <SwmToken path="HotelManagementSystem/Modules/modProcedure.bas" pos="47:12:12" line-data="Public Sub FillListView(ByRef sListView As ListView, ByRef sRecordSource As Recordset, ByVal sNumOfFields As Byte, ByVal sNumIco As Byte, ByVal with_num As Boolean, ByVal show_first_rec As Boolean, Optional srcHiddenField As String)">`ListView`</SwmToken> output consistent and user-friendly.

```visual basic
Public Function FormatRS(ByVal srcField As Field, Optional AllowNewLine As Boolean) As String
    Dim strRet As String
    
    With srcField
        If AllowNewLine = True Then
            strRet = srcField
        Else
            strRet = Replace(srcField, vbCrLf, " ", , , vbTextCompare)
        End If
        
        'If srcField.Type = adCurrency Or srcField.Type = adDouble Then
        If srcField.Type = adCurrency Then
            strRet = Format$(srcField, "#,##0.00")
        ElseIf srcField.Type = adDate Then
            strRet = Format$(srcField, "MMM-dd-yyyy")
        Else
            strRet = srcField
        End If
    End With
    
    FormatRS = strRet
    
    strRet = vbNullString
End Function
```

---

</SwmSnippet>

&nbsp;

*This is an auto-generated document by Swimm 🌊 and has not yet been verified by a human*

<SwmMeta version="3.0.0" repo-id="Z2l0aHViJTNBJTNBY3RzLVZCNi1Qcm9qZWN0cyUzQSUzQVN3aW1tLURlbW8=" repo-name="cts-VB6-Projects"><sup>Powered by [Swimm](https://app.swimm.io/)</sup></SwmMeta>
