---
title: Resetting the Counseling Session Form
---
Clicking the reset button on the counseling session form clears all fields, refreshes the session list, and generates a new session ID. This prepares the form for entering a new counseling session.

# Triggering Form Reset

<SwmSnippet path="/BK App/Form/frmkonseling.frm" line="357">

---

In <SwmToken path="BK App/Form/frmkonseling.frm" pos="357:4:4" line-data="Private Sub Command2_Click()">`Command2_Click`</SwmToken>, the flow starts by calling <SwmToken path="BK App/Form/frmkonseling.frm" pos="358:1:1" line-data="    Form_Load">`Form_Load`</SwmToken> directly. This is a shortcut to reset the form and re-run all the setup logic, instead of duplicating that code here.

```visual basic
Private Sub Command2_Click()
    Form_Load
```

---

</SwmSnippet>

## Preparing Form State and Data

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
    node1["Form is loaded"] --> node2["Clear all fields: counselor, student, IDs, problem, handling, notes"]
    click node1 openCode "BK BK App/Form/frmkonseling.frm:395:405"
    node2 --> node3["Load list of counseling sessions"]
    click node2 openCode "BK BK App/Form/frmkonseling.frm:407:416"
    node3 --> node4["Generate session ID: day + month + year + (session count + 1)"]
    click node3 openCode "BK BK App/Form/frmkonseling.frm:400:400"
    node4 --> node5["Set session ID in form field"]
    click node4 openCode "BK BK App/Form/frmkonseling.frm:401:404"
    click node5 openCode "BK BK App/Form/frmkonseling.frm:404:404"

classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;

%% Swimm:
%% %%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
%% flowchart TD
%%     node1["Form is loaded"] --> node2["Clear all fields: counselor, student, IDs, problem, handling, notes"]
%%     click node1 openCode "BK <SwmPath>[BK App/Form/frmkonseling.frm](BK%20App/Form/frmkonseling.frm)</SwmPath>:395:405"
%%     node2 --> node3["Load list of counseling sessions"]
%%     click node2 openCode "BK <SwmPath>[BK App/Form/frmkonseling.frm](BK%20App/Form/frmkonseling.frm)</SwmPath>:407:416"
%%     node3 --> node4["Generate session ID: day + month + year + (session count + 1)"]
%%     click node3 openCode "BK <SwmPath>[BK App/Form/frmkonseling.frm](BK%20App/Form/frmkonseling.frm)</SwmPath>:400:400"
%%     node4 --> node5["Set session ID in form field"]
%%     click node4 openCode "BK <SwmPath>[BK App/Form/frmkonseling.frm](BK%20App/Form/frmkonseling.frm)</SwmPath>:401:404"
%%     click node5 openCode "BK <SwmPath>[BK App/Form/frmkonseling.frm](BK%20App/Form/frmkonseling.frm)</SwmPath>:404:404"
%% 
%% classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/BK App/Form/frmkonseling.frm" line="395">

---

In <SwmToken path="BK App/Form/frmkonseling.frm" pos="395:4:4" line-data="Private Sub Form_Load()">`Form_Load`</SwmToken>, new instances for the main data handlers are created, then <SwmToken path="BK App/Form/frmkonseling.frm" pos="399:1:1" line-data="    New_data">`New_data`</SwmToken> is called to clear out all the form fields so everything starts fresh.

```visual basic
Private Sub Form_Load()
    Set oBim = New DLLBK.cBK
    Set oSis = New DLLBK.Csiswa
    Set oGuru = New DLLBK.cGuru
    New_data
```

---

</SwmSnippet>

<SwmSnippet path="/BK App/Form/frmkonseling.frm" line="407">

---

<SwmToken path="BK App/Form/frmkonseling.frm" pos="407:4:4" line-data="Private Sub New_data()">`New_data`</SwmToken> just wipes all the text fields so the form is blank and ready for new input.

```visual basic
Private Sub New_data()
    txtid.text = ""
    txtnamaguru.text = ""
    txtnamasiswa.text = ""
    txtnip.text = ""
    txtnis.text = ""
    txtmasalah.text = ""
    txtpenangan.text = ""
    txtket.text = ""
End Sub
```

---

</SwmSnippet>

<SwmSnippet path="/BK App/Form/frmkonseling.frm" line="400">

---

After coming back from <SwmToken path="BK App/Form/frmkonseling.frm" pos="399:1:1" line-data="    New_data">`New_data`</SwmToken>, <SwmToken path="BK App/Form/frmkonseling.frm" pos="358:1:1" line-data="    Form_Load">`Form_Load`</SwmToken> refreshes the main data list and generates a new id for the entry using today's date and the next sequence number, then puts it in the id field.

```visual basic
    oBim.List (True)
    id = Format(Now, "DD")
    id = id & Format(Now, "MM")
    id = id & Format(Now, "YYYY")
    txtid.text = id & oBim.Jumlah + 1
End Sub
```

---

</SwmSnippet>

## Completing the Reset Action

<SwmSnippet path="/BK App/Form/frmkonseling.frm" line="359">

---

Back in <SwmToken path="BK App/Form/frmkonseling.frm" pos="357:4:4" line-data="Private Sub Command2_Click()">`Command2_Click`</SwmToken>, after <SwmToken path="BK App/Form/frmkonseling.frm" pos="358:1:1" line-data="    Form_Load">`Form_Load`</SwmToken> finishes, nothing else happens—so the button just resets the form and that's it.

```visual basic
End Sub
```

---

</SwmSnippet>

&nbsp;

*This is an auto-generated document by Swimm 🌊 and has not yet been verified by a human*

<SwmMeta version="3.0.0" repo-id="Z2l0aHViJTNBJTNBY3RzLVZCNi1Qcm9qZWN0cyUzQSUzQVN3aW1tLURlbW8=" repo-name="cts-VB6-Projects"><sup>Powered by [Swimm](https://app.swimm.io/)</sup></SwmMeta>
