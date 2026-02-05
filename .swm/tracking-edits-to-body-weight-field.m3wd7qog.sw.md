---
title: Tracking Edits to Body Weight Field
---
This document explains how the system tracks edits to the body weight field in the student entry form. When a user makes a change, the system updates the data state to reflect unsaved modifications, enabling prompts or warnings about unsaved data.

# Reacting to User Edits in Weight Field

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
    node1["User changes body weight field"] --> node2{"Is data mode 'new'?"}
    click node1 openCode "BK BK App/Form/frmentrysiswa.frm:2193:2195"
    node2 -->|"Yes"| node3["Mark entry as 'new changed'"]
    click node2 openCode "BK BK App/Form/frmentrysiswa.frm:1962:1963"
    click node3 openCode "BK BK App/Form/frmentrysiswa.frm:1963:1963"
    node2 -->|"No"| node4{"Is data mode 'saved'?"}
    click node4 openCode "BK BK App/Form/frmentrysiswa.frm:1964:1965"
    node4 -->|"Yes"| node5["Mark entry as 'load changed'"]
    click node5 openCode "BK BK App/Form/frmentrysiswa.frm:1965:1965"

classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;

%% Swimm:
%% %%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
%% flowchart TD
%%     node1["User changes body weight field"] --> node2{"Is data mode 'new'?"}
%%     click node1 openCode "BK <SwmPath>[BK App/Form/frmentrysiswa.frm](BK%20App/Form/frmentrysiswa.frm)</SwmPath>:2193:2195"
%%     node2 -->|"Yes"| node3["Mark entry as 'new changed'"]
%%     click node2 openCode "BK <SwmPath>[BK App/Form/frmentrysiswa.frm](BK%20App/Form/frmentrysiswa.frm)</SwmPath>:1962:1963"
%%     click node3 openCode "BK <SwmPath>[BK App/Form/frmentrysiswa.frm](BK%20App/Form/frmentrysiswa.frm)</SwmPath>:1963:1963"
%%     node2 -->|"No"| node4{"Is data mode 'saved'?"}
%%     click node4 openCode "BK <SwmPath>[BK App/Form/frmentrysiswa.frm](BK%20App/Form/frmentrysiswa.frm)</SwmPath>:1964:1965"
%%     node4 -->|"Yes"| node5["Mark entry as 'load changed'"]
%%     click node5 openCode "BK <SwmPath>[BK App/Form/frmentrysiswa.frm](BK%20App/Form/frmentrysiswa.frm)</SwmPath>:1965:1965"
%% 
%% classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/BK App/Form/frmentrysiswa.frm" line="2193">

---

Txtberatbadan_Change kicks off the flow whenever the user edits the weight field. It immediately calls <SwmToken path="BK App/Form/frmentrysiswa.frm" pos="2194:3:3" line-data="    Call ChangeData">`ChangeData`</SwmToken> to flag that the data has been modified, so the system can track unsaved changes from this point.

```visual basic
Private Sub txtberatbadan_Change()
    Call ChangeData
End Sub
```

---

</SwmSnippet>

<SwmSnippet path="/BK App/Form/frmentrysiswa.frm" line="1961">

---

<SwmToken path="BK App/Form/frmentrysiswa.frm" pos="1961:4:4" line-data="Private Sub ChangeData()">`ChangeData`</SwmToken> handles the state switch for <SwmToken path="BK App/Form/frmentrysiswa.frm" pos="1962:2:2" line-data="If DataMode = EN_NEW Then">`DataMode`</SwmToken>. It marks whether the user is editing a new or existing record, so the rest of the app can react accordingly (like enabling save or warning about unsaved changes).

```visual basic
Private Sub ChangeData()
If DataMode = EN_NEW Then
    DataMode = EN_NEW_CHANGED
ElseIf DataMode = EN_SAVED Then
    DataMode = EN_LOAD_CHANGED
End If
End Sub
```

---

</SwmSnippet>

&nbsp;

*This is an auto-generated document by Swimm 🌊 and has not yet been verified by a human*

<SwmMeta version="3.0.0" repo-id="Z2l0aHViJTNBJTNBY3RzLVZCNi1Qcm9qZWN0cyUzQSUzQVN3aW1tLURlbW8=" repo-name="cts-VB6-Projects"><sup>Powered by [Swimm](https://app.swimm.io/)</sup></SwmMeta>
