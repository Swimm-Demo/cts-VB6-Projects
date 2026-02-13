---
title: Resetting the Form for New Data Entry
---
Clicking the reset button clears all previous data from the form and prepares it for new data entry. This enables users to efficiently start a new entry.

# Triggering the Reset Sequence

<SwmSnippet path="/BK App/Form/frmentrysoal.frm" line="176">

---

In <SwmToken path="BK App/Form/frmentrysoal.frm" pos="176:4:4" line-data="Private Sub Command3_Click()">`Command3_Click`</SwmToken>, we kick off the flow by calling <SwmToken path="BK App/Form/frmentrysoal.frm" pos="177:1:1" line-data="    Form_Load">`Form_Load`</SwmToken>. This reuses the form's setup logic, so clicking the button resets everything just like when the form first opens.

```visual basic
Private Sub Command3_Click()
    Form_Load
```

---

</SwmSnippet>

## Reinitializing Form State

<SwmSnippet path="/BK App/Form/frmentrysoal.frm" line="180">

---

In <SwmToken path="BK App/Form/frmentrysoal.frm" pos="180:4:4" line-data="Private Sub Form_Load()">`Form_Load`</SwmToken>, we set up a new <SwmToken path="BK App/Form/frmentrysoal.frm" pos="181:11:11" line-data="    Set oSoal = New DLLBK.cSoal">`cSoal`</SwmToken> object to clear out any old state, then call <SwmToken path="BK App/Form/frmentrysoal.frm" pos="182:3:3" line-data="    Call New_data">`New_data`</SwmToken> to reset the form fields and mode. This keeps things consistent every time the form is loaded or reset.

```visual basic
Private Sub Form_Load()
    Set oSoal = New DLLBK.cSoal
    Call New_data
```

---

</SwmSnippet>

<SwmSnippet path="/BK App/Form/frmentrysoal.frm" line="184">

---

<SwmToken path="BK App/Form/frmentrysoal.frm" pos="184:4:4" line-data="Private Sub New_data()">`New_data`</SwmToken> clears the student name and NIS fields, and switches the form into new entry mode by setting <SwmToken path="BK App/Form/frmentrysoal.frm" pos="187:1:5" line-data="    DataMode = EN_NEW">`DataMode = EN_NEW`</SwmToken>.

```visual basic
Private Sub New_data()
    txtnamasiswa.text = ""
    txtnis.text = ""
    DataMode = EN_NEW
End Sub
```

---

</SwmSnippet>

<SwmSnippet path="/BK App/Form/frmentrysoal.frm" line="183">

---

After coming back from <SwmToken path="BK App/Form/frmentrysoal.frm" pos="182:3:3" line-data="    Call New_data">`New_data`</SwmToken>, <SwmToken path="BK App/Form/frmentrysoal.frm" pos="177:1:1" line-data="    Form_Load">`Form_Load`</SwmToken> ends, so the form is now reset and ready for new input.

```visual basic
End Sub
```

---

</SwmSnippet>

## Completing the Reset Action

<SwmSnippet path="/BK App/Form/frmentrysoal.frm" line="178">

---

After returning from <SwmToken path="BK App/Form/frmentrysoal.frm" pos="177:1:1" line-data="    Form_Load">`Form_Load`</SwmToken>, <SwmToken path="BK App/Form/frmentrysoal.frm" pos="176:4:4" line-data="Private Sub Command3_Click()">`Command3_Click`</SwmToken> ends, so the button click finishes with the form fully reset.

```visual basic
End Sub
```

---

</SwmSnippet>

&nbsp;

*This is an auto-generated document by Swimm 🌊 and has not yet been verified by a human*

<SwmMeta version="3.0.0" repo-id="Z2l0aHViJTNBJTNBY3RzLVZCNi1Qcm9qZWN0cyUzQSUzQVN3aW1tLURlbW8=" repo-name="cts-VB6-Projects"><sup>Powered by [Swimm](https://app.swimm.io/)</sup></SwmMeta>
