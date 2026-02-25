---
title: Saving student data flow
---
This document describes the process triggered by the Save button in the student data entry interface. The flow validates the data, saves it if possible, and resets the form for new input. If saving fails, the user receives feedback and can correct the data.

# Triggering the Save Process

<SwmSnippet path="/BK App/Form/frmentrysoal.frm" line="172">

---

In <SwmToken path="BK App/Form/frmentrysoal.frm" pos="172:4:4" line-data="Private Sub Command2_Click()">`Command2_Click`</SwmToken>, the only thing happening is triggering the save routine by calling <SwmToken path="BK App/Form/frmentrysoal.frm" pos="173:3:3" line-data="    Call mnuSave">`mnuSave`</SwmToken>. This keeps the button handler clean and offloads all the actual work to <SwmToken path="BK App/Form/frmentrysoal.frm" pos="173:3:3" line-data="    Call mnuSave">`mnuSave`</SwmToken>, which centralizes the save logic.

```visual basic
Private Sub Command2_Click()
    Call mnuSave
```

---

</SwmSnippet>

## Validating and Saving the Data

<SwmSnippet path="/BK App/Form/frmentrysoal.frm" line="189">

---

In <SwmToken path="BK App/Form/frmentrysoal.frm" pos="189:4:4" line-data="Private Sub mnuSave()">`mnuSave`</SwmToken>, we first check <SwmToken path="BK App/Form/frmentrysoal.frm" pos="191:2:2" line-data="If DataMode = EN_NEW Then">`DataMode`</SwmToken> to avoid saving when the form is empty or unchanged. If saving is allowed, we call <SwmToken path="BK App/Form/frmentrysoal.frm" pos="198:2:2" line-data="If SaveData(txtnis.text, txtnamasiswa.text) &gt; 0 Then">`SaveData`</SwmToken> to actually persist the data.

```visual basic
Private Sub mnuSave()
On Error GoTo Hell
If DataMode = EN_NEW Then
    MsgBox "Data harus diisi dulu" & vbCrLf & "Simpan data dibatalkan", vbExclamation, "Simpan Data"
    Exit Sub
ElseIf DataMode = EN_SAVED Then
    MsgBox "Tidak ada data yang berubah" & vbCrLf & "Simpan data dibatalkan", vbExclamation, "Simpan Data"
    Exit Sub
End If
If SaveData(txtnis.text, txtnamasiswa.text) > 0 Then
```

---

</SwmSnippet>

<SwmSnippet path="/BK App/Form/frmentrysoal.frm" line="209">

---

<SwmToken path="BK App/Form/frmentrysoal.frm" pos="209:4:4" line-data="Private Function SaveData(pNo As Integer, pSoal As String) As Integer">`SaveData`</SwmToken> checks <SwmToken path="BK App/Form/frmentrysoal.frm" pos="210:2:2" line-data="If DataMode = EN_NEW_CHANGED Then">`DataMode`</SwmToken> to decide if it should add a new entry or update an existing one in <SwmToken path="BK App/Form/frmentrysoal.frm" pos="211:1:1" line-data="    oSoal.Add pNo, pSoal">`oSoal`</SwmToken>, then marks the data as saved and returns success.

```visual basic
Private Function SaveData(pNo As Integer, pSoal As String) As Integer
If DataMode = EN_NEW_CHANGED Then
    oSoal.Add pNo, pSoal
ElseIf DataMode = EN_LOAD_CHANGED Then
    oSoal.Edit pNo, pSoal
End If
DataMode = EN_SAVED
SaveData = 1
End Function
```

---

</SwmSnippet>

<SwmSnippet path="/BK App/Form/frmentrysoal.frm" line="199">

---

Back in <SwmToken path="BK App/Form/frmentrysoal.frm" pos="173:3:3" line-data="    Call mnuSave">`mnuSave`</SwmToken>, after <SwmToken path="BK App/Form/frmentrysoal.frm" pos="198:2:2" line-data="If SaveData(txtnis.text, txtnamasiswa.text) &gt; 0 Then">`SaveData`</SwmToken> returns success, we show a confirmation and reload the form with <SwmToken path="BK App/Form/frmentrysoal.frm" pos="200:1:1" line-data="    Form_Load">`Form_Load`</SwmToken> to reset everything for the next input.

```visual basic
    MsgBox "Data BERHASIL disimpan", vbInformation, "Simpan Data"
    Form_Load
```

---

</SwmSnippet>

### Resetting the Form State

<SwmSnippet path="/BK App/Form/frmentrysoal.frm" line="180">

---

In <SwmToken path="BK App/Form/frmentrysoal.frm" pos="180:4:4" line-data="Private Sub Form_Load()">`Form_Load`</SwmToken>, we start with a new <SwmToken path="BK App/Form/frmentrysoal.frm" pos="181:11:11" line-data="    Set oSoal = New DLLBK.cSoal">`cSoal`</SwmToken> instance and immediately call <SwmToken path="BK App/Form/frmentrysoal.frm" pos="182:3:3" line-data="    Call New_data">`New_data`</SwmToken> to clear the form and set it up for new input.

```visual basic
Private Sub Form_Load()
    Set oSoal = New DLLBK.cSoal
    Call New_data
```

---

</SwmSnippet>

<SwmSnippet path="/BK App/Form/frmentrysoal.frm" line="184">

---

<SwmToken path="BK App/Form/frmentrysoal.frm" pos="184:4:4" line-data="Private Sub New_data()">`New_data`</SwmToken> clears the input fields and sets <SwmToken path="BK App/Form/frmentrysoal.frm" pos="187:1:1" line-data="    DataMode = EN_NEW">`DataMode`</SwmToken> to <SwmToken path="BK App/Form/frmentrysoal.frm" pos="187:5:5" line-data="    DataMode = EN_NEW">`EN_NEW`</SwmToken>, prepping the form for the next new entry.

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

Back in <SwmToken path="BK App/Form/frmentrysoal.frm" pos="180:4:4" line-data="Private Sub Form_Load()">`Form_Load`</SwmToken>, after <SwmToken path="BK App/Form/frmentrysoal.frm" pos="182:3:3" line-data="    Call New_data">`New_data`</SwmToken> runs, the form is clean and ready for the next user action.

```visual basic
End Sub
```

---

</SwmSnippet>

### Handling Save Failures and Cleanup

<SwmSnippet path="/BK App/Form/frmentrysoal.frm" line="201">

---

Back in <SwmToken path="BK App/Form/frmentrysoal.frm" pos="173:3:3" line-data="    Call mnuSave">`mnuSave`</SwmToken>, if <SwmToken path="BK App/Form/frmentrysoal.frm" pos="198:2:2" line-data="If SaveData(txtnis.text, txtnamasiswa.text) &gt; 0 Then">`SaveData`</SwmToken> fails, we show an error message and leave the form unchanged so the user can correct the issue.

```visual basic
Else
    MsgBox "Data GAGAL disimpan", vbCritical, "Simpan Data"
End If
Exit Sub
Hell:
    MsgBox Err.Description, vbCritical, "Internal"
    'Resume Next
End Sub
```

---

</SwmSnippet>

## Completing the Save Trigger

<SwmSnippet path="/BK App/Form/frmentrysoal.frm" line="174">

---

Back in <SwmToken path="BK App/Form/frmentrysoal.frm" pos="172:4:4" line-data="Private Sub Command2_Click()">`Command2_Click`</SwmToken>, after <SwmToken path="BK App/Form/frmentrysoal.frm" pos="173:3:3" line-data="    Call mnuSave">`mnuSave`</SwmToken> finishes, nothing else happens—the button handler is done and control returns to the user.

```visual basic
End Sub
```

---

</SwmSnippet>

&nbsp;

*This is an auto-generated document by Swimm 🌊 and has not yet been verified by a human*

<SwmMeta version="3.0.0" repo-id="Z2l0aHViJTNBJTNBY3RzLVZCNi1Qcm9qZWN0cyUzQSUzQVN3aW1tLURlbW8=" repo-name="cts-VB6-Projects"><sup>Powered by [Swimm](https://app.swimm.io/)</sup></SwmMeta>
