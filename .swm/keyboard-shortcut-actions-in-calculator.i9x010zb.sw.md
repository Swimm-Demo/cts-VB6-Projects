---
title: Keyboard Shortcut Actions in Calculator
---
This document describes how keyboard shortcuts enable users to control calculator actions efficiently. Pressing Backspace, Escape, or Delete updates the calculator's display or state just like pressing the corresponding button, allowing for quick corrections, entry clearing, or resetting using the keyboard.

# Keyboard Shortcut Handling and Action Dispatch

<SwmSnippet path="/warnet/Server/timeronline.frm" line="885">

---

In <SwmToken path="warnet/Server/timeronline.frm" pos="885:4:4" line-data="Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)">`Form_KeyDown`</SwmToken>, we check which key was pressed and immediately route Backspace to <SwmToken path="warnet/Server/timeronline.frm" pos="888:3:3" line-data="        Call CmdBS_Click">`CmdBS_Click`</SwmToken>. This lets users use the keyboard to trigger the same logic as the backspace button, keeping keyboard and UI actions in sync.

```visual basic
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 8 'Backspace
        Call CmdBS_Click
```

---

</SwmSnippet>

<SwmSnippet path="/warnet/Server/timeronline.frm" line="845">

---

<SwmToken path="warnet/Server/timeronline.frm" pos="845:4:4" line-data="Private Sub CmdBS_Click()">`CmdBS_Click`</SwmToken> handles removing the last digit from the display if possible, but blocks the action if we're in an error or operation state, or if the number is already minimal. It also updates the last number and counts backspaces.

```visual basic
Private Sub CmdBS_Click()

If bWasError Then
    Beep
    Exit Sub
End If

If bEqual Or bMEM Or bOp Then
    Beep
    Exit Sub
End If

Static nBSCount As Integer
If (Len(lblOutput.Caption) > 1 And CDbl(lblOutput.Caption) > 0) Or (CDbl(lblOutput.Caption) < 0 And Len(lblOutput.Caption) > 2) Then
    lblOutput.Caption = Left$(lblOutput.Caption, Len(lblOutput.Caption) - 1)
    nBSCount = nBSCount + 1
    Else
    Beep
    lblOutput.Caption = 0
End If
nLastNum = CDbl(lblOutput.Caption)
End Sub
```

---

</SwmSnippet>

<SwmSnippet path="/warnet/Server/timeronline.frm" line="889">

---

Back in <SwmToken path="warnet/Server/timeronline.frm" pos="885:4:4" line-data="Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)">`Form_KeyDown`</SwmToken>, after handling Backspace, Escape triggers <SwmToken path="warnet/Server/timeronline.frm" pos="890:3:3" line-data="        Call cmdC_Click">`cmdC_Click`</SwmToken>. This gives users a quick way to reset the calculator state from the keyboard.

```visual basic
    Case 27 'Escape
        Call cmdC_Click
```

---

</SwmSnippet>

<SwmSnippet path="/warnet/Server/timeronline.frm" line="567">

---

<SwmToken path="warnet/Server/timeronline.frm" pos="567:4:4" line-data="Private Sub cmdC_Click()">`cmdC_Click`</SwmToken> wipes all calculation state and resets the display, so everything is ready for a new calculation.

```visual basic
Private Sub cmdC_Click()
bWasError = False
nLastNum = 0
nResult = 0
bOp = False
nOp = 0
bEqual = False
lblOutput.Caption = "0"

End Sub
```

---

</SwmSnippet>

<SwmSnippet path="/warnet/Server/timeronline.frm" line="891">

---

Back in <SwmToken path="warnet/Server/timeronline.frm" pos="885:4:4" line-data="Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)">`Form_KeyDown`</SwmToken>, after Escape, Delete triggers <SwmToken path="warnet/Server/timeronline.frm" pos="892:3:3" line-data="        Call cmdCE_Click">`cmdCE_Click`</SwmToken>. This lets users clear just the current entry from the keyboard, without wiping the whole state.

```visual basic
    Case 46 'Del
        Call cmdCE_Click
```

---

</SwmSnippet>

## Entry Clearing and Error Reset

<SwmSnippet path="/warnet/Server/timeronline.frm" line="577">

---

In <SwmToken path="warnet/Server/timeronline.frm" pos="577:4:4" line-data="Private Sub cmdCE_Click()">`cmdCE_Click`</SwmToken>, if we're in an error state, we call <SwmToken path="warnet/Server/timeronline.frm" pos="579:3:3" line-data="    Call cmdC_Click">`cmdC_Click`</SwmToken> to do a full reset before clearing just the entry. This keeps things consistent.

```visual basic
Private Sub cmdCE_Click()
If bWasError Then
    Call cmdC_Click
```

---

</SwmSnippet>

<SwmSnippet path="/warnet/Server/timeronline.frm" line="580">

---

After coming back from <SwmToken path="warnet/Server/timeronline.frm" pos="567:4:4" line-data="Private Sub cmdC_Click()">`cmdC_Click`</SwmToken>, <SwmToken path="warnet/Server/timeronline.frm" pos="577:4:4" line-data="Private Sub cmdCE_Click()">`cmdCE_Click`</SwmToken> still clears the error flag, display, and last number to make sure the entry is definitely reset, even if we just did a full clear.

```visual basic
End If
bWasError = False
lblOutput.Caption = 0
nLastNum = 0
End Sub
```

---

</SwmSnippet>

## Finishing Keyboard Event Processing

<SwmSnippet path="/warnet/Server/timeronline.frm" line="893">

---

After returning from <SwmToken path="warnet/Server/timeronline.frm" pos="577:4:4" line-data="Private Sub cmdCE_Click()">`cmdCE_Click`</SwmToken>, <SwmToken path="warnet/Server/timeronline.frm" pos="885:4:4" line-data="Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)">`Form_KeyDown`</SwmToken> just finishes up. The UI is already updated by the button logic, so nothing else happens here.

```visual basic
End Select
End Sub
```

---

</SwmSnippet>

&nbsp;

*This is an auto-generated document by Swimm 🌊 and has not yet been verified by a human*

<SwmMeta version="3.0.0" repo-id="Z2l0aHViJTNBJTNBY3RzLVZCNi1Qcm9qZWN0cyUzQSUzQVN3aW1tLURlbW8=" repo-name="cts-VB6-Projects"><sup>Powered by [Swimm](https://app.swimm.io/)</sup></SwmMeta>
