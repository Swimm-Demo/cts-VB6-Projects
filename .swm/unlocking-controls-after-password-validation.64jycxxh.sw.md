---
title: Unlocking Controls After Password Validation
---
This document outlines how users gain access to the application's main controls and system features after entering the correct password. Once validated, the system enables key controls and removes restrictions, allowing normal use.

# Password Validation and Unlocking Controls

<SwmSnippet path="/warnet/Client/clnonline.frm" line="900">

---

<SwmToken path="warnet/Client/clnonline.frm" pos="900:4:4" line-data="Private Sub Txpass_Change()">`Txpass_Change`</SwmToken> checks if the password input matches the reference value. If it does, it enables the main action buttons so the user can interact with the app beyond the login. WINLOCKOPEN is called right after to unlock system-level restrictions, letting the user fully interact with the system once authenticated.

```visual basic
Private Sub Txpass_Change()
If Txpass.Text = Text1.Text Then
    CMExit.Enabled = True
    CmConnect.Enabled = True
    CmCaptured.Enabled = True
    Cmsetting.Enabled = True
    WINLOCKOPEN
    End If
End Sub
```

---

</SwmSnippet>

<SwmSnippet path="/warnet/Client/clnonline.frm" line="852">

---

<SwmToken path="warnet/Client/clnonline.frm" pos="852:4:4" line-data="Private Sub WINLOCKOPEN()">`WINLOCKOPEN`</SwmToken> flips system-level restrictions back off—Alt+Tab, task switching, Task Manager, and Ctrl+Alt+Del all get re-enabled so the user isn't locked down anymore after logging in.

```visual basic
Private Sub WINLOCKOPEN()
    AltTab2_Enable_Disable 0, True
    TaskSwitching_Enable_Disable (True)
    TaskManager_Enable_Disable (True)
    CtrlAltDel_Enable_Disable (True)
End Sub
```

---

</SwmSnippet>

&nbsp;

*This is an auto-generated document by Swimm 🌊 and has not yet been verified by a human*

<SwmMeta version="3.0.0" repo-id="Z2l0aHViJTNBJTNBY3RzLVZCNi1Qcm9qZWN0cyUzQSUzQVN3aW1tLURlbW8=" repo-name="cts-VB6-Projects"><sup>Powered by [Swimm](https://app.swimm.io/)</sup></SwmMeta>
