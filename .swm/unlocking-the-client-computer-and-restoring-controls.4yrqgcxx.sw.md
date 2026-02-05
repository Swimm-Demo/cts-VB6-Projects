---
title: Unlocking the client computer and restoring controls
---
This document describes how users can unlock a client computer and regain access to controls in the Internet Cafe System. When the unlock action is triggered, the system communicates with the server, updates the locked state, restores keyboard controls, and transitions the interface to allow further configuration.

# Unlocking the Client and Restoring Controls

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
  node1["Send unlock signal to server"]
  click node1 openCode "Internet Cafe Internet Cafe System/cLiEnTe/frmConfig.frm:246:246"
  node1 --> node2["Set computer as unlocked"]
  click node2 openCode "Internet Cafe Internet Cafe System/cLiEnTe/frmConfig.frm:247:247"
  node2 --> node3["Restore keyboard controls"]
  click node3 openCode "Internet Cafe Internet Cafe System/cLiEnTe/frmConfig.frm:248:248"
  node3 --> node4["Enable lock button, disable unlock button"]
  click node4 openCode "Internet Cafe Internet Cafe System/cLiEnTe/frmConfig.frm:249:250"
  node4 --> node5["Unload unlock screen"]
  click node5 openCode "Internet Cafe Internet Cafe System/cLiEnTe/frmConfig.frm:251:251"
  node5 --> node6["Hide main screen"]
  click node6 openCode "Internet Cafe Internet Cafe System/cLiEnTe/frmConfig.frm:252:252"
  node6 --> node7["Show configuration screen"]
  click node7 openCode "Internet Cafe Internet Cafe System/cLiEnTe/frmConfig.frm:253:253"

classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;

%% Swimm:
%% %%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
%% flowchart TD
%%   node1["Send unlock signal to server"]
%%   click node1 openCode "Internet Cafe <SwmPath>[Internet Cafe System/cLiEnTe/frmConfig.frm](Internet%20Cafe%20System/cLiEnTe/frmConfig.frm)</SwmPath>:246:246"
%%   node1 --> node2["Set computer as unlocked"]
%%   click node2 openCode "Internet Cafe <SwmPath>[Internet Cafe System/cLiEnTe/frmConfig.frm](Internet%20Cafe%20System/cLiEnTe/frmConfig.frm)</SwmPath>:247:247"
%%   node2 --> node3["Restore keyboard controls"]
%%   click node3 openCode "Internet Cafe <SwmPath>[Internet Cafe System/cLiEnTe/frmConfig.frm](Internet%20Cafe%20System/cLiEnTe/frmConfig.frm)</SwmPath>:248:248"
%%   node3 --> node4["Enable lock button, disable unlock button"]
%%   click node4 openCode "Internet Cafe <SwmPath>[Internet Cafe System/cLiEnTe/frmConfig.frm](Internet%20Cafe%20System/cLiEnTe/frmConfig.frm)</SwmPath>:249:250"
%%   node4 --> node5["Unload unlock screen"]
%%   click node5 openCode "Internet Cafe <SwmPath>[Internet Cafe System/cLiEnTe/frmConfig.frm](Internet%20Cafe%20System/cLiEnTe/frmConfig.frm)</SwmPath>:251:251"
%%   node5 --> node6["Hide main screen"]
%%   click node6 openCode "Internet Cafe <SwmPath>[Internet Cafe System/cLiEnTe/frmConfig.frm](Internet%20Cafe%20System/cLiEnTe/frmConfig.frm)</SwmPath>:252:252"
%%   node6 --> node7["Show configuration screen"]
%%   click node7 openCode "Internet Cafe <SwmPath>[Internet Cafe System/cLiEnTe/frmConfig.frm](Internet%20Cafe%20System/cLiEnTe/frmConfig.frm)</SwmPath>:253:253"
%% 
%% classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/Internet Cafe System/cLiEnTe/frmConfig.frm" line="244">

---

CmdUnlock_Click kicks off the unlock process: it sends an unlock signal to the server, updates the local Locked state, re-enables Ctrl+Alt+Del, and swaps the enabled state of the lock/unlock buttons. The form transitions (unload/hide/show) refresh the UI to reflect the unlocked state. We call into <SwmPath>[Internet Cafe System/cLiEnTe/Module1.bas](Internet%20Cafe%20System/cLiEnTe/Module1.bas)</SwmPath> next to actually re-enable Ctrl+Alt+Del at the OS level.

```visual basic
Private Sub cmdUnlock_Click()
On Error Resume Next
  frmMain.Sucket.SendData "UL" 'sends unlock signal to server
  Locked = False
  BlockCtrl_Alt_Del False
  cmdLock.Enabled = True
  cmdUnlock.Enabled = False
  Unload Me
  frmMain.Hide
  frmConfig.Show
End Sub
```

---

</SwmSnippet>

<SwmSnippet path="/Internet Cafe System/cLiEnTe/Module1.bas" line="88">

---

<SwmToken path="Internet Cafe System/cLiEnTe/Module1.bas" pos="88:2:2" line-data="Sub BlockCtrl_Alt_Del(bDisabled As Boolean)">`BlockCtrl_Alt_Del`</SwmToken> handles toggling the Ctrl+Alt+Del key sequence by calling <SwmToken path="Internet Cafe System/cLiEnTe/Module1.bas" pos="90:5:5" line-data="  X = SystemParametersInfo(97, bDisabled, CStr(1), 0)">`SystemParametersInfo`</SwmToken> with action code 97. It uses the <SwmToken path="Internet Cafe System/cLiEnTe/Module1.bas" pos="88:4:4" line-data="Sub BlockCtrl_Alt_Del(bDisabled As Boolean)">`bDisabled`</SwmToken> flag to decide whether to block or unblock, but the use of <SwmToken path="Internet Cafe System/cLiEnTe/Module1.bas" pos="90:13:13" line-data="  X = SystemParametersInfo(97, bDisabled, CStr(1), 0)">`CStr`</SwmToken>(1) is a bit hacky and might not be robust everywhere.

```visual basic
Sub BlockCtrl_Alt_Del(bDisabled As Boolean)
  Dim X As Long
  X = SystemParametersInfo(97, bDisabled, CStr(1), 0)

End Sub
```

---

</SwmSnippet>

&nbsp;

*This is an auto-generated document by Swimm 🌊 and has not yet been verified by a human*

<SwmMeta version="3.0.0" repo-id="Z2l0aHViJTNBJTNBY3RzLVZCNi1Qcm9qZWN0cyUzQSUzQVN3aW1tLURlbW8=" repo-name="cts-VB6-Projects"><sup>Powered by [Swimm](https://app.swimm.io/)</sup></SwmMeta>
