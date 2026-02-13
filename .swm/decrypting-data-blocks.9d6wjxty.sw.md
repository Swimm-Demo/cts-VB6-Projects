---
title: Decrypting data blocks
---
This document describes how encrypted data blocks are restored to their original form using Blowfish decryption. The process involves preparing the block, applying S-box mixing and arithmetic to reverse encryption, and updating the block state after each round.

# Block Decryption Loop Setup

<SwmSnippet path="/HotelManagementSystem/Modules/clsBlowfish.cls" line="320">

---

In `DecryptBlock`, we start by swapping and XORing the block halves with the last two P-box values, setting up the state for decryption. The loop then alternates between applying f to Xr and Xl, which is how Blowfish reverses the encryption process. We need to call f next because each round of decryption depends on the output of f to properly unwind the encryption steps.

```apex
Private Static Sub DecryptBlock(Xl As Long, Xr As Long)
    Dim I As Long, j As Long, K As Long
    K = Xr
    Xr = Xl Xor m_pBox(Rounds + 1)
    Xl = K Xor m_pBox(Rounds)
    j = Rounds - 2
    For I = 0 To (Rounds \ 2 - 1)
        Xl = Xl Xor f(Xr)
        Xr = Xr Xor m_pBox(j + 1)
        Xr = Xr Xor f(Xl)
```

---

</SwmSnippet>

## S-Box Mixing and Arithmetic

<SwmSnippet path="/HotelManagementSystem/Modules/clsBlowfish.cls" line="423">

---

In `f`, we split the input into bytes and use them for S-box lookups and arithmetic. If we're not running compiled, we call UnsignedAdd to make sure the addition behaves like unsigned math, which is what Blowfish expects. This avoids issues with signed overflow in VB6.

```apex
Private Static Function f(ByVal X As Long) As Long
    Dim xb(0 To 3) As Byte
    Call CopyMem(xb(0), X, 4)
    If (m_RunningCompiled) Then f = (((m_sBox(0, xb(3)) + m_sBox(1, xb(2))) Xor m_sBox(2, xb(1))) + m_sBox(3, xb(0))) Else f = UnsignedAdd((UnsignedAdd(m_sBox(0, xb(3)), m_sBox(1, xb(2))) Xor m_sBox(2, xb(1))), m_sBox(3, xb(0)))
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Modules/clsBlowfish.cls" line="444">

---

`UnsignedAdd` does the addition byte by byte, carrying overflow between bytes, so the result matches what unsigned addition would do. This is needed because VB6 only has signed types.

```apex
Private Static Function UnsignedAdd(ByVal Data1 As Long, Data2 As Long) As Long
    Dim x1(0 To 3) As Byte, x2(0 To 3) As Byte, xx(0 To 3) As Byte, Rest As Long, Value As Long, a As Long
    Call CopyMem(x1(0), Data1, 4)
    Call CopyMem(x2(0), Data2, 4)
    Rest = 0
    For a = 0 To 3
        Value = CLng(x1(a)) + CLng(x2(a)) + Rest
        xx(a) = Value And 255
        Rest = Value \ 256
    Next
    Call CopyMem(UnsignedAdd, xx(0), 4)
End Function
```

---

</SwmSnippet>

<SwmSnippet path="/HotelManagementSystem/Modules/clsBlowfish.cls" line="427">

---

We just got back from UnsignedAdd, so the result returned from `f` is guaranteed to match unsigned addition rules. This keeps the decryption logic consistent with Blowfish, regardless of how VB6 handles integer overflow.

```apex
End Function
```

---

</SwmSnippet>

## Final Block State Update

<SwmSnippet path="/HotelManagementSystem/Modules/clsBlowfish.cls" line="330">

---

We just returned from f, so now in `DecryptBlock` we finish the round by XORing with the next P-box value and updating the loop index. This step ensures each decryption round is fully reversed, restoring the original block state.

```apex
        Xl = Xl Xor m_pBox(j)
        j = j - 2
    Next
End Sub
```

---

</SwmSnippet>

&nbsp;

*This is an auto-generated document by Swimm 🌊 and has not yet been verified by a human*

<SwmMeta version="3.0.0" repo-id="Z2l0aHViJTNBJTNBY3RzLVZCNi1Qcm9qZWN0cyUzQSUzQVN3aW1tLURlbW8=" repo-name="cts-VB6-Projects"><sup>Powered by [Swimm](https://app.swimm.io/)</sup></SwmMeta>
