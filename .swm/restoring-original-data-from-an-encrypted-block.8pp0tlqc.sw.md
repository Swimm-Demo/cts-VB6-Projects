---
title: Restoring original data from an encrypted block
---
This document describes the process of restoring original data from an encrypted block. The flow takes an encrypted block, applies a series of transformations to reverse the encryption, and outputs the original data.

# Block Decryption Loop Setup

<SwmSnippet path="/HotelManagementSystem/Modules/clsBlowfish.cls" line="320">

---

In <SwmToken path="HotelManagementSystem/Modules/clsBlowfish.cls" pos="320:6:6" line-data="Private Static Sub DecryptBlock(Xl As Long, Xr As Long)">`DecryptBlock`</SwmToken>, we set up the initial state for decryption by swapping and XORing Xl and Xr with the last two P-box values. The loop then repeatedly applies f to mix the halves and undoes the encryption rounds in reverse. Calling f here is what actually scrambles the halves back toward their original values.

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

In <SwmToken path="HotelManagementSystem/Modules/clsBlowfish.cls" pos="423:6:6" line-data="Private Static Function f(ByVal X As Long) As Long">`f`</SwmToken>, we break the input into bytes and use them to index into S-boxes, mixing the values with addition and XOR. If we're not running compiled, we call <SwmToken path="HotelManagementSystem/Modules/clsBlowfish.cls" pos="426:75:75" line-data="    If (m_RunningCompiled) Then f = (((m_sBox(0, xb(3)) + m_sBox(1, xb(2))) Xor m_sBox(2, xb(1))) + m_sBox(3, xb(0))) Else f = UnsignedAdd((UnsignedAdd(m_sBox(0, xb(3)), m_sBox(1, xb(2))) Xor m_sBox(2, xb(1))), m_sBox(3, xb(0)))">`UnsignedAdd`</SwmToken> to make sure the math matches unsigned 32-bit behavior, which is needed for correct decryption.

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

<SwmToken path="HotelManagementSystem/Modules/clsBlowfish.cls" pos="444:6:6" line-data="Private Static Function UnsignedAdd(ByVal Data1 As Long, Data2 As Long) As Long">`UnsignedAdd`</SwmToken> handles 32-bit unsigned addition by adding each byte with carry, so the result matches what Blowfish expects, even though VB6 doesn't support unsigned longs.

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

Back in <SwmToken path="HotelManagementSystem/Modules/clsBlowfish.cls" pos="327:9:9" line-data="        Xl = Xl Xor f(Xr)">`f`</SwmToken>, we return the mixed value, using <SwmToken path="HotelManagementSystem/Modules/clsBlowfish.cls" pos="426:75:75" line-data="    If (m_RunningCompiled) Then f = (((m_sBox(0, xb(3)) + m_sBox(1, xb(2))) Xor m_sBox(2, xb(1))) + m_sBox(3, xb(0))) Else f = UnsignedAdd((UnsignedAdd(m_sBox(0, xb(3)), m_sBox(1, xb(2))) Xor m_sBox(2, xb(1))), m_sBox(3, xb(0)))">`UnsignedAdd`</SwmToken> if needed, so the output always matches Blowfish's unsigned math regardless of VB6 mode.

```apex
End Function
```

---

</SwmSnippet>

## Final Block Unmixing

<SwmSnippet path="/HotelManagementSystem/Modules/clsBlowfish.cls" line="330">

---

Back in <SwmToken path="HotelManagementSystem/Modules/clsBlowfish.cls" pos="320:6:6" line-data="Private Static Sub DecryptBlock(Xl As Long, Xr As Long)">`DecryptBlock`</SwmToken>, after each call to f, we XOR Xl with the next P-box value and update the loop counter. This step reverses the last layer of mixing for each round, moving the block closer to its original state.

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
