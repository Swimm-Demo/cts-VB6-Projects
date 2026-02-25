---
title: Block Encryption Flow
---
This document describes how a block of data is encrypted using the Blowfish algorithm. The input block undergoes multiple rounds of transformation, mixing with S-boxes and subkeys, resulting in a securely encrypted output block.

# Block Encryption Loop

<SwmSnippet path="/HotelManagementSystem/Modules/clsBlowfish.cls" line="334">

---

In <SwmToken path="HotelManagementSystem/Modules/clsBlowfish.cls" pos="334:6:6" line-data="Private Static Sub EncryptBlock(Xl As Long, Xr As Long)">`EncryptBlock`</SwmToken>, we're running the main Feistel rounds for Blowfish. Each iteration alternates between updating Xl and Xr using XOR with subkeys from <SwmToken path="HotelManagementSystem/Modules/clsBlowfish.cls" pos="338:9:9" line-data="        Xl = Xl Xor m_pBox(j)">`m_pBox`</SwmToken> and the output of f. Calling f here is what actually mixes the data using the S-boxes, so each round isn't just a simple XOR but a more complex transformation. This is what makes the block encryption strong.

```apex
Private Static Sub EncryptBlock(Xl As Long, Xr As Long)
    Dim I As Long, j As Long, Temp As Long
    j = 0
    For I = 0 To (Rounds \ 2 - 1)
        Xl = Xl Xor m_pBox(j)
        Xr = Xr Xor f(Xl)
        Xr = Xr Xor m_pBox(j + 1)
        Xl = Xl Xor f(Xr)
```

---

</SwmSnippet>

## S-Box Mixing Logic

<SwmSnippet path="/HotelManagementSystem/Modules/clsBlowfish.cls" line="423">

---

In <SwmToken path="HotelManagementSystem/Modules/clsBlowfish.cls" pos="423:6:6" line-data="Private Static Function f(ByVal X As Long) As Long">`f`</SwmToken>, we split X into bytes and use them to pull values from the S-boxes. Depending on <SwmToken path="HotelManagementSystem/Modules/clsBlowfish.cls" pos="426:4:4" line-data="    If (m_RunningCompiled) Then f = (((m_sBox(0, xb(3)) + m_sBox(1, xb(2))) Xor m_sBox(2, xb(1))) + m_sBox(3, xb(0))) Else f = UnsignedAdd((UnsignedAdd(m_sBox(0, xb(3)), m_sBox(1, xb(2))) Xor m_sBox(2, xb(1))), m_sBox(3, xb(0)))">`m_RunningCompiled`</SwmToken>, we either use regular addition or call <SwmToken path="HotelManagementSystem/Modules/clsBlowfish.cls" pos="426:75:75" line-data="    If (m_RunningCompiled) Then f = (((m_sBox(0, xb(3)) + m_sBox(1, xb(2))) Xor m_sBox(2, xb(1))) + m_sBox(3, xb(0))) Else f = UnsignedAdd((UnsignedAdd(m_sBox(0, xb(3)), m_sBox(1, xb(2))) Xor m_sBox(2, xb(1))), m_sBox(3, xb(0)))">`UnsignedAdd`</SwmToken> to combine the S-box values. <SwmToken path="HotelManagementSystem/Modules/clsBlowfish.cls" pos="426:75:75" line-data="    If (m_RunningCompiled) Then f = (((m_sBox(0, xb(3)) + m_sBox(1, xb(2))) Xor m_sBox(2, xb(1))) + m_sBox(3, xb(0))) Else f = UnsignedAdd((UnsignedAdd(m_sBox(0, xb(3)), m_sBox(1, xb(2))) Xor m_sBox(2, xb(1))), m_sBox(3, xb(0)))">`UnsignedAdd`</SwmToken> is needed to make sure addition doesn't break due to signed/unsigned overflow differences in VB6.

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

<SwmToken path="HotelManagementSystem/Modules/clsBlowfish.cls" pos="444:6:6" line-data="Private Static Function UnsignedAdd(ByVal Data1 As Long, Data2 As Long) As Long">`UnsignedAdd`</SwmToken> does unsigned 32-bit addition by splitting the inputs into bytes, adding each with carry, and reassembling the result. This avoids issues with VB6's signed math, so the cryptographic logic stays correct.

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

Back in <SwmToken path="HotelManagementSystem/Modules/clsBlowfish.cls" pos="339:9:9" line-data="        Xr = Xr Xor f(Xl)">`f`</SwmToken>, after returning from <SwmToken path="HotelManagementSystem/Modules/clsBlowfish.cls" pos="426:75:75" line-data="    If (m_RunningCompiled) Then f = (((m_sBox(0, xb(3)) + m_sBox(1, xb(2))) Xor m_sBox(2, xb(1))) + m_sBox(3, xb(0))) Else f = UnsignedAdd((UnsignedAdd(m_sBox(0, xb(3)), m_sBox(1, xb(2))) Xor m_sBox(2, xb(1))), m_sBox(3, xb(0)))">`UnsignedAdd`</SwmToken>, we finalize the S-box mixing and return the result. This ensures the output is consistent and correct for the rest of the encryption process.

```apex
End Function
```

---

</SwmSnippet>

## Block Finalization

<SwmSnippet path="/HotelManagementSystem/Modules/clsBlowfish.cls" line="342">

---

Finally, in <SwmToken path="HotelManagementSystem/Modules/clsBlowfish.cls" pos="334:6:6" line-data="Private Static Sub EncryptBlock(Xl As Long, Xr As Long)">`EncryptBlock`</SwmToken>, after the last call to f, we increment j, swap Xl and Xr, and apply the final XORs with the last two subkeys. This wraps up the block encryption, making sure both halves are fully processed and ready for output.

```apex
        j = j + 2
    Next
    Temp = Xr
    Xr = Xl Xor m_pBox(Rounds)
    Xl = Temp Xor m_pBox(Rounds + 1)
End Sub
```

---

</SwmSnippet>

&nbsp;

*This is an auto-generated document by Swimm 🌊 and has not yet been verified by a human*

<SwmMeta version="3.0.0" repo-id="Z2l0aHViJTNBJTNBY3RzLVZCNi1Qcm9qZWN0cyUzQSUzQVN3aW1tLURlbW8=" repo-name="cts-VB6-Projects"><sup>Powered by [Swimm](https://app.swimm.io/)</sup></SwmMeta>
