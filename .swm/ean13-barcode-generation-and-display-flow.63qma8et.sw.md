---
title: EAN13 Barcode Generation and Display Flow
---
This document describes the process of generating and displaying an <SwmToken path="BARCODE/arisBarcode.cls" pos="22:2:2" line-data="Function EAN13(zObj As Object, zBarH As Integer, BarText As String, Optional ByVal HasCaption As Boolean = False)">`EAN13`</SwmToken> barcode image. Users provide barcode parameters, which are used to render the barcode and optional captions in a visual control. The barcode is drawn according to <SwmToken path="BARCODE/arisBarcode.cls" pos="22:2:2" line-data="Function EAN13(zObj As Object, zBarH As Integer, BarText As String, Optional ByVal HasCaption As Boolean = False)">`EAN13`</SwmToken> standards and displayed for further use.

The main steps are:

- Trigger barcode generation
- Prepare encoding tables and rendering surface
- Render barcode bars and captions
- Finalize and display the barcode image
- Return control to the caller

# Triggering <SwmToken path="BARCODE/arisBarcode.cls" pos="22:2:2" line-data="Function EAN13(zObj As Object, zBarH As Integer, BarText As String, Optional ByVal HasCaption As Boolean = False)">`EAN13`</SwmToken> Barcode Generation

<SwmSnippet path="/BARCODE/arisBarcode.cls" line="22">

---

In <SwmToken path="BARCODE/arisBarcode.cls" pos="22:2:2" line-data="Function EAN13(zObj As Object, zBarH As Integer, BarText As String, Optional ByVal HasCaption As Boolean = False)">`EAN13`</SwmToken>, we just delegate to <SwmToken path="BARCODE/arisBarcode.cls" pos="23:1:1" line-data="    BarEAN &quot;EAN13&quot;, zObj, zBarH, BarText, HasCaption">`BarEAN`</SwmToken>, passing all the relevant parameters and explicitly specifying <SwmToken path="BARCODE/arisBarcode.cls" pos="22:2:2" line-data="Function EAN13(zObj As Object, zBarH As Integer, BarText As String, Optional ByVal HasCaption As Boolean = False)">`EAN13`</SwmToken> as the barcode type. This means <SwmToken path="BARCODE/arisBarcode.cls" pos="22:2:2" line-data="Function EAN13(zObj As Object, zBarH As Integer, BarText As String, Optional ByVal HasCaption As Boolean = False)">`EAN13`</SwmToken> is just a convenience wrapper for generating <SwmToken path="BARCODE/arisBarcode.cls" pos="22:2:2" line-data="Function EAN13(zObj As Object, zBarH As Integer, BarText As String, Optional ByVal HasCaption As Boolean = False)">`EAN13`</SwmToken> barcodes, and the actual work happens in <SwmToken path="BARCODE/arisBarcode.cls" pos="23:1:1" line-data="    BarEAN &quot;EAN13&quot;, zObj, zBarH, BarText, HasCaption">`BarEAN`</SwmToken>.

```apex
Function EAN13(zObj As Object, zBarH As Integer, BarText As String, Optional ByVal HasCaption As Boolean = False)
    BarEAN "EAN13", zObj, zBarH, BarText, HasCaption
```

---

</SwmSnippet>

## Preparing Barcode Rendering and Encoding Tables

<SwmSnippet path="/BARCODE/EAN13.bas" line="16">

---

In <SwmToken path="BARCODE/EAN13.bas" pos="16:2:2" line-data="Sub BarEAN(BarType As String, Obj As Object, BarH As Integer, BarText As String, Optional ByVal HasCaption As Boolean = False)">`BarEAN`</SwmToken>, we set up the target object and immediately call <SwmToken path="BARCODE/EAN13.bas" pos="18:1:1" line-data="   Init_Table">`Init_Table`</SwmToken> to make sure all encoding tables are ready for barcode rendering. Without this, the barcode logic wouldn't have the necessary data to encode digits.

```visual basic
Sub BarEAN(BarType As String, Obj As Object, BarH As Integer, BarText As String, Optional ByVal HasCaption As Boolean = False)
   Set zObj = Obj
   Init_Table
```

---

</SwmSnippet>

<SwmSnippet path="/BARCODE/ModCode128.bas" line="59">

---

<SwmToken path="BARCODE/ModCode128.bas" pos="59:4:4" line-data="Private Sub Init_Table()">`Init_Table`</SwmToken> here sets up the character set and encoding patterns for Code 128 barcodes. Even though we're focused on <SwmToken path="BARCODE/arisBarcode.cls" pos="22:2:2" line-data="Function EAN13(zObj As Object, zBarH As Integer, BarText As String, Optional ByVal HasCaption As Boolean = False)">`EAN13`</SwmToken>, this setup is shared infrastructure for all supported barcode types.

```visual basic
Private Sub Init_Table()
    CharSet = " !""#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}~"
    arrEncoding = Array( _
             "11011001100", "11001101100", "11001100110", "10010011000", "10010001100", "10001001100", _
             "10011001000", "10011000100", "10001100100", "11001001000", "11001000100", "11000100100", _
             "10110011100", "10011011100", "10011001110", "10111001100", "10011101100", "10011100110", _
             "11001110010", "11001011100", "11001001110", "11011100100", "11001110100", "11101101110", _
             "11101001100", "11100101100", "11100100110", "11101100100", "11100110100", "11100110010", _
             "11011011000", "11011000110", "11000110110", "10100011000", "10001011000", "10001000110", _
             "10110001000", "10001101000", "10001100010", "11010001000", "11000101000", "11000100010", _
             "10110111000", "10110001110", "10001101110", "10111011000", "10111000110", "10001110110", _
             "11101110110", "11010001110", "11000101110", "11011101000", "11011100010", "11011101110", _
             "11101011000", "11101000110", "11100010110", "11101101000", "11101100010", "11100011010", _
             "11101111010", "11001000010", "11110001010", "10100110000", "10100001100", "10010110000", _
             "10010000110", "10000101100", "10000100110", "10110010000", "10110000100", "10011010000", _
             "10011000010", "10000110100", "10000110010", "11000010010", "11001010000", "11110111010", _
             "11000010100", "10001111010", "10100111100", "10010111100", "10010011110", "10111100100", _
             "10011110100", "10011110010", "11110100100", "11110010100", "11110010010", "11011011110", _
             "11011110110", "11110110110", "10101111000", "10100011110", "10001011110", "10111101000", _
             "10111100010", "11110101000", "11110100010", "10111011110", "10111101110", "11101011110", _
             "11110101110", "11010000100", "11010010000", "11010011100", "11000111010", "11" _
             )
End Sub
```

---

</SwmSnippet>

<SwmSnippet path="/BARCODE/EAN13.bas" line="19">

---

Back in <SwmPath>[BARCODE/EAN13.bas](BARCODE/EAN13.bas)</SwmPath>, after setting up the encoding tables, we validate the barcode text, prep the drawing surface, and size the control based on the barcode and caption. Then, depending on the barcode type, we call <SwmToken path="BARCODE/EAN13.bas" pos="43:1:1" line-data="      Paint_Bar13 zBarText">`Paint_Bar13`</SwmToken> to actually render the <SwmToken path="BARCODE/EAN13.bas" pos="42:8:8" line-data="   If zBarType = &quot;EAN13&quot; Then">`EAN13`</SwmToken> barcode.

```visual basic
   
   zBarText = BarText
   zHasCaption = HasCaption
   zBarType = BarType
   zBarH = BarH * 72
   
   If Not CheckCode Then Exit Sub
   
   zObj.Picture = Nothing
   zObj.BackColor = vbWhite
   zObj.AutoRedraw = True
   zObj.ScaleMode = 3
   
   If zHasCaption Then
      zObj.Height = (zObj.TextHeight(zBarText) + zBarH + 5) * Screen.TwipsPerPixelY
   Else
      zObj.Height = zBarH * Screen.TwipsPerPixelY
   End If
   
   zObj.Height = zObj.Height + 10 ' Border
   zObj.Width = (Len(zBarText) * 7) + (zObj.TextWidth(Mid(zBarText, 1, 1)) * 2) + 30
   zObj.Width = zObj.Width * Screen.TwipsPerPixelX
   
   If zBarType = "EAN13" Then
      Paint_Bar13 zBarText
```

---

</SwmSnippet>

### Rendering <SwmToken path="BARCODE/arisBarcode.cls" pos="22:2:2" line-data="Function EAN13(zObj As Object, zBarH As Integer, BarText As String, Optional ByVal HasCaption As Boolean = False)">`EAN13`</SwmToken> Bars and Captions

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
    node1["Start: Prepare to draw EAN-13 barcode"] --> node2{"Show human-readable caption?"}
    click node1 openCode "BARCODE/EAN13.bas:70:73"
    node2 -->|"Yes"| node3["Draw left caption above barcode"]
    click node2 openCode "BARCODE/EAN13.bas:76:81"
    node2 -->|"No"| node4
    node3 --> node4["Draw start guard pattern"]
    click node3 openCode "BARCODE/EAN13.bas:82:83"
    node4 --> subgraph loop1["For each digit in barcode number"]
        node5{"Is this the middle digit?"}
        click node5 openCode "BARCODE/EAN13.bas:89:97"
        node5 -->|"Yes"| node6["Draw middle guard pattern"]
        click node6 openCode "BARCODE/EAN13.bas:96:97"
        node6 --> node7{"Is digit on left (2-7) or right (8-13)?"}
        node5 -->|"No"| node7
        node7 -->|"Left (2-7)"| node8["Draw left digit with parity"]
        click node8 openCode "BARCODE/EAN13.bas:99:100"
        node7 -->|"Right (8-13)"| node9["Draw right digit"]
        click node9 openCode "BARCODE/EAN13.bas:101:102"
        node8 --> node10["Next digit"]
        node9 --> node10
        node10 --> node5
    end
    node10 --> node11["Calculate and draw checksum digit"]
    click node11 openCode "BARCODE/EAN13.bas:105:110"
    node11 --> node12["Draw end guard pattern"]
    click node12 openCode "BARCODE/EAN13.bas:113:113"
    node12 --> node13{"Show human-readable caption below
barcode?"}
    click node13 openCode "BARCODE/EAN13.bas:116:126"
    node13 -->|"Yes"| node14["Draw center and right captions below
barcode"]
    click node14 openCode "BARCODE/EAN13.bas:117:125"
    node13 -->|"No"| node15["Finish"]
    node14 --> node15
classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;

%% Swimm:
%% %%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
%% flowchart TD
%%     node1["Start: Prepare to draw EAN-13 barcode"] --> node2{"Show human-readable caption?"}
%%     click node1 openCode "<SwmPath>[BARCODE/EAN13.bas](BARCODE/EAN13.bas)</SwmPath>:70:73"
%%     node2 -->|"Yes"| node3["Draw left caption above barcode"]
%%     click node2 openCode "<SwmPath>[BARCODE/EAN13.bas](BARCODE/EAN13.bas)</SwmPath>:76:81"
%%     node2 -->|"No"| node4
%%     node3 --> node4["Draw start guard pattern"]
%%     click node3 openCode "<SwmPath>[BARCODE/EAN13.bas](BARCODE/EAN13.bas)</SwmPath>:82:83"
%%     node4 --> subgraph loop1["For each digit in barcode number"]
%%         node5{"Is this the middle digit?"}
%%         click node5 openCode "<SwmPath>[BARCODE/EAN13.bas](BARCODE/EAN13.bas)</SwmPath>:89:97"
%%         node5 -->|"Yes"| node6["Draw middle guard pattern"]
%%         click node6 openCode "<SwmPath>[BARCODE/EAN13.bas](BARCODE/EAN13.bas)</SwmPath>:96:97"
%%         node6 --> node7{"Is digit on left (2-7) or right (8-13)?"}
%%         node5 -->|"No"| node7
%%         node7 -->|"Left (2-7)"| node8["Draw left digit with parity"]
%%         click node8 openCode "<SwmPath>[BARCODE/EAN13.bas](BARCODE/EAN13.bas)</SwmPath>:99:100"
%%         node7 -->|"Right (8-13)"| node9["Draw right digit"]
%%         click node9 openCode "<SwmPath>[BARCODE/EAN13.bas](BARCODE/EAN13.bas)</SwmPath>:101:102"
%%         node8 --> node10["Next digit"]
%%         node9 --> node10
%%         node10 --> node5
%%     end
%%     node10 --> node11["Calculate and draw checksum digit"]
%%     click node11 openCode "<SwmPath>[BARCODE/EAN13.bas](BARCODE/EAN13.bas)</SwmPath>:105:110"
%%     node11 --> node12["Draw end guard pattern"]
%%     click node12 openCode "<SwmPath>[BARCODE/EAN13.bas](BARCODE/EAN13.bas)</SwmPath>:113:113"
%%     node12 --> node13{"Show human-readable caption below
%% barcode?"}
%%     click node13 openCode "<SwmPath>[BARCODE/EAN13.bas](BARCODE/EAN13.bas)</SwmPath>:116:126"
%%     node13 -->|"Yes"| node14["Draw center and right captions below
%% barcode"]
%%     click node14 openCode "<SwmPath>[BARCODE/EAN13.bas](BARCODE/EAN13.bas)</SwmPath>:117:125"
%%     node13 -->|"No"| node15["Finish"]
%%     node14 --> node15
%% classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/BARCODE/EAN13.bas" line="70">

---

In <SwmToken path="BARCODE/EAN13.bas" pos="70:4:4" line-data="Private Sub Paint_Bar13(ByVal xstr As String)">`Paint_Bar13`</SwmToken>, we handle the actual drawing of the <SwmToken path="BARCODE/arisBarcode.cls" pos="22:2:2" line-data="Function EAN13(zObj As Object, zBarH As Integer, BarText As String, Optional ByVal HasCaption As Boolean = False)">`EAN13`</SwmToken> barcode. The function figures out the parity pattern, draws the guard bars, loops through each digit to draw the correct bar pattern, and prints captions if needed. The actual bar drawing is delegated to <SwmToken path="BARCODE/EAN13.bas" pos="82:1:1" line-data="    Draw_Bar &quot;101&quot;, True">`Draw_Bar`</SwmToken>, which is why we need to call into <SwmPath>[BARCODE/Bar39.cls](BARCODE/Bar39.cls)</SwmPath> next.

```visual basic
Private Sub Paint_Bar13(ByVal xstr As String)
    Dim ii As Long, jj As Integer, ctr As Integer, xEven As Boolean, xParity As String
 
    xTotal = 0
    xPos = 11
    
    If zHasCaption Then
        zObj.CurrentX = xPos
        zObj.CurrentY = zBarH - zObj.TextHeight(zBarText)
        zObj.Print Mid(xstr, 1, 1)
        xPos = xPos + zObj.TextWidth(Mid(xstr, 1, 1)) + 1
    End If
    Draw_Bar "101", True
    StartX = zObj.CurrentX
    
    zObj.CurrentY = 15 + zBarH
    xParity = Parity(CInt(Mid(xstr, 1, 1)))
    
    
    For ii = 1 To Len(xstr)
        If ((Len(xstr) + 1) - ii) Mod 2 = 0 Then
           xTotal = xTotal + (CInt(Mid(xstr, ii, 1)))
        Else ' Odd
           xTotal = xTotal + CInt(Mid(xstr, ii, 1) * 3)
        End If
        If ii = 8 Then
           Draw_Bar "01010", True
        End If
        jj = CInt(Mid(xstr, ii, 1))
        If ii > 1 And ii < 8 Then
           Draw_Bar CStr(IIf(Mid(xParity, ii - 1, 1) = "E", LeftHand_Even(jj), LeftHand_Odd(jj))), False
        ElseIf ii > 1 And ii >= 8 Then
           Draw_Bar CStr(Right_Hand(jj)), False
        End If
    Next
    chkSum = 0
    jj = xTotal Mod 10
    If jj <> 0 Then
       chkSum = 10 - jj
    End If
    Draw_Bar CStr(Right_Hand(chkSum)), False
    
    EndX = zObj.CurrentX
    Draw_Bar "101", True
    
   
   If zHasCaption Then
        Dim centerX As Double
        centerX = ((EndX - StartX) / 2)
        zObj.CurrentX = ((centerX - zObj.TextWidth(Mid(xstr, 2, 6))) / 2) + StartX  '23
        zObj.CurrentY = 2 + zBarH
        zObj.Print Mid(xstr, 2, 6)
        
        zObj.CurrentX = ((centerX - zObj.TextWidth(Mid(xstr, 8, 6))) / 2) + StartX + centerX
        zObj.CurrentY = 2 + zBarH
        zObj.Print Mid(xstr, 8, 6) & chkSum
    End If
End Sub
Private Sub Paint_Bar8(ByVal xstr As String)
    Dim ii As Long, jj As Integer, ctr As Integer, xEven As Boolean, xParity As String
 
    xTotal = 0
    xPos = 11
    
    
    Draw_Bar "101", True ' Start
    StartX = zObj.CurrentX
    
    zObj.CurrentX = xPos
    zObj.CurrentY = zBarH + 15
    xParity = Parity(7) 'CInt(Mid(xstr, 1, 1)))
    
    
    For ii = 1 To Len(xstr)
        If ((Len(xstr) + 1) - ii) Mod 2 = 0 Then 'EVEN
           xTotal = xTotal + (CInt(Mid(xstr, ii, 1)))
        Else ' Odd
           xTotal = xTotal + CInt(Mid(xstr, ii, 1) * 3)
        End If
        If ii = 5 Then
           Draw_Bar "01010", True ' Middle
        End If
        jj = CInt(Mid(xstr, ii, 1))
        If ii < 5 Then
           Draw_Bar CStr(LeftHand_Odd(jj)), False
        ElseIf ii >= 5 Then
           Draw_Bar CStr(Right_Hand(jj)), False
        End If
    Next
    chkSum = 0
    jj = xTotal Mod 10
    If jj <> 0 Then
       chkSum = 10 - jj
    End If
    Draw_Bar CStr(Right_Hand(chkSum)), False
    
    EndX = zObj.CurrentX
    Draw_Bar "101", True
    
    If zHasCaption Then
        Dim centerX As Double
        centerX = ((EndX - StartX) / 2)
        zObj.CurrentX = ((centerX - zObj.TextWidth(Mid(xstr, 1, 4))) / 2) + StartX
        zObj.CurrentY = 2 + zBarH
        zObj.Print Mid(xstr, 1, 4)
        
        zObj.CurrentX = ((centerX - zObj.TextWidth(Mid(xstr, 4, 4))) / 2) + StartX + centerX
        zObj.CurrentY = 2 + zBarH
        zObj.Print Mid(xstr, 5, 3) & chkSum
    End If
End Sub

Private Sub Draw_Bar(encoding As String, Guard As Boolean)
```

---

</SwmSnippet>

<SwmSnippet path="/BARCODE/Bar39.cls" line="87">

---

<SwmToken path="BARCODE/Bar39.cls" pos="87:4:4" line-data="Private Sub Draw_Bar(Encoding As String)">`Draw_Bar`</SwmToken> in <SwmPath>[BARCODE/Bar39.cls](BARCODE/Bar39.cls)</SwmPath> loops through the encoding string, drawing a vertical line for each character, black for '1' and white for '0', incrementing the x-position each time. This is the low-level rendering of the barcode pattern.

```apex
Private Sub Draw_Bar(Encoding As String)
    Dim ii As Integer
    For ii = 1 To Len(Encoding)
        xPos = xPos + 1
        xObj.Line (xPos + 10, xtop)-(xPos + 10, xtop + BarH), IIf(Mid(Encoding, ii, 1), vbBlack, vbWhite)
    Next
End Sub
```

---

</SwmSnippet>

<SwmSnippet path="/BARCODE/EAN13.bas" line="183">

---

Back in <SwmPath>[BARCODE/EAN13.bas](BARCODE/EAN13.bas)</SwmPath>, after drawing the bars, we set up the digit encoding tables (<SwmToken path="BARCODE/EAN13.bas" pos="100:35:35" line-data="           Draw_Bar CStr(IIf(Mid(xParity, ii - 1, 1) = &quot;E&quot;, LeftHand_Even(jj), LeftHand_Odd(jj))), False">`LeftHand_Odd`</SwmToken>, <SwmToken path="BARCODE/EAN13.bas" pos="100:29:29" line-data="           Draw_Bar CStr(IIf(Mid(xParity, ii - 1, 1) = &quot;E&quot;, LeftHand_Even(jj), LeftHand_Odd(jj))), False">`LeftHand_Even`</SwmToken>, <SwmToken path="BARCODE/EAN13.bas" pos="102:5:5" line-data="           Draw_Bar CStr(Right_Hand(jj)), False">`Right_Hand`</SwmToken>, Parity) so <SwmToken path="BARCODE/EAN13.bas" pos="43:1:1" line-data="      Paint_Bar13 zBarText">`Paint_Bar13`</SwmToken> can map digits to the right bar patterns.

```visual basic
    Dim ii As Integer
    For ii = 1 To Len(encoding)
        xPos = xPos + 1
        zObj.Line (xPos, 10)-(xPos, zBarH + IIf(Guard, 5, 0)), IIf(Mid(encoding, ii, 1), vbBlack, vbWhite)
    Next
End Sub
Private Sub Init_Table()
```

---

</SwmSnippet>

<SwmSnippet path="/BARCODE/EAN13.bas" line="190">

---

<SwmToken path="BARCODE/EAN13.bas" pos="18:1:1" line-data="   Init_Table">`Init_Table`</SwmToken> here fills the arrays for left-hand odd, left-hand even, right-hand, and parity encodings. These are used by <SwmToken path="BARCODE/EAN13.bas" pos="43:1:1" line-data="      Paint_Bar13 zBarText">`Paint_Bar13`</SwmToken> to translate each digit into the correct bar pattern for <SwmToken path="BARCODE/arisBarcode.cls" pos="22:2:2" line-data="Function EAN13(zObj As Object, zBarH As Integer, BarText As String, Optional ByVal HasCaption As Boolean = False)">`EAN13`</SwmToken>.

```visual basic
    LeftHand_Odd = Array("0001101", "0011001", "0010011", "0111101", "0100011", "0110001", "0101111", "0111011", "0110111", "0001011")
    LeftHand_Even = Array("0100111", "0110011", "0011011", "0100001", "0011101", "0111001", "0000101", "0010001", "0001001", "0010111")
    Right_Hand = Array("1110010", "1100110", "1101100", "1000010", "1011100", "1001110", "1010000", "1000100", "1001000", "1110100")
    Parity = Array("OOOOOO", "OOEOEE", "OOEEOE", "OOEEEO", "OEOOEE", "OEEOOE", "OEEEOO", "OEOEOE", "OEOEEO", "OEEOEO")
End Sub
```

---

</SwmSnippet>

### Finalizing and Displaying the Barcode Image

<SwmSnippet path="/BARCODE/EAN13.bas" line="44">

---

Back in <SwmPath>[BARCODE/EAN13.bas](BARCODE/EAN13.bas)</SwmPath>, after drawing the barcode (either <SwmToken path="BARCODE/arisBarcode.cls" pos="22:2:2" line-data="Function EAN13(zObj As Object, zBarH As Integer, BarText As String, Optional ByVal HasCaption As Boolean = False)">`EAN13`</SwmToken> or <SwmToken path="BARCODE/arisBarcode.cls" pos="25:2:2" line-data="Function EAN8(zObj As Object, zBarH As Integer, BarText As String, Optional ByVal HasCaption As Boolean = False)">`EAN8`</SwmToken>), we assign the rendered image to the Picture property so it actually shows up in the control.

```visual basic
   Else
      Paint_Bar8 zBarText
   End If
   
   zObj.Picture = zObj.Image
End Sub
```

---

</SwmSnippet>

## Returning Control to the Caller

<SwmSnippet path="/BARCODE/arisBarcode.cls" line="24">

---

Finally, in <SwmToken path="BARCODE/arisBarcode.cls" pos="22:2:2" line-data="Function EAN13(zObj As Object, zBarH As Integer, BarText As String, Optional ByVal HasCaption As Boolean = False)">`EAN13`</SwmToken>, after <SwmToken path="BARCODE/arisBarcode.cls" pos="23:1:1" line-data="    BarEAN &quot;EAN13&quot;, zObj, zBarH, BarText, HasCaption">`BarEAN`</SwmToken> finishes, we just return. There's nothing else to do here—the barcode is already rendered.

```apex
End Function
```

---

</SwmSnippet>

&nbsp;

*This is an auto-generated document by Swimm 🌊 and has not yet been verified by a human*

<SwmMeta version="3.0.0" repo-id="Z2l0aHViJTNBJTNBY3RzLVZCNi1Qcm9qZWN0cyUzQSUzQVN3aW1tLURlbW8=" repo-name="cts-VB6-Projects"><sup>Powered by [Swimm](https://app.swimm.io/)</sup></SwmMeta>
