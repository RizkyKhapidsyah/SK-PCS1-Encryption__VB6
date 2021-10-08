Attribute VB_Name = "PSC1"

Public n As Long

Public PrvI As Long
Public PubI As Long

Public Prv As Long
Public Pub As Long

Public ValueIndex As Long

Sub GenKey(ByVal NMin As Long, ByVal NMax As Long)

Dim tPub As Long

Randomize

Top:

n = Int((NMax * Rnd) + NMax)
Prv = RndPrime(1, n)
Pub = Int((n * Rnd) + 1)

tPub = Pub
Do Until Pub * Prv Mod n = 1
   Pub = Pub + 1
   If Pub = tPub Then GoTo Top
   If Pub > n Then Pub = 1
Loop

PrvI = 1
PubI = n - PrvI
ValueIndex = 1

End Sub

Function RndPrime(Min As Long, Max As Long) As Long

LoopBig:
    
    RndPrime = Int((Max * Rnd) + Min)

loopSmall:
    
    RndPrime = RndPrime + 1
    If RndPrime > Max Then GoTo LoopBig
    If IsPrime(RndPrime) = False Then GoTo loopSmall
    If RndPrime = 0 Or RndPrime = 1 Then GoTo LoopBig

End Function

Private Function IsPrime(lngNumber) As Boolean

Dim lngCount As Long
Dim lngSqr As Long
Dim X As Long
lngSqr = Sqr(lngNumber)

If lngNumber < 2 Then
    IsPrime = False
    Exit Function
End If

lngCount = 2
IsPrime = True

If lngNumber Mod lngCount = 0& Then
    IsPrime = False
    Exit Function
End If

lngCount = 3

For X& = lngCount To lngSqr Step 2
    If lngNumber Mod X& = 0 Then
        IsPrime = False
        Exit Function
    End If
Next

End Function

Function Encrypt(m As Long) As Long

    Encrypt = ((m + PubI) * Pub) Mod n
    PubI = (PubI * (ValueIndex * m + 1)) Mod n
    
End Function

Function Decrypt(C As Long) As Long

    Decrypt = ((C * Prv) + PrvI) Mod n
    PrvI = (PrvI * (ValueIndex * Decrypt + 1)) Mod n
   
    ValueIndex = ValueIndex Mod n

End Function

Function EncryptBt(b As String) As String

    EncryptBt = Hex(Encrypt(Asc(Mid(b, 1, 1))))

End Function

Function DecryptBt(b As String) As String

    DecryptBt = Chr(Decrypt(Val("&H" + b)))

End Function

Function EncryptBk(Block As String) As String

Dim Length As Long
Dim iDX As Long

Length = Len(Block) + 1
iDX = 1
EncryptBk = ""

Do Until iDX = Length
    
    EncryptBk = EncryptBk + EncryptBt(Mid(Block, iDX, 1)) + " "
    iDX = iDX + 1

Loop

End Function

Function DecryptBk(Block As String) As String

Dim temp As String
Dim iDX As Long

temp = Block
iDX = 1
DecryptBk = ""


Do Until InStr(1, temp, " ") = 0

    DecryptBk = DecryptBk + DecryptBt(Mid(temp, 1, InStr(1, temp, " ")))
    temp = Mid(temp, InStr(1, temp, " ") + 1, Len(temp) - InStr(1, temp, " "))
    iDX = iDX + 1

Loop

End Function
