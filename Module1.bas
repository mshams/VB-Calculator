Attribute VB_Name = "Module1"
'---------------------------------------------
'-             VB Programming Project        -
'-                                           -
'-  Project Name  : Calculator               -
'-  Coded By      : Mohammad Shams Javi      -
'-  E-Mail        : info@mshams.ir           -
'-                                           -
'-  Copyright (c) 1384/8/9                   -
'---------------------------------------------

Public Declare Function PM Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function AW Lib "user32" Alias "AnimateWindow" (ByVal hwnd As Long, ByVal dwTime As Long, ByVal dwFlags As Long) As Boolean
Public c As Control    'object usage
Public opr As String   'operation
Public lopr As String  'last operation
Public nm As Double    'nm=current number
Public lnm As Double   'lnm=last number
Public mem As Double   'memory
Public ldo As Boolean  'last do =? operator
Public Const pi = 3.14159265358979
'Pi nummber can determined by  4*Atn(1)
                 
Public Function bin(n As Double) As String
'Note : Its funny that VB havnt Binary convertor function :)
'Convert to base2
Dim minus As Boolean 'if n is less than 0

If n < 0 Then minus = True
n = Int(Abs(n))
If n = 0 Then bin = "0"

Do While n > 0   'trim used to format string
   bin = Trim(Str(modh(n, 2))) + bin
   n = div(n, 2)
Loop
If minus Then bin = "-" + bin
End Function

Public Function octt(n As Double) As String
'Note : Oct function in VB cant convert huge numbers
'Convert to base8
Dim minus As Boolean 'if n is less than 0

If n < 0 Then minus = True
n = Int(Abs(n))
If n = 0 Then octt = "0"

Do While n > 0   'trim used to format string
   octt = Trim(Str(modh(n, 8))) + octt
   n = div(n, 8)
Loop
If minus Then octt = "-" + octt
End Function

Public Function hext(n As Double) As String
'Note : HEX function in VB cant convert huge numbers
'Convert to base16
Dim minus As Boolean 'if n is less than 0
Dim h As String * 2
Dim m As Byte

If n < 0 Then minus = True
n = Int(Abs(n))
If n = 0 Then hext = "0"

Do While n > 0
  m = modh(n, 16)
  Select Case m
    Case 10: h = "A"
    Case 11: h = "B"
    Case 12: h = "C"
    Case 13: h = "D"
    Case 14: h = "E"
    Case 15: h = "F"
    Case Else: h = Str(m)
  End Select
  hext = Trim(h) + hext
  n = div(n, 16)
Loop
If minus Then hext = "-" + hext
End Function

Public Function Dec(s As String, base As Byte) As Double
'Convert any base to Decimal
Dim minus As Boolean 'if s is less than 0
Dim X As String
Dim i As Byte, n As Byte

If Left(s, 1) = "-" Then     'Del minus sign "-"
    s = Right(s, Len(s) - 1)
    minus = True
End If
If s = "" Then Dec = "0"

s = Trim(s)
Do While Len(s) > 0
  X = Trim(Right(s, 1))
  s = Left(s, Len(s) - 1)
  Select Case X             'For hex numbers
    Case "A": n = 10
    Case "B": n = 11
    Case "C": n = 12
    Case "D": n = 13
    Case "E": n = 14
    Case "F": n = 15
    Case "0" To "9": n = Val(X)
  End Select
  Dec = Dec + n * base ^ i
  i = i + 1
Loop
If minus Then Dec = Dec * -1
End Function

Public Function modh(X, Y As Double) As Double
'Note : You cant use Mod operator on huge number
modh = (X - (Y * Fix(X / Y)))
End Function

Public Function div(X, Y As Byte) As Double
'Note : You cant use \ operator on huge number
div = Int(X / Y)
End Function

Public Function Lsh(X As Double, Y As Double) As Double
Dim r As String
r = Trim(bin(X))
For i = 1 To Y
    r = Right(r, Len(r) - 1)
    r = r + "0"
Next
Lsh = Dec(r, 2)
End Function

Public Function Rsh(X As Double, Y As Double) As Double
Dim r As String
r = Trim(bin(X))
For i = 1 To Y
    r = Left(r, Len(r) - 1)
    r = "0" + r
Next
Rsh = Dec(r, 2)
End Function

Public Function fact(n As Double) As Double
'Factorial function
n = Int(Abs(n))
If n <> 0 Then fact = 1
While n > 1
    fact = fact * n
    n = n - 1
Wend
End Function

Public Sub disp(ByVal n As Double, Optin As Byte, Optional base As Byte = 0)
'This is display formatter subroutine
'I used Base parameter to display numbers that is none decimal
'   Base=0 Decimal   Base<>0 Bin/Oct/Hex
Dim t As String 'text of display
Dim idx As Byte  'index of enabled option button

For Each c In calc.Optn  'get enabled option index
    If c.Value = True Then idx = c.Index
Next

t = calc.t(idx)
If t = "0" Then t = ""

Select Case Optin
    Case 0: t = "0" 'clear
    Case 1:
        If idx = 2 And n > 9 And base = 0 Then
            t = Chr(n + 53) 'Hexadecimal
            Else: t = Trim(Str(n)) 'set to a number
        End If
        
    Case 2:
        If idx = 2 And n > 9 And base = 0 Then
            t = t + Chr(n + 53) 'Hexadecimal
            Else: t = t + Trim(Str(n)) 'add a number
        End If
    
    Case 3:                         'backspace
        If Len(t) > 1 Then
            t = Left$(t, (Len(t) - 1))
        Else
            disp 0, 0: t = 0
        End If
    
    Case 4:                         ' operator -/+
        If n = 10 And Val(t) <> "0" Then
            If Left$(t, 1) <> "-" Then
                t = "-" + t
                Else: t = Right$(t, Len(t) - 1)
            End If
        End If                      'operator "."
        If n = 11 And InStr(t, ".") = 0 Then t = Trim(Str(Val(t))) & "."
End Select

'Note : I used the Tag of OptionButtons to
'        determine base of number.
'        calc.Optn(idx).Tag = Base Number

'if number is none deciml convert it and put it in t(0)
If idx <> 0 And base = 0 Then
    calc.t(0) = Dec(t, Val(calc.Optn(idx).Tag))
    t = calc.t(0)
End If

'display other base numbers
calc.t(1) = octt(Val(t))
calc.t(2) = hext(Val(t))
calc.t(3) = bin(Val(t))
If t = "" Then t = "0"
calc.t(0) = t       'Set decimal display
nm = Val(calc.t(0)) 'Set numeric variable
End Sub

Public Sub proc(operator As String)
'This is proccessor subroutine that proccess operators
Dim r As Double 'r=result
On Error GoTo 10

Select Case operator
    Case "+": r = lnm + nm
    Case "-": r = lnm - nm
    Case "/": r = lnm / nm
    Case "*": r = lnm * nm
    
    Case "Mod": r = modh(lnm, nm)
    Case "And": r = lnm And nm
    Case "Or": r = lnm Or nm
    Case "Xor": r = lnm Xor nm
    Case "Eqv": r = lnm Eqv nm
    Case "Imp": r = lnm Imp nm
    Case "Lsh": r = Lsh(lnm, nm)
    Case "Rsh": r = Rsh(lnm, nm)
    
    Case "x^y": r = lnm ^ nm
    Case "%": r = lnm * nm / 100
    
'Semi operators

    Case "M+": mem = mem + nm: r = nm
    Case "MR": r = mem: lnm = nm
    Case "MC": mem = 0: r = nm
    Case "MS": mem = nm: r = nm
    
    Case "Sqrt": r = Sqr(nm)
    Case "Sin": r = Sin(nm * pi / 180)
    Case "Cos": r = Cos(nm * pi / 180)
    Case "Tan": r = Tan(nm * pi / 180)
    Case "Pi": r = pi
    Case "Fix": r = Fix(nm)
    Case "Int": r = Round(nm)
    Case "x^2": r = nm * nm
    Case "n!": r = fact(nm)
    Case "Rnd": r = Rnd
    Case "Not": r = Not nm
    Case "1/x":
        If nm = 0 Then Exit Sub
        r = 1 / nm
    Case "Log":
        If nm = 0 Then Exit Sub
        r = Log(nm)
End Select

Call disp(r, 1, 1)
'-----------------------Error-handling--------------
Exit Sub
10: MsgBox Error$(Err), vbCritical, "Error Reporter"
End Sub

Public Sub Ab_me()
'About Me
MsgBox "             VB Programming Project        " + vbCrLf + _
 vbCrLf + _
"  Project Name  : Calculator" + vbCrLf + _
"  Coded By        : Mohammad Shams Javi" + vbCrLf + _
"  Mail                 : info@mshams.ir" + vbCrLf + vbCrLf + _
vbTab + "Copyright (c) 1384/8/9", , "About Me"
End Sub

Public Sub F_act()
'Form activate animation
calc.Cls
For Each c In calc.Controls
    If c <> TextBox Then AW c.hwnd, 100, &H40002
    AW c.hwnd, 10, &H10
Next
End Sub

Public Sub F_unl()
'Form unload animation
For Each c In calc.Controls
    AW c.hwnd, 15, &H50010
Next
AW hwnd, 250, &H10010
Unload calc
End Sub
