VERSION 5.00
Begin VB.Form calc 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator  By  M.SH"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   6135
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.TextBox t 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Height          =   315
      HideSelection   =   0   'False
      Index           =   3
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "0"
      Top             =   1170
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.TextBox t 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Height          =   315
      HideSelection   =   0   'False
      Index           =   2
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "0"
      Top             =   810
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.TextBox t 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Height          =   315
      HideSelection   =   0   'False
      Index           =   1
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   48
      Text            =   "0"
      Top             =   450
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.TextBox t 
      Alignment       =   1  'Right Justify
      Height          =   315
      HideSelection   =   0   'False
      Index           =   0
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "0"
      Top             =   90
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.OptionButton Optn 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Bin"
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   49
      Tag             =   "2"
      Top             =   1170
      Visible         =   0   'False
      Width           =   800
   End
   Begin VB.OptionButton Optn 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Hex"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   50
      Tag             =   "16"
      Top             =   810
      Visible         =   0   'False
      Width           =   800
   End
   Begin VB.OptionButton Optn 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Oct"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   51
      Tag             =   "8"
      Top             =   450
      Visible         =   0   'False
      Width           =   800
   End
   Begin VB.OptionButton Optn 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Dec"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   52
      Tag             =   "10"
      Top             =   90
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   800
   End
   Begin VB.CommandButton num 
      BackColor       =   &H00BFD7DF&
      Caption         =   "F"
      Height          =   375
      Index           =   17
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   3975
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton num 
      BackColor       =   &H00BFD7DF&
      Caption         =   "E"
      Height          =   375
      Index           =   16
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   3975
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton num 
      BackColor       =   &H00BFD7DF&
      Caption         =   "D"
      Height          =   375
      Index           =   15
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   3975
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton num 
      BackColor       =   &H00BFD7DF&
      Caption         =   "A"
      Height          =   375
      Index           =   12
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   3975
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton num 
      BackColor       =   &H00BFD7DF&
      Caption         =   "B"
      Height          =   375
      Index           =   13
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   3975
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton num 
      BackColor       =   &H00BFD7DF&
      Caption         =   "C"
      Height          =   375
      Index           =   14
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   3975
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton op 
      BackColor       =   &H00BFD7DF&
      Caption         =   "Pi"
      Height          =   375
      Index           =   35
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   54
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   3495
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton op 
      BackColor       =   &H00BFD7DF&
      Caption         =   "Rnd"
      Height          =   375
      Index           =   34
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   53
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   3975
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton op 
      BackColor       =   &H00BFD7DF&
      Caption         =   "n!"
      Height          =   375
      Index           =   31
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   47
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   3975
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton op 
      BackColor       =   &H00BFD7DF&
      Caption         =   "Int"
      Height          =   375
      Index           =   30
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   46
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   3975
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton op 
      BackColor       =   &H00BFD7DF&
      Caption         =   "Fix"
      Height          =   375
      Index           =   29
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   45
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   3975
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton op 
      BackColor       =   &H00E0E0E0&
      Caption         =   "MC"
      Height          =   375
      Index           =   27
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   44
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   2055
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton op 
      BackColor       =   &H00E0E0E0&
      Caption         =   "MR"
      Height          =   375
      Index           =   26
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   43
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   2535
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton op 
      BackColor       =   &H00E0E0E0&
      Caption         =   "MS"
      Height          =   375
      Index           =   25
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   42
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   3015
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton op 
      BackColor       =   &H00E0E0E0&
      Caption         =   "M+"
      Height          =   375
      Index           =   24
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   41
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   3495
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton op 
      BackColor       =   &H00BFD7DF&
      Caption         =   "Rsh"
      Height          =   375
      Index           =   23
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   3495
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton op 
      BackColor       =   &H00BFD7DF&
      Caption         =   "Imp"
      Height          =   375
      Index           =   22
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   3015
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton op 
      BackColor       =   &H00BFD7DF&
      Caption         =   "Eqv"
      Height          =   375
      Index           =   21
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   2535
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton op 
      BackColor       =   &H00BFD7DF&
      Caption         =   "Xor"
      Height          =   375
      Index           =   20
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   2055
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton op 
      BackColor       =   &H00BFD7DF&
      Caption         =   "Lsh"
      Height          =   375
      Index           =   19
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   3495
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton op 
      BackColor       =   &H00BFD7DF&
      Caption         =   "Not"
      Height          =   375
      Index           =   18
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   35
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   3015
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton op 
      BackColor       =   &H00BFD7DF&
      Caption         =   "Or"
      Height          =   375
      Index           =   17
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   2535
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton op 
      BackColor       =   &H00BFD7DF&
      Caption         =   "And"
      Height          =   375
      Index           =   16
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   2055
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton op 
      BackColor       =   &H00BFD7DF&
      Caption         =   "1/x"
      Height          =   375
      Index           =   15
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   32
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   3015
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton op 
      BackColor       =   &H00BFD7DF&
      Caption         =   "x^y"
      Height          =   375
      Index           =   13
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   2535
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton op 
      BackColor       =   &H00BFD7DF&
      Caption         =   "x^2"
      Height          =   375
      Index           =   12
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   30
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   2055
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton op 
      BackColor       =   &H00BFD7DF&
      Caption         =   "Log"
      Height          =   375
      Index           =   11
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   29
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   3495
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton op 
      BackColor       =   &H00BFD7DF&
      Caption         =   "Tan"
      Height          =   375
      Index           =   10
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   28
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   3015
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton op 
      BackColor       =   &H00BFD7DF&
      Caption         =   "Cos"
      Height          =   375
      Index           =   9
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   2535
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton op 
      BackColor       =   &H00BFD7DF&
      Caption         =   "Sin"
      Height          =   375
      Index           =   8
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   26
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   2055
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton op 
      BackColor       =   &H00E0E0E0&
      Caption         =   "="
      Height          =   375
      Index           =   7
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3495
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton op 
      BackColor       =   &H00E0E0E0&
      Caption         =   "%"
      Height          =   375
      Index           =   6
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   24
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   3015
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton op 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mod"
      Height          =   375
      Index           =   5
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2535
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton op 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Sqrt"
      Height          =   375
      Index           =   4
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   2055
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton op 
      BackColor       =   &H00E0E0E0&
      Caption         =   "-"
      Height          =   375
      Index           =   3
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3015
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton op 
      BackColor       =   &H00E0E0E0&
      Caption         =   "/"
      Height          =   375
      Index           =   2
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2055
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton op 
      BackColor       =   &H00E0E0E0&
      Caption         =   "+"
      Height          =   375
      Index           =   1
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3495
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton op 
      BackColor       =   &H00E0E0E0&
      Caption         =   "*"
      Height          =   375
      Index           =   0
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2535
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton num 
      BackColor       =   &H00E0E0E0&
      Caption         =   "."
      Height          =   375
      Index           =   11
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3495
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton num 
      BackColor       =   &H00E0E0E0&
      Caption         =   "+/-"
      Height          =   375
      Index           =   10
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3495
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton num 
      BackColor       =   &H00E0E0E0&
      Caption         =   "9"
      Height          =   375
      Index           =   9
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2055
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton num 
      BackColor       =   &H00E0E0E0&
      Caption         =   "8"
      Height          =   375
      Index           =   8
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2055
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton num 
      BackColor       =   &H00E0E0E0&
      Caption         =   "7"
      Height          =   375
      Index           =   7
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2055
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton num 
      BackColor       =   &H00E0E0E0&
      Caption         =   "6"
      Height          =   375
      Index           =   6
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2535
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton num 
      BackColor       =   &H00E0E0E0&
      Caption         =   "5"
      Height          =   375
      Index           =   5
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2535
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton num 
      BackColor       =   &H00E0E0E0&
      Caption         =   "4"
      Height          =   375
      Index           =   4
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2535
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton num 
      BackColor       =   &H00E0E0E0&
      Caption         =   "3"
      Height          =   375
      Index           =   3
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3015
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton num 
      BackColor       =   &H00E0E0E0&
      Caption         =   "2"
      Height          =   375
      Index           =   2
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3015
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton num 
      BackColor       =   &H00E0E0E0&
      Caption         =   "1"
      Height          =   375
      Index           =   1
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3015
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton num 
      BackColor       =   &H00E0E0E0&
      Caption         =   "0"
      Height          =   375
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3495
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Remv 
      BackColor       =   &H00BFD7DF&
      Caption         =   "C"
      Height          =   375
      Index           =   2
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1575
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Remv 
      BackColor       =   &H00BFD7DF&
      Caption         =   "CE"
      Height          =   375
      Index           =   1
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1575
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Remv 
      BackColor       =   &H00BFD7DF&
      Caption         =   "Backspace"
      Height          =   375
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1575
      Visible         =   0   'False
      Width           =   2295
   End
End
Attribute VB_Name = "calc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------
'-             VB Programming Project        -
'-                                           -
'-  Project Name  : Calculator               -
'-  Coded By      : Mohammad Shams Javi      -
'-                                           -
'-  Copyright (c) 1384/8/9                   -
'---------------------------------------------

Private Sub Form_Activate()
Call F_act
End Sub

Private Sub Form_DblClick()
'About Me
Call Ab_me
End Sub

Private Sub Form_KeyDown(K As Integer, Shift As Integer)
'Its better that , you read Module and then try to understand this code. ;)
'8=BS  46=Del  48=0  65=A  96=Pad(0)  106=Operators  110='.'

Select Case K
 Case 8: PM Remv(0).hwnd, &H100, 32, 0: Remv_Click (0)
 Case 46: PM Remv(1).hwnd, &H100, 32, 0: Remv_Click (1)
 Case 48 To 57: PM num(K - 48).hwnd, &H100, 32, 0: num_Click (K - 48)
 Case 65 To 70: PM num(K - 53).hwnd, &H100, 32, 0: num_Click (K - 53)
 Case 96 To 105: PM num(K - 96).hwnd, &H100, 32, 0: num_Click (K - 96)
 Case 106 To 109: PM op(K - 106).hwnd, &H100, 32, 0: op_Click (K - 106)
 Case 110: PM num(11).hwnd, &H100, 32, 0: num_Click (11)
 'Case 13: PM op(7).hwnd, &H100, 32, 0: op_Click (7): K = 0
 'Case 16:PM op(7).hwnd, &H100, 32, 0: op_Click (7): K = 0
     '  ^-- You can use (16)SHIFT key for Equal sign "="
End Select

'Note By Author : [Mohammad.Sh.J]
'If you want to handle equal '=' sign with keyboard, you
'   must create new class of buttons that can handle key
'   preview of ENTER button.
'It means that Button not pushed when have focus and
'   Enter or Space are pressed.
'I solved this problem in my projects by creating  a new
'    class of command buttons that havnt focus , but not
'    in this educational project.  Have Fun :)

End Sub

Private Sub Form_Load()
AW hwnd, 100, &H20010
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Form_DblClick
F_unl
End Sub

Private Sub num_Click(Index As Integer)

'This lines do exit sub , if entered number 'n' is
'   illegal or above than maximum in base
' EX : if Base=8 (Oct) then  'n' must be less than 8
'-----------------Range--Check------------------------
For Each c In Optn  'get enabled option index
    If c.Value = True Then idx = c.Index
Next
Select Case idx           'index 10="+/-"  index 11="."
    Case 0: If Index > 11 Then Exit Sub
    Case 2: If Index = 11 Then Exit Sub
    Case 1: If Index > 7 And Index <> 10 Then Exit Sub
    Case 3: If Index > 1 And Index <> 10 Then Exit Sub
End Select
'-----------------------------------------------------

If opr <> "" And ldo Then
    t(0) = ""      'clear display to show new number
    t(1) = "": t(2) = "": t(3) = ""
    lnm = nm    'set last number
End If

Select Case Index
    Case 0 To 9: disp Index, 2     '0..9
    Case 10 To 11: disp Index, 4   '. + -
    Case 12 To 17: disp Index, 2   'A..F
End Select
ldo = False
End Sub

Private Sub op_Click(Index As Integer)
Dim tmp As Double
'For handle 0 number and semi operators
If nm = 0 And op(Index).Tag <> "1" Then Exit Sub

If opr <> "=" Then lopr = opr      'set last operator
'it is a trick, set operator name :)  ... And only
'   works on none semi operators
If op(Index).Tag = "" Then opr = op(Index).Caption

If opr = "=" And lopr <> "" And lopr <> "=" And op(Index).Tag <> "1" Then
    If lnm <> 0 Then   'Normal operator proccessing , '=' used
        tmp = nm
        proc (lopr)
        lnm = nm
        nm = tmp
    Else               'Semi USING and loop equal
        If lopr = "+" Or lopr = "-" Then
            tmp = nm
            proc (lopr)     'in +,- we can do a loop
            lnm = nm        '  with nm=X , lnm=0
            nm = tmp        '  (y=nm+lnm)<>0
        Else                'But in *,/ we cant use
            tmp = nm        '  0 in operands
            lnm = nm  '<<----- make lnm<>0
            proc (lopr)     '  nm,lnm must be <>0
            lnm = nm
            nm = tmp
        End If
    End If
ElseIf op(Index).Tag = "1" Then
    'With this line , semi operators can proccess last operation
    'If lopr <> "" Then proc (lopr)
    'Beacause of semi operator
    proc (op(Index).Caption)
    
'Do last operator. used when operator pressed together without using "="
ElseIf lopr <> "" And Not ldo Then proc (lopr)
End If

ldo = True
End Sub

Private Sub Optn_Click(Index As Integer)
For Each c In t
    c.BackColor = &HC0C0C0
Next
t(Index).BackColor = &H80000005 'set active display
End Sub

Private Sub Remv_Click(Index As Integer)
'Subroutine for BS,C,CE

Select Case Index
    Case 0: disp 0, 3   'Backspace
    Case 1:
        lnm = 0: opr = "": lopr = "": nm = 0
        mem = 0: ldo = False
        disp 0, 0
    Case 2: disp 0, 0
End Select
End Sub
