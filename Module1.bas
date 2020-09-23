Attribute VB_Name = "vkbmod"
Option Explicit
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer
Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer
Private Keystate As Long
Private KeyPressState As Long
Private A As Boolean
Private B As Boolean
Private C As Boolean
Private D As Boolean
Private E As Boolean
Private F As Boolean
Private G As Boolean
Private H As Boolean
Private I As Boolean
Private J As Boolean
Private K As Boolean
Private L As Boolean
Private M As Boolean
Private N As Boolean
Private O As Boolean
Private P As Boolean
Private Q As Boolean
Private R As Boolean
Private S As Boolean
Private T As Boolean
Private U As Boolean
Private V As Boolean
Private W As Boolean
Private X As Boolean
Private Y As Boolean
Private Z As Boolean
Private ESCAPE As Boolean
Private F1 As Boolean
Private F2 As Boolean
Private F3 As Boolean
Private F4 As Boolean
Private F5 As Boolean
Private F6 As Boolean
Private F7 As Boolean
Private F8 As Boolean
Private F9 As Boolean
Private F10 As Boolean
Private F11 As Boolean
Private F12 As Boolean
Private PRINTSCREEN As Boolean
Private SCROLL As Boolean
Private PAUSE As Boolean
Private TABKEY As Boolean
Private CAPS As Boolean
Private SHIFT As Boolean
Private CONTROL As Boolean
Private MENU As Boolean
Private OPT As Boolean
Private ALT As Boolean
Private SPACE As Boolean
Private ENTER As Boolean
Private BACKSPACE As Boolean
Private INSERT As Boolean
Private HOME As Boolean
Private PAGEUP As Boolean
Private PAGEDOWN As Boolean
Private DELETE As Boolean
Private ENDKEY As Boolean
Private UP As Boolean
Private DOWN As Boolean
Private LEFT As Boolean
Private RIGHT As Boolean
Private NUMLOCK As Boolean
Private NUM0 As Boolean
Private NUM1 As Boolean
Private NUM2 As Boolean
Private NUM3 As Boolean
Private NUM4 As Boolean
Private NUM5 As Boolean
Private NUM6 As Boolean
Private NUM7 As Boolean
Private NUM8 As Boolean
Private NUM9 As Boolean
Private NUMDIV As Boolean
Private NUMMULT As Boolean
Private NUMSUB As Boolean
Private NUMADD As Boolean
Private NUMDEC As Boolean
Private TIDLE As Boolean
Private B1 As Boolean
Private B2 As Boolean
Private B3 As Boolean
Private B4 As Boolean
Private B5 As Boolean
Private B6 As Boolean
Private B7 As Boolean
Private B8 As Boolean
Private B9 As Boolean
Private B0 As Boolean
Private BSUB As Boolean
Private BEQ As Boolean
Private BRBRA As Boolean
Private BLBRA As Boolean
Private BQUOTES As Boolean
Private BCOLIN As Boolean
Private BQUEST As Boolean
Private BPER As Boolean
Private BCOM As Boolean
Private BDIV As Boolean

Public Const VK_LBUTTON = &H1
Public Const VK_RBUTTON = &H2
Public Const VK_CANCEL = &H3
Public Const VK_MBUTTON = &H4
Public Const VK_BACK = &H8
Public Const VK_TAB = &H9
Public Const VK_CLEAR = &HC
Public Const VK_RETURN = &HD
Public Const VK_SHIFT = &H10
Public Const VK_CONTROL = &H11
Public Const VK_MENU = &H12
Public Const VK_PAUSE = &H13
Public Const VK_CAPITAL = &H14
Public Const VK_ESCAPE = &H1B
Public Const VK_SPACE = &H20
Public Const VK_PRIOR = &H21
Public Const VK_NEXT = &H22
Public Const VK_END = &H23
Public Const VK_HOME = &H24
Public Const VK_LEFT = &H25
Public Const VK_UP = &H26
Public Const VK_RIGHT = &H27
Public Const VK_DOWN = &H28
Public Const VK_SELECT = &H29
Public Const VK_PRINT = &H2A
Public Const VK_EXECUTE = &H2B
Public Const VK_SNAPSHOT = &H2C
Public Const VK_INSERT = &H2D
Public Const VK_DELETE = &H2E
Public Const VK_HELP = &H2F
Public Const VK_0 = &H30
Public Const VK_1 = &H31
Public Const VK_2 = &H32
Public Const VK_3 = &H33
Public Const VK_4 = &H34
Public Const VK_5 = &H35
Public Const VK_6 = &H36
Public Const VK_7 = &H37
Public Const VK_8 = &H38
Public Const VK_9 = &H39
Public Const VK_A = &H41
Public Const VK_B = &H42
Public Const VK_C = &H43
Public Const VK_D = &H44
Public Const VK_E = &H45
Public Const VK_F = &H46
Public Const VK_G = &H47
Public Const VK_H = &H48
Public Const VK_I = &H49
Public Const VK_J = &H4A
Public Const VK_K = &H4B
Public Const VK_L = &H4C
Public Const VK_M = &H4D
Public Const VK_N = &H4E
Public Const VK_O = &H4F
Public Const VK_P = &H50
Public Const VK_Q = &H51
Public Const VK_R = &H52
Public Const VK_S = &H53
Public Const VK_T = &H54
Public Const VK_U = &H55
Public Const VK_V = &H56
Public Const VK_W = &H57
Public Const VK_X = &H58
Public Const VK_Y = &H59
Public Const VK_Z = &H5A
Public Const VK_STARTKEY = &H5B
Public Const VK_CONTEXTKEY = &H5D
Public Const VK_NUMPAD0 = &H60
Public Const VK_NUMPAD1 = &H61
Public Const VK_NUMPAD2 = &H62
Public Const VK_NUMPAD3 = &H63
Public Const VK_NUMPAD4 = &H64
Public Const VK_NUMPAD5 = &H65
Public Const VK_NUMPAD6 = &H66
Public Const VK_NUMPAD7 = &H67
Public Const VK_NUMPAD8 = &H68
Public Const VK_NUMPAD9 = &H69
Public Const VK_MULTIPLY = &H6A
Public Const VK_ADD = &H6B
Public Const VK_SEPARATOR = &H6C
Public Const VK_SUBTRACT = &H6D
Public Const VK_DECIMAL = &H6E
Public Const VK_DIVIDE = &H6F
Public Const VK_F1 = &H70
Public Const VK_F2 = &H71
Public Const VK_F3 = &H72
Public Const VK_F4 = &H73
Public Const VK_F5 = &H74
Public Const VK_F6 = &H75
Public Const VK_F7 = &H76
Public Const VK_F8 = &H77
Public Const VK_F9 = &H78
Public Const VK_F10 = &H79
Public Const VK_F11 = &H7A
Public Const VK_F12 = &H7B
Public Const VK_F13 = &H7C
Public Const VK_F14 = &H7D
Public Const VK_F15 = &H7E
Public Const VK_F16 = &H7F
Public Const VK_F17 = &H80
Public Const VK_F18 = &H81
Public Const VK_F19 = &H82
Public Const VK_F20 = &H83
Public Const VK_F21 = &H84
Public Const VK_F22 = &H85
Public Const VK_F23 = &H86
Public Const VK_F24 = &H87
Public Const VK_NUMLOCK = &H90
Public Const VK_OEM_SCROLL = &H91
Public Const VK_OEM_1 = &HBA
Public Const VK_OEM_PLUS = &HBB
Public Const VK_OEM_COMMA = &HBC
Public Const VK_OEM_MINUS = &HBD
Public Const VK_OEM_PERIOD = &HBE
Public Const VK_OEM_2 = &HBF
Public Const VK_OEM_3 = &HC0
Public Const VK_OEM_4 = &HDB
Public Const VK_OEM_5 = &HDC
Public Const VK_OEM_6 = &HDD
Public Const VK_OEM_7 = &HDE
Public Const VK_OEM_8 = &HDF
Public Const VK_ICO_F17 = &HE0
Public Const VK_ICO_F18 = &HE1
Public Const VK_OEM102 = &HE2
Public Const VK_ICO_HELP = &HE3
Public Const VK_ICO_00 = &HE4
Public Const VK_ICO_CLEAR = &HE6
Public Const VK_OEM_RESET = &HE9
Public Const VK_OEM_JUMP = &HEA
Public Const VK_OEM_PA1 = &HEB
Public Const VK_OEM_PA2 = &HEC
Public Const VK_OEM_PA3 = &HED
Public Const VK_OEM_WSCTRL = &HEE
Public Const VK_OEM_CUSEL = &HEF
Public Const VK_OEM_ATTN = &HF0
Public Const VK_OEM_FINNISH = &HF1
Public Const VK_OEM_COPY = &HF2
Public Const VK_OEM_AUTO = &HF3
Public Const VK_OEM_ENLW = &HF4
Public Const VK_OEM_BACKTAB = &HF5
Public Const VK_ATTN = &HF6
Public Const VK_CRSEL = &HF7
Public Const VK_EXSEL = &HF8
Public Const VK_EREOF = &HF9
Public Const VK_PLAY = &HFA
Public Const VK_ZOOM = &HFB
Public Const VK_NONAME = &HFC
Public Const VK_PA1 = &HFD
Public Const VK_OEM_CLEAR = &HFE
Public Const KEYEVENTF_EXTENDEDKEY = &H1
Public Const KEYEVENTF_KEYUP = &H2

Public Function CheckKeys()
Keystate = GetKeyState(VK_A)
If (Keystate And &H80) = &H80 Then vkb.A.BackColor = &HFFFF& Else vkb.A.BackColor = &H8000000A
Keystate = GetKeyState(VK_B)
If (Keystate And &H80) = &H80 Then vkb.B.BackColor = &HFFFF& Else vkb.B.BackColor = &H8000000A
Keystate = GetKeyState(VK_C)
If (Keystate And &H80) = &H80 Then vkb.C.BackColor = &HFFFF& Else vkb.C.BackColor = &H8000000A
Keystate = GetKeyState(VK_D)
If (Keystate And &H80) = &H80 Then vkb.D.BackColor = &HFFFF& Else vkb.D.BackColor = &H8000000A
Keystate = GetKeyState(VK_E)
If (Keystate And &H80) = &H80 Then vkb.E.BackColor = &HFFFF& Else vkb.E.BackColor = &H8000000A
Keystate = GetKeyState(VK_F)
If (Keystate And &H80) = &H80 Then vkb.F.BackColor = &HFFFF& Else vkb.F.BackColor = &H8000000A
Keystate = GetKeyState(VK_G)
If (Keystate And &H80) = &H80 Then vkb.G.BackColor = &HFFFF& Else vkb.G.BackColor = &H8000000A
Keystate = GetKeyState(VK_H)
If (Keystate And &H80) = &H80 Then vkb.H.BackColor = &HFFFF& Else vkb.H.BackColor = &H8000000A
Keystate = GetKeyState(VK_I)
If (Keystate And &H80) = &H80 Then vkb.I.BackColor = &HFFFF& Else vkb.I.BackColor = &H8000000A
Keystate = GetKeyState(VK_J)
If (Keystate And &H80) = &H80 Then vkb.J.BackColor = &HFFFF& Else vkb.J.BackColor = &H8000000A
Keystate = GetKeyState(VK_K)
If (Keystate And &H80) = &H80 Then vkb.K.BackColor = &HFFFF& Else vkb.K.BackColor = &H8000000A
Keystate = GetKeyState(VK_L)
If (Keystate And &H80) = &H80 Then vkb.L.BackColor = &HFFFF& Else vkb.L.BackColor = &H8000000A
Keystate = GetKeyState(VK_M)
If (Keystate And &H80) = &H80 Then vkb.M.BackColor = &HFFFF& Else vkb.M.BackColor = &H8000000A
Keystate = GetKeyState(VK_N)
If (Keystate And &H80) = &H80 Then vkb.N.BackColor = &HFFFF& Else vkb.N.BackColor = &H8000000A
Keystate = GetKeyState(VK_O)
If (Keystate And &H80) = &H80 Then vkb.O.BackColor = &HFFFF& Else vkb.O.BackColor = &H8000000A
Keystate = GetKeyState(VK_P)
If (Keystate And &H80) = &H80 Then vkb.P.BackColor = &HFFFF& Else vkb.P.BackColor = &H8000000A
Keystate = GetKeyState(VK_Q)
If (Keystate And &H80) = &H80 Then vkb.Q.BackColor = &HFFFF& Else vkb.Q.BackColor = &H8000000A
Keystate = GetKeyState(VK_R)
If (Keystate And &H80) = &H80 Then vkb.R.BackColor = &HFFFF& Else vkb.R.BackColor = &H8000000A
Keystate = GetKeyState(VK_S)
If (Keystate And &H80) = &H80 Then vkb.S.BackColor = &HFFFF& Else vkb.S.BackColor = &H8000000A
Keystate = GetKeyState(VK_T)
If (Keystate And &H80) = &H80 Then vkb.T.BackColor = &HFFFF& Else vkb.T.BackColor = &H8000000A
Keystate = GetKeyState(VK_U)
If (Keystate And &H80) = &H80 Then vkb.U.BackColor = &HFFFF& Else vkb.U.BackColor = &H8000000A
Keystate = GetKeyState(VK_V)
If (Keystate And &H80) = &H80 Then vkb.V.BackColor = &HFFFF& Else vkb.V.BackColor = &H8000000A
Keystate = GetKeyState(VK_W)
If (Keystate And &H80) = &H80 Then vkb.W.BackColor = &HFFFF& Else vkb.W.BackColor = &H8000000A
Keystate = GetKeyState(VK_X)
If (Keystate And &H80) = &H80 Then vkb.X.BackColor = &HFFFF& Else vkb.X.BackColor = &H8000000A
Keystate = GetKeyState(VK_Y)
If (Keystate And &H80) = &H80 Then vkb.Y.BackColor = &HFFFF& Else vkb.Y.BackColor = &H8000000A
Keystate = GetKeyState(VK_Z)
If (Keystate And &H80) = &H80 Then vkb.Z.BackColor = &HFFFF& Else vkb.Z.BackColor = &H8000000A
Keystate = GetKeyState(VK_1)
If (Keystate And &H80) = &H80 Then vkb.N1.BackColor = &HFFFF& Else vkb.N1.BackColor = &H8000000A
Keystate = GetKeyState(VK_2)
If (Keystate And &H80) = &H80 Then vkb.N2.BackColor = &HFFFF& Else vkb.N2.BackColor = &H8000000A
Keystate = GetKeyState(VK_3)
If (Keystate And &H80) = &H80 Then vkb.N3.BackColor = &HFFFF& Else vkb.N3.BackColor = &H8000000A
Keystate = GetKeyState(VK_4)
If (Keystate And &H80) = &H80 Then vkb.N4.BackColor = &HFFFF& Else vkb.N4.BackColor = &H8000000A
Keystate = GetKeyState(VK_5)
If (Keystate And &H80) = &H80 Then vkb.N5.BackColor = &HFFFF& Else vkb.N5.BackColor = &H8000000A
Keystate = GetKeyState(VK_6)
If (Keystate And &H80) = &H80 Then vkb.N6.BackColor = &HFFFF& Else vkb.N6.BackColor = &H8000000A
Keystate = GetKeyState(VK_7)
If (Keystate And &H80) = &H80 Then vkb.N7.BackColor = &HFFFF& Else vkb.N7.BackColor = &H8000000A
Keystate = GetKeyState(VK_8)
If (Keystate And &H80) = &H80 Then vkb.N8.BackColor = &HFFFF& Else vkb.N8.BackColor = &H8000000A
Keystate = GetKeyState(VK_9)
If (Keystate And &H80) = &H80 Then vkb.N9.BackColor = &HFFFF& Else vkb.N9.BackColor = &H8000000A
Keystate = GetKeyState(VK_0)
If (Keystate And &H80) = &H80 Then vkb.N0.BackColor = &HFFFF& Else vkb.N0.BackColor = &H8000000A
Keystate = GetKeyState(VK_BACK)
If (Keystate And &H80) = &H80 Then vkb.Back.BackColor = &HFFFF& Else vkb.Back.BackColor = &H8000000A
Keystate = GetKeyState(VK_OEM_3)
If (Keystate And &H80) = &H80 Then vkb.OEM_3.BackColor = &HFFFF& Else vkb.OEM_3.BackColor = &H8000000A
Keystate = GetKeyState(VK_OEM_MINUS)
If (Keystate And &H80) = &H80 Then vkb.OEM_MINUS.BackColor = &HFFFF& Else vkb.OEM_MINUS.BackColor = &H8000000A
Keystate = GetKeyState(VK_OEM_PLUS)
If (Keystate And &H80) = &H80 Then vkb.OEM_PLUS.BackColor = &HFFFF& Else vkb.OEM_PLUS.BackColor = &H8000000A
Keystate = GetKeyState(VK_OEM_5)
If (Keystate And &H80) = &H80 Then vkb.OEM_5.BackColor = &HFFFF& Else vkb.OEM_5.BackColor = &H8000000A
Keystate = GetKeyState(VK_TAB)
If (Keystate And &H80) = &H80 Then vkb.TAB.BackColor = &HFFFF& Else vkb.TAB.BackColor = &H8000000A
Keystate = GetKeyState(VK_RETURN)
If (Keystate And &H80) = &H80 Then
vkb.ENTER.BackColor = &HFFFF&
vkb.ENTER2.BackColor = &HFFFF&
Else
vkb.ENTER.BackColor = &H8000000A
vkb.ENTER2.BackColor = &H8000000A
End If
Keystate = GetKeyState(VK_OEM_4)
If (Keystate And &H80) = &H80 Then vkb.OEM_4.BackColor = &HFFFF& Else vkb.OEM_4.BackColor = &H8000000A
Keystate = GetKeyState(VK_OEM_6)
If (Keystate And &H80) = &H80 Then vkb.OEM_6.BackColor = &HFFFF& Else vkb.OEM_6.BackColor = &H8000000A
Keystate = GetKeyState(VK_OEM_1)
If (Keystate And &H80) = &H80 Then vkb.OEM_1.BackColor = &HFFFF& Else vkb.OEM_1.BackColor = &H8000000A
Keystate = GetKeyState(VK_OEM_7)
If (Keystate And &H80) = &H80 Then vkb.OEM_7.BackColor = &HFFFF& Else vkb.OEM_7.BackColor = &H8000000A
Keystate = GetKeyState(VK_OEM_COMMA)
If (Keystate And &H80) = &H80 Then vkb.OEM_COMMA.BackColor = &HFFFF& Else vkb.OEM_COMMA.BackColor = &H8000000A
Keystate = GetKeyState(VK_OEM_PERIOD)
If (Keystate And &H80) = &H80 Then vkb.OEM_PERIOD.BackColor = &HFFFF& Else vkb.OEM_PERIOD.BackColor = &H8000000A
Keystate = GetKeyState(VK_OEM_2)
If (Keystate And &H80) = &H80 Then vkb.OEM_2.BackColor = &HFFFF& Else vkb.OEM_2.BackColor = &H8000000A
Keystate = GetKeyState(VK_OEM_2)
If (Keystate And &H80) = &H80 Then vkb.OEM_2.BackColor = &HFFFF& Else vkb.OEM_2.BackColor = &H8000000A
Keystate = GetKeyState(VK_SHIFT)
If (Keystate And &H80) = &H80 Then vkb.SHIFT1.BackColor = &HFFFF& Else vkb.SHIFT1.BackColor = &H8000000A
Keystate = GetKeyState(VK_SHIFT)
If (Keystate And &H80) = &H80 Then vkb.SHIFT2.BackColor = &HFFFF& Else vkb.SHIFT2.BackColor = &H8000000A
Keystate = GetKeyState(VK_CAPITAL)
If (Keystate And &H80) = &H80 Then vkb.CAPS.BackColor = &HFFFF& Else vkb.CAPS.BackColor = &H8000000A
Keystate = GetKeyState(VK_CONTROL)
If (Keystate And &H80) = &H80 Then vkb.CONTROL1.BackColor = &HFFFF& Else vkb.CONTROL1.BackColor = &H8000000A
Keystate = GetKeyState(VK_CONTROL)
If (Keystate And &H80) = &H80 Then vkb.CONTROL2.BackColor = &HFFFF& Else vkb.CONTROL2.BackColor = &H8000000A
Keystate = GetKeyState(VK_STARTKEY)
If (Keystate And &H80) = &H80 Then
vkb.START1.BackColor = &HFFFF&
vkb.START1.Picture = LoadPicture("MNU2.bmp")
Else
vkb.START1.BackColor = &H8000000A
vkb.START1.Picture = LoadPicture("MNU1.bmp")
End If
If (Keystate And &H80) = &H80 Then
vkb.START2.BackColor = &HFFFF&
vkb.START2.Picture = LoadPicture("MNU2.bmp")
Else
vkb.START2.BackColor = &H8000000A
vkb.START2.Picture = LoadPicture("MNU1.bmp")
End If
Keystate = GetKeyState(VK_CONTEXTKEY)
If (Keystate And &H80) = &H80 Then
vkb.OPT.BackColor = &HFFFF&
vkb.OPT.Picture = LoadPicture("opt2.bmp")
Else
vkb.OPT.BackColor = &H8000000A
vkb.OPT.Picture = LoadPicture("opt1.bmp")
End If
Keystate = GetKeyState(VK_MENU)
If (Keystate And &H80) = &H80 Then vkb.ALT1.BackColor = &HFFFF& Else vkb.ALT1.BackColor = &H8000000A
Keystate = GetKeyState(VK_MENU)
If (Keystate And &H80) = &H80 Then vkb.ALT2.BackColor = &HFFFF& Else vkb.ALT2.BackColor = &H8000000A
Keystate = GetKeyState(VK_SPACE)
If (Keystate And &H80) = &H80 Then vkb.SPACEBAR.BackColor = &HFFFF& Else vkb.SPACEBAR.BackColor = &H8000000A
Keystate = GetKeyState(VK_F1)
If (Keystate And &H80) = &H80 Then vkb.F1.BackColor = &HFFFF& Else vkb.F1.BackColor = &H8000000A
Keystate = GetKeyState(VK_F2)
If (Keystate And &H80) = &H80 Then vkb.F2.BackColor = &HFFFF& Else vkb.F2.BackColor = &H8000000A
Keystate = GetKeyState(VK_F3)
If (Keystate And &H80) = &H80 Then vkb.F3.BackColor = &HFFFF& Else vkb.F3.BackColor = &H8000000A
Keystate = GetKeyState(VK_F4)
If (Keystate And &H80) = &H80 Then vkb.F4.BackColor = &HFFFF& Else vkb.F4.BackColor = &H8000000A
Keystate = GetKeyState(VK_F5)
If (Keystate And &H80) = &H80 Then vkb.F5.BackColor = &HFFFF& Else vkb.F5.BackColor = &H8000000A
Keystate = GetKeyState(VK_F6)
If (Keystate And &H80) = &H80 Then vkb.F6.BackColor = &HFFFF& Else vkb.F6.BackColor = &H8000000A
Keystate = GetKeyState(VK_F7)
If (Keystate And &H80) = &H80 Then vkb.F7.BackColor = &HFFFF& Else vkb.F7.BackColor = &H8000000A
Keystate = GetKeyState(VK_F8)
If (Keystate And &H80) = &H80 Then vkb.F8.BackColor = &HFFFF& Else vkb.F8.BackColor = &H8000000A
Keystate = GetKeyState(VK_F9)
If (Keystate And &H80) = &H80 Then vkb.F9.BackColor = &HFFFF& Else vkb.F9.BackColor = &H8000000A
Keystate = GetKeyState(VK_F10)
If (Keystate And &H80) = &H80 Then vkb.F10.BackColor = &HFFFF& Else vkb.F10.BackColor = &H8000000A
Keystate = GetKeyState(VK_F11)
If (Keystate And &H80) = &H80 Then vkb.F11.BackColor = &HFFFF& Else vkb.F11.BackColor = &H8000000A
Keystate = GetKeyState(VK_F12)
If (Keystate And &H80) = &H80 Then vkb.F12.BackColor = &HFFFF& Else vkb.F12.BackColor = &H8000000A
Keystate = GetKeyState(VK_ESCAPE)
If (Keystate And &H80) = &H80 Then vkb.ESCAPE.BackColor = &HFFFF& Else vkb.ESCAPE.BackColor = &H8000000A
Keystate = GetKeyState(VK_SNAPSHOT)
If (Keystate And &H80) = &H80 Then vkb.PRINTSCR.BackColor = &HFFFF& Else vkb.PRINTSCR.BackColor = &H8000000A
Keystate = GetKeyState(VK_OEM_SCROLL)
If (Keystate And &H80) = &H80 Then vkb.SCROLL.BackColor = &HFFFF& Else vkb.SCROLL.BackColor = &H8000000A
Keystate = GetKeyState(VK_PAUSE)
If (Keystate And &H80) = &H80 Then vkb.PAUSEKEY.BackColor = &HFFFF& Else vkb.PAUSEKEY.BackColor = &H8000000A
Keystate = GetKeyState(VK_INSERT)
If (Keystate And &H80) = &H80 Then vkb.INSERT.BackColor = &HFFFF& Else vkb.INSERT.BackColor = &H8000000A
Keystate = GetKeyState(VK_HOME)
If (Keystate And &H80) = &H80 Then vkb.HOME.BackColor = &HFFFF& Else vkb.HOME.BackColor = &H8000000A
Keystate = GetKeyState(VK_DELETE)
If (Keystate And &H80) = &H80 Then vkb.DEL.BackColor = &HFFFF& Else vkb.DEL.BackColor = &H8000000A
Keystate = GetKeyState(VK_END)
If (Keystate And &H80) = &H80 Then vkb.END.BackColor = &HFFFF& Else vkb.END.BackColor = &H8000000A
Keystate = GetKeyState(VK_PRIOR)
If (Keystate And &H80) = &H80 Then vkb.PGUP.BackColor = &HFFFF& Else vkb.PGUP.BackColor = &H8000000A
Keystate = GetKeyState(VK_NEXT)
If (Keystate And &H80) = &H80 Then vkb.PGDOWN.BackColor = &HFFFF& Else vkb.PGDOWN.BackColor = &H8000000A


Keystate = GetKeyState(VK_NUMPAD0)
If (Keystate And &H80) = &H80 Then vkb.NUMPAD0.BackColor = &HFFFF& Else vkb.NUMPAD0.BackColor = &H8000000A
Keystate = GetKeyState(VK_NUMPAD1)
If (Keystate And &H80) = &H80 Then vkb.NUMPAD1.BackColor = &HFFFF& Else vkb.NUMPAD1.BackColor = &H8000000A
Keystate = GetKeyState(VK_NUMPAD2)
If (Keystate And &H80) = &H80 Then vkb.NUMPAD2.BackColor = &HFFFF& Else vkb.NUMPAD2.BackColor = &H8000000A
Keystate = GetKeyState(VK_NUMPAD3)
If (Keystate And &H80) = &H80 Then vkb.NUMPAD3.BackColor = &HFFFF& Else vkb.NUMPAD3.BackColor = &H8000000A
Keystate = GetKeyState(VK_NUMPAD4)
If (Keystate And &H80) = &H80 Then vkb.NUMPAD4.BackColor = &HFFFF& Else vkb.NUMPAD4.BackColor = &H8000000A
Keystate = GetKeyState(VK_NUMPAD5)
If (Keystate And &H80) = &H80 Then vkb.NUMPAD5.BackColor = &HFFFF& Else vkb.NUMPAD5.BackColor = &H8000000A
Keystate = GetKeyState(VK_NUMPAD6)
If (Keystate And &H80) = &H80 Then vkb.NUMPAD6.BackColor = &HFFFF& Else vkb.NUMPAD6.BackColor = &H8000000A
Keystate = GetKeyState(VK_NUMPAD7)
If (Keystate And &H80) = &H80 Then vkb.NUMPAD7.BackColor = &HFFFF& Else vkb.NUMPAD7.BackColor = &H8000000A
Keystate = GetKeyState(VK_NUMPAD8)
If (Keystate And &H80) = &H80 Then vkb.NUMPAD8.BackColor = &HFFFF& Else vkb.NUMPAD8.BackColor = &H8000000A
Keystate = GetKeyState(VK_NUMPAD9)
If (Keystate And &H80) = &H80 Then vkb.NUMPAD9.BackColor = &HFFFF& Else vkb.NUMPAD9.BackColor = &H8000000A
Keystate = GetKeyState(VK_UP)
If (Keystate And &H80) = &H80 Then vkb.UP.BackColor = &HFFFF& Else vkb.UP.BackColor = &H8000000A
Keystate = GetKeyState(VK_DOWN)
If (Keystate And &H80) = &H80 Then vkb.DOWN.BackColor = &HFFFF& Else vkb.DOWN.BackColor = &H8000000A
Keystate = GetKeyState(VK_LEFT)
If (Keystate And &H80) = &H80 Then vkb.LEFTB.BackColor = &HFFFF& Else vkb.LEFTB.BackColor = &H8000000A
Keystate = GetKeyState(VK_RIGHT)
If (Keystate And &H80) = &H80 Then vkb.RIGHTB.BackColor = &HFFFF& Else vkb.RIGHTB.BackColor = &H8000000A
Keystate = GetKeyState(VK_NUMLOCK)
If (Keystate And &H80) = &H80 Then vkb.NUM.BackColor = &HFFFF& Else vkb.NUM.BackColor = &H8000000A
Keystate = GetKeyState(VK_DIVIDE)
If (Keystate And &H80) = &H80 Then vkb.DIV.BackColor = &HFFFF& Else vkb.DIV.BackColor = &H8000000A
Keystate = GetKeyState(VK_MULTIPLY)
If (Keystate And &H80) = &H80 Then vkb.MULT.BackColor = &HFFFF& Else vkb.MULT.BackColor = &H8000000A
Keystate = GetKeyState(VK_ADD)
If (Keystate And &H80) = &H80 Then vkb.ADDI.BackColor = &HFFFF& Else vkb.ADDI.BackColor = &H8000000A
Keystate = GetKeyState(VK_DECIMAL)
If (Keystate And &H80) = &H80 Then vkb.DEC.BackColor = &HFFFF& Else vkb.DEC.BackColor = &H8000000A
Keystate = GetKeyState(VK_SUBTRACT)
If (Keystate And &H80) = &H80 Then vkb.SUBT.BackColor = &HFFFF& Else vkb.SUBT.BackColor = &H8000000A
End Function

Sub main()
Load splash
End Sub
