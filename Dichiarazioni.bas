Attribute VB_Name = "Dichiarazioni"
'Permette di spostare il controllo all'interno di un form
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'Dichiarazioni per il ridimensionamento dinanico a run-time
Public Const HTLEFT = 10
Public Const HTRIGHT = 11
Public Const HTTOP = 12
Public Const HTBOTTOM = 15

'Variabili per la colorazione dei form
Public Assex As Integer, Assey As Integer, Lungh As Integer, Angol As Integer
Public X As Integer, Y As Integer
Public C1 As Integer, C2 As Integer

'Variabili per iconizzare il controllo per UseAsForm
Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long

'API Decalaration for FlashWindow
Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, _
   ByVal bInvert As Long) As Long
   
'Declarations for ExplodeForm
Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
   lpRect As RECT) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
   ByVal hdc As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
   ByVal hObject As Long) As Long  'note error in declare
Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, _
   ByVal x1 As Long, ByVal y1 As Long, _
   ByVal x2 As Long, ByVal y2 As Long) As Long
  
