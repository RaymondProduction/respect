Attribute VB_Name = "Declare"
'деклараци€ функций и констант Win32API
'    BOOL SetWindowPos(
'
'        HWND hWnd,            // дескриптор окна
'        HWND hWndInsertAfter, // дескриптор пор€дка размещени€
'        int x,                // позици€ по горизонтали
'        int y,                // позици€ по вертикали
'        int cx,               // ширина
'        int cy,               // высота
'        UINT uFlags           // флажки позиционировани€ окна
'
'    );

'[in] ƒескриптор окна, которое предшествует установленному окну в  Z-последовательности.
'Ётот параметр должен быть дескриптор окна или одним из следующих значений:
'HWND_BOTTOM
'ѕомещает окно внизу Z-последовательности.
'≈сли параметр hWnd идентифицирует самое верхнее окно, окно тер€ет свой статус
'самого верхнего и помещаетс€ внизу всех других окон.
'HWND_NOTOPMOST
'ѕомещает окно перед всеми не самыми верхними окнами
'(то есть позади всех самых верхних окон).
'Ётот флажок не имеет никакого вли€ни€, если окно - уже не самое верхнее окно.
'HWND_TOP
'ѕомещает окно наверху Z-последовательности.
'HWND_TOPMOST
'ѕомещает окно перед не самыми верхними окнами.
'ќкно сохран€ет свою самую верхнюю позицию даже тогда, когда оно тер€ет активность.


   Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

   Public Const HWND_TOPMOST = -1
   Public Const HWND_NOTOPMOST = -2
   Public Const SWP_NOMOVE = &H2
   Public Const SWP_NOSIZE = &H1
   Public Const SWP_NOACTIVATE = &H10
   Public Const SWP_SHOWWINDOW = &H40
