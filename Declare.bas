Attribute VB_Name = "Declare"
'���������� ������� � �������� Win32API
'    BOOL SetWindowPos(
'
'        HWND hWnd,            // ���������� ����
'        HWND hWndInsertAfter, // ���������� ������� ����������
'        int x,                // ������� �� �����������
'        int y,                // ������� �� ���������
'        int cx,               // ������
'        int cy,               // ������
'        UINT uFlags           // ������ ���������������� ����
'
'    );

'[in] ���������� ����, ������� ������������ �������������� ���� �  Z-������������������.
'���� �������� ������ ���� ���������� ���� ��� ����� �� ��������� ��������:
'HWND_BOTTOM
'�������� ���� ����� Z-������������������.
'���� �������� hWnd �������������� ����� ������� ����, ���� ������ ���� ������
'������ �������� � ���������� ����� ���� ������ ����.
'HWND_NOTOPMOST
'�������� ���� ����� ����� �� ������ �������� ������
'(�� ���� ������ ���� ����� ������� ����).
'���� ������ �� ����� �������� �������, ���� ���� - ��� �� ����� ������� ����.
'HWND_TOP
'�������� ���� ������� Z-������������������.
'HWND_TOPMOST
'�������� ���� ����� �� ������ �������� ������.
'���� ��������� ���� ����� ������� ������� ���� �����, ����� ��� ������ ����������.


   Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

   Public Const HWND_TOPMOST = -1
   Public Const HWND_NOTOPMOST = -2
   Public Const SWP_NOMOVE = &H2
   Public Const SWP_NOSIZE = &H1
   Public Const SWP_NOACTIVATE = &H10
   Public Const SWP_SHOWWINDOW = &H40
