Attribute VB_Name = "RGBmod"
'*****************************************************'
'*          This code was made by Nick Smith         *'
'*                  Copyright 2000                   *'
'*                                                   *'
'*        Questions or comments?  Send mail to       *'
'*               CCSkater@mailcity.com               *'
'*****************************************************'


Public X_1 As Long
Public X_2 As Long
Public Y_1 As Long
Public Y_2 As Long

Public lastX As Long
Public lastY As Long

Public tmpRad As Long
Public rd As Single
Public gr As Single
Public bl As Single
Public pAPI As POINTAPI
Public RCT As RECT
Public Counter1 As Long
Public YN As Integer
Public tmpBMP As Long
Public brushBMP As Long

Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Declare Function SetPixelV& Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long)


Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function TextOut& Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long)
Declare Function GetCursorPos& Lib "user32" (lpPoint As POINTAPI)
Declare Function GetWindowRect& Lib "user32" (ByVal hwnd As Long, lpRect As RECT)

Declare Function CreateDIBPatternBrushPt& Lib "gdi32" (lpPackedDIB As Any, ByVal wUsage As Long)
Declare Function FillPath& Lib "gdi32" (ByVal hDC As Long)
Declare Function BeginPath& Lib "gdi32" (ByVal hDC As Long)
Declare Function EndPath& Lib "gdi32" (ByVal hDC As Long)
Declare Function Arc& Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long)
Declare Function Chord& Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long)
Declare Function CloseClipboard& Lib "user32" ()
Declare Function CloseMetaFile& Lib "gdi32" (ByVal hMF As Long)
Declare Function CreateHatchBrush& Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long)
Declare Function CreateMetaFile& Lib "gdi32" Alias "CreateMetaFileA" (ByVal lpString As String)
Declare Function CreatePen& Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long)
Declare Function CreateSolidBrush& Lib "gdi32" (ByVal crColor As Long)
Declare Function DeleteMetaFile& Lib "gdi32" (ByVal hMF As Long)
Declare Function DeleteObject& Lib "gdi32" (ByVal hObject As Long)
Declare Function DrawFocusRect& Lib "user32" (ByVal hDC As Long, lpRect As RECT)
Declare Function Ellipse& Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)
Declare Function EmptyClipboard& Lib "user32" ()
Declare Function EnumMetaFile Lib "gdi32" (ByVal hDC As Long, ByVal hMF As Long, ByVal lpCallbackFunc As Long, ByVal lpClientData As Long) As Long
Declare Function GetClientRect& Lib "user32" (ByVal hwnd As Long, lpRect As RECT)
Declare Function GetMetaFileBitsEx& Lib "gdi32" (ByVal hMF As Long, ByVal nSize As Long, lpvData As Any)
Declare Function GlobalAlloc& Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long)
Declare Function GlobalFree& Lib "kernel32" (ByVal hMem As Long)
Declare Function GlobalLock& Lib "kernel32" (ByVal hMem As Long)
Declare Function GetObjectType& Lib "gdi32" (ByVal hgdiobj As Long)
Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalUnlock& Lib "kernel32" (ByVal hMem As Long)
Declare Function InflateRect& Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long)
Declare Function LineTo& Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long)
Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Declare Function OpenClipboard& Lib "user32" (ByVal hwnd As Long)
Declare Function Pie& Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long)
Declare Function PlayMetaFile& Lib "gdi32" (ByVal hDC As Long, ByVal hMF As Long)
Declare Function PlayMetaFileRecord& Lib "gdi32" (ByVal hDC As Long, ByVal lpHandletable As Long, lpMetaRecord As Any, ByVal nHandles As Long)
Declare Function Polyline& Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long)
Declare Function Polygon& Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long)
Declare Function Rectangle& Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)
Declare Function RestoreDC& Lib "gdi32" (ByVal hDC As Long, ByVal nSavedDC As Long)
Declare Function SaveDC& Lib "gdi32" (ByVal hDC As Long)
Declare Function SelectObject& Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long)
Declare Function SetClipboardData& Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long)
Declare Function SetMapMode& Lib "gdi32" (ByVal hDC As Long, ByVal nMapMode As Long)
Declare Function SetMetaFileBitsEx& Lib "gdi32" (ByVal nSize As Long, lpData As Byte)
Declare Function SetMetaFileBitsBuffer& Lib "gdi32" Alias "SetMetaFileBitsEx" (ByVal nSize As Long, ByVal lpData As Long)
Declare Function BitBlt& Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long)

Declare Function GetDC& Lib "user32" (ByVal hwnd As Long)
Declare Function GetDesktopWindow& Lib "user32" ()

Declare Function SetPolyFillMode& Lib "gdi32" (ByVal hDC As Long, ByVal nPolyFillMode As Long)
Declare Function SetRect& Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)
Declare Function SetViewportExtEx& Lib "gdi32" (ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, lpSize As SIZE)
Declare Function SetViewportOrgEx& Lib "gdi32" (ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI)
Declare Function SetWindowOrgEx& Lib "gdi32" (ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI)
Declare Function SetWindowExtEx& Lib "gdi32" (ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, lpSize As SIZE)

Declare Function lopen Lib "kernel32" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Long) As Long
Declare Function lclose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long
Declare Function lcreat Lib "kernel32" Alias "_lcreat" (ByVal lpPathName As String, ByVal iAttribute As Long) As Long
Declare Function llseek Lib "kernel32" Alias "_llseek" (ByVal hFile As Long, ByVal lOffset As Long, ByVal iOrigin As Long) As Long
Declare Function lread Lib "kernel32" Alias "_lread" (ByVal hFile As Long, lpBuffer As Any, ByVal wBytes As Long) As Long
Declare Function lwrite Lib "kernel32" Alias "_lwrite" (ByVal hFile As Long, lpBuffer As Any, ByVal wBytes As Long) As Long

Declare Function hread Lib "kernel32" Alias "_hread" (ByVal hFile As Long, lpBuffer As Any, ByVal lBytes As Long) As Long
Declare Function hwrite Lib "kernel32" Alias "_hwrite" (ByVal hFile As Long, lpBuffer As Any, ByVal lBytes As Long) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long

Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40

Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Const SRCCOPY = &HCC0020

Public Const OBJ_PEN = 1
Public Const OBJ_BRUSH = 2
Public Const OBJ_DC = 3
Public Const OBJ_METADC = 4
Public Const OBJ_PAL = 5
Public Const OBJ_FONT = 6
Public Const OBJ_BITMAP = 7
Public Const OBJ_REGION = 8
Public Const OBJ_METAFILE = 9
Public Const OBJ_MEMDC = 10
Public Const OBJ_EXTPEN = 11
Public Const OBJ_ENHMETADC = 12
Public Const OBJ_ENHMETAFILE = 13
Public Const PS_USERSTYLE = 7
Public Const PS_TYPE_MASK = &HF0000
Public Const PS_STYLE_MASK = &HF
Public Const PS_SOLID = 0
Public Const PS_NULL = 5
Public Const PS_JOIN_ROUND = &H0
Public Const PS_JOIN_MITER = &H2000
Public Const PS_JOIN_MASK = &HF000
Public Const PS_JOIN_BEVEL = &H1000
Public Const PS_INSIDEFRAME = 6
Public Const PS_GEOMETRIC = &H10000
Public Const PS_ENDCAP_SQUARE = &H100
Public Const PS_ENDCAP_ROUND = &H0
Public Const PS_ENDCAP_MASK = &HF00
Public Const PS_ENDCAP_FLAT = &H200
Public Const PS_DOT = 2
Public Const PS_DASHDOTDOT = 4
Public Const PS_DASHDOT = 3
Public Const PS_DASH = 1
Public Const PS_COSMETIC = &H0
Public Const PS_ALTERNATE = 8

Type BITMAP '14 bytes
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type

Type RECT
        Left As Integer
        Top As Integer
        Right As Integer
        Bottom As Integer
End Type

Type POINTAPI
        X As Integer
        Y As Integer
End Type

Type SIZE
        cx As Integer
        cy As Integer
End Type

Type METAFILEPICT    '8 Bytes
    mm As Integer
    xExt As Integer
    yExt As Integer
    hMF As Integer
End Type

Public Type LOGBRUSH
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type




Public Function Get_RGB(RGBValue As Long, val As Integer)
'note:  for the val argument, 1 = red, 2 = green,'
'3 = blue'

'check for valid rgb value, then execute math depending'
'on if the user is requesting reg, green, or blue'
If RGBValue > -1 And val > 0 And val < 4 Then
    Select Case val
        Case 1
            Get_RGB = (RGBValue And &HFF&)
        Case 2
            Get_RGB = (RGBValue And &HFF00&) / &H100
        Case 3
            Get_RGB = (RGBValue And &HFF0000) / &H10000
    End Select
End If
End Function

Public Sub Draw_Gradient(pctbox As PictureBox, tmpColor1 As Long, tmpColor2 As Long, style As Long)
'if style is set to 0 or 2, the gradient is vertical'
'so the for...next statement will be set to repeat'
'as many times as there are pixels wide in the picture'
'box.'
'if style is set to 1 or 3, the gradient is horizontal'
'so the for...next statement will be set to repeat'
'as many times as there are pixels high in the picture'
'box.'
    If style = 1 Or style = 3 Then tmpDir% = pctbox.ScaleWidth
    If style = 0 Or style = 2 Then tmpDir% = pctbox.ScaleHeight

'if the style is 0 or 1, then the starting color is on'
'the left or top, so set the first RGBs to the first'
'color, else set the first RGBs to the second color'
If style = 0 Or style = 1 Then
tmpRed1& = Get_RGB(tmpColor1&, 1)
tmpGreen1& = Get_RGB(tmpColor1&, 2)
tmpBlue1& = Get_RGB(tmpColor1&, 3)
tmpRed2& = Get_RGB(tmpColor2&, 1)
tmpGreen2& = Get_RGB(tmpColor2&, 2)
tmpBlue2& = Get_RGB(tmpColor2&, 3)

'get the difference between each RGB of the two colors'
'then them by tmpDir% which was set earlier.  This '
'sets the step of color to go up for each column or '
'row of pixels                               ^
rd = (tmpRed2& - tmpRed1&) / tmpDir%        '|'
gr = (tmpGreen2& - tmpGreen1&) / tmpDir%    '|'
bl = (tmpBlue2& - tmpBlue1&) / tmpDir%      '|'
Else                                        '|'
tmpRed1& = Get_RGB(tmpColor2&, 1)           '|'
tmpGreen1& = Get_RGB(tmpColor2&, 2)         '|'
tmpBlue1& = Get_RGB(tmpColor2&, 3)          '|'
tmpRed2& = Get_RGB(tmpColor1&, 1)           '|'
tmpGreen2& = Get_RGB(tmpColor1&, 2)         '|'
tmpBlue2& = Get_RGB(tmpColor1&, 3)          '|'
'same as above ------------------------------'
rd = (tmpRed2& - tmpRed1&) / tmpDir%
gr = (tmpGreen2& - tmpGreen1&) / tmpDir%
bl = (tmpBlue2& - tmpBlue1&) / tmpDir%
End If

    If style = 1 Or style = 3 Then
        For Counter1 = 0 To tmpDir%
        'set the final RGBs of the column of pixels'
        'indexed at Counter1.  Multiply the step by'
        'Counter1 which is the index of columns'
        finalRed& = tmpRed1& + (rd * Counter1)
        finalGreen& = tmpGreen1& + (gr * Counter1)
        finalBlue& = tmpBlue1& + (bl * Counter1)

        'convert RGB values to get a color'
        finalCol! = RGB(finalRed&, finalGreen&, finalBlue&)
        
        'draw a line on column Counter1, letting Counter1'
        'be the index of columns'
        pctbox.Line (Counter1, 0)-(Counter1, tmpDir%), finalCol!
            
        Next Counter1
    ElseIf style = 0 Or style = 2 Then
        For Counter1 = 0 To tmpDir%
        'set the final RGBs of the row of pixels'
        'indexed at Counter1.  Multiply the step by'
        'Counter1 which is the index of columns'
        finalRed& = tmpRed1& + (rd * Counter1)
        finalGreen& = tmpGreen1& + (gr * Counter1)
        finalBlue& = tmpBlue1& + (bl * Counter1)
        
        'convert RGB values to get a color'
        finalCol! = RGB(finalRed&, finalGreen&, finalBlue&)
        
        'draw a line on row Counter1, letting Counter1'
        'be the index of rows'
        pctbox.Line (0, Counter1)-(tmpDir%, Counter1), finalCol!
            
        Next Counter1
    End If

End Sub

Sub hold(timetohold)
Current = Timer
Do While Timer - Current < val(timetohold)
DoEvents
Loop
End Sub

Public Sub Draw_GradientCircle(pctbox As PictureBox, X As Long, Y As Long, color1 As Long, color2 As Long)
Dim SQNum As Double
If X > (pctbox.ScaleWidth / 2) Then
    If Y > (pctbox.ScaleHeight / 2) Then
    SQNum = (X * X) + (Y * Y)
    tmpDir% = Sqr(SQNum)
    Else
    SQNum = (X * X) + ((pctbox.ScaleHeight - Y) * (pctbox.ScaleHeight - Y))
    tmpDir% = Sqr(SQNum)
    End If
ElseIf X < (pctbox.ScaleWidth / 2) Then
    If Y > (pctbox.ScaleHeight / 2) Then
    SQNum = ((pctbox.ScaleWidth - X) * (pctbox.ScaleWidth - X)) + (Y * Y)
    tmpDir% = Sqr(SQNum)
    Else
    SQNum = ((pctbox.ScaleWidth - X) * (pctbox.ScaleWidth - X)) + ((pctbox.ScaleHeight - Y) * (pctbox.ScaleHeight - Y))
    tmpDir% = Sqr(SQNum)
    End If
Else
SQNum = ((pctbox.ScaleWidth / 2) * (pctbox.ScaleWidth / 2)) + ((pctbox.ScaleHeight / 2) * (pctbox.ScaleHeight / 2))
End If


tmpRed1& = Get_RGB(color1&, 1)
tmpGreen1& = Get_RGB(color1&, 2)
tmpBlue1& = Get_RGB(color1&, 3)
tmpRed2& = Get_RGB(color2&, 1)
tmpGreen2& = Get_RGB(color2&, 2)
tmpBlue2& = Get_RGB(color2&, 3)

rd = (tmpRed2& - tmpRed1&) / tmpDir%
gr = (tmpGreen2& - tmpGreen1&) / tmpDir%
bl = (tmpBlue2& - tmpBlue1&) / tmpDir%

pctbox.DrawWidth = 2
pctbox.DrawMode = 13

For Counter1 = 0 To tmpDir% - 1
finalRed& = tmpRed1& + (rd * Counter1)
finalGreen& = tmpGreen1& + (gr * Counter1)
finalBlue& = tmpBlue1& + (bl * Counter1)

finalCol! = RGB(finalRed&, finalGreen&, finalBlue&)
pctbox.Circle (X, Y), Counter1, finalCol!
Next Counter1
End Sub

Sub KeepOnTop(frm As Form)
SetWindowPos frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
End Sub
