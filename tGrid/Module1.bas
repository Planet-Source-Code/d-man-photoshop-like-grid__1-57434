Attribute VB_Name = "Module1"
Option Explicit

'Apis
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Public Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function SetRect Lib "user32.dll" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Public Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Public Declare Function CreatePatternBrush Lib "gdi32.dll" (ByVal hBitmap As Long) As Long
Public Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long

Public Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long

Public Function GetGridBrush(ByVal qCellSize As Integer, qHdc As Long) As Long

Dim brWhite As Long
Dim brGray As Long
Dim brDC As Long
Dim brBmp As Long

Dim qRect As RECT

brWhite = CreateSolidBrush(vbWhite)
brGray = CreateSolidBrush(RGB(200, 200, 200))

brDC = CreateCompatibleDC(qHdc)
brBmp = CreateCompatibleBitmap(qHdc, 2 * qCellSize, 2 * qCellSize)
SelectObject brDC, brBmp


SetRect qRect, 0, 0, qCellSize, qCellSize
FillRect brDC, qRect, brWhite

SetRect qRect, qCellSize, 0, 2 * qCellSize, qCellSize
FillRect brDC, qRect, brGray

SetRect qRect, 0, qCellSize, qCellSize, 2 * qCellSize
FillRect brDC, qRect, brGray

SetRect qRect, qCellSize, qCellSize, 2 * qCellSize, 2 * qCellSize
FillRect brDC, qRect, brWhite

Dim bmpBr As Long
bmpBr = CreatePatternBrush(brBmp)

DeleteDC brDC
DeleteObject brBmp
DeleteObject brWhite
DeleteObject brGray

GetGridBrush = bmpBr

End Function





