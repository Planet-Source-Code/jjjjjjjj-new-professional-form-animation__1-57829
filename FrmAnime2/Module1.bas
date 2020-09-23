Attribute VB_Name = "mdlAnime"
'============================
'By     Jim Jose
'email  jimjosev33@yahoo.com
'============================

'PLEASE READ THIS

'If you ( Feel Satisfactory )
'   Please 'Rate' this code.
'Else
'   Give feedback to improve this code.
'End If
'Good luck
'============================

Option Explicit

Private Type RECT   'Rectangle coordinates
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type POINTAPI    'Cursor Pos co-ordinate
        X As Long
        y As Long
End Type

Public Enum AnimeEvent  'Determines the Animation on Loading/Unloading
    aUnload = 0
    aLoad = 1
End Enum

Public Enum AnimeSpeed  'Determines the Speed of animation
    aFast = 1
    aMedium = 10
    aSlow = 50
End Enum

Private DrawCol  As Long    'Determines Draw color

'Controll/Info API's Used
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long  'Gets the hdc of Desktop
Private Declare Function GetDesktopWindow Lib "user32" () As Long   'Gets the hwnd of Desktop
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Draw/Clear API's Used
Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As Long, ByVal bErase As Long) As Long     'Clear up the screen
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DrawAnimatedRects Lib "user32" (ByVal hwnd As Long, ByVal idAni As Long, lprcFrom As RECT, lprcTo As RECT) As Long     'Draw Animated rectangles( Using as the last event of animation )
'----------------------------------------------------------------------------------------------------------------------
'< Ref >
    'This Module mainly uses 'DrawAnimatedRects' API
    'This was my primary project(API) and all other are later Addings
    'The new feature is that the function can automatically detect the controll which calls to load the form and
    'the form orginates  and terminates on  the same controll
'< Info >
    '1)The Static variable (Static CurPos)is used to store cursor position
    '2)There is more possible 'Styles' that you can add.
    '3)I don't know if sub 'TransRectangle" is a standard way.
    'I  have to do that since there is no direct linedrawing API/ The hollow rectangle drawing (by API) is more complex
'< Tips >
    '1)Change the 'DrawWidth' for the sub 'TransRectangle' for some different effects
    'now it is set to the default 'One'
    '2)Change the default 'RctCount' (Rectangle count) in the sub  'PrivateAnime'
'----------------------------------------------------------------------------------------------------------------------

'Animtion using 'DrawAnimatedRects' API
Public Sub AnimateForm(Frm As Form, aEvent As AnimeEvent, Optional aSpeed As AnimeSpeed = 10, _
                                                                    Optional SleepTime As Integer = 1)
Dim ScrX        As Long    'Determines the 'Screen.TwipsPerPixelX'
Dim ScrY        As Long    'Determines the 'Screen.TwipsPerPixelY'
Dim Rct1        As RECT    'The ending rect in 'Load' event
Dim Rct2        As RECT    'The starting rect in 'Load' event
Static CurPos   As POINTAPI 'The 'Static' stores the values for unload event

ScrX = Screen.TwipsPerPixelX    'Setting value
ScrY = Screen.TwipsPerPixelY    'Setting value
DrawCol = Frm.FillColor
If aEvent = aLoad Then GetCursorPos CurPos  'Reset cur pos on Load event

With Rct1   'Setting the First(Starting) rectangle as the dimensions of the form
    .Left = Frm.Left / ScrX     'Setting value
    .Top = Frm.Top / ScrY       'Setting value
    .Right = (Frm.Left + Frm.Width) / ScrX  'Setting value
    .Bottom = (Frm.Top + Frm.Height) / ScrY 'Setting value
End With

With Rct2
    .Left = CurPos.X
    .Right = CurPos.X
    .Top = CurPos.y
    .Bottom = CurPos.y
End With

If aEvent = aLoad Then 'Load
    PrivateAnime Rct2, Rct1, aSpeed, 10    'The Animation coded by me ( not API animation ) to draw with hollow rectangles
    DrawAnimatedRects Frm.hwnd, 3, Rct2, Rct1   'The API animation
End If

If aEvent = aUnload Then 'Unload
    PrivateAnime Rct1, Rct2, aSpeed, 10    'The Animation coded by me ( not API animation ) to draw with hollow rectangles
    DrawAnimatedRects Frm.hwnd, 3, Rct1, Rct2   'The API animation
    Unload Frm  'Unloading the form in the case of 'Unload' event
End If

ClearScreen 'Clearing the Screen before exiting
End Sub

'Returns the Desktop HDC
Private Function DeskDc()
    DeskDc = GetWindowDC(GetDesktopWindow)
End Function

'Returns the DeskTop Hwnd
Private Function DeskHwnd()
    DeskHwnd = GetDesktopWindow
End Function

'Clearing the sceen
Public Sub ClearScreen()
   InvalidateRect 0&, 0&, True
End Sub

'My Animation
Public Function PrivateAnime(sRct As RECT, eRct As RECT, ByVal aSpeed As AnimeSpeed, Optional ByVal RctCount = 25)
Dim X As Integer
Dim XIncr As Double
Dim YIncr As Double
Dim HIncr As Double
Dim WIncr As Double
Dim TempRect As RECT    'Declaring a 'Temporary rectagle' the dimensions in b/w the starting and ending rectangles

    XIncr = (eRct.Left - sRct.Left) / RctCount    'Determines Amount of change in each loop for the 'Left' property
    YIncr = (eRct.Top - sRct.Top) / RctCount    'Determines Amount of change in each loop for the 'Top' property
    HIncr = ((eRct.Bottom - eRct.Top) - (sRct.Bottom - sRct.Top)) / RctCount   'Determines Amount of change in each loop for the 'Height' of rectagle
    WIncr = ((eRct.Right - eRct.Left) - (sRct.Right - sRct.Left)) / RctCount    'Determines Amount of change in each loop for the 'Width' of rectagle
    TempRect = sRct
    
    For X = 1 To RctCount 'Doing the animation
        Sleep aSpeed    'Controlling the speed
        'Setting the Temporary rectangle's dimensions
        TempRect.Left = TempRect.Left + XIncr: TempRect.Right = TempRect.Right + XIncr + WIncr
        TempRect.Top = TempRect.Top + YIncr: TempRect.Bottom = TempRect.Bottom + YIncr + HIncr
        TransRectangle DeskDc, TempRect 'Drawing the Hollow rectangle
    Next X
End Function

'My Hollow rectangle drawing method ( I don't know if there is a standard method(API) )
'I have to do this because there was no direct line drawing API ,I could find.

'This sub creates four other rectangles as the sides of the 'Required Rectangle'
'drawing all the four rectangle will result in the 'Required Rectangle'
Public Sub TransRectangle(Dhdc As Long, VRct As RECT, Optional ByVal DrawWidth As Long = 1)
Dim X As Integer
Dim hBrush  As Long
Dim TempRect(1 To 4) As RECT
    For X = 1 To 4
        TempRect(X) = VRct
        If X = 1 Then TempRect(X).Bottom = TempRect(X).Top + DrawWidth
        If X = 2 Then TempRect(X).Left = TempRect(X).Right - DrawWidth
        If X = 3 Then TempRect(X).Top = TempRect(X).Bottom - DrawWidth
        If X = 4 Then TempRect(X).Right = TempRect(X).Left + DrawWidth
        hBrush = CreateSolidBrush(DrawCol)
        FillRect DeskDc, TempRect(X), hBrush
        DeleteObject hBrush
    Next X
End Sub
