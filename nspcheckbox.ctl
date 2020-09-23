VERSION 5.00
Begin VB.UserControl nspCheckBox 
   ClientHeight    =   1485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3330
   LockControls    =   -1  'True
   ScaleHeight     =   99
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   222
   Begin VB.PictureBox picButton 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   0
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   14
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   210
   End
End
Attribute VB_Name = "nspCheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************
'* Small but attractive checkbox                                  *
'* Author: John Underhill (Steppenwolfe)                          *
'* Modified for: Heriberto Mantilla Santamaría (only a little) :) *
'******************************************************************
'* NOTE: The intellectual author of the code authorizes myself to *
'* make it public.                                                *
'******************************************************************

Option Explicit

Private Const ALL_MESSAGES                  As Long = -1
Private Const MSG_ENTRIES                   As Long = 32
Private Const WNDPROC_OFF                   As Long = &H38
Private Const GWL_WNDPROC                   As Long = -4
Private Const IDX_SHUTDOWN                  As Long = 1
Private Const IDX_HWND                      As Long = 2
Private Const IDX_WNDPROC                   As Long = 9
Private Const IDX_BTABLE                    As Long = 11
Private Const IDX_ATABLE                    As Long = 12
Private Const IDX_PARM_USER                 As Long = 13
Private Const WM_MOUSEMOVE                  As Long = &H200
Private Const WM_MOUSELEAVE                 As Long = &H2A3
Private Const COLOR_GRAYTEXT                As Integer = 17
Private Const defBackColor                  As Long = &H8000000F
Private Const defBorderColor                As Long = vbHighlight
Private Const DC_TEXT                       As Long = &H8
Private Const Version                       As String = "NSPowertool CheckBox Control"
Private Const VER_PLATFORM_WIN32_NT         As Integer = 2
Private Const GRADIENT_FILL_RECT_H          As Long = 0
Private Const GRADIENT_FILL_RECT_V          As Long = 1

Private Type TRIVERTEX
    X                                       As Long
    Y                                       As Long
    Red                                     As Integer
    Green                                   As Integer
    Blue                                    As Integer
    alpha                                   As Integer
End Type

Private Type GRADIENT_RECT
    UpperLeft                               As Long
    LowerRight                              As Long
End Type

Private Type POINTAPI
    X                                       As Long
    Y                                       As Long
End Type

Private Type RECT
    Left                                    As Long
    Top                                     As Long
    Right                                   As Long
    Bottom                                  As Long
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize                     As Long
    dwMajorVersion                          As Long
    dwMinorVersion                          As Long
    dwBuildNumber                           As Long
    dwPlatformId                            As Long
    szCSDVersion                            As String * 128
End Type

Private Type RGBQUAD
    rgbBlue                                 As Byte
    rgbGreen                                As Byte
    rgbRed                                  As Byte
    rgbReserved                             As Byte
End Type

Private Enum eMsgWhen
    MSG_BEFORE = 1
    MSG_AFTER = 2
    MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER
End Enum

Private Type TRACKMOUSEEVENT_STRUCT
    cbSize                                  As Long
    dwFlags                                 As TRACKMOUSEEVENT_FLAGS
    hwndTrack                               As Long
    dwHoverTime                             As Long
End Type

Private Enum TRACKMOUSEEVENT_FLAGS
    TME_HOVER = &H1&
    TME_LEAVE = &H2&
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum

Public Enum eAlign
    AlignLeft = &H0
    AlignRight = &H1
End Enum

Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, _
                                                 lpRect As RECT, _
                                                 ByVal hBrush As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function GradientFill Lib "Msimg32.dll" (ByVal hdc As Long, _
                                                         pVertex As TRIVERTEX, _
                                                         ByVal dwNumVertex As Long, _
                                                         pMesh As GRADIENT_RECT, _
                                                         ByVal dwNumMesh As Long, _
                                                         ByVal dwMode As Long) As Long

Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, _
                                                               ByVal HPALETTE As Long, _
                                                               pccolorref As Long) As Long

Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, _
                                                       ByVal hWnd As Long, _
                                                       ByVal Msg As Long, _
                                                       ByVal wParam As Long, _
                                                       ByVal lParam As Long) As Long

Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long

Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, _
                                                        ByVal lpProcName As String) As Long

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, _
                                                                lpdwProcessId As Long) As Long

Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long

Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, _
                                                      ByVal nIndex As Long, _
                                                      ByVal dwNewLong As Long) As Long

Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, _
                                                      ByVal dwSize As Long, _
                                                      ByVal flAllocationType As Long, _
                                                      ByVal flProtect As Long) As Long

Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, _
                                                     ByVal dwSize As Long, _
                                                     ByVal dwFreeType As Long) As Long

Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, _
                                                  ByVal Source As Long, _
                                                  ByVal Length As Long)

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long

Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, _
                                                                     pSrc As Any, _
                                                                     ByVal ByteLen As Long)

Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, _
                                                ByVal nWidth As Long, _
                                                ByVal crColor As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function DrawTextA Lib "user32" (ByVal hdc As Long, _
                                                 ByVal lpStr As String, _
                                                 ByVal nCount As Long, _
                                                 lpRect As RECT, _
                                                 ByVal wFormat As Long) As Long

Private Declare Function DrawTextW Lib "user32" (ByVal hdc As Long, _
                                                 ByVal lpStr As Long, _
                                                 ByVal nCount As Long, _
                                                 lpRect As RECT, _
                                                 ByVal wFormat As Long) As Long

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, _
                                             ByVal X As Long, _
                                             ByVal Y As Long) As Long

Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, _
                                               ByVal X As Long, _
                                               ByVal Y As Long, _
                                               lpPoint As Any) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal hObject As Long) As Long

Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal crColor As Long) As Long

Public Event Click()
Public Event MouseEnter()
Public Event MouseLeave()
Public Event Status(ByVal sStatus As String)

'/* subclass
Private m_bTrack                        As Boolean
Private m_bTrackUser32                  As Boolean
Private m_bInCtrl                       As Boolean
Private z_ScMem                         As Long
Private z_Sc(64)                        As Long
Private z_Funk                          As Collection
'/* control
Private m_bChecked                      As Boolean
Private m_bEnabled                      As Boolean
Private m_oTopColor                     As OLE_COLOR
Private m_bFocus                        As Boolean
Private m_stdFont                       As StdFont
Private m_oBottomColor                  As OLE_COLOR
Private m_tRect                         As RECT
Private m_sCaption                      As String
Private m_iState                        As Integer
Private m_eAlign                        As eAlign
Private m_oBackColor                    As OLE_COLOR
Private m_oBorderColor                  As OLE_COLOR
Private m_objObject                     As Object
Private m_oFocusColor                   As OLE_COLOR
Private m_oForeColor                    As OLE_COLOR
Private m_oRadiusColor                  As OLE_COLOR
Private m_bIsNT                         As Boolean


Public Property Get Alignment() As eAlign
    Alignment = m_eAlign
End Property

Public Property Let Alignment(ByVal New_Align As eAlign)

    m_eAlign = New_Align
    PropertyChanged "Alignment"
    DrawCheckBox m_iState, m_bChecked

End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_oBackColor
End Property

Public Property Let BackColor(ByVal New_Color As OLE_COLOR)

    m_oBackColor = ConvertSystemColor(New_Color)
    PropertyChanged "BackColor"
    DrawCheckBox m_iState, m_bChecked

End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = m_oBorderColor
End Property

Public Property Let BorderColor(ByVal New_Color As OLE_COLOR)

    m_oBorderColor = ConvertSystemColor(New_Color)
    PropertyChanged "BorderColor"
    DrawCheckBox m_iState, m_bChecked

End Property

Public Property Get Caption() As String
    Caption = m_sCaption
End Property

Public Property Let Caption(ByVal New_Caption As String)

    m_sCaption = New_Caption
    PropertyChanged "Caption"
    DrawCheckBox m_iState, m_bChecked

End Property

Public Property Get Enabled() As Boolean
    Enabled = m_bEnabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)

    UserControl.Enabled = New_Enabled
    m_bEnabled = New_Enabled
    PropertyChanged "Enabled"

    If Not m_bEnabled Then
        DrawCheckBox -1, m_bChecked
    Else
        DrawCheckBox m_iState, m_bChecked
    End If

End Property

Public Property Get FocusColor() As OLE_COLOR
    FocusColor = m_oFocusColor
End Property

Public Property Let FocusColor(ByVal NewColor As OLE_COLOR)

    m_oFocusColor = ConvertSystemColor(NewColor)
    PropertyChanged "FocusColor"

End Property

Public Property Get Font() As StdFont
    Set Font = m_stdFont
End Property

Public Property Set Font(ByVal New_Font As StdFont)

On Error Resume Next

    With m_stdFont
        .Name = New_Font.Name
        .Size = New_Font.Size
        .Bold = New_Font.Bold
        .Italic = New_Font.Italic
        .Underline = New_Font.Underline
        .Strikethrough = New_Font.Strikethrough
    End With

    PropertyChanged "Font"
    DrawCheckBox m_iState, m_bChecked

On Error GoTo 0

End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_oForeColor
End Property

Public Property Let ForeColor(ByVal NewColor As OLE_COLOR)

    m_oForeColor = ConvertSystemColor(NewColor)
    PropertyChanged "ForeColor"
    DrawCheckBox m_iState, m_bChecked

End Property

Public Property Get GetControlVersion() As String
    GetControlVersion = Version & " © " & Year(Now)
End Property

Public Property Get MouseIcon() As StdPicture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property

Public Property Let mspMousePointer(ByVal mspMousePointer As MousePointerConstants)

    UserControl.MousePointer = mspMousePointer
    PropertyChanged "MousePointer"

End Property


Public Property Set stdMouseIcon(ByVal stdMouseIcon As StdPicture)

    Set UserControl.MouseIcon = stdMouseIcon
    PropertyChanged "MouseIcon"

End Property

Private Sub DrawCheckBox(ByVal iState As Integer, Optional ByVal Focus As Boolean = False)
'/* draw main

Dim lHdc            As Long
Dim oForeColor      As OLE_COLOR
Dim oColor          As OLE_COLOR
Dim tRect           As RECT
Dim tSel            As RECT

On Error Resume Next

    If iState = 0 Then
        iState = 1
    End If
    With UserControl
        .Cls
        Set .Font = m_stdFont
        m_iState = iState
        .BackColor = m_oBackColor
        .Height = .TextHeight("Qr") * Screen.TwipsPerPixelY + 35
    End With

    With picButton
        .Cls
        .Top = (UserControl.ScaleHeight / 2) - 8
        If m_eAlign = 0 Then
            .Left = 0
        Else
            .Left = UserControl.ScaleWidth - 16
        End If
        .BackColor = m_oBackColor
        lHdc = .hdc
    End With

    '/* base rectangle
    With tRect
        .Bottom = picButton.Left + 14
        .Left = picButton.Left
        .Right = picButton.Top + 14
        .Top = picButton.Top
    End With

    Select Case m_iState
        '/* normal
    Case 1
        oColor = ConvertSystemColor(m_oBorderColor)
        oForeColor = m_oForeColor
        Draw_Gradient tRect, lHdc, ShiftColorXP(oColor, 43), ShiftColorXP(oColor, 143)

        '/* focus
    Case 2
        oColor = ShiftColorXP(m_oBorderColor, 38)
        oForeColor = ShiftColorXP(oColor, 10)
        Draw_Gradient tRect, lHdc, ShiftColorXP(oColor, 80), ShiftColorXP(oColor, 160)
        
        '/* disabled
    Case -1
        If Not m_bChecked Then
            oColor = ConvertSystemColor(oColor)
            oForeColor = m_oForeColor
            Draw_Gradient tRect, lHdc, ShiftColorXP(oColor, 120), &HAAAAAA
        Else
            oColor = ShiftColorXP(m_oBorderColor, 180)
            oForeColor = m_oForeColor
            Draw_Gradient tRect, lHdc, ShiftColorXP(oColor, 43), &HAAAAAA
            oColor = ShiftColorXP(oForeColor, 180)
            With tSel
                .Bottom = 10
                .Left = tRect.Left + 3
                .Right = 10
                .Top = picButton.Top + 5
            End With
            Draw_Gradient tSel, lHdc, oColor, &H0
        End If
        
    End Select
    If Focus = True And m_iState = 2 Then '/* checked focus
       oColor = ShiftColorXP(m_oBorderColor, 38)
       oForeColor = ShiftColorXP(oColor, 10)
       Draw_Gradient tRect, lHdc, oColor, ShiftColorXP(oColor, 150)
       With tSel
           .Bottom = 10
           .Left = tRect.Left + 3
           .Right = 10
           .Top = picButton.Top + 5
       End With
       Draw_Gradient tSel, lHdc, m_oFocusColor, &H0
    ElseIf Focus = True And m_iState = 1 Then '/* checked
       oColor = ShiftColorXP(m_oBorderColor, -53)
       oForeColor = m_oForeColor
       Draw_Gradient tRect, lHdc, ShiftColorXP(oColor, 43), ShiftColorXP(oColor, 143)
       oColor = ShiftColorXP(oColor, 20)
       With tSel
           .Bottom = 10
           .Left = tRect.Left + 3
           .Right = 10
           .Top = picButton.Top + 5
       End With
       Draw_Gradient tSel, lHdc, m_oFocusColor, &H0
    End If
    '/* draw frame and caption
    DrawCaption m_sCaption, oForeColor
    DrawFrame tRect, picButton.hdc, m_oBorderColor

On Error GoTo 0

End Sub

Private Function ConvertSystemColor(ByVal theColor As Long) As Long

    OleTranslateColor theColor, 0, ConvertSystemColor

End Function

Private Sub Draw_Gradient(ByRef tRect As RECT, _
                          ByVal lHdc As Long, _
                          ByVal oGradTop As OLE_COLOR, _
                          ByVal oGradBottom As OLE_COLOR, _
                          Optional ByVal bVertical As Boolean)


Dim tVert(0 To 1) As TRIVERTEX
Dim oColor        As OLE_COLOR
Dim tGradRect     As GRADIENT_RECT


    oColor = TranslateColor(oGradTop)
    With tVert(0)
        .X = tRect.Left
        .Y = tRect.Top
        .Red = GetRed(oColor)
        .Green = GetGreen(oColor)
        .Blue = GetBlue(oColor)
    End With

    oColor = TranslateColor(oGradBottom)
    With tVert(1)
        .X = tRect.Right
        .Y = tRect.Bottom
        .Red = GetRed(oColor)
        .Green = GetGreen(oColor)
        .Blue = GetBlue(oColor)
    End With

    With tGradRect
        .UpperLeft = 0
        .LowerRight = 1
    End With

    GradientFill lHdc, tVert(0), 2, tGradRect, 1, IIf(Not bVertical, GRADIENT_FILL_RECT_H, GRADIENT_FILL_RECT_V)

End Sub

Private Sub DrawCaption(ByVal lCaption As String, _
                        Optional ByVal lColor As OLE_COLOR = &HF0)

    If Not m_bEnabled Then
        lColor = GetSysColor(COLOR_GRAYTEXT)
    End If
    
    SetTextColor UserControl.hdc, lColor
    m_tRect.Bottom = UserControl.ScaleHeight
    m_tRect.Top = 1
    
    If m_eAlign = 0 Then
        m_tRect.Left = 18
    Else
        m_tRect.Left = 0
    End If
    
    m_tRect.Right = UserControl.ScaleWidth
    If m_bIsNT Then
        DrawTextW UserControl.hdc, StrPtr(lCaption), Len(lCaption), m_tRect, DC_TEXT
    Else
        DrawTextA UserControl.hdc, lCaption, Len(lCaption), m_tRect, DC_TEXT
    End If

End Sub

Private Sub DrawFrame(ByRef tRect As RECT, _
                      ByVal lHdc As Long, _
                      Optional ByVal lColor As Long)


Dim lBrush As Long
Dim tFrame As RECT

On Error Resume Next

    With tFrame
        .Bottom = tRect.Bottom
        .Left = 0
        .Right = tRect.Right
        .Top = 1
    End With

    '/* create brush
    lBrush = CreateSolidBrush(TranslateColor(lColor))
    '/* draw the frame
    FrameRect lHdc, tFrame, lBrush
    '/* cleanup
    DeleteObject lBrush

On Error GoTo 0

End Sub


Private Function GetBlue(ByVal oColor As OLE_COLOR) As Long

    GetBlue = ((oColor \ &H10000) And &HFF) * &H100&
    If GetBlue >= &H8000& Then
        GetBlue = GetBlue - &H10000
    End If

End Function

Private Function GetGreen(ByVal oColor As OLE_COLOR) As Long

    GetGreen = ((oColor \ &H100) And &HFF) * &H100&
    If GetGreen >= &H8000& Then
        GetGreen = GetGreen - &H10000
    End If

End Function

Private Function GetRed(ByVal oColor As OLE_COLOR) As Long

    GetRed = ((oColor \ &H1) And &HFF) * &H100&
    If GetRed >= &H8000& Then
        GetRed = GetRed - &H10000
    End If

End Function

Private Sub picButton_GotFocus()

    UserControl_GotFocus
    
End Sub

Public Property Get hWnd() As Long

    hWnd = UserControl.hWnd

End Property

Private Function IsFunctionExported(ByVal sFunction As String, _
                                    ByVal sModule As String) As Boolean

Dim hMod            As Long
Dim bLibLoaded      As Boolean

    hMod = GetModuleHandleA(sModule)

    If hMod = 0 Then
        hMod = LoadLibraryA(sModule)
        If hMod Then
            bLibLoaded = True
        End If
    End If

    If hMod Then
        If GetProcAddress(hMod, sFunction) Then
            IsFunctionExported = True
        End If
    End If

    If bLibLoaded Then
        FreeLibrary hMod
    End If

End Function

Private Sub picButton_KeyDown(KeyCode As Integer, _
                              Shift As Integer)

    If KeyCode = vbKeySpace Then
        DrawCheckBox m_iState, m_bChecked
    End If

End Sub

Private Sub picButton_LostFocus()

    UserControl_LostFocus

End Sub

Private Sub picButton_MouseDown(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)

    UserControl_MouseDown Button, Shift, X, Y

End Sub

Private Sub picButton_MouseUp(Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)

    m_bFocus = True
    If m_bChecked Then
        DrawCheckBox 2, True
    Else
        DrawCheckBox 2, m_bChecked
    End If
    
End Sub

Private Sub picButton_Paint()

    DrawCheckBox m_iState, m_bChecked

End Sub

Private Sub SetAccessKeys()


Dim AmperSandPos        As Long


    UserControl.AccessKeys = ""
    If Len(Caption) > 1 Then
        AmperSandPos = InStr(1, Caption, "&", vbTextCompare)
        If AmperSandPos < Len(Caption) Then
            If AmperSandPos > 0 Then
                If (Mid$(Caption, AmperSandPos + 1, 1) <> "&") Then
                    UserControl.AccessKeys = LCase$(Mid$(Caption, AmperSandPos + 1, 1))
                Else
                    AmperSandPos = InStr(AmperSandPos + 2, Caption, "&", vbTextCompare)
                    If (Mid$(Caption, AmperSandPos + 1, 1) <> "&") Then
                        UserControl.AccessKeys = LCase$(Mid$(Caption, AmperSandPos + 1, 1))
                    End If
                End If
            End If
        End If
    End If

End Sub

Private Function ShiftColorXP(ByVal theColor As Long, _
                              Optional ByVal Base As Long = &HB0) As Long


Dim cRed        As Long
Dim cBlue       As Long
Dim Delta       As Long
Dim cGreen      As Long


    cBlue = ((theColor \ &H10000) Mod &H100)
    cGreen = ((theColor \ &H100) Mod &H100)
    cRed = (theColor And &HFF)
    Delta = &HFF - Base
    cBlue = Base + cBlue * Delta \ &HFF
    cGreen = Base + cGreen * Delta \ &HFF
    cRed = Base + cRed * Delta \ &HFF
    If cRed > 255 Then cRed = 255
    If cGreen > 255 Then cGreen = 255
    If cBlue > 255 Then cBlue = 255
    ShiftColorXP = cRed + 256& * cGreen + 65536 * cBlue

End Function

Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)

Dim tme         As TRACKMOUSEEVENT_STRUCT

    If m_bTrack Then
        With tme
            .cbSize = Len(tme)
            .dwFlags = TME_LEAVE
            .hwndTrack = lng_hWnd
        End With

        If m_bTrackUser32 Then
            TrackMouseEvent tme
        Else
            TrackMouseEventComCtl tme
        End If
    End If

End Sub

Private Function TranslateColor(ByVal oClr As OLE_COLOR, _
                                Optional hPal As Long = 0) As Long

    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = -1
    End If

End Function

Private Sub UserControl_GotFocus()

    m_bFocus = True
    If m_bChecked Then
        DrawCheckBox 2, True
    Else
        DrawCheckBox 2, m_bChecked
    End If

End Sub

Private Sub UserControl_Initialize()

Dim OS      As OSVERSIONINFO

    OS.dwOSVersionInfoSize = Len(OS)
    GetVersionEx OS
    m_bIsNT = ((OS.dwPlatformId And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)

End Sub

Private Sub UserControl_InitProperties()

    m_eAlign = 0
    m_bEnabled = True
    ForeColor = &H80000012
    FocusColor = &H80000012
    m_sCaption = Ambient.DisplayName
    m_iState = 1
    m_oBackColor = ConvertSystemColor(defBackColor)
    m_oBorderColor = ConvertSystemColor(defBorderColor)
    Set m_stdFont = Ambient.Font
    m_oRadiusColor = vb3DHighlight
    Value = False
    DrawCheckBox m_iState, m_bChecked
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeySpace Then
        DrawCheckBox m_iState, m_bChecked
    End If

End Sub

Private Sub UserControl_LostFocus()

    m_bFocus = False
    If m_bChecked Then
        DrawCheckBox 1, True
    Else
        DrawCheckBox 1, m_bChecked
    End If

End Sub

Private Sub UserControl_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)

    If Button = vbLeftButton Then
        m_bChecked = Not m_bChecked
        If m_bChecked Then
            DrawCheckBox 1, True
        Else
            DrawCheckBox 1, m_bChecked
        End If
        RaiseEvent Click
    End If

End Sub

Private Sub UserControl_Paint()

    DrawCheckBox m_iState, m_bChecked

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        Set m_stdFont = .ReadProperty("Font", Ambient.Font)
        BackColor = .ReadProperty("BackColor", ConvertSystemColor(defBackColor))
        BorderColor = .ReadProperty("BorderColor", ConvertSystemColor(defBorderColor))
        m_sCaption = .ReadProperty("Caption", Ambient.DisplayName)
        Enabled = .ReadProperty("Enabled", True)
        FocusColor = .ReadProperty("FocusColor", &H80000012)
        Alignment = .ReadProperty("Alignment", 0)
        ForeColor = .ReadProperty("ForeColor", &H80000012)
        UserControl.MousePointer = .ReadProperty("MousePointer", vbDefault)
        Set UserControl.MouseIcon = .ReadProperty("MouseIcon", Nothing)
        Value = .ReadProperty("Value", False)
    End With

    SetAccessKeys
    If Ambient.UserMode Then
        m_bTrack = True
        m_bTrackUser32 = IsFunctionExported("TrackMouseEvent", "User32")
        If Not m_bTrackUser32 Then
            If Not IsFunctionExported("_TrackMouseEvent", "Comctl32") Then
                m_bTrack = False
            End If
        End If
        If m_bTrack Then
            With picButton
                sc_Subclass .hWnd
                sc_AddMsg .hWnd, WM_MOUSEMOVE
                sc_AddMsg .hWnd, WM_MOUSELEAVE
            End With
            With UserControl
                sc_Subclass .hWnd
                sc_AddMsg .hWnd, WM_MOUSEMOVE
                sc_AddMsg .hWnd, WM_MOUSELEAVE
            End With
        End If
    End If

End Sub

Private Sub UserControl_Resize()

    If Not Ambient.UserMode Then
        DrawCheckBox m_iState, m_bChecked
    End If

End Sub

Private Sub UserControl_Terminate()

    On Error GoTo Catch

    sc_Terminate

Catch:

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        .WriteProperty "Alignment", m_eAlign, 0
        .WriteProperty "BackColor", m_oBackColor, ConvertSystemColor(defBackColor)
        .WriteProperty "BorderColor", m_oBorderColor, ConvertSystemColor(defBorderColor)
        .WriteProperty "Caption", m_sCaption, Ambient.DisplayName
        .WriteProperty "Enabled", m_bEnabled, True
        .WriteProperty "FocusColor", m_oFocusColor, &H80000012
        .WriteProperty "Font", m_stdFont, Ambient.Font
        .WriteProperty "ForeColor", m_oForeColor, &H80000012
        .WriteProperty "MousePointer", MousePointer, vbDefault
        .WriteProperty "MouseIcon", MouseIcon, Nothing
        .WriteProperty "Value", m_bChecked, False
    End With

End Sub

Public Property Get Value() As Boolean

    Value = m_bChecked

End Property

Public Property Let Value(ByVal lChecked As Boolean)

    m_bChecked = lChecked
    PropertyChanged "Value"
    DrawCheckBox m_iState, m_bChecked

End Property

'-The following routines are exclusively for the sc_ subclass routines----------------------------

Private Sub zAddMsg(ByVal uMsg As Long, _
                    ByVal nTable As Long)

Dim nCount      As Long
Dim nBase       As Long
Dim i           As Long

    nBase = z_ScMem                                                         'Remember z_ScMem so that we can restore its value on exit
    z_ScMem = zData(nTable)                                                 'Map zData() to the specified table

    If uMsg = ALL_MESSAGES Then                                             'If ALL_MESSAGES are being added to the table...
        nCount = ALL_MESSAGES                                               'Set the table entry count to ALL_MESSAGES
    Else
        nCount = zData(0)                                                   'Get the current table entry count
        If nCount >= MSG_ENTRIES Then                                       'Check for message table overflow
            zError "zAddMsg", "Message table overflow. Either increase the value" & _
            "of Const MSG_ENTRIES or use ALL_MESSAGES instead of specific message values"
            GoTo Bail
        End If

        For i = 1 To nCount                                                 'Loop through the table entries
            If zData(i) = 0 Then                                            'If the element is free...
                zData(i) = uMsg                                             'Use this element
                GoTo Bail
            ElseIf zData(i) = uMsg Then                                     'If the message is already in the table...
                GoTo Bail
            End If
        Next i                                                              'Next message table entry

        nCount = i                                                          'On drop through: i = nCount + 1, the new table entry count
        zData(nCount) = uMsg                                                'Store the message in the appended table entry
    End If

    zData(0) = nCount                                                       'Store the new table entry count
Bail:
    z_ScMem = nBase                                                         'Restore the value of z_ScMem

End Sub

Private Function zAddressOf(ByVal oCallback As Object, _
                            ByVal nOrdinal As Long) As Long

Dim bSub        As Byte
Dim bVal        As Byte
Dim nAddr       As Long
Dim i           As Long
Dim j           As Long


    RtlMoveMemory VarPtr(nAddr), ObjPtr(oCallback), 4                       'Get the address of the callback object's instance
    If Not zProbe(nAddr + &H1C, i, bSub) Then                               'Probe for a Class method
        If Not zProbe(nAddr + &H6F8, i, bSub) Then                          'Probe for a Form method
            If Not zProbe(nAddr + &H7A4, i, bSub) Then                      'Probe for a UserControl method
                Exit Function
            End If
        End If
    End If

    i = i + 4                                                               'Bump to the next entry
    j = i + 1024                                                            'Set a reasonable limit, scan 256 vTable entries
    Do While i < j
        RtlMoveMemory VarPtr(nAddr), i, 4                                   'Get the address stored in this vTable entry

        If IsBadCodePtr(nAddr) Then                                         'Is the entry an invalid code address?
            RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4         'Return the specified vTable entry address
            Exit Do                                                         'Bad method signature, quit loop
        End If

        RtlMoveMemory VarPtr(bVal), nAddr, 1                                'Get the byte pointed to by the vTable entry
        If Not bVal = bSub Then                                             'If the byte doesn't match the expected value...
            RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4         'Return the specified vTable entry address
            Exit Do                                                         'Bad method signature, quit loop
        End If

        i = i + 4                                                           'Next vTable entry
    Loop

End Function

Private Property Get zData(ByVal nIndex As Long) As Long

    RtlMoveMemory VarPtr(zData), z_ScMem + (nIndex * 4), 4

End Property

Private Property Let zData(ByVal nIndex As Long, _
                           ByVal nValue As Long)

    RtlMoveMemory z_ScMem + (nIndex * 4), VarPtr(nValue), 4

End Property

Private Sub zDelMsg(ByVal uMsg As Long, _
                    ByVal nTable As Long)

Dim nCount      As Long
Dim nBase       As Long
Dim i           As Long

    nBase = z_ScMem                                             'Remember z_ScMem so that we can restore its value on exit
    z_ScMem = zData(nTable)                                     'Map zData() to the specified table

    If uMsg = ALL_MESSAGES Then                                 'If ALL_MESSAGES are being deleted from the table...
        zData(0) = 0                                            'Zero the table entry count
    Else
        nCount = zData(0)                                       'Get the table entry count
        For i = 1 To nCount                                     'Loop through the table entries
            If zData(i) = uMsg Then                             'If the message is found...
                zData(i) = 0                                    'Null the msg value -- also frees the element for re-use
                GoTo Bail
            End If
        Next i                                                  'Next message table entry
        zError "zDelMsg", "Message &H" & Hex$(uMsg) & " not found in table"
    End If

Bail:
    z_ScMem = nBase                                             'Restore the value of z_ScMem

End Sub

'Error handler
Private Sub zError(ByVal sRoutine As String, _
                   ByVal sMsg As String)

    App.LogEvent TypeName(Me) & "." & sRoutine & ": " & sMsg, vbLogEventTypeError
    MsgBox sMsg & ".", vbExclamation + vbApplicationModal, "Error in " & TypeName(Me) & "." & sRoutine

End Sub

Private Function zFnAddr(ByVal sDLL As String, _
                         ByVal sProc As String) As Long

    zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)
    'Get the specified procedure address
    Debug.Assert zFnAddr

End Function

Private Function zMap_hWnd(ByVal lng_hWnd As Long) As Long

    If z_Funk Is Nothing Then                           'Ensure that subclassing has been started
        zError "zMap_hWnd", "Subclassing hasn't been started"
    Else
        On Error GoTo Catch                             'Catch unsubclassed window handles
        z_ScMem = z_Funk("h" & lng_hWnd)                'Get the thunk address
        zMap_hWnd = z_ScMem
    End If

Exit Function                                           'Exit returning the thunk address

Catch:
    zError "zMap_hWnd", "Window handle isn't subclassed"

End Function

'Probe at the specified start address for a method signature
Private Function zProbe(ByVal nStart As Long, _
                        ByRef nMethod As Long, _
                        ByRef bSub As Byte) As Boolean

Dim bVal        As Byte
Dim nAddr       As Long
Dim nLimit      As Long
Dim nEntry      As Long


    nAddr = nStart                                          'Start address
    nLimit = nAddr + 32                                     'Probe eight entries
    Do While nAddr < nLimit                                 'While we've not reached our probe depth
        RtlMoveMemory VarPtr(nEntry), nAddr, 4              'Get the vTable entry
        If Not nEntry = 0 Then                              'If not an implemented interface
            RtlMoveMemory VarPtr(bVal), nEntry, 1           'Get the value pointed at by the vTable entry
            If bVal = &H33 Or bVal = &HE9 Then              'Check for a native or pcode method signature
                nMethod = nAddr
                                                            'Store the vTable entry
                bSub = bVal                                 'Store the found method signature
                zProbe = True                               'Indicate success
                Exit Function                               'Return
            End If
        End If
        nAddr = nAddr + 4                                   'Next vTable entry
    Loop

End Function



Private Sub sc_AddMsg(ByVal lng_hWnd As Long, _
                      ByVal uMsg As Long, _
                      Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)

    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then       'Ensure that the thunk hasn't already released its memory
        If When And MSG_BEFORE Then                     'If the message is to be added to the before original WndProc table...
            zAddMsg uMsg, IDX_BTABLE                    'Add the message to the before table
        End If
        If When And MSG_AFTER Then                      'If message is to be added to the after original WndProc table...
            zAddMsg uMsg, IDX_ATABLE                    'Add the message to the after table
        End If
    End If

End Sub

'Call the original WndProc
Private Function sc_CallOrigWndProc(ByVal lng_hWnd As Long, _
                                    ByVal uMsg As Long, _
                                    ByVal wParam As Long, _
                                    ByVal lParam As Long) As Long

    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then
        'Ensure that the thunk hasn't already released its memory
        sc_CallOrigWndProc = CallWindowProcA(zData(IDX_WNDPROC), lng_hWnd, uMsg, wParam, lParam)
    End If

End Function

Private Sub sc_DelMsg(ByVal lng_hWnd As Long, _
                      ByVal uMsg As Long, _
                      Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)

    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then           'Ensure that the thunk hasn't already released its memory
        If When And MSG_BEFORE Then                         'If the message is to be deleted from the before original WndProc table...
            zDelMsg uMsg, IDX_BTABLE                        'Delete the message from the before table
        End If
        If When And MSG_AFTER Then                          'If the message is to be deleted from the after original WndProc table...
            zDelMsg uMsg, IDX_ATABLE                        'Delete the message from the after table
        End If
    End If

End Sub

Private Property Get sc_lParamUser(ByVal lng_hWnd As Long) As Long

    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then           'Ensure that the thunk hasn't already released its memory
        sc_lParamUser = zData(IDX_PARM_USER)                'Get the lParamUser callback parameter
    End If

End Property

Private Property Let sc_lParamUser(ByVal lng_hWnd As Long, _
                                   ByVal NewValue As Long)

    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then           'Ensure that the thunk hasn't already released its memory
        zData(IDX_PARM_USER) = NewValue                     'Set the lParamUser callback parameter
    End If

End Property

'-SelfSub code------------------------------------------------------------------------------------
'*************************************************************************************************
'* lng_hWnd   - Handle of the window to subclass
'* lParamUser - Optional, user-defined callback parameter
'* nOrdinal   - Optional, ordinal index of the callback procedure. 1 = last private method, 2 = second last private method, etc.
'* oCallback  - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
'* bIdeSafety - Optional, enable/disable IDE safety measures. NB: you should really only disable IDE safety in a UserControl for design-time subclassing
'*************************************************************************************************
Private Function sc_Subclass(ByVal lng_hWnd As Long, _
                             Optional ByVal lParamUser As Long = 0, _
                             Optional ByVal nOrdinal As Long = 1, _
                             Optional ByVal oCallback As Object = Nothing, _
                             Optional ByVal bIdeSafety As Boolean = True) As Boolean

Const CODE_LEN     As Long = 260                                'Thunk length in bytes
Const MEM_LEN      As Long = CODE_LEN + (8 * (MSG_ENTRIES + 1))
Const PAGE_RWX     As Long = &H40&                              'Allocate executable memory
Const MEM_COMMIT   As Long = &H1000&                            'Commit allocated memory
Const MEM_RELEASE  As Long = &H8000&                            'Release allocated memory flag
Const IDX_EBMODE   As Long = 3                                  'Thunk data index of the EbMode function address
Const IDX_CWP      As Long = 4                                  'Thunk data index of the CallWindowProc function address
Const IDX_SWL      As Long = 5                                  'Thunk data index of the SetWindowsLong function address
Const IDX_FREE     As Long = 6                                  'Thunk data index of the VirtualFree function address
Const IDX_BADPTR   As Long = 7                                  'Thunk data index of the IsBadCodePtr function address
Const IDX_OWNER    As Long = 8                                  'Thunk data index of the Owner object's vTable address
Const IDX_CALLBACK As Long = 10                                 'Thunk data index of the callback method address
Const IDX_EBX      As Long = 16                                 'Thunk code patch index of the thunk data
Const SUB_NAME     As String = "sc_Subclass"                    'This routine's name
Dim nAddr          As Long
Dim nID            As Long
Dim nMyID          As Long


    If IsWindow(lng_hWnd) = 0 Then                              'Ensure the window handle is valid
        zError SUB_NAME, "Invalid window handle"
    Else
        nMyID = GetCurrentProcessId                             'Get this process's ID
        GetWindowThreadProcessId lng_hWnd, nID                  'Get the process ID associated with the window handle
        If Not nID = nMyID Then                                 'Ensure that the window handle doesn't belong to another process
            zError SUB_NAME, "Window handle belongs to another process"
            Exit Function
        End If

        If oCallback Is Nothing Then                            'If the user hasn't specified the callback owner
            Set oCallback = Me                                  'Then it is me
        End If

        nAddr = zAddressOf(oCallback, nOrdinal)                 'Get the address of the specified ordinal method
        If nAddr = 0 Then                                       'Ensure that we've found the ordinal method
            zError SUB_NAME, "Callback method not found"
            Exit Function
        End If

        If z_Funk Is Nothing Then                               'If this is the first time through, do the one-time initialization
            Set z_Funk = New Collection                         'Create the hWnd/thunk-address collection
            z_Sc(14) = &HD231C031
            z_Sc(15) = &HBBE58960
            z_Sc(17) = &H4339F631
            z_Sc(18) = &H4A21750C
            z_Sc(19) = &HE82C7B8B
            z_Sc(20) = &H74&
            z_Sc(21) = &H75147539
            z_Sc(22) = &H21E80F
            z_Sc(23) = &HD2310000
            z_Sc(24) = &HE8307B8B
            z_Sc(25) = &H60&
            z_Sc(26) = &H10C261
            z_Sc(27) = &H830C53FF
            z_Sc(28) = &HD77401F8
            z_Sc(29) = &H2874C085
            z_Sc(30) = &H2E8&
            z_Sc(31) = &HFFE9EB00
            z_Sc(32) = &H75FF3075
            z_Sc(33) = &H2875FF2C
            z_Sc(34) = &HFF2475FF
            z_Sc(35) = &H3FF2473
            z_Sc(36) = &H891053FF
            z_Sc(37) = &HBFF1C45
            z_Sc(38) = &H73396775
            z_Sc(39) = &H58627404
            z_Sc(40) = &H6A2473FF
            z_Sc(41) = &H873FFFC
            z_Sc(42) = &H891453FF
            z_Sc(43) = &H7589285D
            z_Sc(44) = &H3045C72C
            z_Sc(45) = &H8000&
            z_Sc(46) = &H8920458B
            z_Sc(47) = &H4589145D
            z_Sc(48) = &HC4836124
            z_Sc(49) = &H1862FF04
            z_Sc(50) = &H35E30F8B
            z_Sc(51) = &HA78C985
            z_Sc(52) = &H8B04C783
            z_Sc(53) = &HAFF22845
            z_Sc(54) = &H73FF2775
            z_Sc(55) = &H1C53FF28
            z_Sc(56) = &H438D1F75
            z_Sc(57) = &H144D8D34
            z_Sc(58) = &H1C458D50
            z_Sc(59) = &HFF3075FF
            z_Sc(60) = &H75FF2C75
            z_Sc(61) = &H873FF28
            z_Sc(62) = &HFF525150
            z_Sc(63) = &H53FF2073
            z_Sc(64) = &HC328&

            z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcA")    'Store CallWindowProc function address in the thunk data
            z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongA")     'Store the SetWindowLong function address in the thunk data
            z_Sc(IDX_FREE) = zFnAddr("kernel32", "VirtualFree")     'Store the VirtualFree function address in the thunk data
            z_Sc(IDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr")  'Store the IsBadCodePtr function address in the thunk data
        End If

        z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX)    'Allocate executable memory

        If z_ScMem <> 0 Then                                        'Ensure the allocation succeeded
            On Error GoTo CatchDoubleSub                            'Catch double subclassing
            z_Funk.Add z_ScMem, "h" & lng_hWnd                      'Add the hWnd/thunk-address to the collection
            On Error GoTo 0

            If bIdeSafety Then                                      'If the user wants IDE protection
                z_Sc(IDX_EBMODE) = zFnAddr("vba6", "EbMode")        'Store the EbMode function address in the thunk data
            End If

            z_Sc(IDX_EBX) = z_ScMem                                 'Patch the thunk data address
            z_Sc(IDX_HWND) = lng_hWnd                               'Store the window handle in the thunk data
            z_Sc(IDX_BTABLE) = z_ScMem + CODE_LEN                   'Store the address of the before table in the thunk data
            z_Sc(IDX_ATABLE) = z_ScMem + CODE_LEN + ((MSG_ENTRIES + 1) * 4) 'Store the address of the after table in the thunk data
            z_Sc(IDX_OWNER) = ObjPtr(oCallback)                     'Store the callback owner's object address in the thunk data
            z_Sc(IDX_CALLBACK) = nAddr                              'Store the callback address in the thunk data
            z_Sc(IDX_PARM_USER) = lParamUser                        'Store the lParamUser callback parameter in the thunk data

            nAddr = SetWindowLongA(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF)    'Set the new WndProc, return the address of the original WndProc
            If nAddr = 0 Then                                       'Ensure the new WndProc was set correctly
                zError SUB_NAME, "SetWindowLong failed, error #" & Err.LastDllError
                GoTo ReleaseMemory
            End If

            z_Sc(IDX_WNDPROC) = nAddr                               'Store the original WndProc address in the thunk data
            RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), CODE_LEN        'Copy the thunk code/data to the allocated memory
            sc_Subclass = True                                      'Indicate success
        Else
            zError SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError
        End If

        Exit Function                                               'Exit sc_Subclass

CatchDoubleSub:
        zError SUB_NAME, "Window handle is already subclassed"

ReleaseMemory:
        VirtualFree z_ScMem, 0, MEM_RELEASE                         'sc_Subclass has failed after memory allocation, so release the memory
    End If

End Function

Private Sub sc_Terminate()

Dim i       As Long

    If Not (z_Funk Is Nothing) Then                                 'Ensure that subclassing has been started
        With z_Funk
            For i = .Count To 1 Step -1                             'Loop through the collection of window handles in reverse order
                z_ScMem = .Item(i)                                  'Get the thunk address
                If IsBadCodePtr(z_ScMem) = 0 Then                   'Ensure that the thunk hasn't already released its memory
                    sc_UnSubclass zData(IDX_HWND)                   'UnSubclass
                End If
            Next i                                                  'Next member of the collection
        End With
        Set z_Funk = Nothing                                        'Destroy the hWnd/thunk-address collection
    End If

End Sub

Private Sub sc_UnSubclass(ByVal lng_hWnd As Long)

    If z_Funk Is Nothing Then                                       'Ensure that subclassing has been started
        zError "sc_UnSubclass", "Window handle isn't subclassed"
    Else
        If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then               'Ensure that the thunk hasn't already released its memory
            zData(IDX_SHUTDOWN) = -1                                'Set the shutdown indicator
            zDelMsg ALL_MESSAGES, IDX_BTABLE                        'Delete all before messages
            zDelMsg ALL_MESSAGES, IDX_ATABLE                        'Delete all after messages
        End If
        z_Funk.Remove "h" & lng_hWnd                                'Remove the specified window handle from the collection
    End If

End Sub

'*************************************************************************************************
'* bBefore    - Indicates whether the callback is before or after the original WndProc. Usually
'*              you will know unless the callback for the uMsg value is specified as
'*              MSG_BEFORE_AFTER (both before and after the original WndProc).
'* bHandled   - In a before original WndProc callback, setting bHandled to True will prevent the
'*              message being passed to the original WndProc and (if set to do so) the after
'*              original WndProc callback.
'* lReturn    - WndProc return value. Set as per the MSDN documentation for the message value,
'*              and/or, in an after the original WndProc callback, act on the return value as set
'*              by the original WndProc.
'* lng_hWnd   - Window handle.
'* uMsg       - Message value.
'* wParam     - Message related data.
'* lParam     - Message related data.
'* lParamUser - User-defined callback parameter
'*************************************************************************************************
'-Subclass callback, usually ordinal #1, the last method in this source file----------------------
Private Sub zWndProc1(ByVal bBefore As Boolean, _
                      ByRef bHandled As Boolean, _
                      ByRef lReturn As Long, _
                      ByVal lng_hWnd As Long, _
                      ByVal uMsg As Long, _
                      ByVal wParam As Long, _
                      ByVal lParam As Long, _
                      ByRef lParamUser As Long)

    Select Case uMsg
    Case WM_MOUSEMOVE
        If Not m_bInCtrl Then
            m_bInCtrl = True
            TrackMouseLeave lng_hWnd
            If m_bChecked Then
                DrawCheckBox 2, True
            Else
                DrawCheckBox 2, m_bChecked
            End If
            RaiseEvent MouseEnter
        End If

    Case WM_MOUSELEAVE
        m_bInCtrl = False
        RaiseEvent MouseLeave
        If m_bChecked Then
            DrawCheckBox 1, True
        Else
            DrawCheckBox 1, m_bChecked
        End If
    End Select

End Sub

