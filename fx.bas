Attribute VB_Name = "fx"
Option Explicit

'Ê¹ÓÃ±¾³ÌÐòÓ¦×ñÊØµÄÌõ¿î
'
'1¡¢ÔÚ·¢²¼±¾³ÌÐòÊ±£¬frmAbout´°Ìå±ØÐë±£ÁôÔ­×÷ÕßºÍÊý¾ÝÌá¹©ÕßÐÅÏ¢£»
'2¡¢Ô­×÷Õß²»³Ðµ£Ê¹ÓÃ±¾´úÂëËù·¢Õ¹³öÀ´µÄÈÎºÎºó¹û£»
'
'
'   ========================================================================
'    FX Module Information
'   ========================================================================
'
'    File Version:      1.03
'    Description:       FX.DLL Procedure Declarations, Constant and Tag
'                       Definitions
'    Copyright:         Copyright © Martins Skujenieks 2003
'    Product Name:      FX.DLL
'    Product Version:   1.03
'
'
'   ========================================================================
'    End User License Agreement (EULA)
'   ========================================================================
'
'    This product is provided "as is", with no guarantee of completeness or
'    accuracy and without warranty of any kind, express or implied.
'
'    In no event will developer be liable for damages of any kind that may
'    be incurred with your hardware, peripherals or software programs.
'
'    This product and all of its parts may not be copied, emulated, cloned,
'    rented, leased, sold, reproduced, modified, decompiled, disassembled,
'    otherwise reverse engineered, republished, uploaded, posted,
'    transmitted or distributed in any way, without prior written consent
'    of the developer.
'
'
'   ========================================================================
'    Contact Developer
'   ========================================================================
'
'    Website:           http://www.exe.times.lv
'    E-Mail:            martins_s@mail.teliamtc.lv
'
'   ========================================================================





    '/* Ternary Raster Operations */
    Public Const SRCCOPY = &HCC0020
    Public Const SRCPAINT = &HEE0086
    Public Const SRCAND = &H8800C6
    Public Const SRCINVERT = &H660046
    Public Const SRCERASE = &H440328
    Public Const NOTSRCCOPY = &H330008
    Public Const NOTSRCERASE = &H1100A6
    Public Const MERGECOPY = &HC000CA
    Public Const MERGEPAINT = &HBB0226
    Public Const PATCOPY = &HF00021
    Public Const PATPAINT = &HFB0A09
    Public Const PATINVERT = &H5A0049
    Public Const DSTINVERT = &H550009
    Public Const BLACKNESS = &H42
    Public Const WHITENESS = &HFF0062


    '/* Text Alignment Options */
    Public Const TA_NOUPDATECP = 0
    Public Const TA_UPDATECP = 1
    Public Const TA_LEFT = 0
    Public Const TA_RIGHT = 2
    Public Const TA_CENTER = 6
    Public Const TA_TOP = 0
    Public Const TA_BOTTOM = 8
    Public Const TA_BASELINE = 24
    Public Const TA_RTLREADING = 256
    Public Const TA_MASK = (TA_BASELINE + TA_CENTER + TA_UPDATECP + TA_RTLREADING)
    
    
    '/* Vertical Text Alignment Options */
    Public Const VTA_BASELINE = TA_BASELINE
    Public Const VTA_LEFT = TA_BOTTOM
    Public Const VTA_RIGHT = TA_TOP
    Public Const VTA_CENTER = TA_CENTER
    Public Const VTA_BOTTOM = TA_RIGHT
    Public Const VTA_TOP = TA_LEFT


    '/* struct tagPOINT */
    Public Type POINT
        X       As Long
        Y       As Long
    End Type

    
    '/* struct tagRECT */
    Public Type RECT
        Left    As Long
        Top     As Long
        Right   As Long
        Bottom  As Long
    End Type


    '/* Function prototypes */
    Public Declare Function fxRender Lib "fx.dll" (ByVal DestDC As Long, ByVal CenterX As Long, ByVal CenterY As Long, ByVal SrcDC As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal Blending As Long, ByVal Angle As Double, ByVal Zoom As Double, ByVal TransparentColor As Long, Optional ByVal Flags As Long = 0) As Long

    'Demo variabls:
    Public fxIndex As Long
    Public fxX As Long
    Public fxY As Long
    Public fxColor As Long
    Public fxMaskColor As Long
    
    'Handle of font for DrawText:
    Public hFont As Long
 



