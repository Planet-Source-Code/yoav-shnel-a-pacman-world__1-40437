Attribute VB_Name = "modODMenus"
' ********************************************************************************
'
'  Project: Menu bitmap module
'  Author:  G. D. Sever    aka "The Hand" (thehand@elitevb.com)
'
' ********************************************************************************
'
'  Description:
'
'    This little module allows you to easily add in bitmaps and textures to your
'    menus. And if you laugh at the "little" part, you should check out the source
'    code for the OD menu solution at www.vbaccelerator.com - this is nothing
'    compared to that beast!! Also, nearly 50% of the length of this module is pure
'    comments to help you understand what is going on. It would be much shorter if
'    it were pure code.
'
'    Hopefully this module will show you that if you want something bad enough,
'    its possible to realize your desires thru sheer willpower.
'
'  Terms of use:
'    Ahhh.... the sticky part. Here's the deal: Use this code in your projects.
'    For the love of god, use it and make your apps prettier. We need to start
'    showing people that VB has every bit as much potential as the C++ apps out
'    there.
'
'    Change the code if you want to! Don't like the way we use a resource file to
'    store the images? CHANGE IT! Want to have different fonts in the menus with
'    hideous and ugly colors? GO NUTS! Gradient effects for selected items? ITS
'    ALL YOU!
'
'    HOWEVER: If you happen to post this code or a portion of it somewhere, give
'    us credit for the parts we are responsible for. Saying that we was
'    "an inspiration" for the code when 70% of it was cut & paste from here is NOT
'    adequate to us. Put a 1 or two line comment at the beginning of the subs and
'    functions you use and name us as the authors! And let us tell you something...
'    Doing a global "Replace all" on a couple of variable names and function names
'    does not suddenly make the code something you wrote.
'
'    I really do not think we're asking too much - it all boils down to one simple
'    principle: Give credit where its due. We do wherever necessary and possible.
'
'    That being said, API declarations are almost 100% from the ALLAPI.NET guide
'    which is a fantastic resource. Go out and download it immediately from
'    www.allapi.net
'
' ********************************************************************************
'      Visit http://www.elitevb.com for more high-powered solutions!!
' ********************************************************************************

Option Explicit

' ********************************************************************************
'     User accessible variables - used to change the look of the menus
' ********************************************************************************

' Use the custom background in the menus
Public gbCustomTexture As Boolean
' Texture to be used in the menu background
Public gMenuBG As StdPicture
' Collection of bitmaps to be used in the menu drawing
Public gMenuBmps As Collection
' Percent translucent
Public gMenuFontColor As OLE_COLOR
' Menu font
Public gMenuFont As StdFont
' Individual menu fonts
Public gCustomFonts As Collection
' Whether to use custom fonts or not
Public gUseCustomFonts As Boolean

' ********************************************************************************
'     Couple of things the module uses to keep track of stuff
' ********************************************************************************

' Default menu height
Private Const gMnuHeight = 20
' Menu height
Private Const gMnuWidth = 20
' Menu item captions - stored in an array
Private gMenuCaps As Collection
' ID for the last top level menu
Private gLastTopMenuID As Long
' ID numbers for the top level menus - For some blasphemous reason I always start with 666
Private Const gStartTopID = 666&

' ********************************************************************************
'     Here come the API declarations... and whoa nelly, are there alot of them
' ********************************************************************************

' A rectangle. DUH!
Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

' Get the windows's dimensions using its handle - used when drawing a texture on the
'  top level menus.
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
' Used to get the system's 3D object border width - subtracted from the overall value
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Const SM_CXBORDER = 5

' Used to get caption widths & heights
Private Const SM_CYMENUSIZE = 55
' Used by SystemParametersInfo to get the non client metrics
Private Const SPI_GETNONCLIENTMETRICS = 41
' Logical font type used to size a font with CreateFont
Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(1 To 32) As Byte
End Type
' Duh! Non client metrics!
Private Type NONCLIENTMETRICS
    cbSize As Long
    iBorderWidth As Long
    iScrollWidth As Long
    iScrollHeight As Long
    iCaptionWidth As Long
    iCaptionHeight As Long
    lfCaptionFont As LOGFONT
    iSMCaptionWidth As Long
    iSMCaptionHeight As Long
    lfSMCaptionFont As LOGFONT
    iMenuWidth As Long
    iMenuHeight As Long
    lfMenuFont As LOGFONT
    lfStatusFont As LOGFONT
    lfMessageFont As LOGFONT
End Type
' Used to get various system parameters and settings
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
' A POINT! YEAH! You know... x and y?
Private Type POINTAPI
    x As Long
    Y As Long
End Type
' Gets the width & height of text in a DC using that DC's currently selected font
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As POINTAPI) As Long
' Creates a new... uh.. font
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal I As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
' Gets display info
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Const LOGPIXELSY = 90

' Type that says how big & wide the menu items will be
Private Type MEASUREITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    itemWidth As Long
    itemHeight As Long
    ItemData As Long
End Type

' Structure used when WM_DRAWITEM is passed that says
'  which part of the form/menu/etc will be worked on
Private Type DRAWITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    itemAction As Long
    itemState As Long
    hwndItem As Long
    hdc As Long
    rcItem As RECT
    ItemData As Long
End Type

' Used to copy information from pointers into structures
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

' ******************************************************************************
' MENU DECLARES - Used to get / set information for menu items
' ******************************************************************************
' Used to set the "Owner drawn" functionality of the menu item
Private Type MENUITEMINFO
  cbSize As Long
  fMask As Long
  fType As Long
  fState As Long
  wid As Long
  hSubMenu As Long
  hbmpChecked As Long
  hbmpUnchecked As Long
  dwItemData As Long
  dwTypeData As String
  cch As Long
End Type
Private Const MIIM_STATE = &H1&
Private Const MIIM_ID = &H2&
Private Const MIIM_SUBMENU = &H4&
Private Const MIIM_CHECKMARKS = &H8&
Private Const MIIM_TYPE = &H10&
Private Const MIIM_DATA = &H20&
' Used to set the "Owner drawn" functionality of the menu item
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Long) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long

Private Const MF_DISABLED = &H2
Private Const MF_CHECKED = &H8
Private Const MF_MENUBREAK = &H40&
Private Const MF_BYCOMMAND = &H0&
Private Const MF_BYPOSITION = &H400&
Private Const MF_BITMAP = &H4&
Private Const MF_OWNERDRAW = &H100&
Private Const MF_SEPARATOR = &H800&

' These are used to find a menu item on the form
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long

Private Const ODS_SELECTED = &H1  ' whether an item is highlighted or not.
Private Const ODS_GRAYED = &H2    ' * The Hand shrugs
Private Const ODS_DISABLED = &H4  ' Whether an item is disabled or not
Private Const ODS_CHECKED = &H8   ' Whether an item is "checked" or not
Private Const ODS_HOTTRACK = &H40 ' Whether the menubar is hottracking

' Used to make sure we're not trying to modify the control box menus.
'  why? Because it will cause a great big GPF error.
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
' Used to get the information for the menu item
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpmii As MENUITEMINFO) As Long
' Used to redraw the menu bar when we switch between textured and non-textured
Public Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long

' ******************************************************************************
' SUBCLASSING ROUTINES - All messages are sent directly to the parent form
' ******************************************************************************
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const GWL_WNDPROC = (-4)
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)

Private Const WM_COMMAND = &H111
Private Const WM_SYSCOMMAND = &H112
Private Const WM_ERASEBKGND = &H14
Private Const WM_NCPAINT = &H85
Private Const WM_NCACTIVATE = &H86
Private Const WM_DRAWITEM = &H2B     ' Used to actually draw the owner-drawn item
Private Const WM_MEASUREITEM = &H2C  ' Used to return the size of the area in which
                                    '  we will be drawing. This would be extremely
                                    '  useful if we wanted to create a custom seperator
Private Const WM_INITMENUPOPUP = &H117 ' Used to grab the menus and make them owner-drawn

' ******************************************************************************
' GRAPHICS DECLARES (GDI32 & USER32) - for drawing edge and pictures and text
' ******************************************************************************
' The following are used to draw the upraised square on the menu item:
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_SUNKENOUTER = &H2
Private Const BDR_SUNKENINNER = &H8
Private Const BF_BOTTOM = &H8
Private Const BF_LEFT = &H1
Private Const BF_RIGHT = &H4
Private Const BF_TOP = &H2
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
'Used to paint solid colors into the menu item's area
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
'Used to move the origin of our pattern brush so the background is painted correctly
Private Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lppt As POINTAPI) As Long
'Used to manipulate the GDI32 objects we create / use
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
'Used to paint the pictures into our menu items
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
'Used to print the item text
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long

' Used to store the original procedure addresses by individual window handle
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long

' Message used for cleaning up resources when menu loop is exited
Private Const WM_ENTERMENULOOP = &H211
Private Const WM_EXITMENULOOP As Long = &H212
Private Const WM_ACTIVATEAPP As Long = &H1C

' Used to get width and height dimensions for a bitmap
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

' Used for some hwnd/hDC trickiness
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function WindowFromDC Lib "user32" (ByVal hdc As Long) As Long
Private Declare Function GetUpdateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Private Const WM_MDIACTIVATE = &H222
' *************************
'   For disabled items:
' *************************
Const DST_PREFIXTEXT = &H2
Const DST_BITMAP = &H4
Const DSS_NORMAL = &H0
Const DSS_DISABLED = &H20
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal flags As Long) As Long
Private Declare Function DrawStateText Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lString As String, ByVal wParam As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal flags As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Const ODT_MENU = 1

Public Sub startODMenus(aControl As Object, bMenuBar As Boolean, Optional debugMode As Boolean)

    Dim origProc As Long
    Dim aVer As Long
    Dim aPt As POINTAPI
    Dim hwnd As Long
        
    If gMenuBmps Is Nothing Then Set gMenuBmps = New Collection
    If gMenuCaps Is Nothing Then Set gMenuCaps = New Collection
    If gCustomFonts Is Nothing Then Set gCustomFonts = New Collection
    'If TypeOf aControl Is Toolbar Then
    '    ' We want the child of the toolbar wrapper
    '    hwnd = FindWindowEx(aControl.hwnd, ByVal 0&, "msvb_lib_toolbar", vbNullString)
    If TypeOf aControl Is Form Then
        ' If the user specifies to make the menubar ownerdrawn, do it!
        If bMenuBar Then makeTopMenusOD aControl
        hwnd = aControl.hwnd
    End If
    
    ' Start the subclassing
    origProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf ODWindowProc)
    
    ' Store the original process address in Windows' catalog using the form's handle
    SetProp hwnd, "ODMenuOrigProc", origProc
    
End Sub
Public Sub stopODMenus(aControl As Object)
    
    Dim origProc As Long
    Dim I As Long
    Dim abrush As Long
    Dim aStyle As Long
    Dim hwnd As Long
    
    'If TypeOf aControl Is Toolbar Then
    '    hwnd = FindWindowEx(aControl.hwnd, ByVal 0&, "msvb_lib_toolbar", vbNullString)
    If TypeOf aControl Is Form Then
        ' Get the original process address using the form's handle
        If Forms.Count = 1 Then
            'last form - unhook stuff and clear out the crap
            Set gMenuCaps = Nothing
            Set gMenuBmps = Nothing
            Set gCustomFonts = Nothing
        End If
        hwnd = aControl.hwnd
    End If
    
    origProc = GetProp(hwnd, "ODMenuOrigProc")
    ' Unsubclass the form by replacing the original process address
    SetWindowLong hwnd, GWL_WNDPROC, origProc
    ' Remove the property entry from the Windows' catalog
    RemoveProp hwnd, "ODMenuOrigProc"
    
End Sub

Public Sub makeTopMenusOD(aForm As Form)
    
    Dim hMenubar As Long
    Dim aMenu As Long
    Dim numTopMnus As Long
    Dim aMII As MENUITEMINFO
    Dim I As Long
    Dim sCap As String
    Dim aStart As Long
    
    ' Grab the form's menubar
    hMenubar = GetMenu(aForm.hwnd)
    ' Get the number of top-level menubar items
    numTopMnus = GetMenuItemCount(hMenubar)
    ' store the last ID number just for quick reference in the drawing routine
    gLastTopMenuID = gStartTopID + numTopMnus - 1
    aStart = IIf(aForm.WindowState = vbMaximized, 1, 0)
    For I = aStart To numTopMnus - 1
        ' initialize our menu item info structure to get data
        aMII.fMask = MIIM_DATA Or MIIM_ID Or MIIM_STATE Or MIIM_SUBMENU Or MIIM_TYPE
        aMII.cch = 127
        aMII.dwTypeData = String$(128, 0)
        aMII.cbSize = Len(aMII)
        ' Actually go get the menu item data
        GetMenuItemInfo hMenubar, I, True, aMII
        ' Save the captions in our memory collection
        On Error Resume Next
        sCap = VBA.left$(aMII.dwTypeData, aMII.cch)
        If sCap <> "" Then
            gMenuCaps.Remove "M" & CStr(gStartTopID + I) & "-" & aForm.hwnd
            gMenuCaps.Add sCap, "M" & CStr(gStartTopID + I) & "-" & aForm.hwnd
        End If
        On Error GoTo 0
        'Get the state of the menu item
        aMII.fMask = MIIM_STATE
        GetMenuItemInfo hMenubar, I, True, aMII
        ' Turn the menubar item into an owner-drawn one
        ModifyMenu hMenubar, I, MF_OWNERDRAW Or MF_BYPOSITION Or aMII.fState, gStartTopID + I, ByVal 2&
        'SetMenuItemInfo hMenubar, I, True, aMII
    Next I

End Sub

Private Function ODWindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Dim aMeas As MEASUREITEMSTRUCT
    Dim aMnuDim As POINTAPI
    Dim oldWndProc As Long
    Dim systemAnim As Boolean
    Dim aDIS As DRAWITEMSTRUCT
    Dim abrush As Long
    Dim aRect As RECT
    Dim aHWNDTmp As Long
    Dim sClassName As String
    Dim MII As MENUITEMINFO
    Dim mnuHwnd As Long
    Dim isSep As Boolean
    Dim prgPtr As Long
    
    oldWndProc = GetProp(hwnd, "ODMenuOrigProc")
    If uMsg = WM_DRAWITEM Then
        CopyMemory aDIS, ByVal lParam, Len(aDIS)
        If aDIS.CtlType = ODT_MENU Then
            'Use our custom drawing subroutine to draw the menu item
            drawMnuBitmap hwnd, lParam
            'Don't do any other processing
            ODWindowProc = False
        Else
            ODWindowProc = CallWindowProc(oldWndProc, hwnd, uMsg, wParam, lParam)
        End If
    ElseIf uMsg = WM_ERASEBKGND Then
        ' Figure out if this is a toolbar.
        aHWNDTmp = WindowFromDC(wParam)
        sClassName = String$(128, Chr$(0))
        GetClassName aHWNDTmp, sClassName, Len(sClassName)
        ' If it IS a toolbar and we have a custom menu background selected
        If InStr(sClassName, "toolbar") > 1 Then
            GetWindowRect aHWNDTmp, aRect
            aRect.right = aRect.right - aRect.left
            aRect.bottom = aRect.bottom - aRect.top
            aRect.top = 0
            aRect.left = 0
            If gbCustomTexture And Not (gMenuBG Is Nothing) Then
                abrush = CreatePatternBrush(gMenuBG.Handle)
            Else
                abrush = CreateSolidBrush(getAppColor("menubg"))
            End If
            FillRect wParam, aRect, abrush
            DeleteObject abrush
            ODWindowProc = True
        Else
            ODWindowProc = CallWindowProc(oldWndProc, hwnd, uMsg, wParam, lParam)
        End If
    ElseIf uMsg = WM_MEASUREITEM Then
        Dim aWid As Long
        'copy the information from pointer lparam to our structure
        CopyMemory aMeas, ByVal lParam, Len(aMeas)
        ' determine the width of our menu item
        getMenuDimensions hwnd, aMeas.itemID, aMeas.ItemData, aMnuDim
        aWid = aMnuDim.x
        ' if this item value is bigger than the previous one, store the bigger
        '  value. This allows the menu to be properly sized for the biggest item.
        'If aMeas.itemWidth < aWid Then aMeas.itemWidth = aWid
        aMeas.itemWidth = aWid
        ' Make each item either 6 or gMnuHeight pixels high. We check the
        '  item data value to determine whether it is a seperator or a
        '  regular menu (0 = seperator, 1 = normal menu)
        aMeas.itemHeight = IIf(isSep, 6, aMnuDim.Y)
        'Copy the structure back to the one located at pointer location
        CopyMemory ByVal lParam, aMeas, Len(aMeas)
        'Don't do any other processing
        ODWindowProc = False
    ElseIf uMsg = WM_INITMENUPOPUP Then
        Dim hSysMenu As Long
        Dim aForm As Form
        Dim isSysMenu As Boolean
        ' Make sure that we're not trying to set OD styles on the system menu...
        '  (you REALLY REALLY REALLY don't want to try to do that)
        isSysMenu = False
        For Each aForm In Forms
            hSysMenu = GetSystemMenu(aForm.hwnd, False)
            ' if its not the systemmenu then set all the styles to ownerdrawn and
            '  pop it open!
            isSysMenu = isSysMenu Or (wParam = hSysMenu)
        Next aForm
        ' Invoke whatever it was going to do
         ODWindowProc = CallWindowProc(oldWndProc, hwnd, uMsg, wParam, lParam)
        If Not isSysMenu Then setPopupStyleOD wParam, hwnd
    Else
        ' Invoke the default window procedure
        ODWindowProc = CallWindowProc(oldWndProc, hwnd, uMsg, wParam, lParam)
    End If
End Function

Private Sub setPopupStyleOD(aHwnd As Long, wndHwnd As Long, Optional anItemInd As Long)

    Dim I As Long
    Dim hSubMenu As Long
    Dim hSubMenuID As Long
    Dim MII As MENUITEMINFO
    Dim isSep As Boolean
    Dim done As Long
    Dim capStr As String
    Dim startInd As Long
    Dim endInd As Long
    Dim temp As Long
    Dim lNewStyle As Long
    
    ' Determine whether we are going set the ownerdrawn for one individual item
    '  or for a whole submenu.
    If anItemInd > 0 Then
        startInd = anItemInd
        endInd = anItemInd
    Else
        startInd = 0
        endInd = GetMenuItemCount(aHwnd) - 1
    End If
    
    ' Loop thru from startInd to endInd
    For I = startInd To endInd
        ' initialize our menu item info structure to get data
       ' MII.fMask = MIIM_DATA Or MIIM_ID Or MIIM_SUBMENU Or MIIM_STATE Or MIIM_TYPE
        MII.fMask = MIIM_TYPE Or MIIM_ID
        MII.cch = 127
        MII.dwTypeData = String$(128, 0)
        MII.cbSize = Len(MII)
        ' get the menu item information
        GetMenuItemInfo aHwnd, I, True, MII
        ' get the ID number for the menu item
        hSubMenuID = GetMenuItemID(aHwnd, I)
        ' determine if the item is a seperator or not
        isSep = ((MII.fType And MF_SEPARATOR) = MF_SEPARATOR)
        
        lNewStyle = MII.fType Or MF_OWNERDRAW Or 0& 'Or mii.fState
    
        ' trim extra null characters out of the caption
        If InStr(MII.dwTypeData, Chr(0)) > 0 Then
            capStr = left$(MII.dwTypeData, InStr(MII.dwTypeData, Chr(0)) - 1)
            ' Split menu item if first char is a pipe
            If left$(capStr, 1) = "|" Then
                capStr = right$(capStr, Len(capStr) - 1)
                lNewStyle = lNewStyle Or MF_MENUBREAK
            End If
        Else
            capStr = ""
        End If
        
        MII.fType = lNewStyle
        MII.fMask = MIIM_TYPE Or MIIM_ID 'Or MIIM_DATA
        SetMenuItemInfo aHwnd, I, True, MII
        
        ' provided there is a caption, store it. This is a weird requirement
        '  I ran into while testing out the OD menus in NT. For some reason,
        '  MII.dwTypeData doesn't always have the caption string... sometimes
        '  it disappears! That's why we can't just use the information in the
        '  DRAWITEMSTRUCT in our ownerdrawn procedure.
        
        If capStr <> "" Then
            On Error Resume Next
            ' Store an item unique to each WINDOW AND ITEM ID
            gMenuCaps.Remove "M" & CStr(MII.wid) & "-" & wndHwnd
            gMenuCaps.Add capStr, "M" & CStr(MII.wid) & "-" & wndHwnd
            On Error GoTo 0
        End If
            
    Next I
End Sub

Public Function getMenuDimensions(hwnd As Long, subItemID As Long, itemType As Long, aPt As POINTAPI) As Boolean

    Dim aPt2        As POINTAPI
    Dim aCap        As String
    Dim formDC      As Long
    Dim origFont    As Long
    Dim mnuFont     As Long
    Dim lPic        As StdPicture
    Dim aBmp        As BITMAP
    Dim lPicInds    As Variant
    Dim aWid        As Long
    
    ' Determine an estimate for the menu width based on the stored menu
    '  caption string.

    On Error Resume Next
    aCap = gMenuCaps("M" & CStr(subItemID) & "-" & hwnd)
    
    ' Replace tab character with a space.
    If InStr(aCap, vbTab) > 0 Then aCap = VBA.left$(aCap, InStr(aCap, vbTab) - 1) & " " & VBA.right$(aCap, Len(aCap) - InStr(aCap, vbTab))
    ' Get the form's DC using its handle
    formDC = GetDC(hwnd)
    ' Get the proper font size
    mnuFont = getFontForItem(aCap, formDC)
    ' select the font into the device context
    origFont = SelectObject(formDC, mnuFont)
    
    ' Get the text width using the form's DC as a reference
    GetTextExtentPoint32 formDC, aCap, Len(aCap), aPt2
    ' Replace the font with the original
    SelectObject formDC, origFont
    ' Release the form's DC back to itself
    ReleaseDC hwnd, formDC
    ' Delete the temporary font
    DeleteObject mnuFont
    
    Err.Clear
    ' Check and see if the bitmap is taller than the font
    aCap = gMenuCaps(CStr(subItemID))
    If InStr(aCap, vbTab) > 0 Then aCap = VBA.left$(aCap, InStr(aCap, vbTab) - 1)
    lPicInds = gMenuBmps(aCap)
    aWid = gMnuWidth
    If Err.Number = 0 Then
        Set lPic = LoadResPicture(lPicInds(0), vbResBitmap)
        GetObject lPic.Handle, Len(aBmp), aBmp
        DeleteObject lPic.Handle
    End If
    
    '  Make the width = text width plus 2 times the menu height (and picture width ;) )
    aPt.x = (aPt2.x + IIf(itemType = 2, 0, aWid) + 6)
    '  Calculate the height
    aPt.Y = IIf(aBmp.bmHeight > aPt2.Y, aBmp.bmHeight, aPt2.Y) + 6
    On Error GoTo 0
    
End Function
Private Function getFontForItem(anID As String, hdc As Long, Optional getDefault As Boolean) As Long
   Dim aFont As StdFont
   Dim aNCMETRIC As NONCLIENTMETRICS
   Dim logPixConv As Double

   Dim systemMenuFont As String
   Dim systemMenuSize As Long

   ' Calculate a logical pixels conversion factor
   logPixConv = GetDeviceCaps(hdc, LOGPIXELSY) / 72
   
   On Error Resume Next
   ' Get the nonclient metrics, including system menu height
   aNCMETRIC.cbSize = Len(aNCMETRIC)
   SystemParametersInfo SPI_GETNONCLIENTMETRICS, aNCMETRIC.cbSize, aNCMETRIC, 0
   ' If we're not using custom fonts, or they are not set just use the normal fonts.
   If gMenuFont Is Nothing Or Not gUseCustomFonts Then
       ' Create a font with the system menu parameters
       With aNCMETRIC.lfMenuFont
           systemMenuFont = StrConv(.lfFaceName, vbUnicode)
           systemMenuFont = left$(systemMenuFont, InStr(systemMenuFont, Chr$(0)) - 1)
           getFontForItem = CreateFont(-1 * .lfHeight * logPixConv, .lfWidth, .lfEscapement, .lfOrientation, .lfWeight, .lfItalic, .lfUnderline, .lfStrikeOut, .lfCharSet, .lfOutPrecision, .lfClipPrecision, .lfQuality, .lfPitchAndFamily, systemMenuFont)
       End With
   Else
       ' use a custom font
       Set aFont = gCustomFonts(anID)
       With aNCMETRIC.lfMenuFont
            If aFont Is Nothing Or getDefault Then
                getFontForItem = CreateFont(-logPixConv * gMenuFont.Size, 0, .lfEscapement, .lfOrientation, gMenuFont.Weight, gMenuFont.Italic, gMenuFont.Underline, gMenuFont.Strikethrough, .lfCharSet, .lfOutPrecision, .lfClipPrecision, .lfQuality, .lfPitchAndFamily, gMenuFont.Name)
            Else
                getFontForItem = CreateFont(-logPixConv * aFont.Size, 0, .lfEscapement, .lfOrientation, aFont.Weight, aFont.Italic, aFont.Underline, aFont.Strikethrough, .lfCharSet, .lfOutPrecision, .lfClipPrecision, .lfQuality, .lfPitchAndFamily, aFont.Name)
            End If
       End With
   End If
End Function

Private Function drawMnuBitmap(ByVal hwnd As Long, ByVal drawInfoPtr As Long) As Long

    ' this function is only invoked when uMsg = WM_DRAWITEM, or when
    Dim aDIS As DRAWITEMSTRUCT
    
    Dim lPic        As StdPicture       ' our picture
    Dim lMask       As StdPicture       ' our picture's mask
    Dim picDC       As Long             ' picture device context
    Dim maskDC      As Long             ' mask device context
    Dim sCap        As String           ' Caption string
    Dim lPicInds    As Variant          ' picture indices (array where 0=picID, 1=maskID)
    Dim boxRect     As RECT             ' a rectangle used to paint stuff
    Dim abrush      As Long             ' brush object
    Dim noPics      As Boolean          ' true if no pictures for this menu item
    Dim colDC       As Long             ' a color's DC
    Dim colBmp      As Long             ' a Colors bitmap - used for adjusting stuff
    Dim aRect       As RECT             ' another rectangle. imagine that.
    Dim sAcc        As String           ' accelerator string
    Dim aPt         As POINTAPI         ' a user defined structure for x,y
    Dim xAcc        As Long             ' x location of the accelerator text
    Dim aPen        As Long             ' a pen object
    Dim tempWid     As Long             ' temporary width - used for adjusting wid of last top level menu item
    Dim testBrush   As Long             ' temporary brush object.
    Dim tempDC      As Long             ' temporary DC
    Dim tempBmp     As Long             ' temporary bitmap
    Dim customFont  As Long             ' custom font (if specified)
    Dim origFont    As Long             ' original font (if customFont is used)
    Dim aBmp        As BITMAP
    Dim aStyle      As Long
    Dim mnuItemWid  As Long
    Dim isSep       As Boolean
    
    Static lastDISRect As RECT
    Static lastDISID   As Long
    
    Dim backBuffDC  As Long
    Dim backBuffBmp As Long
    Dim aBrushOrg As POINTAPI
    Dim backBuffRECT As RECT
    Dim aTempBrush As Long
    Dim MII As MENUITEMINFO
    Dim aStr As String
    If drawInfoPtr = 0 Then Exit Function
    
    ' get the drawing structure information
    CopyMemory aDIS, ByVal drawInfoPtr, LenB(aDIS)

    aStr = String$(128, 0)
    MII.cbSize = Len(MII)
    MII.fMask = MIIM_TYPE Or MIIM_ID
    MII.cch = 127
    MII.dwTypeData = aStr
    GetMenuItemInfo aDIS.hwndItem, aDIS.itemID, False, MII
    Debug.Print "X" & MII.dwTypeData & "X"
    isSep = ((MII.fType And MF_SEPARATOR) = MF_SEPARATOR)
    
    ' Create a back buffer to draw stuff on - prevents flickering
    backBuffRECT.right = aDIS.rcItem.right - aDIS.rcItem.left
    backBuffRECT.bottom = aDIS.rcItem.bottom - aDIS.rcItem.top
    backBuffDC = CreateCompatibleDC(aDIS.hdc)
    backBuffBmp = CreateCompatibleBitmap(aDIS.hdc, backBuffRECT.right, backBuffRECT.bottom)
    DeleteObject SelectObject(backBuffDC, backBuffBmp)
    
    mnuItemWid = gMnuWidth
    ' create a brush of color/texture appropriate for the settings the user has defined
    If (aDIS.itemState And ODS_CHECKED) = ODS_CHECKED And (aDIS.itemState And ODS_SELECTED) = ODS_SELECTED And Not (aDIS.itemState And ODS_DISABLED) = ODS_DISABLED Then
        ' selected and checked items need the "selected" color behind the checkmark
    '    aBrush = CreateSolidBrush(GetSysColor(13))
        abrush = CreateSolidBrush(getAppColor("Selected"))
    Else
        ' Otherwise check whether we use the system menu color (4) or the custom texture
        If gbCustomTexture Then
            ' Create a texture brush
            abrush = CreatePatternBrush(gMenuBG.Handle)
        Else
        '    aBrush = CreateSolidBrush(GetSysColor(4))
            abrush = CreateSolidBrush(getAppColor("menubg"))
        End If
    End If
        
    ' If this is a top level menu and its the first one on a new row,
    '  go back and paint from the end of the previous item all the way to
    '  the right side. This is different from the very last item.
    If aDIS.ItemData = 2 Then
        If (lastDISRect.top < aDIS.rcItem.top And aDIS.itemID > gStartTopID And lastDISRect.right <> 0 And aDIS.itemID = lastDISID + 1) Then
            boxRect.left = lastDISRect.right
            boxRect.top = lastDISRect.top
            GetWindowRect hwnd, aRect
            boxRect.right = aRect.right - aRect.left - GetSystemMetrics(SM_CXBORDER) * 2 - 2
            boxRect.bottom = lastDISRect.bottom
            FillRect aDIS.hdc, boxRect, abrush
        End If
        CopyMemory lastDISRect, aDIS.rcItem, Len(lastDISRect)
        lastDISID = aDIS.itemID
    End If
    
    ' If this is a top level menu and its the last one, then we should
    '  temporarily change the item width so it paints all the way to the right
    '  on the form
    If (aDIS.ItemData = 2 And aDIS.itemID = gLastTopMenuID) Then
        GetWindowRect hwnd, aRect
        boxRect.left = aDIS.rcItem.right
        boxRect.right = aRect.right - aRect.left - GetSystemMetrics(SM_CXBORDER) * 2 - 2
        boxRect.top = aDIS.rcItem.top
        boxRect.bottom = aDIS.rcItem.bottom
        FillRect aDIS.hdc, boxRect, abrush
    End If
    
    ' Adjust our brush's point of origin so it paints correctly
    SetBrushOrgEx backBuffDC, -aDIS.rcItem.left, -aDIS.rcItem.top, aBrushOrg
    aTempBrush = SelectObject(backBuffDC, abrush)
    FillRect backBuffDC, backBuffRECT, abrush
    SelectObject backBuffDC, aTempBrush
    DeleteObject abrush
    
    ' get the caption of the menu item we need to draw.
    '  We will need this not only to draw in the menu, but
    '  to retrieve the resource IDs for the bitmap images
    
    If isSep Then ' seperator
        boxRect.left = backBuffRECT.left + 5
        boxRect.top = (backBuffRECT.bottom - 2) / 2
        boxRect.right = backBuffRECT.right - 5
        boxRect.bottom = backBuffRECT.bottom - 2
        DrawEdge backBuffDC, boxRect, BDR_RAISEDINNER Or BDR_SUNKENOUTER, BF_TOP
        'DrawEdge backBuffDC, boxRect, BDR_SUNKENINNER Or BDR_SUNKENOUTER, BF_TOP
        GoTo drawMnuBitmap_exitFunction
    Else
        On Error Resume Next
        sCap = gMenuCaps("M" & CStr(aDIS.itemID) & "-" & hwnd)
        If Err.Number <> 0 Then sCap = ""
        ' Check and see if there is a key accelerator
        If InStr(sCap, vbTab) > 0 Then
            sAcc = VBA.right(sCap, Len(sCap) - InStr(sCap, vbTab))
            sCap = VBA.left$(sCap, InStr(sCap, vbTab) - 1)
        End If
    End If
    customFont = getFontForItem(sCap, backBuffDC)
    If customFont <> 0 Then origFont = SelectObject(backBuffDC, customFont)
    
    If aDIS.ItemData = 2 Then GoTo drawMnuBitmap_MenuBarMenu
    
    ' Get the bitmap image indices if they exist.
    '  if not, then skip the picture drawing stuff
    On Error Resume Next
    lPicInds = gMenuBmps(sCap)
    If Err.Number <> 0 And (aDIS.itemState And ODS_CHECKED) <> ODS_CHECKED Then
        noPics = True
        Err.Clear
        GoTo drawMnuBitmap_picsDone
    End If
    
    ' Get the pictures from the resource file
    If (aDIS.itemState And ODS_CHECKED) = ODS_CHECKED Then
        Set lPic = LoadResPicture("checked", vbResBitmap)
        Set lMask = LoadResPicture("checked", vbResBitmap)
        GetObject lPic, Len(aBmp), aBmp
        ' The following is purely to get the checkmark the color it
        '  is supposed to be. Crazy, eh?
        colDC = CreateCompatibleDC(aDIS.hdc)
        colBmp = CreateCompatibleBitmap(aDIS.hdc, aBmp.bmWidth, aBmp.bmHeight)
        DeleteObject SelectObject(colDC, colBmp)
        aRect.right = aBmp.bmWidth
        aRect.bottom = aBmp.bmHeight
        abrush = CreateSolidBrush(IIf((aDIS.itemState And ODS_SELECTED) = ODS_SELECTED, getAppColor("menutext"), IIf(gUseCustomFonts, gMenuFontColor, getAppColor("menutext"))))
        
        FillRect colDC, aRect, abrush
        DeleteObject abrush
    Else
        Set lPic = LoadResPicture(lPicInds(0), vbResBitmap)
        Set lMask = LoadResPicture(lPicInds(1), vbResBitmap)
        GetObject lPic, Len(aBmp), aBmp
    End If
    
    ' Create a compatible device context for both of the bitmaps
    picDC = CreateCompatibleDC(aDIS.hdc)
    maskDC = CreateCompatibleDC(aDIS.hdc)
    
    ' select the bitmaps into our device context and delete the temporary 1x1
    '  that's created with the DC
    
    If (aDIS.itemState And ODS_DISABLED) <> ODS_DISABLED Then
        DeleteObject SelectObject(picDC, lPic.Handle)
        DeleteObject SelectObject(maskDC, lMask.Handle)
    End If
    
    If (aDIS.itemState And ODS_CHECKED) = ODS_CHECKED Then
        'Make the checkmark the right color
        BitBlt picDC, 0, 0, aBmp.bmWidth, aBmp.bmHeight, colDC, 0, 0, vbSrcPaint
        DeleteDC colDC
        DeleteObject colBmp
   '     DrawEdge backBuffDC, backBuffRECT, BDR_SUNKENOUTER, BF_RECT
    End If
    
    ' set up a rectangle to draw the upraised edge
    boxRect.top = 0
    boxRect.left = 0
    boxRect.right = boxRect.left + mnuItemWid
    boxRect.bottom = backBuffRECT.bottom 'boxRect.top + gMnuHeight
    If (aDIS.itemState And ODS_SELECTED) = ODS_SELECTED And (aDIS.itemState And ODS_CHECKED) <> ODS_CHECKED And (aDIS.itemState And ODS_DISABLED) <> ODS_DISABLED Then DrawEdge backBuffDC, boxRect, BDR_RAISEDINNER, BF_RECT

drawMnuBitmap_picsDone:
    
    ' If the item is in a "highlighted" state, then
    If (aDIS.itemState And ODS_SELECTED) = ODS_SELECTED And Not (aDIS.itemState And ODS_DISABLED) = ODS_DISABLED Then
        ' Draw the edge
        ' set up a rectangle to draw the "highlight" color
        boxRect.left = IIf(noPics, backBuffRECT.left, boxRect.right)
        boxRect.top = backBuffRECT.top
        boxRect.bottom = backBuffRECT.bottom
        boxRect.right = backBuffRECT.right
        ' create a brush in the highlight color
        abrush = CreateSolidBrush(getAppColor("Selected"))
        ' color the rectangular area
        FillRect backBuffDC, boxRect, abrush
        ' delete the brush object (clear up resources)
        DeleteObject abrush
    End If
    

    'Set our text colors appropriately, depending on whether we are
    ' in a highlighted state or not
    If (aDIS.itemState And ODS_SELECTED) = ODS_SELECTED Then
        SetTextColor backBuffDC, getAppColor("SelectedText")
        SetBkColor backBuffDC, getAppColor("Selected")
    Else
        SetTextColor backBuffDC, IIf(gMenuFontColor = 0, getAppColor("menutext"), gMenuFontColor)
        SetBkColor backBuffDC, getAppColor("menubg")
    End If
    
    'Print the text
    SetBkMode backBuffDC, 0
    GetTextExtentPoint32 backBuffDC, sCap, Len(sCap), aPt
    boxRect.top = backBuffRECT.top + ((backBuffRECT.bottom - backBuffRECT.top - aPt.Y) / 2)
    boxRect.bottom = backBuffRECT.bottom
    boxRect.right = backBuffRECT.right - 10
    boxRect.left = mnuItemWid + 6
    DrawStateText backBuffDC, 0, 0, sCap, Len(sCap), boxRect.left, boxRect.top, 0, 0, DST_PREFIXTEXT Or IIf((aDIS.itemState And ODS_DISABLED) = ODS_DISABLED, DSS_DISABLED, DSS_NORMAL)
    'Print the accelerator (Ctl-Whatever)
    If sAcc <> "" Then
        DeleteObject SelectObject(backBuffDC, origFont)
        customFont = getFontForItem(sCap, backBuffDC, True)
        origFont = SelectObject(backBuffDC, customFont)
        GetTextExtentPoint32 backBuffDC, sAcc, Len(sAcc), aPt
        DrawStateText backBuffDC, 0, 0, sAcc, Len(sAcc), boxRect.right - aPt.x, boxRect.top, 0, 0, DST_PREFIXTEXT Or IIf((aDIS.itemState And ODS_DISABLED) = ODS_DISABLED, DSS_DISABLED, DSS_NORMAL)
    End If
    
    If lPic Is Nothing Then GoTo drawMnuBitmap_exitFunction
    If lPic.Handle <> 0 Then
        If (aDIS.itemState And ODS_DISABLED) = ODS_DISABLED Then
            DrawState backBuffDC, 0, 0, lPic.Handle, 0, (mnuItemWid - aBmp.bmWidth) / 2, backBuffRECT.top + (backBuffRECT.bottom - backBuffRECT.top - aBmp.bmHeight) / 2, aBmp.bmHeight, aBmp.bmHeight, DST_BITMAP Or DSS_DISABLED
        Else
            ' Blt the mask
            BitBlt backBuffDC, 2, backBuffRECT.top + (backBuffRECT.bottom - backBuffRECT.top - aBmp.bmHeight) / 2, aBmp.bmWidth, aBmp.bmHeight, maskDC, 0, 0, vbMergePaint
            ' Blt the picture
            BitBlt backBuffDC, 2, backBuffRECT.top + (backBuffRECT.bottom - backBuffRECT.top - aBmp.bmHeight) / 2, aBmp.bmWidth, aBmp.bmHeight, picDC, 0, 0, vbSrcAnd
        End If
        ' Clean up our graphics resources.
        DeleteDC picDC
        DeleteObject lPic.Handle
        DeleteDC maskDC
        DeleteObject lMask.Handle
    End If

drawMnuBitmap_exitFunction:
    
    ' Check to see if we are in a 2nd or 3rd column... if so, then
    '  we need to draw over a little bit of the menu.
    If aDIS.ItemData <> 2 And (aDIS.rcItem.left > 0) Then
        ' Calculate the rectangle for 4 pixels to the left
        aRect.left = aDIS.rcItem.left - 4
        aRect.right = aDIS.rcItem.left
        aRect.top = aDIS.rcItem.top
        aRect.bottom = aDIS.rcItem.bottom
        DeleteObject abrush
        ' Create a new brush object
        If gbCustomTexture And Not (gMenuBG Is Nothing) Then
            ' Pattern brush
            abrush = CreatePatternBrush(gMenuBG.Handle)
        Else
            ' Solid brush
            abrush = CreateSolidBrush(getAppColor("menubg"))
        End If
        ' Fill in the area
        FillRect aDIS.hdc, aRect, abrush
        ' Clean up our brush resource.
        DeleteObject abrush
    End If
    
    ' Clean up our graphics resources to free up memory
    If origFont <> 0 Then
        SelectObject backBuffDC, origFont
        DeleteObject customFont
    End If
    DeleteObject abrush
    DeleteDC picDC
    DeleteDC maskDC
    
    ' Transfer the menu item from our back buffer into the menu DC
    BitBlt aDIS.hdc, aDIS.rcItem.left, aDIS.rcItem.top, backBuffRECT.right, backBuffRECT.bottom, backBuffDC, 0, 0, vbSrcCopy
    DeleteDC backBuffDC
    DeleteObject backBuffBmp
    
    On Error GoTo 0

    Exit Function
    
drawMnuBitmap_MenuBarMenu:
    ' Top level menus... These things are so freakin easy its funny.
    If (aDIS.itemState And ODS_HOTTRACK) = ODS_HOTTRACK Then
        ' This little style bit is courtesy of VolteFace from www.visualbasicforum.com
        DrawEdge backBuffDC, backBuffRECT, BDR_RAISEDINNER, BF_RECT
    ElseIf (aDIS.itemState And ODS_SELECTED) = ODS_SELECTED And (aDIS.itemState And ODS_DISABLED) <> ODS_DISABLED Then
        ' If its a selected item, paint the background with the systems 'Highlighted' color
        abrush = SelectObject(backBuffDC, CreateSolidBrush(getAppColor("Selected")))
        ' Also make the text print out in the highlighted text color
        SetTextColor backBuffDC, getAppColor("SelectedText")
        aPen = SelectObject(backBuffDC, CreatePen(0, 1, getAppColor("BorderColor")))
        Rectangle backBuffDC, backBuffRECT.left, backBuffRECT.top, backBuffRECT.right, backBuffRECT.bottom
        DeleteObject SelectObject(backBuffDC, aPen)
        DeleteObject SelectObject(backBuffDC, abrush)
    Else
        ' otherwise just make it the system menu text color
        'SetTextColor aDIS.hdc, GetSysColor(IIf((aDIS.itemState And ODS_DISABLED) = ODS_DISABLED, 17, 7))
        SetTextColor backBuffDC, IIf(gMenuFontColor = 0, getAppColor("GeneralText"), gMenuFontColor)
    End If
    ' Make the text print transparently
    SetBkMode backBuffDC, 0
    ' Get text dimensions
    GetTextExtentPoint32 backBuffDC, sCap, Len(sCap), aPt
    ' Draw the text!
    DrawStateText backBuffDC, 0, 0, sCap, Len(sCap), backBuffRECT.top + (backBuffRECT.right - backBuffRECT.left - aPt.x) / 2, backBuffRECT.top + (backBuffRECT.bottom - backBuffRECT.top - aPt.Y) / 2, 0, 0, DST_PREFIXTEXT Or IIf((aDIS.itemState And ODS_DISABLED) = ODS_DISABLED, DSS_DISABLED, DSS_NORMAL)
    GoTo drawMnuBitmap_exitFunction
End Function

