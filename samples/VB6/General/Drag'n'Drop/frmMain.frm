VERSION 5.00
Object = "{E7BB2F30-C5DD-4370-B7E2-19A7EDF169EE}#1.10#0"; "TabStripCtlU.ocx"
Begin VB.Form frmMain 
   Caption         =   "TabStrip 1.10 - Drag'n'Drop Sample"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6450
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   265
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   430
   StartUpPosition =   2  'Bildschirmmitte
   Begin TabStripCtlLibUCtl.TabStrip TabStripU 
      Align           =   3  'Links ausrichten
      Height          =   3975
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4755
      _cx             =   8387
      _cy             =   7011
      AllowDragDrop   =   -1  'True
      Appearance      =   0
      BorderStyle     =   0
      CloseableTabs   =   0   'False
      CloseableTabsMode=   1
      DisabledEvents  =   262370
      DragActivateTime=   -1
      DragScrollTimeBase=   -1
      Enabled         =   -1  'True
      FixedTabWidth   =   96
      FocusOnButtonDown=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HorizontalPadding=   6
      HotTracking     =   0   'False
      HoverTime       =   -1
      InsertMarkColor =   0
      MinTabWidth     =   -1
      MousePointer    =   0
      MultiRow        =   0   'False
      MultiSelect     =   0   'False
      OLEDragImageStyle=   0
      OwnerDrawn      =   0   'False
      ProcessContextMenuKeys=   -1  'True
      RaggedTabRows   =   -1  'True
      RegisterForOLEDragDrop=   2
      RightToLeft     =   0
      ScrollToOpposite=   0   'False
      ShowButtonSeparators=   -1  'True
      ShowToolTips    =   -1  'True
      Style           =   0
      SupportOLEDragImages=   -1  'True
      TabBoundingBoxDefinition=   8198
      TabCaptionAlignment=   0
      TabHeight       =   18
      TabPlacement    =   0
      UseFixedTabWidth=   0   'False
      UseSystemFont   =   -1  'True
      VerticalPadding =   3
   End
   Begin VB.CheckBox chkOLEDragDrop 
      Caption         =   "&OLE Drag'n'Drop"
      Height          =   255
      Left            =   4860
      TabIndex        =   1
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4980
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuAeroDragImages 
         Caption         =   "&Aero OLE Drag Images"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "dummy"
      Visible         =   0   'False
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy Item"
      End
      Begin VB.Menu mnuMove 
         Caption         =   "&Move Item"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "C&ancel"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Implements ISubclassedWindow


  Private Enum ChosenMenuItemConstants
    ciCopy
    ciMove
    ciCancel
  End Enum


  Private Const CF_DIBV5 = 17
  Private Const CF_OEMTEXT = 7
  Private Const CF_UNICODETEXT = 13

  Private Const strCLSID_RecycleBin = "{645FF040-5081-101B-9F08-00AA002F954E}"


  Private Type DLLVERSIONINFO
    cbSize As Long
    dwMajor As Long
    dwMinor As Long
    dwBuildNumber As Long
    dwPlatformId As Long
  End Type

  Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
  End Type

  Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
  End Type

  Private Type Size
    cx As Long
    cy As Long
  End Type
  
  Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As Long
    cch As Long
  End Type
  
  Private Type MENUBARINFO
    cbSize As Long
    rcBar As RECT
    hMenu As Long
    hwndMenu As Long
    fFlags As Long
  End Type

  Private Type DRAWITEMSTRUCT
    ctlType As Long
    ctlID As Long
    itemID As Long
    itemAction As Long
    itemState As Long
    hwndItem As Long
    hDC As Long
    rcItem As RECT
    itemData As Long
  End Type
  
  Private Type UAHMENU
    hMenu As Long
    hDC As Long
    dwFlags As Long
  End Type
  
  Private Type UAHMENUITEMMETRICS
    rgsizeBar1 As Size
    rgsizeBar2 As Size
    rgsizePopup3 As Size
    rgsizePopup4 As Size
  End Type
  
  Private Type UAHMENUPOPUPMETRICS
    rgcx(0 To 3) As Long
    fUpdateMaxWidths As Integer
  End Type
  
  Private Type UAHMENUITEM
    iPosition As Long
    umim As UAHMENUITEMMETRICS
    umpm As UAHMENUPOPUPMETRICS
  End Type
  
  Private Type UAHDRAWMENUITEM
    dis As DRAWITEMSTRUCT
    um As UAHMENU
    uni As UAHMENUITEM
  End Type

  Private CFSTR_HTML As Long
  Private CFSTR_HTML2 As Long
  Private CFSTR_LOGICALPERFORMEDDROPEFFECT As Long
  Private CFSTR_MOZILLAHTMLCONTEXT As Long
  Private CFSTR_PERFORMEDDROPEFFECT As Long
  Private CFSTR_TARGETCLSID As Long
  Private CFSTR_TEXTHTML As Long

  Private bComctl32Version600OrNewer As Boolean
  Private bDarkModeSupported As Boolean
  Private bDarkModeActivated As Boolean
  Private bRightDrag As Boolean
  Private chosenMenuItem As ChosenMenuItemConstants
  Private CLSID_RecycleBin As GUID
  Private hImgLst As Long
  Private lastDropTarget As Long
  Private ToolTip As clsToolTip
  Private themeableOS As Boolean
  Private hFormBkBrush As Long

  Private Declare Function AllowDarkModeForWindow Lib "uxtheme.dll" Alias "#133" (ByVal hWnd As Long, ByVal bAllow As Long) As Long
  Private Declare Function AllowDarkModeForWindowWithParentFallback Lib "uxtheme.dll" Alias "#145" (ByVal hWnd As Long, ByVal bAutoThemeChange As Long) As Long
  Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
  Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal pString As Long, CLSID As GUID) As Long
  Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hrgnDest As Long, ByVal hrgnSrc1 As Long, ByVal hrgnSrc2 As Long, ByVal fnCombineMode As Long) As Long
  Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal sz As Long)
  Private Declare Function CreateRectRgnIndirect Lib "gdi32.dll" (lprc As RECT) As Long
  Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
  Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
  Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
  Private Declare Function DllGetVersion_comctl32 Lib "comctl32.dll" Alias "DllGetVersion" (Data As DLLVERSIONINFO) As Long
  Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpchText As Long, ByVal nCount As Long, ByVal lpRect As Long, ByVal uFormat As Long) As Long
  Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, ByVal pClipRect As Long) As Long
  Private Declare Function DrawThemeText Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal pszText As Long, ByVal cchText As Long, ByVal dwTextFlags As Long, ByVal dwTextFlags2 As Long, pRect As RECT) As Long
  Private Declare Function DwmSetWindowAttribute Lib "dwmapi.dll" (ByVal hWnd As Long, ByVal dwAttribute As Long, ByVal pvAttribute As Long, ByVal cbAttribute As Long) As Long
  Private Declare Function ExcludeClipRect Lib "gdi32.dll" (ByVal hDC As Long, ByVal nLeftRect As Long, ByVal nTopRect As Long, ByVal nRightRect As Long, ByVal nBottomRect As Long) As Long
  Private Declare Function FillRect Lib "user32.dll" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
  Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
  Private Declare Function GetClientRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As RECT) As Long
  Private Declare Function GetMenuBarInfo Lib "user32.dll" (ByVal hWnd As Long, ByVal idObject As Long, ByVal idItem As Long, pmbi As MENUBARINFO) As Long
  Private Declare Function GetMenuItemInfo Lib "user32.dll" Alias "GetMenuItemInfoW" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
  Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As Long) As Long
  Private Declare Function GetStockObject Lib "gdi32.dll" (ByVal i As Long) As Long
  Private Declare Function GetThemeSysColorBrush Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal iColorId As Long) As Long
  Private Declare Function GetWindowDC Lib "user32.dll" (ByVal hWnd As Long) As Long
  Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As RECT) As Long
  Private Declare Function ImageList_AddIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal hIcon As Long) As Long
  Private Declare Function ImageList_Create Lib "comctl32.dll" (ByVal cx As Long, ByVal cy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
  Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal himl As Long) As Long
  Private Declare Function ImageList_GetImageCount Lib "comctl32.dll" (ByVal himl As Long) As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
  Private Declare Function LoadImage Lib "user32.dll" Alias "LoadImageW" (ByVal hinst As Long, ByVal lpszName As Long, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
  Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
  Private Declare Function MapWindowPoints Lib "user32.dll" (ByVal hwndFrom As Long, ByVal hWndTo As Long, ByVal lpPoints As Long, ByVal cPoints As Long) As Long
  Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
  Private Declare Function RegisterClipboardFormat Lib "user32.dll" Alias "RegisterClipboardFormatW" (ByVal lpString As Long) As Long
  Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long
  Private Declare Function SendMessageAsLong Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Private Declare Function SetBkMode Lib "gdi32.dll" (ByVal hDC As Long, ByVal mode As Long) As Long
  Private Declare Function SetDCBrushColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal color As Long) As Long
  Private Declare Function SetPreferredAppMode Lib "uxtheme.dll" Alias "#135" (ByVal lOption As Long) As Long
  Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
  Private Declare Function SetWindowTheme Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pSubAppName As Long, ByVal pSubIDList As Long) As Long
  Private Declare Function ShouldAppsUseDarkMode Lib "uxtheme.dll" Alias "#132" () As Long


Private Function ISubclassedWindow_HandleMessage(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal eSubclassID As EnumSubclassID, bCallDefProc As Boolean) As Long
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  Select Case eSubclassID
    Case EnumSubclassID.escidFrmMain
      lRet = HandleMessage_Form(hWnd, uMsg, wParam, lParam, bCallDefProc)
    Case EnumSubclassID.escidFrmMainTabStripU
      lRet = HandleMessage_TabStripU(hWnd, uMsg, wParam, lParam, bCallDefProc)
    Case Else
      Debug.Print "frmMain.ISubclassedWindow_HandleMessage: Unknown Subclassing ID " & CStr(eSubclassID)
  End Select

StdHandler_Ende:
  ISubclassedWindow_HandleMessage = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in frmMain.ISubclassedWindow_HandleMessage (SubclassID=" & CStr(eSubclassID) & ": ", Err.Number, Err.Description
  Resume StdHandler_Ende
End Function

Private Function HandleMessage_Form(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bCallDefProc As Boolean) As Long
  Const DC_BRUSH = 18
  Const DT_CENTER = &H1
  Const DT_HIDEPREFIX = &H100000
  Const DT_SINGLELINE = &H20
  Const DT_VCENTER = &H4
  Const MENU_BARBACKGROUND = 7
  Const MENU_BARITEM = 8
  Const MIIM_STRING = &H40
  Const MB_ACTIVE = 1
  Const MB_INACTIVE = 2
  Const MBI_NORMAL = 1
  Const MBI_HOT = 2
  Const MBI_PUSHED = 3
  Const MBI_DISABLED = 4
  Const MBI_DISABLEDHOT = 5
  Const MBI_DISABLEDPUSHED = 6
  Const OBJID_MENU = &HFFFFFFFD
  Const ODS_SELECTED = &H1
  Const ODS_GRAYED = &H2
  Const ODS_DISABLED = &H4
  Const ODS_DEFAULT = &H20
  Const ODS_HOTLIGHT = &H40
  Const ODS_INACTIVE = &H80
  Const ODS_NOACCEL = &H100
  Const TMT_MENUHILIGHT = 1630
  Const TMT_MENUBAR = 1631
  Const WM_ACTIVATE = &H6
  Const WM_CTLCOLORBTN = &H135
  Const WM_CTLCOLORSTATIC = &H138
  Const WM_NOTIFYFORMAT = &H55
  Const WM_SETTINGCHANGE = &H1A
  Const WM_SIZE = &H5
  Const WM_THEMECHANGED = &H31A
  Const WM_UAHDRAWMENU = &H91
  Const WM_UAHDRAWMENUITEM = &H92
  Const WM_USER = &H400
  Const WM_WINDOWPOSCHANGED = &H47
  Const OCM__BASE = WM_USER + &H1C00
  Dim lRet As Long
  Dim mnuData As UAHMENU
  Dim mnuItemData As UAHDRAWMENUITEM
  Dim mnuBarInfo As MENUBARINFO
  Dim miiData As MENUITEMINFO
  Dim rcWindow As RECT
  Dim hTheme As Long
  Dim iBackgroundState As Long
  Dim iTextState As Long
  Dim textFlags As Long
  Dim bytBuffer(0 To 512) As Byte
  Dim previousColor As Long
  Dim previousDarkModeActivated As Boolean

  On Error GoTo StdHandler_Error
  Select Case uMsg
    Case WM_ACTIVATE, WM_SIZE, WM_WINDOWPOSCHANGED
      If bDarkModeActivated Then
        lRet = DefSubclassProc(hWnd, uMsg, wParam, lParam)
        bCallDefProc = False
        DrawMenuBarLine hWnd
      End If
    
    Case WM_CTLCOLORBTN, WM_CTLCOLORSTATIC
      If bDarkModeActivated Then
        lRet = hFormBkBrush
        bCallDefProc = False
      End If
    
    Case WM_NOTIFYFORMAT
      ' give the control a chance to request Unicode notifications
      lRet = SendMessageAsLong(wParam, OCM__BASE + uMsg, wParam, lParam)
      bCallDefProc = False
      
    Case WM_SETTINGCHANGE, WM_THEMECHANGED
      lRet = DefSubclassProc(hWnd, uMsg, wParam, lParam)
      bCallDefProc = False
      If bDarkModeSupported Then
        previousDarkModeActivated = bDarkModeActivated
        If ShouldAppsUseDarkMode Then
          bDarkModeActivated = True
        Else
          bDarkModeActivated = False
        End If
        If previousDarkModeActivated <> bDarkModeActivated Then
          If bDarkModeActivated And hFormBkBrush = 0 Then
            hFormBkBrush = CreateSolidBrush(&H141414)
          End If
          SetupDarkMode
        End If
      End If
    
    Case WM_UAHDRAWMENU
      If bDarkModeActivated Then
        CopyMemory VarPtr(mnuData), lParam, LenB(mnuData)
        mnuBarInfo.cbSize = LenB(mnuBarInfo)
        If GetMenuBarInfo(hWnd, OBJID_MENU, 0, mnuBarInfo) Then
          bCallDefProc = False
          GetWindowRect hWnd, rcWindow
          mnuBarInfo.rcBar.Left = mnuBarInfo.rcBar.Left - rcWindow.Left
          mnuBarInfo.rcBar.Right = mnuBarInfo.rcBar.Right - rcWindow.Left
          mnuBarInfo.rcBar.Top = mnuBarInfo.rcBar.Top - rcWindow.Top - 1
          mnuBarInfo.rcBar.Bottom = mnuBarInfo.rcBar.Bottom - rcWindow.Top
'          hTheme = OpenThemeData(hWnd, StrPtr("Menu"))
'          If hTheme Then
'            'DrawThemeBackground hTheme, mnuData.hDC, MENU_BARBACKGROUND, MB_ACTIVE, mnuBarInfo.rcBar, 0
'            FillRect mnuData.hDC, mnuBarInfo.rcBar, GetThemeSysColorBrush(hTheme, TMT_MENUBAR)
'            CloseThemeData hTheme
'          End If
          previousColor = SetDCBrushColor(mnuData.hDC, &H141414)
          FillRect mnuData.hDC, mnuBarInfo.rcBar, GetStockObject(DC_BRUSH)
          SetDCBrushColor mnuData.hDC, previousColor
        End If
      End If
    
    Case WM_UAHDRAWMENUITEM
      If bDarkModeActivated Then
        CopyMemory VarPtr(mnuItemData), lParam, LenB(mnuItemData)
        miiData.cbSize = LenB(miiData)
        miiData.fMask = MIIM_STRING
        miiData.dwTypeData = VarPtr(bytBuffer(0))
        miiData.cch = 256
        If GetMenuItemInfo(mnuItemData.um.hMenu, mnuItemData.uni.iPosition, 1, miiData) Then
          bCallDefProc = False
          iBackgroundState = 0
          If mnuItemData.dis.itemState And (ODS_INACTIVE Or ODS_DEFAULT) Then
            iBackgroundState = MBI_NORMAL
          End If
          If mnuItemData.dis.itemState And ODS_HOTLIGHT Then
            iBackgroundState = MBI_HOT
          End If
          If mnuItemData.dis.itemState And ODS_SELECTED Then
            iBackgroundState = MBI_PUSHED
          End If
          If mnuItemData.dis.itemState And (ODS_GRAYED Or ODS_DISABLED) Then
            iBackgroundState = MBI_DISABLED
            If mnuItemData.dis.itemState And ODS_HOTLIGHT Then
              iBackgroundState = MBI_DISABLEDHOT
            End If
            If mnuItemData.dis.itemState And ODS_SELECTED Then
              iBackgroundState = MBI_DISABLEDPUSHED
            End If
          End If
          iTextState = iBackgroundState
          textFlags = DT_CENTER Or DT_SINGLELINE Or DT_VCENTER
          If mnuItemData.dis.itemState And ODS_NOACCEL Then
            textFlags = textFlags Or DT_HIDEPREFIX
          End If
'          hTheme = OpenThemeData(hWnd, StrPtr("Menu"))
'          If hTheme Then
'            'DrawThemeBackground hTheme, mnuItemData.um.hDC, MENU_BARITEM, iBackgroundState, mnuItemData.dis.rcItem, 0
'            If iBackgroundState = MBI_HOT Or iBackgroundState = MBI_PUSHED Then
'              FillRect mnuItemData.um.hDC, mnuItemData.dis.rcItem, GetThemeSysColorBrush(hTheme, TMT_MENUHILIGHT)
'            Else
'              FillRect mnuItemData.um.hDC, mnuItemData.dis.rcItem, GetThemeSysColorBrush(hTheme, TMT_MENUBAR)
'            End If
'            SetTextColor mnuItemData.um.hDC, &HDEDEDE
'            SetBkMode mnuItemData.um.hDC, 1
'            'DrawThemeText hTheme, mnuItemData.um.hDC, MENU_BARITEM, iTextState, miiData.dwTypeData, miiData.cch, textFlags, 0, mnuItemData.dis.rcItem
'            DrawText mnuItemData.um.hDC, miiData.dwTypeData, miiData.cch, VarPtr(mnuItemData.dis.rcItem), textFlags
'            CloseThemeData hTheme
'          End If
          If iBackgroundState = MBI_HOT Or iBackgroundState = MBI_PUSHED Then
            previousColor = SetDCBrushColor(mnuItemData.um.hDC, &H414141)
          Else
            previousColor = SetDCBrushColor(mnuItemData.um.hDC, &H141414)
          End If
          FillRect mnuItemData.um.hDC, mnuItemData.dis.rcItem, GetStockObject(DC_BRUSH)
          SetDCBrushColor mnuItemData.um.hDC, previousColor
          previousColor = SetTextColor(mnuItemData.um.hDC, &HDEDEDE)
          SetBkMode mnuItemData.um.hDC, 1
          DrawText mnuItemData.um.hDC, miiData.dwTypeData, miiData.cch, VarPtr(mnuItemData.dis.rcItem), textFlags
          SetTextColor mnuItemData.um.hDC, previousColor
        End If
      End If
  End Select

StdHandler_Ende:
  HandleMessage_Form = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in frmMain.HandleMessage_Form: ", Err.Number, Err.Description
  Resume StdHandler_Ende
End Function

Private Function HandleMessage_TabStripU(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bCallDefProc As Boolean) As Long
  Const WM_PAINT As Long = &HF
  Dim lRet As Long
  Dim hDC As Long
  Dim rcClient As RECT
  Dim rcTabPage As RECT
  Dim hRgn As Long
  Dim i As Long
  Dim tsTab As TabStripTab
  Dim isLastTab As Boolean

  On Error GoTo StdHandler_Error
  Select Case uMsg
    Case WM_PAINT
      If bDarkModeActivated Then
        lRet = DefSubclassProc(hWnd, uMsg, wParam, lParam)
        GetClientRect hWnd, rcClient
        hDC = GetWindowDC(hWnd)
        If hDC Then
          rcTabPage.Bottom = rcClient.Bottom
          rcTabPage.Left = rcClient.Left
          rcTabPage.Top = -999999
          rcTabPage.Right = rcClient.Right
          isLastTab = True
          For i = TabStripU.Tabs.Count - 1 To 0 Step -1
            Set tsTab = TabStripU.Tabs.Item(i)
            ExcludeClipRect hDC, tsTab.Left - IIf(tsTab.Active, 2, 0), tsTab.Top - IIf(tsTab.Active, 2, 0), tsTab.Left + tsTab.Width - IIf(isLastTab And Not tsTab.Active, 2, -2), tsTab.Top + tsTab.Height
            If tsTab.Top + tsTab.Height > rcTabPage.Top Then
              rcTabPage.Top = tsTab.Top + tsTab.Height
            End If
            isLastTab = False
          Next i
          ExcludeClipRect hDC, rcTabPage.Left, rcTabPage.Top, rcTabPage.Right, rcTabPage.Bottom
          FillRect hDC, rcClient, hFormBkBrush
          ReleaseDC hWnd, hDC
        End If
        bCallDefProc = False
      End If
  End Select

StdHandler_Ende:
  HandleMessage_TabStripU = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in frmMain.HandleMessage_TabStripU: ", Err.Number, Err.Description
  Resume StdHandler_Ende
End Function


Private Sub cmdAbout_Click()
  TabStripU.About
End Sub

Private Sub Form_Initialize()
  Const ILC_COLOR24 = &H18
  Const ILC_COLOR32 = &H20
  Const ILC_MASK = &H1
  Const IMAGE_ICON = 1
  Const LR_DEFAULTSIZE = &H40
  Const LR_LOADFROMFILE = &H10
  Const AllowDark = 1
  Dim DLLVerData As DLLVERSIONINFO
  Dim hIcon As Long
  Dim hMod As Long
  Dim iconsDir As String
  Dim iconPath As String

  InitCommonControls

  CFSTR_HTML = RegisterClipboardFormat(StrPtr("HTML Format"))
  CFSTR_HTML2 = RegisterClipboardFormat(StrPtr("HTML (HyperText Markup Language)"))
  CFSTR_LOGICALPERFORMEDDROPEFFECT = RegisterClipboardFormat(StrPtr("Logical Performed DropEffect"))
  CFSTR_MOZILLAHTMLCONTEXT = RegisterClipboardFormat(StrPtr("text/_moz_htmlcontext"))
  CFSTR_PERFORMEDDROPEFFECT = RegisterClipboardFormat(StrPtr("Performed DropEffect"))
  CFSTR_TARGETCLSID = RegisterClipboardFormat(StrPtr("TargetCLSID"))
  CFSTR_TEXTHTML = RegisterClipboardFormat(StrPtr("text/html"))

  CLSIDFromString StrPtr(strCLSID_RecycleBin), CLSID_RecycleBin
  lastDropTarget = -1

  With DLLVerData
    .cbSize = LenB(DLLVerData)
    DllGetVersion_comctl32 DLLVerData
    bComctl32Version600OrNewer = (.dwMajor >= 6)
  End With

  hMod = LoadLibrary(StrPtr("uxtheme.dll"))
  If hMod Then
    themeableOS = True
    If GetProcAddress(hMod, 135) <> 0 And GetProcAddress(hMod, 132) <> 0 Then
      bDarkModeSupported = True
    End If
    FreeLibrary hMod
  End If

  If bDarkModeSupported Then
    SetPreferredAppMode AllowDark
    If ShouldAppsUseDarkMode Then
      bDarkModeActivated = True
      hFormBkBrush = CreateSolidBrush(&H141414)
    End If
  End If

  hImgLst = ImageList_Create(16, 16, IIf(bComctl32Version600OrNewer, ILC_COLOR32, ILC_COLOR24) Or ILC_MASK, 14, 0)
  If Right$(App.Path, 3) = "bin" Then
    iconsDir = App.Path & "\..\res\"
  Else
    iconsDir = App.Path & "\res\"
  End If
  iconsDir = iconsDir & IIf(bComctl32Version600OrNewer, "16x16x32bpp\", "16x16x8bpp\")
  iconPath = Dir$(iconsDir & "*.ico")
  While iconPath <> ""
    hIcon = LoadImage(0, StrPtr(iconsDir & iconPath), IMAGE_ICON, 16, 16, LR_LOADFROMFILE Or LR_DEFAULTSIZE)
    If hIcon Then
      ImageList_AddIcon hImgLst, hIcon
      DestroyIcon hIcon
    End If
    iconPath = Dir$
  Wend
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, shift As Integer)
  If KeyCode = KeyCodeConstants.vbKeyEscape Then
    If Not (TabStripU.DraggedTabs Is Nothing) Then TabStripU.EndDrag True
  End If
End Sub

Private Sub Form_Load()
  Set ToolTip = New clsToolTip
  With ToolTip
    .Create Me.hWnd
    .Activate
    .AddTool chkOLEDragDrop.hWnd, "Check to start OLE Drag'n'Drop operations instead of normal ones.", , , False
  End With

  ' this is required to make the control work as expected
  Subclass

  InsertTabs
  SetupDarkMode
End Sub

Private Sub Form_Terminate()
  If hImgLst Then ImageList_Destroy hImgLst
  Set ToolTip = Nothing
  If hFormBkBrush Then
    DeleteObject hFormBkBrush
    hFormBkBrush = 0
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not UnSubclassWindow(Me.hWnd, EnumSubclassID.escidFrmMain) Then
    Debug.Print "UnSubclassing failed!"
  End If
End Sub

Private Sub mnuAeroDragImages_Click()
  TabStripU.OLEDragImageStyle = IIf(Not mnuAeroDragImages.Checked, OLEDragImageStyleConstants.odistAeroIfAvailable, OLEDragImageStyleConstants.odistClassic)
  mnuAeroDragImages.Checked = (TabStripU.OLEDragImageStyle = OLEDragImageStyleConstants.odistAeroIfAvailable)
End Sub

Private Sub mnuCancel_Click()
  chosenMenuItem = ChosenMenuItemConstants.ciCancel
End Sub

Private Sub mnuCopy_Click()
  chosenMenuItem = ChosenMenuItemConstants.ciCopy
End Sub

Private Sub mnuMove_Click()
  chosenMenuItem = ChosenMenuItemConstants.ciMove
End Sub

Private Sub TabStripU_AbortedDrag()
  'Set TabStripU.DropHilitedTab = Nothing
  TabStripU.SetInsertMarkPosition InsertMarkPositionConstants.impNowhere, Nothing
  lastDropTarget = -1
End Sub

Private Sub TabStripU_DragMouseMove(dropTarget As TabStripCtlLibUCtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibUCtl.HitTestConstants, autoActivateTab As Boolean, autoScrollVelocity As Long)
  Dim newDropTarget As Long
  Dim newInsertMarkRelativePosition As InsertMarkPositionConstants

  autoActivateTab = False

  TabStripU.GetClosestInsertMarkPosition ScaleX(x, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels), ScaleY(y, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels), newInsertMarkRelativePosition, dropTarget
  If dropTarget Is Nothing Then
    newDropTarget = -1
  Else
    newDropTarget = dropTarget.Index
  End If
  If newDropTarget <> lastDropTarget Then
    'Set TabStripU.DropHilitedTab = dropTarget
    lastDropTarget = newDropTarget
  End If
  TabStripU.SetInsertMarkPosition newInsertMarkRelativePosition, dropTarget
End Sub

Private Sub TabStripU_Drop(ByVal dropTarget As TabStripCtlLibUCtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibUCtl.HitTestConstants)
  Dim insertAt As Long
  Dim insertionMark As InsertMarkPositionConstants
  Dim tsTab As TabStripTab

  If bRightDrag Then
    x = Me.ScaleX(x, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)
    y = Me.ScaleY(y, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)
    Me.PopupMenu mnuPopup, , x, y, mnuMove
    Select Case chosenMenuItem
      Case ChosenMenuItemConstants.ciCancel
        TabStripU_AbortedDrag
        Exit Sub
      Case ChosenMenuItemConstants.ciCopy
        MsgBox "ToDo: Copy the tab.", VbMsgBoxStyle.vbInformation, "Command not implemented"
        TabStripU_AbortedDrag
        Exit Sub
      Case ChosenMenuItemConstants.ciMove
        ' fall through
    End Select
  End If

  ' move the tab

  TabStripU.GetInsertMarkPosition insertionMark, tsTab
  Select Case insertionMark
    Case InsertMarkPositionConstants.impAfter
      insertAt = tsTab.Index + 1
    Case InsertMarkPositionConstants.impBefore
      insertAt = tsTab.Index
    Case InsertMarkPositionConstants.impNowhere
      insertAt = -1
  End Select
  If insertAt <> -1 Then
    For Each tsTab In TabStripU.DraggedTabs
      If tsTab.Index < insertAt Then
        tsTab.Index = insertAt - 1
      Else
        tsTab.Index = insertAt
        insertAt = insertAt + 1
      End If
    Next tsTab
  End If

  'Set TabStripU.DropHilitedTab = Nothing
  TabStripU.SetInsertMarkPosition InsertMarkPositionConstants.impNowhere, Nothing
  lastDropTarget = -1
End Sub

Private Sub TabStripU_KeyDown(KeyCode As Integer, ByVal shift As Integer)
  If KeyCode = KeyCodeConstants.vbKeyEscape Then
    If Not (TabStripU.DraggedTabs Is Nothing) Then TabStripU.EndDrag True
  End If
End Sub

Private Sub TabStripU_MouseUp(ByVal tsTab As TabStripCtlLibUCtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibUCtl.HitTestConstants)
  If chkOLEDragDrop.Value = CheckBoxConstants.vbUnchecked Then
    If ((button = MouseButtonConstants.vbLeftButton) And (Not bRightDrag)) Or ((button = MouseButtonConstants.vbRightButton) And bRightDrag) Then
      If Not (TabStripU.DraggedTabs Is Nothing) Then
        ' Are we within the (extended) client area?
        If lastDropTarget <> -1 Then
          ' yes
          TabStripU.EndDrag False
        Else
          ' no
          TabStripU.EndDrag True
        End If
      End If
    End If
  End If
End Sub

Private Sub TabStripU_OLECompleteDrag(ByVal Data As TabStripCtlLibUCtl.IOLEDataObject, ByVal performedEffect As TabStripCtlLibUCtl.OLEDropEffectConstants)
  Dim bytArray() As Byte
  Dim CLSIDOfTarget As GUID

  If performedEffect = TabStripCtlLibUCtl.OLEDropEffectConstants.odeMove Then
    ' remove the dragged tabs
    TabStripU.DraggedTabs.RemoveAll True
  ElseIf Data.GetFormat(CFSTR_TARGETCLSID) Then
    bytArray = Data.GetData(CFSTR_TARGETCLSID)
    CopyMemory VarPtr(CLSIDOfTarget), VarPtr(bytArray(LBound(bytArray))), LenB(CLSIDOfTarget)
    If IsEqualGUID(CLSIDOfTarget, CLSID_RecycleBin) Then
      ' remove the dragged tabs
      TabStripU.DraggedTabs.RemoveAll True
    End If
  End If
End Sub

Private Sub TabStripU_OLEDragDrop(ByVal Data As TabStripCtlLibUCtl.IOLEDataObject, effect As TabStripCtlLibUCtl.OLEDropEffectConstants, ByVal dropTarget As TabStripCtlLibUCtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibUCtl.HitTestConstants)
  Dim bytArray() As Byte
  Dim files() As String
  Dim i As Long
  Dim insertAt As Long
  Dim insertionMark As InsertMarkPositionConstants
  Dim p As Long
  Dim str As String
  Dim tsTab As TabStripTab

  If Data.GetFormat(CF_UNICODETEXT) Then
    str = Data.GetData(CF_UNICODETEXT)
  ElseIf Data.GetFormat(vbCFText) Then
    str = Data.GetData(vbCFText)
  ElseIf Data.GetFormat(CF_OEMTEXT) Then
    str = Data.GetData(CF_OEMTEXT)
  End If
  str = Replace$(str, vbNewLine, "")

  If str <> "" Then
    ' insert a new tab

    TabStripU.GetInsertMarkPosition insertionMark, tsTab
    Select Case insertionMark
      Case InsertMarkPositionConstants.impAfter
        insertAt = tsTab.Index + 1
      Case InsertMarkPositionConstants.impBefore
        insertAt = tsTab.Index
      Case InsertMarkPositionConstants.impNowhere
        insertAt = -1
    End Select
    Set TabStripU.CaretTab = TabStripU.Tabs.Add(str, insertAt, 1)

  ElseIf Data.GetFormat(vbCFFiles) Then
    ' insert a new tab for each file/folder

    TabStripU.GetInsertMarkPosition insertionMark, tsTab
    Select Case insertionMark
      Case InsertMarkPositionConstants.impAfter
        insertAt = tsTab.Index + 1
      Case InsertMarkPositionConstants.impBefore
        insertAt = tsTab.Index
      Case InsertMarkPositionConstants.impNowhere
        insertAt = -1
    End Select

    On Error GoTo NoFiles
    files = Data.GetData(vbCFFiles)
    For i = LBound(files) To UBound(files)
      p = InStrRev(files(i), "\")
      If p > 0 Then files(i) = Mid$(files(i), p + 1)
      Set tsTab = TabStripU.Tabs.Add(files(i), insertAt, 1)
      Set TabStripU.CaretTab = tsTab
      insertAt = tsTab.Index + 1
    Next i
NoFiles:
  End If

  'Set TabStripU.DropHilitedTab = Nothing
  TabStripU.SetInsertMarkPosition InsertMarkPositionConstants.impNowhere, Nothing
  lastDropTarget = -1

  ' Be careful with odeMove!! Some sources delete the data themselves if the Move effect was
  ' chosen. So you may lose data!
  effect = TabStripCtlLibUCtl.OLEDropEffectConstants.odeMove
  If shift And vbShiftMask Then effect = TabStripCtlLibUCtl.OLEDropEffectConstants.odeMove
  If shift And vbCtrlMask Then effect = TabStripCtlLibUCtl.OLEDropEffectConstants.odeCopy
  If shift And vbAltMask Then effect = TabStripCtlLibUCtl.OLEDropEffectConstants.odeLink

'  If Data.GetFormat(vbCFText) Then MsgBox "Dropped Text: " & Data.GetData(vbCFText)
'  If Data.GetFormat(CF_OEMTEXT) Then MsgBox "Dropped OEM-Text: " & Data.GetData(7)
'  If Data.GetFormat(CF_UNICODETEXT) Then MsgBox "Dropped Unicode-Text: " & Data.GetData(13)
'  If Data.GetFormat(vbCFRTF) Then MsgBox "Dropped RTF-Text: " & Data.GetData(vbCFRTF)
'  If Data.GetFormat(CFSTR_HTML) Then MsgBox "Dropped HTML: " & Data.GetData(CFSTR_HTML)
'  If Data.GetFormat(CFSTR_HTML2) Then MsgBox "Dropped HTML2: " & Data.GetData(CFSTR_HTML2)
'  If Data.GetFormat(CFSTR_MOZILLAHTMLCONTEXT) Then MsgBox "Dropped text/_moz_htmlcontext: " & Data.GetData(CFSTR_MOZILLAHTMLCONTEXT)
'  If Data.GetFormat(CFSTR_TEXTHTML) Then MsgBox "Dropped text/html: " & Data.GetData(CFSTR_TEXTHTML)
'  If data.GetFormat(vbCFFiles) Then
'    str = ""
'    For Each file In data.Files
'      str = str & file & vbNewLine
'    Next file
'    MsgBox "Dropped files:" & vbNewLine & str
'  End If
'
'  ReDim bytArray(1 To LenB(CLSID_RecycleBin)) As Byte
'  CopyMemory VarPtr(bytArray(LBound(bytArray))), VarPtr(CLSID_RecycleBin), LenB(CLSID_RecycleBin)
'  Data.SetData CFSTR_TARGETCLSID, bytArray
'
'  ' another way to inform the source of the performed drop effect
'  Data.SetData CFSTR_PERFORMEDDROPEFFECT, Effect
End Sub

Private Sub TabStripU_OLEDragEnter(ByVal Data As TabStripCtlLibUCtl.IOLEDataObject, effect As TabStripCtlLibUCtl.OLEDropEffectConstants, dropTarget As TabStripCtlLibUCtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibUCtl.HitTestConstants, autoActivateTab As Boolean, autoScrollVelocity As Long)
  Dim dropActionDescription As String
  Dim dropDescriptionIcon As DropDescriptionIconConstants
  Dim dropTargetDescription As String
  Dim newDropTarget As Long
  Dim newInsertMarkRelativePosition As InsertMarkPositionConstants

  autoActivateTab = False

  newDropTarget = -1
  TabStripU.GetClosestInsertMarkPosition ScaleX(x, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels), ScaleY(y, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels), newInsertMarkRelativePosition, dropTarget
  If Not (dropTarget Is Nothing) Then
    dropTargetDescription = dropTarget.Text
    newDropTarget = dropTarget.Index
  End If
  If newDropTarget <> lastDropTarget Then
    'Set TabStripU.DropHilitedTab = dropTarget
    lastDropTarget = newDropTarget
  End If
  TabStripU.SetInsertMarkPosition newInsertMarkRelativePosition, dropTarget

  If TabStripU.DraggedTabs Is Nothing Then
    effect = TabStripCtlLibUCtl.OLEDropEffectConstants.odeCopy
  Else
    effect = TabStripCtlLibUCtl.OLEDropEffectConstants.odeMove
  End If
  If shift And vbShiftMask Then effect = TabStripCtlLibUCtl.OLEDropEffectConstants.odeMove
  If shift And vbCtrlMask Then effect = TabStripCtlLibUCtl.OLEDropEffectConstants.odeCopy
  If shift And vbAltMask Then effect = TabStripCtlLibUCtl.OLEDropEffectConstants.odeLink

  If dropTarget Is Nothing Then
    dropDescriptionIcon = DropDescriptionIconConstants.ddiNoDrop
    dropActionDescription = "Cannot place here"
  Else
    Select Case effect
      Case TabStripCtlLibUCtl.OLEDropEffectConstants.odeMove
        Select Case newInsertMarkRelativePosition
          Case InsertMarkPositionConstants.impAfter
            dropActionDescription = "Move after %1"
          Case InsertMarkPositionConstants.impBefore
            dropActionDescription = "Move before %1"
          Case Else
            dropActionDescription = "Move to %1"
        End Select
        dropDescriptionIcon = DropDescriptionIconConstants.ddiMove
      Case TabStripCtlLibUCtl.OLEDropEffectConstants.odeLink
        Select Case newInsertMarkRelativePosition
          Case InsertMarkPositionConstants.impAfter
            dropActionDescription = "Create link after %1"
          Case InsertMarkPositionConstants.impBefore
            dropActionDescription = "Create link before %1"
          Case Else
            dropActionDescription = "Create link in %1"
        End Select
        dropDescriptionIcon = DropDescriptionIconConstants.ddiLink
      Case Else
        Select Case newInsertMarkRelativePosition
          Case InsertMarkPositionConstants.impAfter
            dropActionDescription = "Insert copy after %1"
          Case InsertMarkPositionConstants.impBefore
            dropActionDescription = "Insert copy before %1"
          Case Else
            dropActionDescription = "Copy to %1"
        End Select
        dropDescriptionIcon = DropDescriptionIconConstants.ddiCopy
    End Select
  End If
  Data.SetDropDescription dropTargetDescription, dropActionDescription, dropDescriptionIcon
End Sub

Private Sub TabStripU_OLEDragLeave(ByVal Data As TabStripCtlLibUCtl.IOLEDataObject, ByVal dropTarget As TabStripCtlLibUCtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibUCtl.HitTestConstants)
  'Set TabStripU.DropHilitedTab = Nothing
  TabStripU.SetInsertMarkPosition InsertMarkPositionConstants.impNowhere, Nothing
  lastDropTarget = -1
End Sub

Private Sub TabStripU_OLEDragMouseMove(ByVal Data As TabStripCtlLibUCtl.IOLEDataObject, effect As TabStripCtlLibUCtl.OLEDropEffectConstants, dropTarget As TabStripCtlLibUCtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibUCtl.HitTestConstants, autoActivateTab As Boolean, autoScrollVelocity As Long)
  Dim dropActionDescription As String
  Dim dropDescriptionIcon As DropDescriptionIconConstants
  Dim dropTargetDescription As String
  Dim newDropTarget As Long
  Dim newInsertMarkRelativePosition As InsertMarkPositionConstants

  autoActivateTab = False

  newDropTarget = -1
  TabStripU.GetClosestInsertMarkPosition ScaleX(x, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels), ScaleY(y, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels), newInsertMarkRelativePosition, dropTarget
  If Not (dropTarget Is Nothing) Then
    dropTargetDescription = dropTarget.Text
    newDropTarget = dropTarget.Index
  End If
  If newDropTarget <> lastDropTarget Then
    'Set TabStripU.DropHilitedTab = dropTarget
    lastDropTarget = newDropTarget
  End If
  TabStripU.SetInsertMarkPosition newInsertMarkRelativePosition, dropTarget

  If TabStripU.DraggedTabs Is Nothing Then
    effect = TabStripCtlLibUCtl.OLEDropEffectConstants.odeCopy
  Else
    effect = TabStripCtlLibUCtl.OLEDropEffectConstants.odeMove
  End If
  If shift And vbShiftMask Then effect = TabStripCtlLibUCtl.OLEDropEffectConstants.odeMove
  If shift And vbCtrlMask Then effect = TabStripCtlLibUCtl.OLEDropEffectConstants.odeCopy
  If shift And vbAltMask Then effect = TabStripCtlLibUCtl.OLEDropEffectConstants.odeLink

  If dropTarget Is Nothing Then
    dropDescriptionIcon = DropDescriptionIconConstants.ddiNoDrop
    dropActionDescription = "Cannot place here"
  Else
    Select Case effect
      Case TabStripCtlLibUCtl.OLEDropEffectConstants.odeMove
        Select Case newInsertMarkRelativePosition
          Case InsertMarkPositionConstants.impAfter
            dropActionDescription = "Move after %1"
          Case InsertMarkPositionConstants.impBefore
            dropActionDescription = "Move before %1"
          Case Else
            dropActionDescription = "Move to %1"
        End Select
        dropDescriptionIcon = DropDescriptionIconConstants.ddiMove
      Case TabStripCtlLibUCtl.OLEDropEffectConstants.odeLink
        Select Case newInsertMarkRelativePosition
          Case InsertMarkPositionConstants.impAfter
            dropActionDescription = "Create link after %1"
          Case InsertMarkPositionConstants.impBefore
            dropActionDescription = "Create link before %1"
          Case Else
            dropActionDescription = "Create link in %1"
        End Select
        dropDescriptionIcon = DropDescriptionIconConstants.ddiLink
      Case Else
        Select Case newInsertMarkRelativePosition
          Case InsertMarkPositionConstants.impAfter
            dropActionDescription = "Insert copy after %1"
          Case InsertMarkPositionConstants.impBefore
            dropActionDescription = "Insert copy before %1"
          Case Else
            dropActionDescription = "Copy to %1"
        End Select
        dropDescriptionIcon = DropDescriptionIconConstants.ddiCopy
    End Select
  End If
  Data.SetDropDescription dropTargetDescription, dropActionDescription, dropDescriptionIcon
End Sub

Private Sub TabStripU_OLESetData(ByVal Data As TabStripCtlLibUCtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim str As String
  Dim tsTab As TabStripTab

  Select Case formatID
    Case vbCFText, CF_OEMTEXT, CF_UNICODETEXT
      For Each tsTab In TabStripU.DraggedTabs
        str = str & tsTab.Text & vbNewLine
      Next tsTab
      Data.SetData formatID, str

    Case vbCFFiles
      ' WARNING! Dragging files to objects like the recycle bin may result in deleted files and
      ' lost data!
'      ReDim files(1 To 2) As String
'      files(1) = "C:\Test1.txt"
'      files(2) = "C:\Test2.txt"
'      Data.SetData formatID, files
  End Select
End Sub

Private Sub TabStripU_OLEStartDrag(ByVal Data As TabStripCtlLibUCtl.IOLEDataObject)
  Data.SetData vbCFText
  Data.SetData CF_OEMTEXT
  Data.SetData CF_UNICODETEXT
  Data.SetData vbCFFiles
End Sub

Private Sub TabStripU_RecreatedControlWindow(ByVal hWnd As Long)
  #If UseSubClassing Then
    If Not SubclassWindow(hWnd, Me, EnumSubclassID.escidFrmMainTabStripU) Then
      Debug.Print "Subclassing failed!"
    End If
  #End If
  InsertTabs
  SetupDarkMode
End Sub

Private Sub TabStripU_TabBeginDrag(ByVal tsTab As TabStripCtlLibUCtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibUCtl.HitTestConstants)
  Dim col As TabStripTabs

  bRightDrag = False
  lastDropTarget = -1
  If TabStripU.MultiSelect Then
    Set col = TabStripU.Tabs
    col.FilterType(FilteredPropertyConstants.fpSelected) = FilterTypeConstants.ftIncluding
    col.Filter(FilteredPropertyConstants.fpSelected) = Array(True)

    If chkOLEDragDrop.Value = CheckBoxConstants.vbChecked Then
      TabStripU.OLEDrag , , , TabStripU.CreateTabContainer(col)
    Else
      TabStripU.BeginDrag TabStripU.CreateTabContainer(col), -1
    End If
  Else
    If chkOLEDragDrop.Value = CheckBoxConstants.vbChecked Then
      TabStripU.OLEDrag , , , TabStripU.CreateTabContainer(tsTab)
    Else
      TabStripU.BeginDrag TabStripU.CreateTabContainer(tsTab), -1
    End If
  End If
End Sub

Private Sub TabStripU_TabBeginRDrag(ByVal tsTab As TabStripCtlLibUCtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibUCtl.HitTestConstants)
  Dim col As TabStripTabs

  bRightDrag = True
  lastDropTarget = -1
  If TabStripU.MultiSelect Then
    Set col = TabStripU.Tabs
    col.FilterType(FilteredPropertyConstants.fpSelected) = FilterTypeConstants.ftIncluding
    col.Filter(FilteredPropertyConstants.fpSelected) = Array(True)

    If chkOLEDragDrop.Value = CheckBoxConstants.vbChecked Then
      TabStripU.OLEDrag , , , TabStripU.CreateTabContainer(col)
    Else
      TabStripU.BeginDrag TabStripU.CreateTabContainer(col), -1
    End If
  Else
    If chkOLEDragDrop.Value = CheckBoxConstants.vbChecked Then
      TabStripU.OLEDrag , , , TabStripU.CreateTabContainer(tsTab)
    Else
      TabStripU.BeginDrag TabStripU.CreateTabContainer(tsTab), -1
    End If
  End If
End Sub


Private Sub InsertTabs()
  Dim cImages As Long
  Dim iIcon As Long

  TabStripU.hImageList = hImgLst
  cImages = ImageList_GetImageCount(hImgLst)

  With TabStripU.Tabs
    .Add "Tab 1", , iIcon, 1
    iIcon = (iIcon + 1) Mod cImages
    .Add "Tab 2", , iIcon, 2
    iIcon = (iIcon + 1) Mod cImages
    .Add "Tab 3", , iIcon, 3
    iIcon = (iIcon + 1) Mod cImages
    .Add "Tab 4", , iIcon, 4
    iIcon = (iIcon + 1) Mod cImages
    .Add "Tab 5", , iIcon, 5
    iIcon = (iIcon + 1) Mod cImages
  End With
End Sub

Private Function IsEqualGUID(IID1 As GUID, IID2 As GUID) As Boolean
  Dim Tmp1 As Currency
  Dim Tmp2 As Currency

  If IID1.Data1 = IID2.Data1 Then
    If IID1.Data2 = IID2.Data2 Then
      If IID1.Data3 = IID2.Data3 Then
        CopyMemory VarPtr(Tmp1), VarPtr(IID1.Data4(0)), 8
        CopyMemory VarPtr(Tmp2), VarPtr(IID2.Data4(0)), 8
        IsEqualGUID = (Tmp1 = Tmp2)
      End If
    End If
  End If
End Function

' subclasses this Form
Private Sub Subclass()
  Const NF_REQUERY = 4
  Const WM_NOTIFYFORMAT = &H55

  #If UseSubClassing Then
    If Not SubclassWindow(Me.hWnd, Me, EnumSubclassID.escidFrmMain) Then
      Debug.Print "Subclassing failed!"
    End If
    ' tell the control to negotiate the correct format with the form
    SendMessageAsLong TabStripU.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
    If Not SubclassWindow(TabStripU.hWnd, Me, EnumSubclassID.escidFrmMainTabStripU) Then
      Debug.Print "Subclassing failed!"
    End If
  #End If
End Sub

Private Sub SetupDarkMode()
  Const DWMWA_USE_IMMERSIVE_DARK_MODE = 20
  Dim useImmersiveDarkMode As Long
  
  If bDarkModeActivated Then
    useImmersiveDarkMode = 1
    DwmSetWindowAttribute Me.hWnd, DWMWA_USE_IMMERSIVE_DARK_MODE, VarPtr(useImmersiveDarkMode), LenB(useImmersiveDarkMode)
    'AllowDarkModeForWindow Me.hWnd, 1
    'AllowDarkModeForWindowWithParentFallback Me.hWnd, 1
    SetWindowTheme TabStripU.hWnd, StrPtr("DarkMode_DarkTheme"), 0
    SetWindowTheme cmdAbout.hWnd, StrPtr("DarkMode_DarkTheme"), 0
    SetWindowTheme chkOLEDragDrop.hWnd, StrPtr("DarkMode_DarkTheme"), 0
    SetWindowTheme ToolTip.hWnd, StrPtr("DarkMode_Explorer"), StrPtr("Tooltip")
    Me.BackColor = &H141414
  ElseIf themeableOS Then
    useImmersiveDarkMode = 0
    DwmSetWindowAttribute Me.hWnd, DWMWA_USE_IMMERSIVE_DARK_MODE, VarPtr(useImmersiveDarkMode), LenB(useImmersiveDarkMode)
    SetWindowTheme TabStripU.hWnd, StrPtr("Tab"), 0
    SetWindowTheme cmdAbout.hWnd, StrPtr("Button"), 0
    SetWindowTheme chkOLEDragDrop.hWnd, StrPtr("Button"), 0
    SetWindowTheme ToolTip.hWnd, StrPtr("Tooltip"), 0
    Me.BackColor = vbButtonFace
  End If
End Sub

Private Sub DrawMenuBarLine(ByVal hWnd As Long)
  Const DC_BRUSH = 18
  Const OBJID_MENU = &HFFFFFFFD
  Dim mnuBarInfo As MENUBARINFO
  Dim rcClient As RECT
  Dim rcWindow As RECT
  Dim rcAnnoyingLine As RECT
  Dim hDC As Long
  Dim previousColor As Long

  If bDarkModeActivated Then
    mnuBarInfo.cbSize = LenB(mnuBarInfo)
    If GetMenuBarInfo(hWnd, OBJID_MENU, 0, mnuBarInfo) Then
      GetClientRect hWnd, rcClient
      MapWindowPoints hWnd, 0, VarPtr(rcClient), 2
      GetWindowRect hWnd, rcWindow
      
      rcAnnoyingLine.Left = rcClient.Left - rcWindow.Left
      rcAnnoyingLine.Right = rcClient.Right - rcWindow.Left
      rcAnnoyingLine.Top = rcClient.Top - rcWindow.Top
      rcAnnoyingLine.Bottom = rcClient.Bottom - rcWindow.Top
      rcAnnoyingLine.Bottom = rcAnnoyingLine.Top
      rcAnnoyingLine.Top = rcAnnoyingLine.Top - 1
      
      hDC = GetWindowDC(hWnd)
      If hDC Then
        previousColor = SetDCBrushColor(hDC, &H191919)
        FillRect hDC, rcAnnoyingLine, GetStockObject(DC_BRUSH)
        SetDCBrushColor hDC, previousColor
        ReleaseDC hWnd, hDC
      End If
    End If
  End If
End Sub
