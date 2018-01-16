VERSION 5.00
Object = "{E7BB2F30-C5DD-4370-B7E2-19A7EDF169EE}#1.10#0"; "TabStripCtlU.ocx"
Begin VB.Form frmMain 
   Caption         =   "TabStrip 1.10 - Closeable Tabs Sample"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6420
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
   ScaleWidth      =   428
   StartUpPosition =   2  'Bildschirmmitte
   Begin TabStripCtlLibUCtl.TabStrip TabStripU 
      Align           =   3  'Links ausrichten
      Height          =   3975
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4695
      _cx             =   8281
      _cy             =   7011
      AllowDragDrop   =   0   'False
      Appearance      =   0
      BorderStyle     =   0
      CloseableTabs   =   -1  'True
      CloseableTabsMode=   1
      DisabledEvents  =   262377
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
      RegisterForOLEDragDrop=   0
      RightToLeft     =   0
      ScrollToOpposite=   0   'False
      ShowButtonSeparators=   -1  'True
      ShowToolTips    =   -1  'True
      Style           =   0
      SupportOLEDragImages=   -1  'True
      TabBoundingBoxDefinition=   8198
      TabCaptionAlignment=   2
      TabHeight       =   18
      TabPlacement    =   0
      UseFixedTabWidth=   -1  'True
      UseSystemFont   =   -1  'True
      VerticalPadding =   3
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
      Left            =   4920
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Implements ISubclassedWindow


  Private Type DLLVERSIONINFO
    cbSize As Long
    dwMajor As Long
    dwMinor As Long
    dwBuildNumber As Long
    dwPlatformId As Long
  End Type

  Private Type POINTAPI
    x As Long
    y As Long
  End Type


  Private bComctl32Version600OrNewer As Boolean
  Private hImgLst As Long


  Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
  Private Declare Function DllGetVersion_comctl32 Lib "comctl32.dll" Alias "DllGetVersion" (Data As DLLVERSIONINFO) As Long
  Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
  Private Declare Function ImageList_AddIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal hIcon As Long) As Long
  Private Declare Function ImageList_Create Lib "comctl32.dll" (ByVal cx As Long, ByVal cy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
  Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal himl As Long) As Long
  Private Declare Function ImageList_GetImageCount Lib "comctl32.dll" (ByVal himl As Long) As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
  Private Declare Function LoadImage Lib "user32.dll" Alias "LoadImageW" (ByVal hinst As Long, ByVal lpszName As Long, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
  Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
  Private Declare Function SendMessageAsLong Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


Private Function ISubclassedWindow_HandleMessage(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal eSubclassID As EnumSubclassID, bCallDefProc As Boolean) As Long
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  Select Case eSubclassID
    Case EnumSubclassID.escidFrmMain
      lRet = HandleMessage_Form(hWnd, uMsg, wParam, lParam, bCallDefProc)
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
  Const WM_NOTIFYFORMAT = &H55
  Const WM_USER = &H400
  Const OCM__BASE = WM_USER + &H1C00
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  Select Case uMsg
    Case WM_NOTIFYFORMAT
      ' give the control a chance to request Unicode notifications
      lRet = SendMessageAsLong(wParam, OCM__BASE + uMsg, wParam, lParam)

      bCallDefProc = False
  End Select

StdHandler_Ende:
  HandleMessage_Form = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in frmMain.HandleMessage_Form: ", Err.Number, Err.Description
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
  Dim DLLVerData As DLLVERSIONINFO
  Dim hIcon As Long
  Dim iconsDir As String
  Dim iconPath As String

  InitCommonControls

  With DLLVerData
    .cbSize = LenB(DLLVerData)
    DllGetVersion_comctl32 DLLVerData
    bComctl32Version600OrNewer = (.dwMajor >= 6)
  End With

  hImgLst = ImageList_Create(16, 16, IIf(bComctl32Version600OrNewer, ILC_COLOR32, ILC_COLOR24) Or ILC_MASK, 14, 0)
  If Right$(App.Path, 3) = "bin" Then
    iconsDir = App.Path & "\..\res\"
  Else
    iconsDir = App.Path & "\res\"
  End If
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

Private Sub Form_Load()
  ' this is required to make the control work as expected
  Subclass

  InsertTabs
End Sub

Private Sub Form_Terminate()
  If hImgLst Then ImageList_Destroy hImgLst
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not UnSubclassWindow(Me.hWnd, EnumSubclassID.escidFrmMain) Then
    Debug.Print "UnSubclassing failed!"
  End If
End Sub

Private Sub TabStripU_Click(ByVal tsTab As TabStripCtlLibUCtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibUCtl.HitTestConstants)
  Dim newActiveTab As Long

  If Not (tsTab Is Nothing) Then
    If hitTestDetails And HitTestConstants.htCloseButton Then
      If tsTab.Active Then
        newActiveTab = tsTab.Index + 1
        If newActiveTab = TabStripU.Tabs.Count Then
          newActiveTab = newActiveTab - 2
        End If
        If newActiveTab <> -1 Then
          Set TabStripU.ActiveTab = TabStripU.Tabs(newActiveTab)
        End If
      End If

      ' close the tab
      TabStripU.Tabs.Remove tsTab.Index
    End If
  End If
End Sub

Private Sub TabStripU_RecreatedControlWindow(ByVal hWnd As Long)
  InsertTabs
End Sub

Private Sub TabStripU_TabGetInfoTipText(ByVal tsTab As TabStripCtlLibUCtl.ITabStripTab, ByVal maxInfoTipLength As Long, infoTipText As String)
  Dim hitTestDetails As HitTestConstants
  Dim pt As POINTAPI

  GetCursorPos pt
  ScreenToClient TabStripU.hWnd, pt

  TabStripU.HitTest pt.x, pt.y, hitTestDetails
  If hitTestDetails And HitTestConstants.htCloseButton Then
    infoTipText = "Close this tab"
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
    .Add "Tab 6", , iIcon, 6
    iIcon = (iIcon + 1) Mod cImages
    .Add "Tab 7", , iIcon, 7
    iIcon = (iIcon + 1) Mod cImages
    .Add "Tab 8", , iIcon, 8
    iIcon = (iIcon + 1) Mod cImages
    .Add "Tab 9", , iIcon, 9
    iIcon = (iIcon + 1) Mod cImages
    .Add "Tab 10", , iIcon, 10
    iIcon = (iIcon + 1) Mod cImages
  End With
End Sub

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
  #End If
End Sub
