VERSION 5.00
Object = "{E7BB2F30-C5DD-4370-B7E2-19A7EDF169EE}#1.10#0"; "TabStripCtlU.ocx"
Object = "{3DF28271-B5ED-46F4-9EFF-E60A4CDD0D02}#1.10#0"; "TabStripCtlA.ocx"
Begin VB.Form frmMain 
   Caption         =   "TabStrip 1.10 - Events Sample"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   437
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   792
   StartUpPosition =   2  'Bildschirmmitte
   Begin TabStripCtlLibACtl.TabStrip TabStripA 
      Height          =   2895
      Left            =   0
      TabIndex        =   4
      Top             =   2880
      Width           =   7215
      _cx             =   12726
      _cy             =   5106
      AllowDragDrop   =   -1  'True
      Appearance      =   0
      BorderStyle     =   0
      CloseableTabs   =   0   'False
      DisabledEvents  =   0
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
      RegisterForOLEDragDrop=   1
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
   Begin VB.CheckBox chkLog 
      Caption         =   "&Log"
      Height          =   255
      Left            =   7320
      TabIndex        =   2
      Top             =   5400
      Width           =   975
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
      Left            =   8430
      TabIndex        =   3
      Top             =   5280
      Width           =   2415
   End
   Begin VB.TextBox txtLog 
      Height          =   4815
      Left            =   7320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
   Begin TabStripCtlLibUCtl.TabStrip TabStripU 
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      _cx             =   12726
      _cy             =   4895
      AllowDragDrop   =   -1  'True
      Appearance      =   0
      BorderStyle     =   0
      CloseableTabs   =   0   'False
      DisabledEvents  =   0
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
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Implements IHook
  Implements ISubclassedWindow


  Private Type DLLVERSIONINFO
    cbSize As Long
    dwMajor As Long
    dwMinor As Long
    dwBuildNumber As Long
    dwPlatformId As Long
  End Type

  Private bComctl32Version600OrNewer As Boolean
  Private bLog As Boolean
  Private hImgLst As Long
  Private objActiveCtl As Object


  Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal sz As Long)
  Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
  Private Declare Function DllGetVersion_comctl32 Lib "comctl32.dll" Alias "DllGetVersion" (Data As DLLVERSIONINFO) As Long
  Private Declare Function ImageList_AddIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal hIcon As Long) As Long
  Private Declare Function ImageList_Create Lib "comctl32.dll" (ByVal cx As Long, ByVal cy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
  Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal himl As Long) As Long
  Private Declare Function ImageList_GetImageCount Lib "comctl32.dll" (ByVal himl As Long) As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
  Private Declare Function LoadImage Lib "user32.dll" Alias "LoadImageW" (ByVal hInst As Long, ByVal lpszName As Long, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
  Private Declare Function SendMessageAsLong Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


Private Function IHook_KeyboardProcAfter(ByVal hookCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  '
End Function

Private Function IHook_KeyboardProcBefore(ByVal hookCode As Long, ByVal wParam As Long, ByVal lParam As Long, eatIt As Boolean) As Long
  Const KF_ALTDOWN = &H2000
  Const UIS_CLEAR = 2
  Const UISF_HIDEACCEL = &H2
  Const UISF_HIDEFOCUS = &H1
  Const VK_DOWN = &H28
  Const VK_LEFT = &H25
  Const VK_RIGHT = &H27
  Const VK_TAB = &H9
  Const VK_UP = &H26
  Const WM_CHANGEUISTATE = &H127

  If HiWord(lParam) And KF_ALTDOWN Then
    ' this will make keyboard and focus cues work
    SendMessageAsLong Me.hWnd, WM_CHANGEUISTATE, MakeDWord(UIS_CLEAR, UISF_HIDEACCEL Or UISF_HIDEFOCUS), 0
  Else
    Select Case wParam
      Case VK_LEFT, VK_RIGHT, VK_UP, VK_DOWN, VK_TAB
        ' this will make focus cues work
        SendMessageAsLong Me.hWnd, WM_CHANGEUISTATE, MakeDWord(UIS_CLEAR, UISF_HIDEFOCUS), 0
    End Select
  End If
End Function

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


Private Sub chkLog_Click()
  bLog = (chkLog.Value = CheckBoxConstants.vbChecked)
End Sub

Private Sub cmdAbout_Click()
  objActiveCtl.About
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
  chkLog.Value = CheckBoxConstants.vbChecked

  ' this is required to make the control work as expected
  Subclass
  InstallKeyboardHook Me

  InsertItemsA
  InsertItemsU
End Sub

Private Sub Form_Resize()
  Dim cx As Long

  If Me.WindowState <> vbMinimized Then
    cx = 0.4 * Me.ScaleWidth
    txtLog.Move Me.ScaleWidth - cx, 0, cx, Me.ScaleHeight - cmdAbout.Height - 10
    cmdAbout.Move txtLog.Left + (cx / 2) - cmdAbout.Width / 2, txtLog.Top + txtLog.Height + 5
    chkLog.Move txtLog.Left, cmdAbout.Top + 5
    TabStripU.Move 0, 0, txtLog.Left - 5, (Me.ScaleHeight - 5) / 2
    TabStripA.Move 0, TabStripU.Top + TabStripU.Height + 5, TabStripU.Width, TabStripU.Height
  End If
End Sub

Private Sub Form_Terminate()
  If hImgLst Then ImageList_Destroy hImgLst
End Sub

Private Sub Form_Unload(Cancel As Integer)
  RemoveKeyboardHook Me
  If Not UnSubclassWindow(Me.hWnd, EnumSubclassID.escidFrmMain) Then
    Debug.Print "UnSubclassing failed!"
  End If
End Sub

Private Sub TabStripA_AbortedDrag()
  AddLogEntry "TabStripA_AbortedDrag"
  Set TabStripA.DropHilitedTab = Nothing
End Sub

Private Sub TabStripA_ActiveTabChanged(ByVal previousActiveTab As TabStripCtlLibACtl.ITabStripTab, ByVal newActiveTab As TabStripCtlLibACtl.ITabStripTab)
  Dim str As String

  str = "TabStripA_ActiveTabChanged: previousActiveTab="
  If previousActiveTab Is Nothing Then
    str = str & "Nothing, newCaretTab="
  Else
    str = str & previousActiveTab.Text & ", newActiveTab="
  End If
  If newActiveTab Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & newActiveTab.Text
  End If

  AddLogEntry str
End Sub

Private Sub TabStripA_ActiveTabChanging(ByVal previousActiveTab As TabStripCtlLibACtl.ITabStripTab, cancelChange As Boolean)
  If previousActiveTab Is Nothing Then
    AddLogEntry "TabStripA_ActiveTabChanging: previousActiveTab=Nothing, cancelChange=" & cancelChange
  Else
    AddLogEntry "TabStripA_ActiveTabChanging: previousActiveTab=" & previousActiveTab.Text & ", cancelChange=" & cancelChange
  End If
End Sub

Private Sub TabStripA_CaretChanged(ByVal previousCaretTab As TabStripCtlLibACtl.ITabStripTab, ByVal newCaretTab As TabStripCtlLibACtl.ITabStripTab)
  Dim str As String

  str = "TabStripA_CaretChanged: previousCaretTab="
  If previousCaretTab Is Nothing Then
    str = str & "Nothing, newCaretTab="
  Else
    str = str & previousCaretTab.Text & ", newCaretTab="
  End If
  If newCaretTab Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & newCaretTab.Text
  End If

  AddLogEntry str
End Sub

Private Sub TabStripA_Click(ByVal tsTab As TabStripCtlLibACtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibACtl.HitTestConstants)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripA_Click: tsTab=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "TabStripA_Click: tsTab=" & tsTab.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub TabStripA_ContextMenu(ByVal tsTab As TabStripCtlLibACtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibACtl.HitTestConstants)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripA_ContextMenu: tsTab=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "TabStripA_ContextMenu: tsTab=" & tsTab.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub TabStripA_DblClick(ByVal tsTab As TabStripCtlLibACtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibACtl.HitTestConstants)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripA_DblClick: tsTab=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "TabStripA_DblClick: tsTab=" & tsTab.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub TabStripA_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "TabStripA_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub TabStripA_DragDrop(Source As Control, x As Single, y As Single)
  AddLogEntry "TabStripA_DragDrop"
End Sub

Private Sub TabStripA_DragMouseMove(dropTarget As TabStripCtlLibACtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibACtl.HitTestConstants, autoActivateTab As Boolean, autoScrollVelocity As Long)
  If dropTarget Is Nothing Then
    AddLogEntry "TabStripA_DragMouseMove: dropTarget=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails & ", autoActivateTab=" & autoActivateTab & ", autoScrollVelocity=" & autoScrollVelocity
  Else
    AddLogEntry "TabStripA_DragMouseMove: dropTarget=" & dropTarget.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails & ", autoActivateTab=" & autoActivateTab & ", autoScrollVelocity=" & autoScrollVelocity
  End If
End Sub

Private Sub TabStripA_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  AddLogEntry "TabStripA_DragOver"
End Sub

Private Sub TabStripA_Drop(ByVal dropTarget As TabStripCtlLibACtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibACtl.HitTestConstants)
  If dropTarget Is Nothing Then
    AddLogEntry "TabStripA_Drop: dropTarget=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "TabStripA_Drop: dropTarget=" & dropTarget.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
  Set TabStripA.DropHilitedTab = Nothing
End Sub

Private Sub TabStripA_FreeTabData(ByVal tsTab As TabStripCtlLibACtl.ITabStripTab)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripA_FreeTabData: tsTab=Nothing"
  Else
    AddLogEntry "TabStripA_FreeTabData: tsTab=" & tsTab.Text
  End If
End Sub

Private Sub TabStripA_GotFocus()
  AddLogEntry "TabStripA_GotFocus"
  Set objActiveCtl = TabStripA
End Sub

Private Sub TabStripA_InsertedTab(ByVal tsTab As TabStripCtlLibACtl.ITabStripTab)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripA_InsertedTab: tsTab=Nothing"
  Else
    AddLogEntry "TabStripA_InsertedTab: tsTab=" & tsTab.Text
  End If
End Sub

Private Sub TabStripA_InsertingTab(ByVal tsTab As TabStripCtlLibACtl.IVirtualTabStripTab, cancelInsertion As Boolean)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripA_InsertingTab: tsTab=Nothing, cancelInsertion=" & cancelInsertion
  Else
    AddLogEntry "TabStripA_InsertingTab: tsTab=" & tsTab.Text & ", cancelInsertion=" & cancelInsertion
  End If
End Sub

Private Sub TabStripA_KeyDown(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "TabStripA_KeyDown: keyCode=" & keyCode & ", shift=" & shift
  If keyCode = KeyCodeConstants.vbKeyEscape Then
    'If Not (ExTVwA.DraggedItems Is Nothing) Then ExTVwA.EndDrag True
  End If
End Sub

Private Sub TabStripA_KeyPress(keyAscii As Integer)
  AddLogEntry "TabStripA_KeyPress: keyAscii=" & keyAscii
End Sub

Private Sub TabStripA_KeyUp(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "TabStripA_KeyUp: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub TabStripA_LostFocus()
  AddLogEntry "TabStripA_LostFocus"
End Sub

Private Sub TabStripA_MClick(ByVal tsTab As TabStripCtlLibACtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibACtl.HitTestConstants)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripA_MClick: tsTab=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "TabStripA_MClick: tsTab=" & tsTab.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub TabStripA_MDblClick(ByVal tsTab As TabStripCtlLibACtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibACtl.HitTestConstants)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripA_MDblClick: tsTab=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "TabStripA_MDblClick: tsTab=" & tsTab.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub TabStripA_MouseDown(ByVal tsTab As TabStripCtlLibACtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibACtl.HitTestConstants)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripA_MouseDown: tsTab=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "TabStripA_MouseDown: tsTab=" & tsTab.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub TabStripA_MouseEnter(ByVal tsTab As TabStripCtlLibACtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibACtl.HitTestConstants)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripA_MouseEnter: tsTab=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "TabStripA_MouseEnter: tsTab=" & tsTab.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub TabStripA_MouseHover(ByVal tsTab As TabStripCtlLibACtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibACtl.HitTestConstants)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripA_MouseHover: tsTab=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "TabStripA_MouseHover: tsTab=" & tsTab.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub TabStripA_MouseLeave(ByVal tsTab As TabStripCtlLibACtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibACtl.HitTestConstants)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripA_MouseLeave: tsTab=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "TabStripA_MouseLeave: tsTab=" & tsTab.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub TabStripA_MouseMove(ByVal tsTab As TabStripCtlLibACtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibACtl.HitTestConstants)
'  If tsTab Is Nothing Then
'    AddLogEntry "TabStripA_MouseMove: tsTab=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
'  Else
'    AddLogEntry "TabStripA_MouseMove: tsTab=" & tsTab.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
'  End If
End Sub

Private Sub TabStripA_MouseUp(ByVal tsTab As TabStripCtlLibACtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibACtl.HitTestConstants)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripA_MouseUp: tsTab=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "TabStripA_MouseUp: tsTab=" & tsTab.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub TabStripA_OLECompleteDrag(ByVal Data As TabStripCtlLibACtl.IOLEDataObject, ByVal performedEffect As TabStripCtlLibACtl.OLEDropEffectConstants)
  Dim files() As String
  Dim str As String

  str = "TabStripA_OLECompleteDrag: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", performedEffect=" & performedEffect

  AddLogEntry str

  If Data.GetFormat(vbCFFiles) Then
    files = Data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    MsgBox "Dragged files:" & vbNewLine & str
  End If
End Sub

Private Sub TabStripA_OLEDragDrop(ByVal Data As TabStripCtlLibACtl.IOLEDataObject, effect As TabStripCtlLibACtl.OLEDropEffectConstants, ByVal dropTarget As TabStripCtlLibACtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibACtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "TabStripA_OLEDragDrop: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str

  If Data.GetFormat(vbCFFiles) Then
    files = Data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    TabStripA.FinishOLEDragDrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub TabStripA_OLEDragEnter(ByVal Data As TabStripCtlLibACtl.IOLEDataObject, effect As TabStripCtlLibACtl.OLEDropEffectConstants, dropTarget As TabStripCtlLibACtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibACtl.HitTestConstants, autoActivateTab As Boolean, autoScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "TabStripA_OLEDragEnter: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails & ", autoActivateTab=" & autoActivateTab & ", autoScrollVelocity=" & autoScrollVelocity

  AddLogEntry str
End Sub

Private Sub TabStripA_OLEDragEnterPotentialTarget(ByVal hWndPotentialTarget As Long)
  AddLogEntry "TabStripA_OLEDragEnterPotentialTarget: hWndPotentialTarget=0x" & Hex(hWndPotentialTarget)
End Sub

Private Sub TabStripA_OLEDragLeave(ByVal Data As TabStripCtlLibACtl.IOLEDataObject, ByVal dropTarget As TabStripCtlLibACtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibACtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "TabStripA_OLEDragLeave: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str
End Sub

Private Sub TabStripA_OLEDragLeavePotentialTarget()
  AddLogEntry "TabStripA_OLEDragLeavePotentialTarget"
End Sub

Private Sub TabStripA_OLEDragMouseMove(ByVal Data As TabStripCtlLibACtl.IOLEDataObject, effect As TabStripCtlLibACtl.OLEDropEffectConstants, dropTarget As TabStripCtlLibACtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibACtl.HitTestConstants, autoActivateTab As Boolean, autoScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "TabStripA_OLEDragMouseMove: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails & ", autoActivateTab=" & autoActivateTab & ", autoScrollVelocity=" & autoScrollVelocity

  AddLogEntry str
End Sub

Private Sub TabStripA_OLEGiveFeedback(ByVal effect As TabStripCtlLibACtl.OLEDropEffectConstants, useDefaultCursors As Boolean)
  AddLogEntry "TabStripA_OLEGiveFeedback: effect=" & effect & ", useDefaultCursors=" & useDefaultCursors
End Sub

Private Sub TabStripA_OLEQueryContinueDrag(ByVal pressedEscape As Boolean, ByVal button As Integer, ByVal shift As Integer, actionToContinueWith As TabStripCtlLibACtl.OLEActionToContinueWithConstants)
  AddLogEntry "TabStripA_OLEQueryContinueDrag: pressedEscape=" & pressedEscape & ", button=" & button & ", shift=" & shift & ", actionToContinueWith=0x" & Hex(actionToContinueWith)
End Sub

Private Sub TabStripA_OLEReceivedNewData(ByVal data As TabStripCtlLibACtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim str As String

  str = "TabStripA_OLEReceivedNewData: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", formatID=" & formatID & ", index=" & Index & ", dataOrViewAspect=" & dataOrViewAspect

  AddLogEntry str
End Sub

Private Sub TabStripA_OLESetData(ByVal Data As TabStripCtlLibACtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim str As String

  str = "TabStripA_OLESetData: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", formatID=" & formatID & ", index=" & Index & ", dataOrViewAspect=" & dataOrViewAspect

  AddLogEntry str
End Sub

Private Sub TabStripA_OLEStartDrag(ByVal Data As TabStripCtlLibACtl.IOLEDataObject)
  Dim files() As String
  Dim str As String

  str = "TabStripA_OLEStartDrag: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If

  AddLogEntry str
End Sub

Private Sub TabStripA_OwnerDrawTab(ByVal tsTab As TabStripCtlLibACtl.ITabStripTab, ByVal tabState As TabStripCtlLibACtl.OwnerDrawTabStateConstants, ByVal hDC As Long, drawingRectangle As TabStripCtlLibACtl.RECTANGLE)
  Dim str As String

  str = "TabStripA_OwnerDrawTab: tsTab="
  If tsTab Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & tsTab.Text
  End If
  str = str & ", tabState=" & tabState & ", hDC=" & hDC & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & ")"

  AddLogEntry str
End Sub

Private Sub TabStripA_RClick(ByVal tsTab As TabStripCtlLibACtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibACtl.HitTestConstants)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripA_RClick: tsTab=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "TabStripA_RClick: tsTab=" & tsTab.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub TabStripA_RDblClick(ByVal tsTab As TabStripCtlLibACtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibACtl.HitTestConstants)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripA_RDblClick: tsTab=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "TabStripA_RDblClick: tsTab=" & tsTab.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub TabStripA_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "TabStripA_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
  InsertItemsA
End Sub

Private Sub TabStripA_RemovedTab(ByVal tsTab As TabStripCtlLibACtl.IVirtualTabStripTab)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripA_RemovedTab: tsTab=Nothing"
  Else
    AddLogEntry "TabStripA_RemovedTab: tsTab=" & tsTab.Text
  End If
End Sub

Private Sub TabStripA_RemovingTab(ByVal tsTab As TabStripCtlLibACtl.ITabStripTab, cancelDeletion As Boolean)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripA_RemovingTab: tsTab=Nothing, cancelDeletion=" & cancelDeletion
  Else
    AddLogEntry "TabStripA_RemovingTab: tsTab=" & tsTab.Text & ", cancelDeletion=" & cancelDeletion
  End If
End Sub

Private Sub TabStripA_ResizedControlWindow()
  AddLogEntry "TabStripA_ResizedControlWindow"
End Sub

Private Sub TabStripA_TabBeginDrag(ByVal tsTab As TabStripCtlLibACtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibACtl.HitTestConstants)
  Dim col As TabStripCtlLibACtl.TabStripTabs

  If tsTab Is Nothing Then
    AddLogEntry "TabStripA_TabBeginDrag: tsTab=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "TabStripA_TabBeginDrag: tsTab=" & tsTab.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If

  On Error Resume Next
  If TabStripA.MultiSelect Then
    Set col = TabStripA.Tabs
    col.FilterType(TabStripCtlLibACtl.FilteredPropertyConstants.fpSelected) = TabStripCtlLibACtl.FilterTypeConstants.ftIncluding
    col.Filter(TabStripCtlLibACtl.FilteredPropertyConstants.fpSelected) = Array(True)

    TabStripA.OLEDrag , , , TabStripA.CreateTabContainer(col)
  Else
    TabStripA.OLEDrag , , , TabStripA.CreateTabContainer(tsTab)
  End If
End Sub

Private Sub TabStripA_TabBeginRDrag(ByVal tsTab As TabStripCtlLibACtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibACtl.HitTestConstants)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripA_TabBeginRDrag: tsTab=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "TabStripA_TabBeginRDrag: tsTab=" & tsTab.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub TabStripA_TabGetInfoTipText(ByVal tsTab As TabStripCtlLibACtl.ITabStripTab, ByVal maxInfoTipLength As Long, infoTipText As String)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripA_TabGetInfoTipText: tsTab=Nothing, maxInfoTipLength=" & maxInfoTipLength & ", infoTipText=" & infoTipText
  Else
    AddLogEntry "TabStripA_TabGetInfoTipText: tsTab=" & tsTab.Text & ", maxInfoTipLength=" & maxInfoTipLength & ", infoTipText=" & infoTipText
    infoTipText = "ID: " & tsTab.ID & ", TabData: " & tsTab.TabData
  End If
End Sub

Private Sub TabStripA_TabMouseEnter(ByVal tsTab As TabStripCtlLibACtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibACtl.HitTestConstants)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripA_TabMouseEnter: tsTab=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "TabStripA_TabMouseEnter: tsTab=" & tsTab.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub TabStripA_TabMouseLeave(ByVal tsTab As TabStripCtlLibACtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibACtl.HitTestConstants)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripA_TabMouseLeave: tsTab=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "TabStripA_TabMouseLeave: tsTab=" & tsTab.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub TabStripA_TabSelectionChanged(ByVal tsTab As TabStripCtlLibACtl.ITabStripTab)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripA_TabSelectionChanged: tsTab=Nothing"
  Else
    AddLogEntry "TabStripA_TabSelectionChanged: tsTab=" & tsTab.Text
  End If
End Sub

Private Sub TabStripA_Validate(Cancel As Boolean)
  AddLogEntry "TabStripA_Validate"
End Sub

Private Sub TabStripA_XClick(ByVal tsTab As TabStripCtlLibACtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibACtl.HitTestConstants)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripA_XClick: tsTab=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "TabStripA_XClick: tsTab=" & tsTab.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub TabStripA_XDblClick(ByVal tsTab As TabStripCtlLibACtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibACtl.HitTestConstants)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripA_XDblClick: tsTab=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "TabStripA_XDblClick: tsTab=" & tsTab.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub TabStripU_AbortedDrag()
  AddLogEntry "TabStripU_AbortedDrag"
  Set TabStripU.DropHilitedTab = Nothing
End Sub

Private Sub TabStripU_ActiveTabChanged(ByVal previousActiveTab As TabStripCtlLibUCtl.ITabStripTab, ByVal newActiveTab As TabStripCtlLibUCtl.ITabStripTab)
  Dim str As String

  str = "TabStripU_ActiveTabChanged: previousActiveTab="
  If previousActiveTab Is Nothing Then
    str = str & "Nothing, newCaretTab="
  Else
    str = str & previousActiveTab.Text & ", newActiveTab="
  End If
  If newActiveTab Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & newActiveTab.Text
  End If

  AddLogEntry str
End Sub

Private Sub TabStripU_ActiveTabChanging(ByVal previousActiveTab As TabStripCtlLibUCtl.ITabStripTab, cancelChange As Boolean)
  If previousActiveTab Is Nothing Then
    AddLogEntry "TabStripU_ActiveTabChanging: previousActiveTab=Nothing, cancelChange=" & cancelChange
  Else
    AddLogEntry "TabStripU_ActiveTabChanging: previousActiveTab=" & previousActiveTab.Text & ", cancelChange=" & cancelChange
  End If
End Sub

Private Sub TabStripU_CaretChanged(ByVal previousCaretTab As TabStripCtlLibUCtl.ITabStripTab, ByVal newCaretTab As TabStripCtlLibUCtl.ITabStripTab)
  Dim str As String

  str = "TabStripU_CaretChanged: previousCaretTab="
  If previousCaretTab Is Nothing Then
    str = str & "Nothing, newCaretTab="
  Else
    str = str & previousCaretTab.Text & ", newCaretTab="
  End If
  If newCaretTab Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & newCaretTab.Text
  End If

  AddLogEntry str
End Sub

Private Sub TabStripU_Click(ByVal tsTab As TabStripCtlLibUCtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibUCtl.HitTestConstants)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripU_Click: tsTab=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "TabStripU_Click: tsTab=" & tsTab.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub TabStripU_ContextMenu(ByVal tsTab As TabStripCtlLibUCtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibUCtl.HitTestConstants)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripU_ContextMenu: tsTab=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "TabStripU_ContextMenu: tsTab=" & tsTab.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub TabStripU_DblClick(ByVal tsTab As TabStripCtlLibUCtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibUCtl.HitTestConstants)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripU_DblClick: tsTab=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "TabStripU_DblClick: tsTab=" & tsTab.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub TabStripU_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "TabStripU_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub TabStripU_DragDrop(Source As Control, x As Single, y As Single)
  AddLogEntry "TabStripU_DragDrop"
End Sub

Private Sub TabStripU_DragMouseMove(dropTarget As TabStripCtlLibUCtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibUCtl.HitTestConstants, autoActivateTab As Boolean, autoScrollVelocity As Long)
  If dropTarget Is Nothing Then
    AddLogEntry "TabStripU_DragMouseMove: dropTarget=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails & ", autoActivateTab=" & autoActivateTab & ", autoScrollVelocity=" & autoScrollVelocity
  Else
    AddLogEntry "TabStripU_DragMouseMove: dropTarget=" & dropTarget.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails & ", autoActivateTab=" & autoActivateTab & ", autoScrollVelocity=" & autoScrollVelocity
  End If
End Sub

Private Sub TabStripU_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  AddLogEntry "TabStripU_DragOver"
End Sub

Private Sub TabStripU_Drop(ByVal dropTarget As TabStripCtlLibUCtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibUCtl.HitTestConstants)
  If dropTarget Is Nothing Then
    AddLogEntry "TabStripU_Drop: dropTarget=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "TabStripU_Drop: dropTarget=" & dropTarget.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
  Set TabStripU.DropHilitedTab = Nothing
End Sub

Private Sub TabStripU_FreeTabData(ByVal tsTab As TabStripCtlLibUCtl.ITabStripTab)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripU_FreeTabData: tsTab=Nothing"
  Else
    AddLogEntry "TabStripU_FreeTabData: tsTab=" & tsTab.Text
  End If
End Sub

Private Sub TabStripU_GotFocus()
  AddLogEntry "TabStripU_GotFocus"
  Set objActiveCtl = TabStripU
End Sub

Private Sub TabStripU_InsertedTab(ByVal tsTab As TabStripCtlLibUCtl.ITabStripTab)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripU_InsertedTab: tsTab=Nothing"
  Else
    AddLogEntry "TabStripU_InsertedTab: tsTab=" & tsTab.Text
  End If
End Sub

Private Sub TabStripU_InsertingTab(ByVal tsTab As TabStripCtlLibUCtl.IVirtualTabStripTab, cancelInsertion As Boolean)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripU_InsertingTab: tsTab=Nothing, cancelInsertion=" & cancelInsertion
  Else
    AddLogEntry "TabStripU_InsertingTab: tsTab=" & tsTab.Text & ", cancelInsertion=" & cancelInsertion
  End If
End Sub

Private Sub TabStripU_KeyDown(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "TabStripU_KeyDown: keyCode=" & keyCode & ", shift=" & shift
  If keyCode = KeyCodeConstants.vbKeyEscape Then
    'If Not (ExTVwA.DraggedItems Is Nothing) Then ExTVwA.EndDrag True
  End If
End Sub

Private Sub TabStripU_KeyPress(keyAscii As Integer)
  AddLogEntry "TabStripU_KeyPress: keyAscii=" & keyAscii
End Sub

Private Sub TabStripU_KeyUp(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "TabStripU_KeyUp: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub TabStripU_LostFocus()
  AddLogEntry "TabStripU_LostFocus"
End Sub

Private Sub TabStripU_MClick(ByVal tsTab As TabStripCtlLibUCtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibUCtl.HitTestConstants)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripU_MClick: tsTab=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "TabStripU_MClick: tsTab=" & tsTab.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub TabStripU_MDblClick(ByVal tsTab As TabStripCtlLibUCtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibUCtl.HitTestConstants)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripU_MDblClick: tsTab=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "TabStripU_MDblClick: tsTab=" & tsTab.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub TabStripU_MouseDown(ByVal tsTab As TabStripCtlLibUCtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibUCtl.HitTestConstants)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripU_MouseDown: tsTab=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "TabStripU_MouseDown: tsTab=" & tsTab.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub TabStripU_MouseEnter(ByVal tsTab As TabStripCtlLibUCtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibUCtl.HitTestConstants)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripU_MouseEnter: tsTab=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "TabStripU_MouseEnter: tsTab=" & tsTab.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub TabStripU_MouseHover(ByVal tsTab As TabStripCtlLibUCtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibUCtl.HitTestConstants)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripU_MouseHover: tsTab=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "TabStripU_MouseHover: tsTab=" & tsTab.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub TabStripU_MouseLeave(ByVal tsTab As TabStripCtlLibUCtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibUCtl.HitTestConstants)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripU_MouseLeave: tsTab=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "TabStripU_MouseLeave: tsTab=" & tsTab.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub TabStripU_MouseMove(ByVal tsTab As TabStripCtlLibUCtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibUCtl.HitTestConstants)
'  If tsTab Is Nothing Then
'    AddLogEntry "TabStripU_MouseMove: tsTab=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
'  Else
'    AddLogEntry "TabStripU_MouseMove: tsTab=" & tsTab.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
'  End If
End Sub

Private Sub TabStripU_MouseUp(ByVal tsTab As TabStripCtlLibUCtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibUCtl.HitTestConstants)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripU_MouseUp: tsTab=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "TabStripU_MouseUp: tsTab=" & tsTab.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub TabStripU_OLECompleteDrag(ByVal Data As TabStripCtlLibUCtl.IOLEDataObject, ByVal performedEffect As TabStripCtlLibUCtl.OLEDropEffectConstants)
  Dim files() As String
  Dim str As String

  str = "TabStripU_OLECompleteDrag: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", performedEffect=" & performedEffect

  AddLogEntry str

  If Data.GetFormat(vbCFFiles) Then
    files = Data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    MsgBox "Dragged files:" & vbNewLine & str
  End If
End Sub

Private Sub TabStripU_OLEDragDrop(ByVal Data As TabStripCtlLibUCtl.IOLEDataObject, effect As TabStripCtlLibUCtl.OLEDropEffectConstants, ByVal dropTarget As TabStripCtlLibUCtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "TabStripU_OLEDragDrop: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str

  If Data.GetFormat(vbCFFiles) Then
    files = Data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    TabStripU.FinishOLEDragDrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub TabStripU_OLEDragEnter(ByVal Data As TabStripCtlLibUCtl.IOLEDataObject, effect As TabStripCtlLibUCtl.OLEDropEffectConstants, dropTarget As TabStripCtlLibUCtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibUCtl.HitTestConstants, autoActivateTab As Boolean, autoScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "TabStripU_OLEDragEnter: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails & ", autoActivateTab=" & autoActivateTab & ", autoScrollVelocity=" & autoScrollVelocity

  AddLogEntry str
End Sub

Private Sub TabStripU_OLEDragEnterPotentialTarget(ByVal hWndPotentialTarget As Long)
  AddLogEntry "TabStripU_OLEDragEnterPotentialTarget: hWndPotentialTarget=0x" & Hex(hWndPotentialTarget)
End Sub

Private Sub TabStripU_OLEDragLeave(ByVal Data As TabStripCtlLibUCtl.IOLEDataObject, ByVal dropTarget As TabStripCtlLibUCtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "TabStripU_OLEDragLeave: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str
End Sub

Private Sub TabStripU_OLEDragLeavePotentialTarget()
  AddLogEntry "TabStripU_OLEDragLeavePotentialTarget"
End Sub

Private Sub TabStripU_OLEDragMouseMove(ByVal Data As TabStripCtlLibUCtl.IOLEDataObject, effect As TabStripCtlLibUCtl.OLEDropEffectConstants, dropTarget As TabStripCtlLibUCtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibUCtl.HitTestConstants, autoActivateTab As Boolean, autoScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "TabStripU_OLEDragMouseMove: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails & ", autoActivateTab=" & autoActivateTab & ", autoScrollVelocity=" & autoScrollVelocity

  AddLogEntry str
End Sub

Private Sub TabStripU_OLEGiveFeedback(ByVal effect As TabStripCtlLibUCtl.OLEDropEffectConstants, useDefaultCursors As Boolean)
  AddLogEntry "TabStripU_OLEGiveFeedback: effect=" & effect & ", useDefaultCursors=" & useDefaultCursors
End Sub

Private Sub TabStripU_OLEQueryContinueDrag(ByVal pressedEscape As Boolean, ByVal button As Integer, ByVal shift As Integer, actionToContinueWith As TabStripCtlLibUCtl.OLEActionToContinueWithConstants)
  AddLogEntry "TabStripU_OLEQueryContinueDrag: pressedEscape=" & pressedEscape & ", button=" & button & ", shift=" & shift & ", actionToContinueWith=0x" & Hex(actionToContinueWith)
End Sub

Private Sub TabStripU_OLEReceivedNewData(ByVal data As TabStripCtlLibUCtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim str As String

  str = "TabStripU_OLEReceivedNewData: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", formatID=" & formatID & ", index=" & Index & ", dataOrViewAspect=" & dataOrViewAspect

  AddLogEntry str
End Sub

Private Sub TabStripU_OLESetData(ByVal Data As TabStripCtlLibUCtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim str As String

  str = "TabStripU_OLESetData: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", formatID=" & formatID & ", index=" & Index & ", dataOrViewAspect=" & dataOrViewAspect

  AddLogEntry str
End Sub

Private Sub TabStripU_OLEStartDrag(ByVal Data As TabStripCtlLibUCtl.IOLEDataObject)
  Dim files() As String
  Dim str As String

  str = "TabStripU_OLEStartDrag: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If

  AddLogEntry str
End Sub

Private Sub TabStripU_OwnerDrawTab(ByVal tsTab As TabStripCtlLibUCtl.ITabStripTab, ByVal tabState As TabStripCtlLibUCtl.OwnerDrawTabStateConstants, ByVal hDC As Long, drawingRectangle As TabStripCtlLibUCtl.RECTANGLE)
  Dim str As String

  str = "TabStripU_OwnerDrawTab: tsTab="
  If tsTab Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & tsTab.Text
  End If
  str = str & ", tabState=" & tabState & ", hDC=" & hDC & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & ")"

  AddLogEntry str
End Sub

Private Sub TabStripU_RClick(ByVal tsTab As TabStripCtlLibUCtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibUCtl.HitTestConstants)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripU_RClick: tsTab=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "TabStripU_RClick: tsTab=" & tsTab.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub TabStripU_RDblClick(ByVal tsTab As TabStripCtlLibUCtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibUCtl.HitTestConstants)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripU_RDblClick: tsTab=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "TabStripU_RDblClick: tsTab=" & tsTab.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub TabStripU_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "TabStripU_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
  InsertItemsU
End Sub

Private Sub TabStripU_RemovedTab(ByVal tsTab As TabStripCtlLibUCtl.IVirtualTabStripTab)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripU_RemovedTab: tsTab=Nothing"
  Else
    AddLogEntry "TabStripU_RemovedTab: tsTab=" & tsTab.Text
  End If
End Sub

Private Sub TabStripU_RemovingTab(ByVal tsTab As TabStripCtlLibUCtl.ITabStripTab, cancelDeletion As Boolean)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripU_RemovingTab: tsTab=Nothing, cancelDeletion=" & cancelDeletion
  Else
    AddLogEntry "TabStripU_RemovingTab: tsTab=" & tsTab.Text & ", cancelDeletion=" & cancelDeletion
  End If
End Sub

Private Sub TabStripU_ResizedControlWindow()
  AddLogEntry "TabStripU_ResizedControlWindow"
End Sub

Private Sub TabStripU_TabBeginDrag(ByVal tsTab As TabStripCtlLibUCtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibUCtl.HitTestConstants)
  Dim col As TabStripCtlLibUCtl.TabStripTabs

  If tsTab Is Nothing Then
    AddLogEntry "TabStripU_TabBeginDrag: tsTab=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "TabStripU_TabBeginDrag: tsTab=" & tsTab.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If

  On Error Resume Next
  If TabStripU.MultiSelect Then
    Set col = TabStripU.Tabs
    col.FilterType(TabStripCtlLibUCtl.FilteredPropertyConstants.fpSelected) = TabStripCtlLibUCtl.FilterTypeConstants.ftIncluding
    col.Filter(TabStripCtlLibUCtl.FilteredPropertyConstants.fpSelected) = Array(True)

    TabStripU.OLEDrag , , , TabStripU.CreateTabContainer(col)
  Else
    TabStripU.OLEDrag , , , TabStripU.CreateTabContainer(tsTab)
  End If
End Sub

Private Sub TabStripU_TabBeginRDrag(ByVal tsTab As TabStripCtlLibUCtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibUCtl.HitTestConstants)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripU_TabBeginRDrag: tsTab=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "TabStripU_TabBeginRDrag: tsTab=" & tsTab.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub TabStripU_TabGetInfoTipText(ByVal tsTab As TabStripCtlLibUCtl.ITabStripTab, ByVal maxInfoTipLength As Long, infoTipText As String)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripU_TabGetInfoTipText: tsTab=Nothing, maxInfoTipLength=" & maxInfoTipLength & ", infoTipText=" & infoTipText
  Else
    AddLogEntry "TabStripU_TabGetInfoTipText: tsTab=" & tsTab.Text & ", maxInfoTipLength=" & maxInfoTipLength & ", infoTipText=" & infoTipText
    infoTipText = "ID: " & tsTab.ID & ", TabData: " & tsTab.TabData
  End If
End Sub

Private Sub TabStripU_TabMouseEnter(ByVal tsTab As TabStripCtlLibUCtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibUCtl.HitTestConstants)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripU_TabMouseEnter: tsTab=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "TabStripU_TabMouseEnter: tsTab=" & tsTab.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub TabStripU_TabMouseLeave(ByVal tsTab As TabStripCtlLibUCtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibUCtl.HitTestConstants)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripU_TabMouseLeave: tsTab=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "TabStripU_TabMouseLeave: tsTab=" & tsTab.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub TabStripU_TabSelectionChanged(ByVal tsTab As TabStripCtlLibUCtl.ITabStripTab)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripU_TabSelectionChanged: tsTab=Nothing"
  Else
    AddLogEntry "TabStripU_TabSelectionChanged: tsTab=" & tsTab.Text
  End If
End Sub

Private Sub TabStripU_Validate(Cancel As Boolean)
  AddLogEntry "TabStripU_Validate"
End Sub

Private Sub TabStripU_XClick(ByVal tsTab As TabStripCtlLibUCtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibUCtl.HitTestConstants)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripU_XClick: tsTab=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "TabStripU_XClick: tsTab=" & tsTab.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub TabStripU_XDblClick(ByVal tsTab As TabStripCtlLibUCtl.ITabStripTab, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TabStripCtlLibUCtl.HitTestConstants)
  If tsTab Is Nothing Then
    AddLogEntry "TabStripU_XDblClick: tsTab=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "TabStripU_XDblClick: tsTab=" & tsTab.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub


Private Sub AddLogEntry(ByVal txt As String)
  Dim pos As Long
  Static cLines As Long
  Static oldtxt As String

  If bLog Then
    cLines = cLines + 1
    If cLines > 50 Then
      ' delete the first line
      pos = InStr(oldtxt, vbNewLine)
      oldtxt = Mid(oldtxt, pos + Len(vbNewLine))
      cLines = cLines - 1
    End If
    oldtxt = oldtxt & (txt & vbNewLine)

    txtLog.Text = oldtxt
    txtLog.SelStart = Len(oldtxt)
  End If
End Sub

' returns the higher 16 bits of <value>
Private Function HiWord(ByVal Value As Long) As Integer
  Dim ret As Integer

  CopyMemory VarPtr(ret), VarPtr(Value) + LenB(ret), LenB(ret)

  HiWord = ret
End Function

Private Sub InsertItemsA()
  Dim cImages As Long
  Dim iIcon As Long

  TabStripA.hImageList = hImgLst
  cImages = ImageList_GetImageCount(hImgLst)

  With TabStripA.Tabs
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

Private Sub InsertItemsU()
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

' makes a 32 bits number out of two 16 bits parts
Private Function MakeDWord(ByVal Lo As Integer, ByVal hi As Integer) As Long
  Dim ret As Long

  CopyMemory VarPtr(ret), VarPtr(Lo), LenB(Lo)
  CopyMemory VarPtr(ret) + LenB(Lo), VarPtr(hi), LenB(hi)

  MakeDWord = ret
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
    SendMessageAsLong TabStripA.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
  #End If
End Sub
