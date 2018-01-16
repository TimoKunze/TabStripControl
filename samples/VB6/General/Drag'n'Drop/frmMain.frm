VERSION 5.00
Object = "{E7BB2F30-C5DD-4370-B7E2-19A7EDF169EE}#1.10#0"; "TabStripCtlU.ocx"
Begin VB.Form frmMain 
   Caption         =   "TabStrip 1.10 - Drag'n'Drop Sample"
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
      TabIndex        =   2
      Top             =   0
      Width           =   4695
      _cx             =   8281
      _cy             =   7011
      AllowDragDrop   =   -1  'True
      Appearance      =   0
      BorderStyle     =   0
      CloseableTabs   =   0   'False
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
      Left            =   4800
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
      Left            =   4920
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


  Private CFSTR_HTML As Long
  Private CFSTR_HTML2 As Long
  Private CFSTR_LOGICALPERFORMEDDROPEFFECT As Long
  Private CFSTR_MOZILLAHTMLCONTEXT As Long
  Private CFSTR_PERFORMEDDROPEFFECT As Long
  Private CFSTR_TARGETCLSID As Long
  Private CFSTR_TEXTHTML As Long

  Private bComctl32Version600OrNewer As Boolean
  Private bRightDrag As Boolean
  Private chosenMenuItem As ChosenMenuItemConstants
  Private CLSID_RecycleBin As GUID
  Private hImgLst As Long
  Private lastDropTarget As Long
  Private ToolTip As clsToolTip


  Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal pString As Long, CLSID As GUID) As Long
  Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal sz As Long)
  Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
  Private Declare Function DllGetVersion_comctl32 Lib "comctl32.dll" Alias "DllGetVersion" (Data As DLLVERSIONINFO) As Long
  Private Declare Function ImageList_AddIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal hIcon As Long) As Long
  Private Declare Function ImageList_Create Lib "comctl32.dll" (ByVal cx As Long, ByVal cy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
  Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal himl As Long) As Long
  Private Declare Function ImageList_GetImageCount Lib "comctl32.dll" (ByVal himl As Long) As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
  Private Declare Function LoadImage Lib "user32.dll" Alias "LoadImageW" (ByVal hinst As Long, ByVal lpszName As Long, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
  Private Declare Function RegisterClipboardFormat Lib "user32.dll" Alias "RegisterClipboardFormatW" (ByVal lpString As Long) As Long
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
End Sub

Private Sub Form_Terminate()
  If hImgLst Then ImageList_Destroy hImgLst
  Set ToolTip = Nothing
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
  InsertTabs
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
  #End If
End Sub
