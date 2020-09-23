VERSION 5.00
Begin VB.UserControl TransForm 
   BackStyle       =   0  'Transparent
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   InvisibleAtRuntime=   -1  'True
   Picture         =   "TransForm.ctx":0000
   ScaleHeight     =   480
   ScaleWidth      =   480
   ToolboxBitmap   =   "TransForm.ctx":0282
End
Attribute VB_Name = "TransForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum AutoDragConstants
    Never
    Always
    WhenTransparent
End Enum

Private Type PointAPI
    X As Long
    Y As Long
End Type

Private Type RectAPI
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

Private meAutoDrag      As AutoDragConstants
Private mbTransparent   As Boolean
Private mbDisabled      As Boolean
Private mlMaskColor     As Long

Private WithEvents mParent  As Form
Attribute mParent.VB_VarHelpID = -1

Private Const meDefAutoDrag     As Integer = AutoDragConstants.WhenTransparent
Private Const mbDefTransparent  As Boolean = True
Private Const mlDefMaskColor    As Long = &HC0C0C0

Private Const NULLREGION        As Long = &H1&
Private Const SIMPLEREGION      As Long = &H2&
Private Const COMPLEXREGION     As Long = &H3&
Private Const HTCAPTION         As Long = &H2&
Private Const HTCLIENT          As Long = 1
Private Const WM_NCLBUTTONDOWN  As Long = &HA1&
Private Const TPM_RETURNCMD     As Long = &H100&
Private Const WM_SYSCOMMAND     As Long = &H112&
Private Const WM_LBUTTONUP      As Long = &H202&
Private Const WM_LBUTTONDOWN    As Long = &H201&

Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As PointAPI) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RectAPI) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Private Declare Function OffsetRgn Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As PointAPI) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hWnd As Long, lprc As RectAPI) As Long


Private Sub SetForm()

Dim bAutoRedraw     As Boolean
Dim lRet            As Long
Dim lErr            As Long
Dim lRgnType        As Long
Dim hWinRgn         As Long
Dim lOffsetX        As Long
Dim lOffsetY        As Long
Dim lScaleWidth     As Long
Dim lScaleHeight    As Long
Dim ptTemp          As PointAPI
Dim rcClient        As RectAPI

    If Not Ambient.UserMode Then
        GoTo NormalExit
    End If
    
    bAutoRedraw = mParent.AutoRedraw
    mParent.AutoRedraw = True
    
    If mbDisabled Or Not mbTransparent Then
        'Unshape the form
        GoTo UnsetForm
    End If
    
    'Make sure the form has a picture
    If mParent.Picture = 0 Or mParent.Picture Is Nothing Then
        GoTo UnsetForm
    Else
        'Get the picture from the form
        With UserControl
            .MaskColor = mlMaskColor
            .Width = .ScaleX(mParent.Picture.Width, vbHimetric, vbTwips)
            .Height = .ScaleY(mParent.Picture.Height, vbHimetric, vbTwips)
            Set .Picture = mParent.Picture
            Set .MaskPicture = mParent.Picture
            DoEvents
        End With
    End If
    
    'Create Dummy region to pass to GetWindowRgn, which
    'only accepts an existing application-defined region.
    hWinRgn = CreateRectRgn(0, 0, 200, 200)

    'Get the MaskedImage's Region
    lRgnType = GetWindowRgn(UserControl.hWnd, hWinRgn)
    Select Case lRgnType
        
        Case SIMPLEREGION, COMPLEXREGION
            With mParent
                
                'Get Form's Client Area
                lRet = GetClientRect(.hWnd, rcClient)
                lScaleWidth = (rcClient.Right - rcClient.Left + 1) * Screen.TwipsPerPixelX
                lScaleHeight = (rcClient.Bottom - rcClient.Top + 1) * Screen.TwipsPerPixelY
                
                'Resize Form to size of UserControl
                .Width = (.Width - lScaleWidth) + UserControl.Width + Screen.TwipsPerPixelX
                .Height = (.Height - lScaleHeight) + UserControl.Height + Screen.TwipsPerPixelY
                
                'Offset Region to allow for Form Border and Title Bar
                ptTemp.X = 0
                ptTemp.Y = 0
                lRet = ClientToScreen(.hWnd, ptTemp)
                lOffsetX = ptTemp.X - (.Left / Screen.TwipsPerPixelX)
                lOffsetY = ptTemp.Y - (.Top / Screen.TwipsPerPixelY)
                lRet = OffsetRgn(hWinRgn, lOffsetX, lOffsetY)
                
                'Shape the Form
                lRet = SetWindowRgn(.hWnd, hWinRgn, True)
                
            End With
        
        Case Else
            'Error or NULLREGION
            GoTo UnsetForm
    
    End Select

NormalExit:
    If Not mParent Is Nothing Then
        mParent.AutoRedraw = bAutoRedraw
    End If
    Exit Sub

UnsetForm:
    'UnShape the Form
    lRet = SetWindowRgn(mParent.hWnd, &H0&, True)
    GoTo NormalExit
    
LocalError:
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume NormalExit

End Sub

Public Property Let MaskColor(ByVal lData As OLE_COLOR)

    mlMaskColor = lData
    PropertyChanged "MaskColor"
    
End Property

Public Property Get MaskColor() As OLE_COLOR

    MaskColor = mlMaskColor
    
End Property

Public Property Let Transparent(ByVal bData As Boolean)

    mbTransparent = bData
    Call SetForm
    PropertyChanged "Transparent"
    
End Property

Public Property Get Transparent() As Boolean

    Transparent = mbTransparent
    
End Property

Public Property Let AutoDrag(ByVal eData As AutoDragConstants)

    meAutoDrag = eData
    PropertyChanged "AutoDrag"
    
End Property

Public Property Get AutoDrag() As AutoDragConstants

    AutoDrag = meAutoDrag
    
End Property

Private Sub mParent_Activate()

    Call SetForm
    
End Sub

Private Sub mParent_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim lCoords As Long
Dim ptStart As PointAPI
Dim ptMove  As PointAPI

    If (Button = vbLeftButton) And ((meAutoDrag = Always) Or _
      (meAutoDrag = WhenTransparent And mbTransparent)) Then
        ptStart.X = mParent.Left
        ptStart.Y = mParent.Top
        Call ReleaseCapture
        Call SendMessage(mParent.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, &H0&)
        'Send the MouseUp.
        Call SetCapture(mParent.hWnd)
        lCoords = mParent.ScaleX(X, mParent.ScaleMode, vbPixels) _
          + (mParent.ScaleY(Y, mParent.ScaleMode, vbPixels) * &H10000)
        Call PostMessage(mParent.hWnd, WM_LBUTTONUP, HTCLIENT, lCoords)
    End If
    
End Sub

Private Sub UserControl_InitProperties()

    meAutoDrag = meDefAutoDrag
    mbTransparent = mbDefTransparent
    mlMaskColor = mlDefMaskColor
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

Dim iCnt    As Integer
Dim lIdx    As Long

    meAutoDrag = PropBag.ReadProperty("AutoDrag", meDefAutoDrag)
    mbTransparent = PropBag.ReadProperty("Transparent", mbDefTransparent)
    mlMaskColor = PropBag.ReadProperty("MaskColor", mlDefMaskColor)
    
    If UserControl.Ambient.UserMode Then
        Set mParent = UserControl.Parent
        For lIdx = 0 To mParent.Controls.Count - 1
            If TypeOf mParent.Controls(lIdx) Is TransForm Then
                If iCnt > 0 Then
                    MsgBox "Only one TransForm control is allowed or " & _
                      "needed on a single form." & vbCrLf & vbCrLf & _
                      "All TransForm functionality will be disabled " & _
                      "for this form.", vbExclamation, "Multiple TransForm Controls"
                    mbDisabled = True
                    Exit For
                End If
                iCnt = iCnt + 1
            End If
        Next
    End If
    
    Call SetForm

End Sub

Private Sub UserControl_Resize()

    If Not Ambient.UserMode Then
        If UserControl.Width <> 480 Then
            UserControl.Width = 480
        ElseIf UserControl.Height <> 480 Then
            UserControl.Height = 480
        Else
            With UserControl
                Set .MaskPicture = .Picture
                .MaskColor = &HC0C0C0
            End With
        End If
    End If
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("AutoDrag", meAutoDrag, meDefAutoDrag)
    Call PropBag.WriteProperty("Transparent", mbTransparent, mbDefTransparent)
    Call PropBag.WriteProperty("MaskColor", mlMaskColor, mlDefMaskColor)
    
End Sub

Public Sub Refresh()

    Call SetForm
    
End Sub
Public Sub PopupSysMenu(Optional ByVal eFlags As MenuControlConstants = vbPopupMenuLeftAlign, Optional ByVal X As Variant, Optional ByVal Y As Variant)

'Syntax: PopupSysMenu(vbPopupMenuCenterAlign, X, Y)

'Shows the System Menu at X, Y or the current cursor position.
'X and Y must be in in mParent form's coordinates.

Dim hMenu   As Long
Dim lMenuID As Long
Dim lFlags  As Long
Dim fScaleX As Single
Dim fScaleY As Single
Dim ptTemp As PointAPI
Dim rcTemp  As RectAPI

    If mbDisabled Then
        Exit Sub
    End If
    
    'Calculate X and Y in screen pixel coordinates.
    If IsMissing(X) And IsMissing(Y) Then
        'If X and Y are not passed in, use the cursor position.
        Call GetCursorPos(ptTemp)
    Else
        If IsMissing(X) Then
            X = 0
        ElseIf IsMissing(Y) Then
            Y = 0
        End If
        ptTemp.X = mParent.ScaleX(X, mParent.ScaleMode, vbPixels)
        ptTemp.Y = mParent.ScaleY(Y, mParent.ScaleMode, vbPixels)
        'Convert X and Y to screen pixel coordinates.
        Call ClientToScreen(mParent.hWnd, ptTemp)
    End If
    
    'Get the System Menu and show it.
    hMenu = GetSystemMenu(mParent.hWnd, &H0&)
    If hMenu <> 0 Then
        lFlags = eFlags Or TPM_RETURNCMD
        lMenuID = TrackPopupMenu(hMenu, lFlags, ptTemp.X, _
          ptTemp.Y, &H0&, mParent.hWnd, rcTemp)
    End If
    
    'Notify the form of the selected menu item.
    If lMenuID <> 0 Then
        Call PostMessage(mParent.hWnd, WM_SYSCOMMAND, lMenuID, hMenu)
    End If
    
End Sub

