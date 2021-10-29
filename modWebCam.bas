Attribute VB_Name = "modWebCam"
Option Explicit

Public hHwnd As Long ' Handle to preview window

Dim iDevice As Long  ' Current device ID

Const WM_CAP As Integer = &H400
Const WM_CAP_DRIVER_CONNECT As Long = WM_CAP + 10
Const WM_CAP_DRIVER_DISCONNECT As Long = WM_CAP + 11
Const WM_CAP_SET_PREVIEW As Long = WM_CAP + 50
Const WM_CAP_SET_PREVIEWRATE As Long = WM_CAP + 52
Const WM_CAP_SET_SCALE As Long = WM_CAP + 53

Const WM_CAP_DLG_VIDEOFORMAT As Long = 1065
Const WM_CAP_DLG_VIDEOSOURCE As Long = 1066

Const WS_CHILD As Long = &H40000000
Const WS_VISIBLE As Long = &H10000000
Const SWP_NOMOVE As Long = &H2
Const SWP_NOSIZE As Integer = 1
Const SWP_NOZORDER As Integer = &H4
Const HWND_BOTTOM As Integer = 1

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hndw As Long) As Boolean

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function capCreateCaptureWindowA Lib "avicap32.dll" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Integer, ByVal hwndParent As Long, ByVal nID As Long) As Long
Private Declare Function capGetDriverDescriptionA Lib "avicap32.dll" (ByVal wDriver As Long, ByVal lpszName As String, ByVal cbName As Long, ByVal lpszVer As String, ByVal cbVer As Long) As Boolean


Public Sub ShowWebCamSize()
    SendMessage hHwnd, WM_CAP_DLG_VIDEOFORMAT, 0, 0
End Sub

Public Sub ShowWebCamSource()
    SendMessage hHwnd, WM_CAP_DLG_VIDEOSOURCE, 0, 0
End Sub


Public Sub LoadDeviceList(mList As ListBox)
    
    ' Load name of all avialable devices into ListBox

    Dim strName As String
    Dim strVer As String
    Dim iReturn As Boolean
    Dim X As Long
    
    X = 0
    strName = Space(100)
    strVer = Space(100)
    
    mList.Clear
    Do
        '   Get Driver name and version
        iReturn = capGetDriverDescriptionA(X, strName, 100, strVer, 100)

        ' If there was a device add device name to the list

        If iReturn Then mList.AddItem Trim$(strName)
        X = X + 1
    Loop Until iReturn = False
End Sub


Public Sub OpenPreviewWindow(DeviceIndex As Long, picCapture As PictureBox)
    
    iDevice = DeviceIndex
    
    ' Open Preview window in picturebox (  320, 240 )
    hHwnd = capCreateCaptureWindowA(iDevice, WS_VISIBLE Or WS_CHILD, 0, 0, 320, 240, picCapture.hwnd, 0)

    ' Connect to device
    If SendMessage(hHwnd, WM_CAP_DRIVER_CONNECT, iDevice, 0) Then
   
        'Set the preview scale
        SendMessage hHwnd, WM_CAP_SET_SCALE, True, 0

        'Set the preview rate in milliseconds
        SendMessage hHwnd, WM_CAP_SET_PREVIEWRATE, 66, 0

        'Start previewing the image from the camera
        SendMessage hHwnd, WM_CAP_SET_PREVIEW, True, 0

        ' Resize window to fit in picturebox
        SetWindowPos hHwnd, HWND_BOTTOM, 0, 0, picCapture.ScaleWidth, picCapture.ScaleHeight, SWP_NOMOVE Or SWP_NOZORDER
      
    Else
        ' Error connecting to device close window
        DestroyWindow hHwnd

     End If
 End Sub

Public Sub ClosePreviewWindow()
    
    ' Disconnect from device
    SendMessage hHwnd, WM_CAP_DRIVER_DISCONNECT, iDevice, 0

    ' close window
    DestroyWindow hHwnd
End Sub





