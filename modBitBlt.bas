Attribute VB_Name = "modBitBlt"

'//////////////////////////////////////////////////////

'    Create a picture object from any DC

'//////////////////////////////////////////////////////

'    KPD-Team 1999
'    URL:        http://www.allapi.net/
'    E-Mail:     KPDTeam@Allapi.net


'    Set Me.Picture = hDCToPicture(GetDC(0), 0, 0, Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY)



Const RC_PALETTE As Long = &H100
Const SIZEPALETTE As Long = 104
Const RASTERCAPS As Long = 38
Private Type PALETTEENTRY
    peRed As Byte
    peGreen As Byte
    peBlue As Byte
    peFlags As Byte
End Type
Private Type LOGPALETTE
    palVersion As Integer
    palNumEntries As Integer
    palPalEntry(255) As PALETTEENTRY ' Enough for 256 colors
End Type
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
Private Type PicBmp
    Size As Long
    Type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type

Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hdc As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture
    Dim R As Long, Pic As PicBmp, IPic As IPicture, IID_IDispatch As GUID

    'Fill GUID info
    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With

    'Fill picture info
    With Pic
        .Size = Len(Pic) ' Length of structure
        .Type = vbPicTypeBitmap ' Type of Picture (bitmap)
        .hBmp = hBmp ' Handle to bitmap
        .hPal = hPal ' Handle to palette (may be null)
    End With

    'Create the picture
    R = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)

    'Return the new picture
    Set CreateBitmapPicture = IPic
End Function

Function hDCToPicture(ByVal hDCSrc As Long, ByVal LeftSrc As Long, ByVal TopSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long) As Picture

    Dim hDCMemory As Long, hBmp As Long, hBmpPrev As Long, R As Long
    Dim hPal As Long, hPalPrev As Long, RasterCapsScrn As Long, HasPaletteScrn As Long
    Dim PaletteSizeScrn As Long, LogPal As LOGPALETTE

    'Create a compatible device context
    hDCMemory = CreateCompatibleDC(hDCSrc)
    'Create a compatible bitmap
    hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
    'Select the compatible bitmap into our compatible device context
    hBmpPrev = SelectObject(hDCMemory, hBmp)

    'Raster capabilities?
    RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS) ' Raster
    'Does our picture use a palette?
    HasPaletteScrn = RasterCapsScrn And RC_PALETTE ' Palette
    'What's the size of that palette?
    PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE) ' Size of

    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        'Set the palette version
        LogPal.palVersion = &H300
        'Number of palette entries
        LogPal.palNumEntries = 256
        'Retrieve the system palette entries
        R = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
        'Create the palette
        hPal = CreatePalette(LogPal)
        'Select the palette
        hPalPrev = SelectPalette(hDCMemory, hPal, 0)
        'Realize the palette
        R = RealizePalette(hDCMemory)
    End If

    'Copy the source image to our compatible device context
    R = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, vbSrcCopy)

    'Restore the old bitmap
    hBmp = SelectObject(hDCMemory, hBmpPrev)

    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        'Select the palette
        hPal = SelectPalette(hDCMemory, hPalPrev, 0)
    End If

    'Delete our memory DC
    R = DeleteDC(hDCMemory)

    Set hDCToPicture = CreateBitmapPicture(hBmp, hPal)
End Function
   





