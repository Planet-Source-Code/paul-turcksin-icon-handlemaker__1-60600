Attribute VB_Name = "MakeIcon"
' I needed to convert a bitmap to an icon without having to save it to a file.
' Just getting a handle to the icon was sufficient, and fncMakeIcon does it.
' A second function fncConvertIconToPic converts an icon handle to an (almost)
' picture. I used it to check the good working of fncMakeIcon. And once it is
' in a picture box it can be easily saved to a file of course.
'
' IMPORTANT        There are a number of code lines used for demo purpose.
'                  These should be removed if the functions are to be used in
'                  another application. They are clearly marked.
'
' Paul Turcksin May 2005

Option Explicit

' DC
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
' BITMAP
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
' OBJECT
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
' ICON
Private Declare Function CreateIconIndirect Lib "user32" (icoinfo As ICONINFO) As Long
' PICTURE
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lppictDesc As pictDesc, riid As Guid, ByVal fown As Long, ipic As IPicture) As Long

Private Type BITMAP
   bmType As Long
   bmWidth As Long
   bmHeight As Long
   bmWidthBytes As Long
   bmPlanes As Integer
   bmBitsPixel As Integer
   bmBits As Long
End Type
Private Type ICONINFO
   fIcon As Boolean
   xHotspot As Long
   yHotspot As Long
   hBMMask As Long
   hBMColor As Long
   End Type

Private Type Guid
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
   End Type

Private Type pictDesc
   cbSizeofStruct As Long
   picType As Long
   hImage As Long
   End Type

Private Const PICTYPE_BITMAP = 1
Private Const PICTYPE_ICON = 3

Public Function fncMakeIcon(frmDC As Long, hBMP As Long, ByVal MaskClr As Long) As Long
     ' where frmDC   (in)  DC of the callng window
     '       hBMP    (in)  handle to a bitmap
     '       MaskClr (in)  if = -1 : pixel(0,0)
     ' Return value is a handle to the icon
     '       ipic    (out) icon picture

   Dim Bitmapdata As BITMAP  ' bitmap dimension
   Dim iWidth As Long
   Dim iHeight As Long
   Dim SrcDC As Long         ' copy of incoming bitmap
   Dim hSrc As Long
   Dim oldSrcObj As Long
   Dim MonoDC As Long        ' Mono mask (XOR)
   Dim MonoBmp As Long
   Dim oldMonoObj As Long
   Dim InvertDC As Long      ' Imverted mask (AND)
   Dim InvertBmp As Long
   Dim oldInvertObj As Long
'
   Dim cBkColor As Long
   Dim icoinfo As ICONINFO

' validate input
   If hBMP = 0 Then
      MsgBox "Invalid bitmap handle.", vbCritical Or vbOKOnly, "fncMakeIcon"
      Exit Function
   End If
   
' get size of bitmap
   If GetObject(hBMP, Len(Bitmapdata), Bitmapdata) = 0 Then
      MsgBox "Cannot get size of bitmap, vbCritical & vbOKOnly", "fncMakeIcon"
      Exit Function
   End If
   With Bitmapdata
      iWidth = .bmWidth
      iHeight = .bmHeight
   End With
   
' create copy of original, we will use it for both masks
   SrcDC = CreateCompatibleDC(0&)
   oldSrcObj = SelectObject(SrcDC, hBMP)
   
' get transparecy color
   If MaskClr = -1 Then
      MaskClr = GetPixel(SrcDC, 0, 0)
   End If
   
' mono mask (XOR) ............................................

' create mono DC/Bitmap for mask (XOR mask)
   MonoDC = CreateCompatibleDC(0&)
   MonoBmp = CreateCompatibleBitmap(MonoDC, iWidth, iHeight)
   oldMonoObj = SelectObject(MonoDC, MonoBmp)
' Set background of source to the mask color
   cBkColor = GetBkColor(SrcDC)   ' preserve original
   SetBkColor SrcDC, MaskClr
' copy bitmap and make monoDC mask in the process
  BitBlt MonoDC, 0, 0, iWidth, iHeight, SrcDC, 0, 0, vbSrcCopy
' restore original backcolor
   SetBkColor SrcDC, cBkColor
' inverted mask (AND) .................................................

' create DC/bitmap for inverted image (AND mask)
   InvertDC = CreateCompatibleDC(frmDC)
   InvertBmp = CreateCompatibleBitmap(frmDC, iWidth, iHeight)
   oldInvertObj = SelectObject(InvertDC, InvertBmp)
' copy bitmap into it
   BitBlt InvertDC, 0, 0, iWidth, iHeight, SrcDC, 0, 0, vbSrcCopy
' Invert background of image to create AND Mask
   SetBkColor InvertDC, vbBlack
   SetTextColor InvertDC, vbWhite
   BitBlt InvertDC, 0, 0, iWidth, iHeight, MonoDC, 0, 0, vbSrcAnd
  
' ========== lines to be removed from final version ==================
' show mono and inverted in calling form                             =
   Form1.Cls                                                     '   =
   BitBlt Form1.hdc, 24, 112, 32, 32, MonoDC, 0, 0, vbSrcCopy    '   =
   BitBlt Form1.hdc, 88, 112, 32, 32, InvertDC, 0, 0, vbSrcCopy  '   =
   Form1.Refresh                                                 '   =
' ====================================================================

' cleanup copy of original
   SelectObject SrcDC, oldSrcObj
   DeleteDC SrcDC
' Release MonoBmp And InvertBMP
   SelectObject MonoDC, oldMonoObj
   SelectObject InvertDC, oldInvertObj

    With icoinfo
      .fIcon = True
      .xHotspot = 16            ' Doesn't matter here
      .yHotspot = 16
      .hBMMask = MonoBmp
      .hBMColor = InvertBmp
      End With
      
' create 'output'
   fncMakeIcon = CreateIconIndirect(icoinfo)
    
' Clean up
    DeleteObject icoinfo.hBMMask
   DeleteObject icoinfo.hBMColor
   DeleteDC MonoDC
   DeleteDC InvertDC
End Function

Public Function fncConvertIconToPic(hIcon As Long) As IPicture
     ' where hIcon   (in)  icon handle
     ' Return value is an interface managing a picture object and its properties
     '          (can be used to set a picture property)

   Dim iGuid As Guid
   Dim pDesc As pictDesc
 
' init GUID
   With iGuid
      .Data1 = &H20400
      .Data4(0) = &HC0
      .Data4(7) = &H46
   End With

' fill picture description type
   With pDesc
      .cbSizeofStruct = Len(pDesc)
      .picType = PICTYPE_ICON
      .hImage = hIcon
       End With
   OleCreatePictureIndirect pDesc, iGuid, 1, fncConvertIconToPic
End Function
