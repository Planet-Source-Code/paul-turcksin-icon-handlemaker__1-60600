VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Convert bmp to icon handle"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   211
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   2160
      Width           =   1935
   End
   Begin VB.PictureBox PictOut 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   3600
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   2
      Top             =   480
      Width           =   510
   End
   Begin VB.PictureBox PictIn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   240
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   480
      Width           =   510
   End
   Begin VB.Label lblDrag 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "You can drag the icon now."
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3360
      TabIndex        =   10
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "AND"
      Height          =   255
      Left            =   1320
      TabIndex        =   9
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "XOR"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "Icon masks (for demo purpose)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "DrawIcon API"
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "OUT (icon)"
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "IN (bmp)"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Demo 'MakeIcon' module
' Paul Turcksin May 2005

Option Explicit
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long

Private hIcon As Long         ' icon handle
Private ipic As IPicture      ' picture interface

Private Sub cmdConvert_Click()
' grab handle from bitmap picture in picture box and produce an icon handle
   hIcon = fncMakeIcon(Me.hdc, PictIn.Picture.Handle, -1)
' draw obtained icon on the form
   DrawIcon Me.hdc, 128, 32, hIcon
   Me.Refresh
End Sub

Private Sub cmdShow_Click()
' produce picture from icon handle
   Set PictOut.Picture = fncConvertIconToPic(hIcon)
   PictOut.Refresh
   lblDrag.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   DestroyIcon hIcon
   Set Form1 = Nothing
End Sub

Private Sub PictOut_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' this enables to drag the icon
   If lblDrag.Visible Then
      PictOut.DragIcon = PictOut.Picture
      PictOut.Drag
   End If
End Sub
