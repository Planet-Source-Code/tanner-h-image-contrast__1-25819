VERSION 5.00
Begin VB.Form frmContrast 
   AutoRedraw      =   -1  'True
   Caption         =   "Tanner's Contrast Example"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   ScaleHeight     =   371
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   419
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtContrast 
      Height          =   285
      Left            =   4560
      TabIndex        =   2
      Text            =   "100"
      Top             =   5040
      Width           =   495
   End
   Begin VB.CommandButton CmdContrast 
      Caption         =   "Change Contrast"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   4920
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   4560
      Left            =   120
      Picture         =   "Contrast.frx":0000
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   0
      Top             =   120
      Width           =   6060
   End
   Begin VB.Label Label1 
      Caption         =   "Contrast Change (-100% to 100%):"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   5040
      Width           =   1335
   End
End
Attribute VB_Name = "frmContrast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Contrast example ©2001 by Tanner Helland

'Here's the only VB example on the net that will show how to correctly
'adjust an image's contrast.  To perfectly determine the contrast, you
'would have to first find the average brightness of the image; I use
'the shortcut method and assume that the average is 127 (close enough
'for our purposes).  The code is pretty straight-forward, but if you
'have any questions feel free to send a message to
'tanner@tannerhelland.com

'The CG graphic in the picture box is ©1998 by SquareSoft
'(it's from Final Fantasy 8, if you care)

'For additional cool code, check out my website at
'tannerhelland.50megs.com

'The Windows API sets and gets pixels a whole lot faster then PSet
'and Point. The format is Picturebox.hdc, destination x, destination y, color to set
Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Byte
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long

Private Sub CmdContrast_Click()
    'variables for contrast, color calculation, positioning
    Dim Contrast As Integer
    Dim NewColor As Long
    Dim x As Integer, y As Integer
    Dim r As Integer, g As Integer, b As Integer
    Contrast = Val(TxtContrast)
    'run a loop through the picture to change every pixel
    For x = 0 To Picture1.ScaleWidth
    For y = 0 To Picture1.ScaleHeight
        'get the current color value
        NewColor = GetPixel(Picture1.hDC, x, y)
        'extract the R,G,B values from the long returned by GetPixel
        r = ExtractR(NewColor)
        g = ExtractG(NewColor)
        b = ExtractB(NewColor)
        'change the RGB settings to their appropriate contrast
        r = r + (((r - 127) * Contrast) \ 100)
        g = g + (((g - 127) * Contrast) \ 100)
        b = b + (((b - 127) * Contrast) \ 100)
        'make sure the new variables aren't too high or too low
        ByteMe r
        ByteMe g
        ByteMe b
        'set the new pixel
        SetPixelV Picture1.hDC, x, y, RGB(r, g, b)
    'continue through the loop
    Next y
        'refresh the picture box every 10 lines (a nice progress bar effect)
        If x Mod 10 = 0 Then Picture1.Refresh
    Next x
    'final picture refresh
    Picture1.Refresh
End Sub

Private Function ExtractR(ByVal CurrentColor As Long) As Integer
    ExtractR = CurrentColor Mod 256
End Function

Private Function ExtractG(ByVal CurrentColor As Long) As Integer
    ExtractG = (CurrentColor \ 256) And 255
End Function

Private Function ExtractB(ByVal CurrentColor As Long) As Integer
    ExtractB = (CurrentColor \ 65536) And 255
End Function

Public Sub ByteMe(ByRef TempVar As Integer)
'Convert to absolute byte values
    If TempVar > 255 Then TempVar = 255
    If TempVar < 0 Then TempVar = 0
End Sub
