<div align="center">

## Animated Glass Cube Effect using StretchBlt


</div>

### Description

This code animates a glass cube moving on a picture in real-time. Almost the same as that Windows 95 screen saver (the moving glass boll). I though this code can help make a screen-saver like that. (You have to e-mail me a copy of that screen-saver please)
 
### More Info
 
Create a new form. Place to picture boxes on it (Picture1 and Picture2). Set the following attributes on them (Appearance=0 - Flat; AutoRedraw=False ;AutoSize=True). Load a nice fair sized BMP graphic into Picture1 and load that same graphic into Picture2. Set Picture2 Visible property to false. Create a command button (Command1) and change it's caption to "Do Effect".

Now copy and paste the code and off you go...

PLEASE SEND ME A COPY IF YOU MAKE ANY NICE CHANGES OR SCREEN-SAVER FROM THIS CODE. If I get the time I'll make one any post it on this site ;)

I only checked this code on Windows 98... I don't know if it will work on any of the other platforms... Please let me know if it work's on any of the other.

A smile on your face ;)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Riaan Aspeling](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/riaan-aspeling.md)
**Level**          |Unknown
**User Rating**    |4.8 (38 globes from 8 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/riaan-aspeling-animated-glass-cube-effect-using-stretchblt__1-1636/archive/master.zip)

### API Declarations

```
'***************************************
'This code must be copied into a module.
'***************************************
'
Option Explicit
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Global Const SRCCOPY = &HCC0020
```


### Source Code

```
'**************************************
'This code must be copied into the form
'**************************************
'
Option Explicit
Dim CompDC As Long, hBmp As Long, CompDCOrg As Long, hBmp2 As Long
Dim SourceHDC As Long, SourceBMP As Long, SourceBMP2 As Long
Dim SourceHDC2 As Long
Dim rtn As Long, xsize As Long, ysize As Long
Dim xbounce As Long, ybounce As Long
Dim aw As Integer, xdir As Integer, ydir As Integer, iloop As Integer
Dim StayInLoop As Boolean
Private Sub Form_Activate()
  Randomize
  'The x and y size of the picture in pixels for the API's
  xsize = Picture1.Width / Screen.TwipsPerPixelX
  ysize = Picture1.Height / Screen.TwipsPerPixelY
  'The aw (Alteration Width) of the glass deformation object
  aw = 20
  'xdir and ydir is the bounce directional variables
  xdir = (Rnd * 5) + 1
  ydir = (Rnd * 5) + 1
  'Make a copy of both picture's into memory DC's
  Call MakeCopyOfImgage
  'Make sure the display picture doesn't redraw itself
  Picture1.AutoRedraw = False
  'The next variable controls the animation loop
  StayInLoop = False
  'Copy the origanal image to the visible picture box
  rtn = BitBlt(Picture1.hdc, 0, 0, xsize, ysize, CompDCOrg, 0, 0, SRCCOPY)
  'xbounce and ybounce is the center-point of the glass object
  'making it aw will display it in the top left-hand corner of the picture box
  xbounce = aw: ybounce = aw
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  'Terminate the animation loop
  StayInLoop = False
  'Free the memory used for the DC's
  Call DeleteCopyOfImage
End Sub
Private Sub Command1_Click()
  StayInLoop = Not StayInLoop
  While StayInLoop
    'Reset the portion of the DC that was deformed
    Call ResetPortion(xbounce, ybounce, aw)
    'Do the movement
    xbounce = xbounce + xdir
    If xbounce > xsize - aw Then xdir = -(Rnd * 5) - 1
    ybounce = ybounce + ydir
    If ybounce > ysize - aw Then ydir = -(Rnd * 5) - 1
    If xbounce < aw Then xdir = (Rnd * 5) + 1
    If ybounce < aw Then ydir = (Rnd * 5) + 1
    'Do the deformation on the memory DC
    Call Stretch(xbounce, ybounce, aw)
    'Copy the memory DC to the visible picture box
    rtn = BitBlt(Picture1.hdc, 0, 0, xsize, ysize, CompDC, 0, 0, SRCCOPY)
    'Let windows do some other stuff (I WILL NOT RECOMEND TO REMOVE THE NEXT LINE)
    DoEvents
  Wend
End Sub
Sub Stretch(ByVal xpos As Long, ByVal ypos As Long, ByVal areawidth As Long)
  Dim Stretchit As Double, i As Double
  Dim rtn As Long
  'The next variable set's the percentage of deformation
  'You can change this variable to get some interesting effects
  Stretchit = 0.9
  For i = 2 To 0.1 Step -0.2
    rtn = StretchBlt(CompDC, _
            xpos - (areawidth * i), _
            ypos - (areawidth * i), _
            (areawidth * i) * 2, _
            (areawidth * i) * 2, _
          CompDC, _
            (xpos - (areawidth * Stretchit * i)), _
            (ypos - (areawidth * Stretchit * i)), _
            (areawidth * Stretchit * i) * 2, _
            (areawidth * Stretchit * i) * 2, _
          SRCCOPY)
  Next i
End Sub
Sub ResetPortion(ByVal xpos As Long, ByVal ypos As Long, ByVal areawidth As Long)
  Dim rtn As Long
  'This next line will reset the are on the DC that was deformed
  rtn = BitBlt(CompDC, _
          xpos - (areawidth * 2), _
          ypos - (areawidth * 2), _
          (areawidth) * 4, _
          (areawidth) * 4, _
          CompDCOrg, _
          xpos - (areawidth * 2), _
          ypos - (areawidth * 2), _
          SRCCOPY)
End Sub
Sub MakeCopyOfImgage()
  'Get the handel to the DC for the two picture boxes
  SourceHDC = Picture1.hdc
  SourceHDC2 = Picture2.hdc
  'Get the pictures
  SourceBMP = Picture1.Picture
  SourceBMP2 = Picture2.Picture
  'Create the to memory DC's
  CompDC = CreateCompatibleDC(SourceHDC)
  CompDCOrg = CreateCompatibleDC(SourceHDC)
  'Copy the pictures to these DC's
  hBmp = SelectObject(CompDC, SourceBMP)
  hBmp2 = SelectObject(CompDCOrg, SourceBMP2)
End Sub
Sub DeleteCopyOfImage()
  'Delete the memory DC's
  rtn = DeleteDC(CompDC)
  rtn = DeleteDC(CompDCOrg)
End Sub
```

