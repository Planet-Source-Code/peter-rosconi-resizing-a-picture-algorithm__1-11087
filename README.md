<div align="center">

## Resizing a picture algorithm


</div>

### Description

It resizes any size picture in a picturebox into another any size picturebox. I made this to show an algorithm of resizing a picture.
 
### More Info
 
startingPic As PictureBox, destinationPic As PictureBox

Must have two picture boxes the 'starting picture box' must have a picture already loaded into it.

Not the most efficient way to resize a picture.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Peter Rosconi](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/peter-rosconi.md)
**Level**          |Advanced
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Graphics](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics__1-46.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/peter-rosconi-resizing-a-picture-algorithm__1-11087/archive/master.zip)





### Source Code

```
'pretty much straight forward just load a picture
'  into 'startingPic' and call this sub
Public Sub ResizePicture(startingPic As PictureBox, destinationPic As PictureBox)
'the horz. and vert. ratios
ratioX = startingPic.ScaleWidth / destinationPic.ScaleWidth
ratioY = startingPic.ScaleHeight / destinationPic.ScaleHeight
'for stats
theTimer = Timer
'go through the startingPic's pixels
For x = 0 To startingPic.ScaleWidth Step ratioX
For y = 0 To startingPic.ScaleHeight Step ratioY
  'get the color of the startingPic
  theColor = startingPic.Point(x, y)
  'find the corresponding x and y values
  ' for the resized destination pic
  realX = ratioX ^ -1 * x
  realY = ratioY ^ -1 * y
  destinationPic.PSet (realX, realY), theColor
Next y
Next x
MsgBox "It took " & Timer - theTimer & " seconds to increase the horizontal size by " & ratioX ^ -1 & " and the vertical size by " & ratioY ^ -1 & "."
End Sub
```

