<div align="center">

## Elastic Resize and Reposition Control


</div>

### Description

This OCX will resize and reposition all of the controls on a Form when you resize the form or change resolutions.This Control is based on Elastic Class by Mikhail Shmukler. This newer version is a very little bit faster than before, and also will not resize status bars if they are present. It also prevents a form from being resized smaller than 700 twips.

Installable compiled version 1.2 can be had here http://www.angelfire.com/band/AMP/files/elastic.zip
 
### More Info
 
Add Control to a Form. Call Init() routine in Form.Load Event with no passed parameters (example: "Elastic1.Init")


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ronald Gladhill](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ronald-gladhill.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[OLE/ COM/ DCOM/ Active\-X](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/ole-com-dcom-active-x__1-29.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ronald-gladhill-elastic-resize-and-reposition-control__1-1850/archive/master.zip)





### Source Code

```
'http://www.angelfire.com/band/AMP/files/elastic.zip
' ****************************************************************************
' * Original Class Programmers Name : Mikhail Shmukler
' * Web Site : www.geocities.com/ResearchTriangle/6311/
' * E-Mail : waty.thierry@usa.net
' * Date : 13/10/98
' * Time : 10:24
' * Module Name : class_Elastic
' * Module Filename : Elastic.cls
' ****************************************************************************
' * Comments :
' * This class can change size and location of controls On your form
' * 1. Resize form
' * 2. Change screen resolution
' * Assumes:1. Add Elastic.cls
' * 2. Add declaration 'Private El as New class_Elastic'
' * 3. Insert string like 'El.init Me' (formload event)
' * 4. Insert string like 'El.FormResize Me' (Resize event)
' * 5. Press 'F5' and resize form ....
' ****************************************************************************
' ****************************************************************************
' * OCX conversion Programming By : Ronald Gladhill
' * E-Mail : cybergar@theramp.net
' * Date : June 27, 1999
' * OCX FileName : Elastic.ocx
' * OCA FileName : Elastic.oca
' ****************************************************************************
' * COMMENTS:
' * This OCX will resize and reposition the controls on a form when you
' * Resize the form or change screen resolutions.
' * INSTRUCTIONS:
' * 1. Add the control to your form
' * 2. Call Init() routine in Form.Load event (example: "Elastic1.Init").
'*****************************************************************************
Option Explicit
Private WithEvents objParent As Form
Private nFormHeight As Long
Private nFormWidth As Long
Private nNumOfControls As Integer
Private nTop() As Long
Private nLeft() As Long
Private nHeight() As Long
Private nWidth() As Long
Private bFirstTime As Boolean
Public Sub Init()
Dim I As Integer
 Set objParent = UserControl.Parent
 With objParent
 nFormHeight = .ScaleHeight
 nFormWidth = .ScaleWidth
 nNumOfControls = .Controls.Count - 1
 bFirstTime = True
 ReDim nTop(nNumOfControls)
 ReDim nLeft(nNumOfControls)
 ReDim nHeight(nNumOfControls)
 ReDim nWidth(nNumOfControls)
 On Error Resume Next
 For I = 0 To nNumOfControls
 Select Case TypeName(.Controls(I))
 Case "Line"
 nTop(I) = .Controls(I).Y1
 nLeft(I) = .Controls(I).X1
 nHeight(I) = .Controls(I).Y2
 nWidth(I) = .Controls(I).X2
 Case "StatusBar"
 'do nothing. Leave it alone
 Case Else
 nTop(I) = .Controls(I).Top
 nLeft(I) = .Controls(I).Left
 nHeight(I) = .Controls(I).Height
 nWidth(I) = .Controls(I).Width
 End Select
 Next I
 End With
End Sub
Private Sub objParent_Resize()
On Error Resume Next ' for comboboxes, timers and other nonsizable controls
Dim I As Integer
Dim nCaptionSize As Integer
Dim dRatioX As Double
Dim dRatioY As Double
Dim nSaveRedraw As Long
 With objParent
 nSaveRedraw = .AutoRedraw
 .AutoRedraw = True
 If .Height <= 700 Then
 .Height = 700
 End If
 If .Width <= 700 Then
 .Width = 700
 End If
 dRatioY = 1# * nFormHeight / .ScaleHeight
 dRatioX = 1# * nFormWidth / .ScaleWidth
 For I = 0 To nNumOfControls
 Select Case TypeName(.Controls(I))
 Case "Line"
 .Controls(I).Y1 = Fix(nTop(I) / dRatioY)
 .Controls(I).X1 = Fix(nLeft(I) / dRatioX)
 .Controls(I).Y2 = Fix(nHeight(I) / dRatioY)
 .Controls(I).X2 = Fix(nWidth(I) / dRatioX)
 Case "StatusBar"
 'Do nothing
 Case Else
 .Controls(I).Top = Fix(nTop(I) / dRatioY)
 .Controls(I).Left = Fix(nLeft(I) / dRatioX)
 .Controls(I).Height = Fix(nHeight(I) / dRatioY)
 .Controls(I).Width = Fix(nWidth(I) / dRatioX)
 End Select
 Next I
 .AutoRedraw = nSaveRedraw
 End With
End Sub
```

