Attribute VB_Name = "modFileIcon"
Public Function GenerateIconForSave(ByRef PicBox As PictureBox) As String
'method for generating the binary data required for TRUE COLOR(!) icons

Dim arrx() As String, isTr As Boolean
Dim x As Integer, y As Integer, s As String
Dim width_in_pixels As Integer, height_in_pixels As Integer
Dim l As Long, r As Integer, g As Integer, b As Integer

PicBox.ScaleMode = vbPixels
width_in_pixels = PicBox.ScaleWidth
height_in_pixels = PicBox.ScaleHeight

ReDim arrx(height_in_pixels - 1, 3) As String

For y = height_in_pixels - 1 To 0 Step -1 'icons are saved from top to bottom
 For x = 0 To width_in_pixels - 1         'and left to right
  l = PicBox.Point(x, y)
  
  r = GetRGB(l).Red
  g = GetRGB(l).Green
  b = GetRGB(l).Blue
  
  If r < 0 Then r = 0
  If g < 0 Then g = 0
  If b < 0 Then b = 0
  
  If l = PicBox.BackColor Then
   s = s & Chr(0) & Chr(0) & Chr(0)
   isTr = True
  Else
   s = s & Chr(b) & Chr(g) & Chr(r)
  End If
 Next x
Next y

If isTr = True Then 'transparent
 For y = height_in_pixels - 1 To 0 Step -1 'populate transparent data array
  For x = 0 To 3                           'to make sure that'll fill properly for 16x16 icons
   arrx(y, x) = "-1"
  Next x
 Next y

For y = 0 To height_in_pixels - 1 'check for transparency
 For x = 0 To width_in_pixels - 1
  l = PicBox.Point(x, y)
  If l = PicBox.BackColor Then 'generate transparent string for transsolution function
    If arrx(y, Int(x / 8)) = "-1" Then arrx(y, Int(x / 8)) = ""
    arrx(y, Int(x / 8)) = arrx(y, Int(x / 8)) & "1"
  Else
    If arrx(y, Int(x / 8)) = "-1" Then arrx(y, Int(x / 8)) = ""
    arrx(y, Int(x / 8)) = arrx(y, Int(x / 8)) & "0"
  End If
  
 Next x
Next y

Dim f As String, e As Integer
 For y = height_in_pixels - 1 To 0 Step -1 'create generated transparent data
  For x = 0 To 3
   If arrx(y, x) = "-1" Then e = 255 Else e = TransSolution(arrx(y, x))
   f = f & Chr(e)
  Next x
 Next y
End If

 s = s & f
 'fill data with chr(0) if there is no transparency
 s = s & String(((width_in_pixels * width_in_pixels) * 3) + ((width_in_pixels / 4) * width_in_pixels) - Len(s), Chr(0))
 
If width_in_pixels = 16 Then '16x16 icon
 GenerateIconForSave = String(2, Chr(0)) & Chr(1) & Chr(0) & Chr(1) & Chr(0) & Chr(width_in_pixels) & Chr(height_in_pixels) & String(6, Chr(0)) & Chr(Val("&H68")) & Chr(3) & String(2, Chr(0)) & Chr(22) & String(3, Chr(0)) & Chr(40) & String(3, Chr(0)) & Chr(16) & String(3, Chr(0)) & Chr(32) & String(3, Chr(0)) & Chr(1) & Chr(0) & Chr(24) & String(5, Chr(0)) & Chr(64) & Chr(3) & String(18, Chr(0)) & s
ElseIf width_in_pixels = 32 Then '32x32 icon
 GenerateIconForSave = String(2, Chr(0)) & Chr(1) & Chr(0) & Chr(1) & Chr(0) & Chr(width_in_pixels) & Chr(height_in_pixels) & String(6, Chr(0)) & Chr(Val("&HA8")) & Chr(12) & String(2, Chr(0)) & Chr(22) & String(3, Chr(0)) & Chr(40) & String(3, Chr(0)) & Chr(32) & String(3, Chr(0)) & Chr(64) & String(3, Chr(0)) & Chr(1) & Chr(0) & Chr(24) & String(6, Chr(0)) & Chr(12) & String(18, Chr(0)) & s
Else 'unsupported icon size
 Err.Raise 10001, , "Unsupported Icon Size"
End If
End Function

Public Function GetIconSize(ByVal FileName As String) As Integer
'check if a file is an icon AND if it is return the icon size
Dim s As String, l As Long
l = FreeFile()
Open FileName For Binary Access Read As #l
 s = Input(LOF(l), #l)
Close #l
If Left(s, 5) <> Chr(0) & Chr(0) & Chr(1) & Chr(0) & Chr(1) Then 'not an icon
 GetIconSize = -1
Else 'is an icon get the size
 GetIconSize = Asc(Mid(s, 7, 1))
End If
End Function

Private Function TransSolution(ByVal s As String) As Integer
'expects a string 8 bytes in length looking something like this: 1100000
'will calculate the correct transparent byte from said string
'to be honest I know there must be a more simple mathematical equation
'to accomplish this but it is beyond my abilities. I spent a good couple hours
'decoding how each 8 pixel area was translated into a single byte and this is
'the result. Ugly? Yes. Works? Yes.

If s = "00000000" Then TransSolution = 0: Exit Function
Dim b As Integer, c As Integer, a As String
Dim arrBase(7) As Single

For i = 7 To 0 Step -1
 a = Mid(s, i + 1, 1)
 If a = "1" Then b = 7 - i
Next i
'//////base 1
arrBase(0) = 1
c = -1
For i = 0 To b
 arrBase(0) = arrBase(0) + arrBase(0)
  If i < b Then
   a = Mid(s, (7 - i) + 1, 1)
   If a = "1" Then c = i
  End If
Next i
'//////base 1

'//////base 2
arrBase(1) = 1
b = -1
For i = 0 To c
 arrBase(1) = arrBase(1) + arrBase(1)
 If i < c Then
  a = Mid(s, (7 - i) + 1, 1)
  If a = "1" Then b = i
 End If
Next i
'//////base 2

'//////base 3
arrBase(2) = 1
c = -1
For i = 0 To b
 arrBase(2) = arrBase(2) + arrBase(2)
  If i < b Then
   a = Mid(s, (7 - i) + 1, 1)
   If a = "1" Then c = i
  End If
Next i
'//////base 3

'//////base 4
arrBase(3) = 1
b = -1
For i = 0 To c
 arrBase(3) = arrBase(3) + arrBase(3)
 If i < c Then
   a = Mid(s, (7 - i) + 1, 1)
   If a = "1" Then b = i
 End If
Next i
'//////base 4

'//////base 5
arrBase(4) = 1
c = -1
For i = 0 To b
 arrBase(4) = arrBase(4) + arrBase(4)
  If i < b Then
   a = Mid(s, (7 - i) + 1, 1)
   If a = "1" Then c = i
  End If
Next i
'//////base 5

'//////base 6
arrBase(5) = 1
b = -1
For i = 0 To c
 arrBase(5) = arrBase(5) + arrBase(5)
 If i < c Then
   a = Mid(s, (7 - i) + 1, 1)
   If a = "1" Then b = i
 End If
Next i
'//////base 6

'//////base 7
arrBase(6) = 1
c = -1
For i = 0 To b
 arrBase(6) = arrBase(6) + arrBase(6)
  If i < b Then
   a = Mid(s, (7 - i) + 1, 1)
   If a = "1" Then c = i
  End If
Next i
If c = 0 Then arrBase(6) = arrBase(6) + 2
'//////base 7

Dim base_number As Integer
Dim base2_number As Integer
Dim base3_number As Integer
Dim base4_number As Integer
Dim base5_number As Integer
Dim base6_number As Integer
Dim base7_number As Integer

base_number = CInt(arrBase(0) / 2)
base2_number = CInt(arrBase(1) / 2)
base3_number = CInt(arrBase(2) / 2)
base4_number = CInt(arrBase(3) / 2)
base5_number = CInt(arrBase(4) / 2)
base6_number = CInt(arrBase(5) / 2)
base7_number = CInt(arrBase(6) / 2)
TransSolution = base_number + base2_number + base3_number + base4_number + base5_number + base6_number + base7_number
End Function
