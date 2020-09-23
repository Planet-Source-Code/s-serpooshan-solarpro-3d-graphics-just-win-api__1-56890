Attribute VB_Name = "ModulePublic"
Option Explicit

'API
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'general definitions for this project
Public Const DefaultHorizontalPlateStr = "Default Horizontal Plate"
Public ShouldRedeawGrShapes As Integer, ShouldRefreshFormGr As Integer, ShouldReCalculate As Integer 'refresh is a fast operation(BitBlt only) but DrawGrShapes is a slow because it should redraw all pages again(is need only when you change some view properties (as scale) or scene objects!)
Public UDefObj As Object3DType
Public Serpoosh As String
Public ComponentG As Integer
Public TargetObjIndex As Long
Public SpCaseObjectCode As Integer, SpCaseObjectChange As Integer
Public HorizPageCol As Long, ColorModeV As Integer
Public UseAbsoluteGMinMax As Integer, AbsoluteGMin As Single, AbsoluteGMax As Single
Public ShouldDrawSun As Integer, ShouldDrawSunBeam As Integer, ShouldDrawAxises As Integer
Public ShouldDrawHorizonPlate As Integer, ShouldDrawInActiveObjects As Integer
Public ShouldRecordImages As Integer, RecordFolder As String, RecordImgID As Long
Public ShouldAddHeader As Integer, HeaderText As String
Public ShouldRStp As Integer
Public FormGrIsVisible As Integer, CancelPressedPublic As Integer 'Cancel.. is true when user cancel calculation on PogressForm
Public iObjTransform As Long

Public Sub PrintHeader(PicObj As Object, sText As String, Optional ByVal x As Single = 0, Optional ByVal y As Single = 0, Optional ByVal col1 As Long = &HA0FFFF, Optional ByVal col2 As Long = &H4040&)
'obj=pic1
With PicObj
    
    .ForeColor = col2
    .CurrentX = x + 1
    .CurrentY = y + 1
    PicObj.Print sText;
    
    .ForeColor = col1
    .CurrentX = x
    .CurrentY = y
    PicObj.Print sText;
    
End With
End Sub

Public Function FindSelectedOption(OptionObj As Object) As Integer
 'return selected option of a option array (return 1 for first!)
 'Note that first one is 1 not 0
 Dim i As Integer
 For i = 0 To OptionObj.UBound
  If OptionObj(i).Value Then FindSelectedOption = i + 1: Exit For
 Next
End Function

Public Sub pPrint(p As Point3DType)
 With p
  Debug.Print Round(.x, 2); ","; Round(.y, 2); ","; Round(.z, 2)
 End With
End Sub

Public Sub vPrint(v As Vector3DType)
 With v
  Debug.Print Round(.dx, 2); ","; Round(.dy, 2); ","; Round(.dz, 2)
 End With
End Sub

Public Sub pgPrint(page1 As Page3DType)
 Dim i As Long
 With page1
  For i = 0 To .NumPoints
   Debug.Print CStr(i); ")";
   pPrint .Point(i)
  Next
 End With
 
End Sub

Sub ApplyTransformsToObj(Obj1 As Object3DType)
   
   Dim s As String, c3 As Variant
   
   On Error Resume Next
   With FormTransforms
     s = .TextTransform.Text: c3 = Split(s, ",")
     If s <> "0,0,0" And UBound(c3) = 2 Then
      TransferSelfObject3D Obj1, Vector3D(c3(0), c3(1), c3(2))
     End If
     s = .TextRotation.Text: c3 = Split(s, ",")
     If s <> "0,0,0" And UBound(c3) = 2 Then
      RotateSelfObject3D Obj1, SinCos3Angles(c3(0), c3(1), c3(2))
     End If
     s = .TextScale.Text
     If s <> "100" Then
      ScaleSelfObject3D Obj1, Val(s) / 100, Point3D(0, 0, 0)
     End If
   End With
   
End Sub
