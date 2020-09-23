Attribute VB_Name = "ModuleDefineTypes"
'these types are defined again in Graphix3D, so remark them here.
Public Type Point3DType
        x As Single
        y As Single
        z As Single
End Type

Public Type Page3DType
        ColRGB As Long 'RGB Color : -1=no specify color => dont change brush (use current brush, so if current be a Null Brush dont fill page)
        NumPoints As Long
        Point() As Point3DType 'Point(0)=Center Of Page , Point(1 to NumPoints)=Points
        SpValue As Single 'any arbitrary value such as magnitude of Radiation on this page
End Type

Public Type Object3DType
        CenterPoint As Point3DType
        NameStr As String
        iCode As Integer '1=Sphere / 2=SpObject-Tower / 4=Find Radiation / 8=Can Make Shadow
                         '16=semiSphere / 32=Dome / 64=SunObject(=>Radius=SpValues(0))
        ColEdge As Long  'color for Edges (-1 = NoEdge)
        SpValues() As Single 'array for Special Values such as Temperature or G_Radiation
        NumPages As Long
        Page() As Page3DType
End Type

Public Type ViewParamsType
        Teta As Single
        fi As Single
        View3DScale As Single
        ViewXshift As Long '2D shift
        ViewYshift As Long '2D shift
        View3DShiftOrig As Point3DType '3D shift (set Target)
        RadMode As Integer
End Type

Public Type View3DAnglesType
        SinTeta As Single
        CosTeta As Single
        SinFi As Single
        CosFi As Single
End Type

Public Type Scene3DType
        NameStr As String
        DefaultView As ViewParamsType
        CenterPoint As Point3DType
        NumObjects As Long
        Object() As Object3DType
        SpValues() As Single 'array for Special Values
End Type

Public Type SinCos3AngType
        SinX As Single
        CosX As Single
        SinY As Single
        CosY As Single
        SinZ As Single
        CosZ As Single
End Type

Public Type Vector3DType
 dx As Single
 dy As Single
 dz As Single
End Type

Public Type Line3DType
        p1 As Point3DType
        p2 As Point3DType
End Type

