VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cylinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Cylinder

Implements RayTraceable

' Position
Private Point1 As Point3D
Private Point2 As Point3D

' Radius
Private Radius As Single

Private HitU As Single

' Lighting values:
' Ambient:
Private AmbKr As Single
Private AmbKg As Single
Private AmbKb As Single
' Diffuse
Private DiffKr As Single
Private DiffKg As Single
Private DiffKb As Single
' Specular
Private Spec_K As Single
Private Spec_N As Single
' Reflected factor
Private ReflKr As Single
Private ReflKg As Single
Private ReflKb As Single

' Culling values
' True if we had a hit on this scanline
Private HadHit As Boolean
' True if we had a hit on previous scanline
Private HadHitPrev As Boolean
' True if we are culled forever
Private ForeverCulled As Boolean
' True if we are done on this scanline
Private ScanlineDone As Boolean

Public Sub RayTraceable_Apply(M() As Single)
    m3Apply Point1.Coord, M, Point1.Trans
    m3Apply Point2.Coord, M, Point2.Trans
End Sub

Public Sub RayTraceable_ApplyFull(M() As Single)
    m3ApplyFull Point1.Coord, M, Point1.Trans
    m3ApplyFull Point2.Coord, M, Point2.Trans
End Sub

Public Sub RayTraceable_CullScanline(ByVal px As Single, ByVal py As Single, ByVal pz As Single, ByVal Nx As Single, ByVal Ny As Single, ByVal Nz As Single)
    ' See if we will ever be visible again.
    If ForeverCulled Then
        ScanlineDone = True
        Exit Sub
    End If

    ' We have not yet had a hit on this scanline.
    HadHit = False

    ' Do not scanline cull.
    ScanlineDone = False
End Sub

Public Sub RayTraceable_FindHitColor(Objects As Collection, ByVal eyeX As Single, ByVal eyeY As Single, ByVal eyeZ As Single, ByVal px As Single, ByVal py As Single, ByVal pz As Single, R As Integer, G As Integer, B As Integer)
    Dim Vx As Single ' Vector V, the axis vector
    Dim Vy As Single
    Dim Vz As Single
    Dim Nx As Single ' Vector N, the normal vector
    Dim Ny As Single
    Dim Nz As Single
    Dim Nlen As Single
    
    ' Get the axis vector
    Vx = Point2.Trans(1) - Point1.Trans(1)
    Vy = Point2.Trans(2) - Point1.Trans(2)
    Vz = Point2.Trans(3) - Point1.Trans(3)
    ' Find the normal
    Nx = px - (Point1.Trans(1) + HitU * Vx)
    Ny = py - (Point1.Trans(2) + HitU * Vy)
    Nz = pz - (Point1.Trans(3) + HitU * Vz)
    ' Normalize
    Nlen = Sqr(Nx * Nx + Ny * Ny + Nz * Nz)
    Nx = Nx / Nlen
    Ny = Ny / Nlen
    Nz = Nz / Nlen
    
    ' Calculate the hit color
    CalculateHitColor Objects, Me, eyeX, eyeY, eyeZ, _
        px, py, pz, _
        Nx, Ny, Nz, _
        DiffKr, DiffKg, DiffKb, _
        AmbKr, AmbKg, AmbKb, _
        Spec_K, Spec_N, _
        ReflKr, ReflKg, ReflKb, _
        R, G, B
End Sub

Public Function RayTraceable_FindT(DirectC As Boolean, px As Single, py As Single, pz As Single, Vx As Single, Vy As Single, Vz As Single) As Single
    Dim X1 As Single
    Dim Y1 As Single
    Dim Z1 As Single
    Dim Wx As Single    ' Vector W: The axis vector
    Dim Wy As Single
    Dim Wz As Single
    Dim WlenSquared As Single
    Dim WdotV As Single
    ' Coefficients for the cylinder equation
    Dim A As Single
    Dim B As Single
    Dim Cx As Single
    Dim Cy As Single
    Dim Cz As Single
    Dim Dx As Single
    Dim Dy As Single
    Dim Dz As Single
    Dim A1 As Single
    Dim B1 As Single
    Dim C1 As Single
    Dim B24AC As Single
    Dim t1 As Single
    Dim t2 As Single
    
    ' Get the axis vector W
    Wx = Point2.Trans(1) - Point1.Trans(1)
    Wy = Point2.Trans(2) - Point1.Trans(2)
    Wz = Point2.Trans(3) - Point1.Trans(3)
    
    ' Find A and B (the cylinder equation coefficients)
    ' Find A
    WlenSquared = Wx * Wx + Wy * Wy + Wz * Wz
    WdotV = Wx * Vx + Wy * Vy + Wz * Vz
    A = WdotV / WlenSquared
    
    ' Find B
    X1 = Point1.Trans(1)
    Y1 = Point1.Trans(2)
    Z1 = Point1.Trans(3)
    B = (Wx * (px - X1) + _
         Wy * (py - Y1) + _
         Wz * (pz - Z1)) / WlenSquared
    
    ' Solve the equation of the cylinder for t.
    
    ' Values:
    ' Cx, Cy, Cz
    Cx = Vx - Wx * A
    Cy = Vy - Wy * A
    Cz = Vz - Wz * A
    ' Dx, Dy, Dz
    Dx = px - X1 - Wx * B
    Dy = py - Y1 - Wy * B
    Dz = pz - Z1 - Wz * B
    ' A1, B1 and C1
    A1 = Cx * Cx + Cy * Cy + Cz * Cz
    B1 = 2 * (Cx * Dx + Cy * Dy + Cz * Dz)
    C1 = Dx * Dx + Dy * Dy + Dz * Dz - Radius * Radius
    
    ' Solve the equation A1*t^2 + B1*t + C1 = 0.
    B24AC = B1 * B1 - 4 * A1 * C1
    If B24AC < 0 Then
        RayTraceable_FindT = -1
        Exit Function
    ElseIf B24AC = 0 Then
        t1 = -B1 / 2 / A1
    Else
        B24AC = Sqr(B24AC)
        t1 = (-B1 + B24AC) / 2 / A1
        t2 = (-B1 - B24AC) / 2 / A1
        ' Use only positive t values.
        If t1 < 0.01 Then t1 = t2
        If t2 < 0.01 Then t2 = t1
        ' Use the smaller t value.
        If t1 > t2 Then t1 = t2
    End If

    ' If there is no positive t value, there's no
    ' intersection in this direction.
    If t1 < 0.01 Then
        RayTraceable_FindT = -1
        Exit Function
    End If

    ' See where on the cylinder this point is.
    HitU = t1 * A + B

    ' If this is not between Point1 and Point2,
    ' ignore it.
    If HitU < 0 Or HitU > 1 Then
        RayTraceable_FindT = -1
    Else
        RayTraceable_FindT = t1
    End If
End Function

Public Sub RayTraceable_GetRminRmax(new_min As Single, new_max As Single, ByVal X As Single, ByVal Y As Single, ByVal Z As Single)
    Dim Dx As Single
    Dim Dy As Single
    Dim Dz As Single
    Dim Dist As Single

    ' Get the distance to the first point.
    Dx = X - Point1.Trans(1)
    Dy = Y - Point1.Trans(2)
    Dz = Z - Point1.Trans(3)
    Dist = Sqr(Dx * Dx + Dy * Dy + Dz * Dz)
    new_max = Dist + Radius
    new_min = Dist - Radius

    ' Get the distance to the second point.
    Dx = X - Point2.Trans(1)
    Dy = Y - Point2.Trans(2)
    Dz = Z - Point2.Trans(3)
    Dist = Sqr(Dx * Dx + Dy * Dy + Dz * Dz)
    If new_max < Dist + Radius Then new_max = Dist + Radius
    If new_min > Dist - Radius Then new_min = Dist - Radius

    If new_min < 0 Then new_min = 0
End Sub

Public Sub SetValues(ByVal X1 As Single, ByVal Y1 As Single, ByVal Z1 As Single, _
    ByVal X2 As Single, ByVal Y2 As Single, ByVal Z2 As Single, _
    ByVal Rad As Single, _
    ByVal AmbientKr As Single, ByVal AmbientKg As Single, ByVal AmbientKb, _
    ByVal DiffuseKr As Single, ByVal DiffuseKg As Single, ByVal DiffuseKb, _
    ByVal SpecularK As Single, ByVal SpecularN As Single, _
    ByVal ReflectedKr As Single, ByVal ReflectedKg As Single, ByVal ReflectedKb As Single)
    
    ' Assign values to local variables
    ' Position and radius
    Point1.Coord(1) = X1
    Point1.Coord(2) = Y1
    Point1.Coord(3) = Z1
    Point1.Coord(4) = 1#
    Point2.Coord(1) = X2
    Point2.Coord(2) = Y2
    Point2.Coord(3) = Z2
    Point2.Coord(4) = 1#
    Radius = Rad
    
    ' Lighting
    ' Ambient
    AmbKr = AmbientKr
    AmbKg = AmbientKg
    AmbKb = AmbientKb
    ' Diffuse
    DiffKr = DiffuseKr
    DiffKg = DiffuseKg
    DiffKb = DiffuseKb
    ' Specular
    Spec_K = SpecularK
    Spec_N = SpecularN
    ' Set reflected factor
    ReflKr = ReflectedKr
    ReflKg = ReflectedKg
    ReflKb = ReflectedKb
End Sub

Public Sub RayTraceable_ResetCulling()
    ForeverCulled = False
    HadHitPrev = False
End Sub
