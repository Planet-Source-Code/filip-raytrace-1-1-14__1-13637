VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Disk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Disk

Implements RayTraceable

' Position
Private Point1 As Point3D ' A point on the plane
Private Point2 As Point3D ' point2-point1 = surface normal
Private Radius As Single ' Radius of the disk

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

' Find the unit surface normal
Public Sub GetUNormal(ByRef Nx As Single, ByRef Ny As Single, ByRef Nz As Single)
    Dim Nlen As Single
    Nx = Point2.Trans(1) - Point1.Trans(1)
    Ny = Point2.Trans(2) - Point1.Trans(2)
    Nz = Point2.Trans(3) - Point1.Trans(3)
    Nlen = Sqr(Nx * Nx + Ny * Ny + Nz * Nz)
    Nx = Nx / Nlen
    Ny = Ny / Nlen
    Nz = Nz / Nlen
End Sub

Public Sub RayTraceable_Apply(M() As Single)
    m3Apply Point1.Coord, M, Point1.Trans
    m3Apply Point2.Coord, M, Point2.Trans
End Sub

Public Sub RayTraceable_ApplyFull(M() As Single)
    m3ApplyFull Point1.Coord, M, Point1.Trans
    m3ApplyFull Point2.Coord, M, Point2.Trans
End Sub

Public Sub RayTraceable_CullScanline(ByVal px As Single, ByVal py As Single, ByVal pz As Single, ByVal Nx As Single, ByVal Ny As Single, ByVal Nz As Single)
    Dim Dx As Single
    Dim Dy As Single
    Dim Dz As Single
    Dim Dist As Single

    ' See if we will ever be visible again.
    If ForeverCulled Then
        ScanlineDone = True
        Exit Sub
    End If

    ' We have not yet had a hit on this scanline.
    HadHit = False

    ' Find the distance from the center of the
    ' disk to the scanline plane.

    ' Get the vector from our center to the point.
    With Point1
        Dx = .Trans(1) - px
        Dy = .Trans(2) - py
        Dz = .Trans(3) - pz
    End With

    ' Take the dot product of this and the normal.
    ' If the resulting distance > Radius, cull.
    ScanlineDone = (Abs(Dx * Nx + Dy * Ny + Dz * Nz) > Radius)

    ' See if we will be culled in the future.
    If ScanlineDone Then
        ' We were not culled on a previous scanline
        ' but we are now. We will be culled on
        ' all later scanlines.
        If HadHitPrev Then ForeverCulled = True
    Else
        ' We are not culled. Remember that.
        HadHitPrev = True
    End If
End Sub

Public Sub RayTraceable_FindHitColor(Objects As Collection, ByVal eyeX As Single, ByVal eyeY As Single, ByVal eyeZ As Single, ByVal px As Single, ByVal py As Single, ByVal pz As Single, R As Integer, G As Integer, B As Integer)
    Dim Nx As Single
    Dim Ny As Single
    Dim Nz As Single
    Dim Vx As Single
    Dim Vy As Single
    Dim Vz As Single
    Dim NdotV As Single
    
    ' Get the unit normal
    GetUNormal Nx, Ny, Nz
    
    ' Make sure the normal points towards the eye
    Vx = Eye_X - px
    Vy = Eye_Y - py
    Vz = Eye_Z - pz
    NdotV = Nx * Vx + Ny * Vy + Nz * Vz
    If NdotV < 0 Then
        Nx = -Nx
        Ny = -Ny
        Nz = -Nz
    End If
    
    ' Calculate the hit color
    CalculateHitColor Objects, Me, _
        eyeX, eyeY, eyeZ, _
        px, py, pz, _
        Nx, Ny, Nz, _
        DiffKr, DiffKg, DiffKb, _
        AmbKr, AmbKg, AmbKb, _
        Spec_K, Spec_N, _
        ReflKr, ReflKg, ReflKb, _
        R, G, B
End Sub

Public Function RayTraceable_FindT(DirectC As Boolean, px As Single, py As Single, pz As Single, Vx As Single, Vy As Single, Vz As Single) As Single
    Dim A As Single
    Dim B As Single
    Dim C As Single
    Dim D As Single
    Dim Nx As Single
    Dim Ny As Single
    Dim Nz As Single
    Dim Denom As Single
    Dim t As Single
    Dim Cx As Single
    Dim Cy As Single
    Dim Cz As Single
    Dim Dx As Single
    Dim Dy As Single
    Dim Dz As Single
    Dim X As Single
    Dim Y As Single
    Dim Z As Single
    
    If DirectC And ScanlineDone Then
        RayTraceable_FindT = -1
        Exit Function
    End If
    
    ' Find the unit normal
    GetUNormal Nx, Ny, Nz
    
    ' Calculate the disk's parameters
    A = Nx
    B = Ny
    C = Nz
    D = -(Nx * Point1.Trans(1) + _
        Ny * Point1.Trans(2) + _
        Nz * Point1.Trans(3))
            
    ' If the denominator in the equation for t equals 0,
    ' the ray is parralel to the disk so there's no
    ' intersection
    Denom = A * Vx + B * Vy + C * Vz
    If Denom = 0 Then
        RayTraceable_FindT = -1
        Exit Function
    End If
    
    ' Solve for t
    t = -(A * px + B * py + C * pz + D) / Denom
    
    ' If there's no positive t value, there's no intersection
    ' in this direction
    If t < 0.01 Then
        RayTraceable_FindT = -1
        Exit Function
    End If
    
    ' Get the coords of the disk's center
    Cx = Point1.Trans(1)
    Cy = Point1.Trans(2)
    Cz = Point1.Trans(3)
    
    ' Get the point of intersection
    X = px + t * Vx
    Y = py + t * Vy
    Z = pz + t * Vz
    
    ' See if the point lies within the disk
    Dx = Cx - X
    Dy = Cy - Y
    Dz = Cz - Z
    If Dx * Dx + Dy * Dy + Dz * Dz > Radius * Radius Then
        ' We are not within the disk
        RayTraceable_FindT = -1
        Exit Function
    End If
    
    ' We had a hit
    If DirectC Then HadHit = True
    RayTraceable_FindT = t
End Function

Public Sub SetValues(ByVal p1x As Single, ByVal p1y As Single, ByVal p1z As Single, _
    ByVal p2x As Single, ByVal p2y As Single, ByVal p2z As Single, _
    ByVal Rad As Single, _
    ByVal AmbientKr As Single, ByVal AmbientKg As Single, ByVal AmbientKb As Single, _
    ByVal DiffuseKr As Single, ByVal DiffuseKg As Single, ByVal DiffuseKb As Single, _
    ByVal SpecularK As Single, ByVal SpecularN As Single, _
    ByVal ReflectedKr As Single, ByVal ReflectedKg As Single, ByVal ReflectedKb As Single)
    
    ' Set point1
    Point1.Coord(1) = p1x
    Point1.Coord(2) = p1y
    Point1.Coord(3) = p1z
    
    ' Set point2
    Point2.Coord(1) = p2x
    Point2.Coord(2) = p2y
    Point2.Coord(3) = p2z
    
    ' Set Radius
    Radius = Rad
    
    ' Set ambient lighting
    AmbKr = AmbientKr
    AmbKg = AmbientKg
    AmbKb = AmbientKb
    
    ' Set diffuse lighting
    DiffKr = DiffuseKr
    DiffKg = DiffuseKg
    DiffKb = DiffuseKb
    
    ' Set specular lighting
    Spec_K = SpecularK
    Spec_N = SpecularN
    
    ' Set reflected factor
    ReflKr = ReflectedKr
    ReflKg = ReflectedKg
    ReflKb = ReflectedKb
End Sub

Private Sub RayTraceable_GetRminRmax(new_min As Single, new_max As Single, ByVal X As Single, ByVal Y As Single, ByVal Z As Single)
    Dim dminx As Single
    Dim dminy As Single
    Dim dminz As Single
    Dim dmaxx As Single
    Dim dmaxy As Single
    Dim dmaxz As Single
    Dim Dist As Single

    new_min = 1E+30
    new_max = -1E+30

    dminx = X - Point1.Trans(1) - Radius
    dminy = Y - Point1.Trans(2) - Radius
    dminz = Z - Point1.Trans(3) - Radius
    dmaxx = X - Point1.Trans(1) + Radius
    dmaxy = Y - Point1.Trans(2) + Radius
    dmaxz = Z - Point1.Trans(3) + Radius
    
    new_max = Sqr(dmaxx * dmaxx + dmaxy * dmaxy + dmaxz * dmaxz)
    new_min = Sqr(dminx * dminx + dminy * dminy + dminz * dminz)
End Sub

Public Sub RayTraceable_ResetCulling()
    ForeverCulled = False
    HadHitPrev = False
End Sub
