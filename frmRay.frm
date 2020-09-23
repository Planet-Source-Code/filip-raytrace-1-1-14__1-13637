VERSION 5.00
Begin VB.Form frmRay 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ray Tracing"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   320
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   456
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "P r e v i e w"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   13
      Top             =   2760
      Width           =   1815
      Begin VB.PictureBox picPreview 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00400000&
         Height          =   1560
         Left            =   120
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   14
         Top             =   240
         Width           =   1560
      End
   End
   Begin VB.TextBox txtStep 
      Height          =   285
      Left            =   840
      TabIndex        =   3
      Text            =   "1"
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdRender 
      Caption         =   "Render"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Viewpoint"
      Height          =   1335
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1815
      Begin VB.TextBox txtEyeTheta 
         Height          =   285
         Left            =   720
         TabIndex        =   7
         Text            =   "-0.3"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtEyePhi 
         Height          =   285
         Left            =   720
         TabIndex        =   6
         Text            =   "-0.6"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtEyeR 
         Height          =   285
         Left            =   720
         TabIndex        =   5
         Text            =   "1000"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Theta:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Phi:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "R:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.PictureBox pic1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00400000&
      Height          =   4800
      Left            =   2040
      ScaleHeight     =   316
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   316
      TabIndex        =   0
      Top             =   0
      Width           =   4800
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Time:"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblTime 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0 sec"
      Height          =   255
      Left            =   840
      TabIndex        =   11
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Step:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   495
   End
End
Attribute VB_Name = "frmRay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' RayTrace V1.1.13

' ---!!!!SPEED IMPROVED!!!!---
' Now this program features scanline culling! It
' checks if a so-called scanline plane intersects
' any objects. If not, it skips tracing all the
' rays on a scanline. This improves speed quite
' a lot! One disappointment of this technique is
' that it is only performed good on spheres, but
' it will be improved with later versions

' If you want to learn the geometry (FindT and
' FindHitColor functions) of the objects, start
' with the sphere. This is the most simple object.
' For further documentation on the sphere, look at
' the HTML page included in the ZIP.

' For more information about Ray Tracing and the
' geometry of objects, look for the book Visual Basic
' Graphics Programming, Second Edition, by Rod
' Stephens, published by Wiley. (ISBN 0-471-35599-2)

' For experts on ray tracing, there is even a new
' technique, called radiosity. This is extremely
' complex, and it generates the most photo-realistic
' images ever generated by computers. It can realize
' real ambient lighting, really generated by the
' reflection of light caused by nearby objects. It
' can also analyze the light spectrum with the colors
' of the rainbow caused by prism's. This is extremely
' difficult, and quite slow for a personal computer.
' (You should let work a Cray supercomputer on radiosity
' for a little speed.)
' For more information about radiosity, consult the
' latest graphics research literature.

Option Explicit

Private Sub cmdRender_Click()
    Dim Obj As RayTraceable
    Dim T As Single
    If Running = False Then
        T = Timer
        For Each Obj In Objects
            Obj.ResetCulling
        Next Obj
        
        ' Show an error message when the inserted values
        ' are not numeric
        If Not ((IsNumeric(txtEyeR)) And (IsNumeric(txtEyePhi)) And (IsNumeric(txtEyeTheta))) Then
            MsgBox "Enter numeric values", , "Ray"
            Exit Sub
        End If
        Running = True
        
        ' Set the eye's position
        EyeR = CSng(txtEyeR.Text)
        EyePhi = CSng(txtEyePhi.Text)
        EyeTheta = CSng(txtEyeTheta.Text)
        
        ' Change the caption of the commandbutton
        cmdRender.Caption = "Stop"
        
        ' Clear the picturebox
        pic1.Cls
        
        ' Render
        Render pic1, txtStep
        
        ' Set the caption of the button back to "Render"
        cmdRender.Caption = "Render"
        lblTime = Timer - T & " sec"
    Else
        ' Stop ray tracing
        Running = False
        cmdRender.Caption = "Render"
        lblTime = Timer - T & " sec"
    End If
End Sub

Private Sub Form_Load()
    Dim Sphere1 As Sphere
    Dim Sphere2 As Sphere
    Dim Sphere3 As Sphere
    Dim Sphere4 As Sphere
    Dim Cyl1 As Cylinder
    Dim Cyl2 As Cylinder
    Dim Cyl3 As Cylinder
    Dim Cyl4 As Cylinder
    Dim Cyl5 As Cylinder
    Dim Cyl6 As Cylinder
    Dim Disk1 As Disk
    Dim Light1 As LightSource
    Dim Light2 As LightSource
    
    ' Show the form
    Me.Show
    DoEvents
    
    ' Set the ambient lighting
    AmbIr = 128
    AmbIg = 128
    AmbIb = 128
    
    ' Set the eye position
    EyeR = 1000
    EyePhi = -0.6
    EyeTheta = -0.3
    
    ' Create new light sources
    Set Light1 = New LightSource
    Set Light2 = New LightSource
        
    ' Set the values of the light sources
    Light1.SetParameters 1000, -500, 1000, 255, 255, 255
    Light2.SetParameters 1000, -500, -1000, 255, 255, 255
    
    ' Add the light sources to the LightSources array
    LightSources.Add Light1
    LightSources.Add Light2
    
    ' Create new spheres
    Set Sphere1 = New Sphere
    Set Sphere2 = New Sphere
    Set Sphere3 = New Sphere
    Set Sphere4 = New Sphere
    
    ' Create new cylinders
    Set Cyl1 = New Cylinder
    Set Cyl2 = New Cylinder
    Set Cyl3 = New Cylinder
    Set Cyl4 = New Cylinder
    Set Cyl5 = New Cylinder
    Set Cyl6 = New Cylinder
    
    ' Create a new disk
    Set Disk1 = New Disk
    
    ' Set the values of the spheres
    Sphere1.SetValues 75, 0, 0, 30, _
        0.6, 0.1, 0.1, _
        0.6, 0.1, 0.1, _
        0.35, 20, _
        0, 0, 0
    Sphere2.SetValues -35, 0, -65, 30, _
        0.1, 0.5, 0.1, _
        0.1, 0.5, 0.1, _
        0.35, 20, _
        0, 0, 0
    Sphere3.SetValues -35, 0, 65, 30, _
        0.1, 0.1, 0.6, _
        0.1, 0.1, 0.6, _
        0.35, 20, _
        0, 0, 0
    Sphere4.SetValues 0, -65, 0, 30, _
        0.6, 0.1, 0.6, _
        0.6, 0.1, 0.6, _
        0.35, 20, _
        0, 0, 0
    ' Set the values of the cylinders
    Cyl1.SetValues 75, 0, 0, -35, 0, -65, 15, _
        0.1, 0.1, 0.6, _
        0.1, 0.1, 0.6, _
        0.35, 20, _
        0, 0, 0
    Cyl2.SetValues -35, 0, -65, -35, 0, 65, 15, _
        0.6, 0.1, 0.1, _
        0.6, 0.1, 0.1, _
        0.35, 20, _
        0, 0, 0
    Cyl3.SetValues -35, 0, 65, 75, 0, 0, 15, _
        0.1, 0.5, 0.1, _
        0.1, 0.5, 0.1, _
        0.35, 20, _
        0, 0, 0
    Cyl4.SetValues 75, 0, 0, 0, -65, 0, 15, _
        0.6, 0.1, 0.6, _
        0.6, 0.1, 0.6, _
        0.35, 20, _
        0, 0, 0
    Cyl5.SetValues -35, 0, -65, 0, -65, 0, 15, _
        0.6, 0.6, 0.1, _
        0.6, 0.6, 0.1, _
        0.35, 20, _
        0, 0, 0
    Cyl6.SetValues -35, 0, 65, 0, -65, 0, 15, _
        0.1, 0.5, 0.5, _
        0.1, 0.5, 0.5, _
        0.35, 20, _
        0, 0, 0
        
    ' Set the values of the disk
    Disk1.SetValues 0, 30, 0, 0, -31, 0, 125, 0.1, 0.1, 0.1, _
        0.1, 0.1, 0.1, 0.35, 20, 0.9, 0.9, 0.9
    
    ' Add the objects to the objects array
    Objects.Add Sphere1
    Objects.Add Sphere2
    Objects.Add Sphere3
    Objects.Add Sphere4
    Objects.Add Cyl1
    Objects.Add Cyl2
    Objects.Add Cyl3
    Objects.Add Cyl4
    Objects.Add Cyl5
    Objects.Add Cyl6
    Objects.Add Disk1
    
    ' Uncomment the following line for a DNA-like structure:
    'DNACreate
    ' If you did, comment the object adding lines
    ' in this sub.
    
    ' Render preview
    PRunning = True
    PreviewRender picPreview, 3
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub txtEyePhi_Change()
    If IsNumeric(txtEyePhi) Then
        EyePhi = CSng(txtEyePhi)
        picPreview.Cls
        PRunning = True
        PreviewRender picPreview, 3
    End If
End Sub

Private Sub txtEyeTheta_Change()
    If IsNumeric(txtEyeTheta) Then
        EyeTheta = CSng(txtEyeTheta)
        picPreview.Cls
        PRunning = True
        PreviewRender picPreview, 3
    End If
End Sub

Private Sub DNACreate()
    Dim Sphere As Sphere
    Dim Bases(1 To 2, 1 To 40) As Sphere
    Dim Helix(1 To 2, 1 To 40) As Sphere
    Dim i
    Dim Rand As Single
    For i = 1 To 40
        Set Bases(1, i) = New Sphere
        Set Bases(2, i) = New Sphere
        Set Helix(1, i) = New Sphere
        Set Helix(2, i) = New Sphere
    Next i
    For i = 1 To 40
        Rand = Rnd
        If Rand < 0.5 Then
            Bases(1, i).SetValues 15 * Sin(i / 3), _
                (i) * 7.5 - 150, 15 * Cos(i / 3), 10, _
                0.6, 0.1, 0.1, _
                0.6, 0.1, 0.1, _
                0.35, 20, _
                0, 0, 0
            Bases(2, i).SetValues 15 * Sin(i / 3 + 1.5707963), _
                (i) * 7.5 - 150, 15 * Cos(i / 3 + 1.5707963), 10, _
                0.1, 0.6, 0.1, _
                0.1, 0.6, 0.1, _
                0.35, 20, _
                0, 0, 0
        ElseIf Rand >= 0.5 Then
            Bases(1, i).SetValues 15 * Sin(i / 3), _
                (i) * 7.5 - 150, 15 * Cos(i / 3), 10, _
                0.1, 0.1, 0.6, _
                0.1, 0.1, 0.6, _
                0.35, 20, _
                0, 0, 0
            Bases(2, i).SetValues 15 * Sin(i / 3 + 1.5707963), _
                (i) * 7.5 - 150, 15 * Cos(i / 3 + 1.5707963), 10, _
                0.6, 0.1, 0.6, _
                0.6, 0.1, 0.6, _
                0.35, 20, _
                0, 0, 0
        End If
    Next i
    For i = 1 To 40
        Helix(1, i).SetValues 25 * Sin(i / 3), _
            (i) * 7.5 - 150, 25 * Cos(i / 3), 10, _
            0.6, 0.6, 0.6, _
            0.6, 0.6, 0.6, _
            0.35, 20, _
            0, 0, 0
        Helix(2, i).SetValues 25 * Sin(i / 3 + 1.5707963), _
            (i) * 7.5 - 150, 25 * Cos(i / 3 + 1.5707963), 10, _
            0.6, 0.6, 0.6, _
            0.6, 0.6, 0.6, _
            0.35, 20, _
            0, 0, 0
    Next i
    For i = 1 To 40
        Objects.Add Bases(1, i)
        Objects.Add Bases(2, i)
    Next i
    For i = 1 To 40
        Objects.Add Helix(1, i)
        Objects.Add Helix(2, i)
    Next i
End Sub
