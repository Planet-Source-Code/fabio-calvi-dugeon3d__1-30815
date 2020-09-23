VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Character Animation-By MartWare-FPS:"
   ClientHeight    =   6810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   454
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   663
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrMain 
      Interval        =   1000
      Left            =   1800
      Top             =   1560
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This program wants only show the high potential of Direct3DRM
'This example don't use 3D Card, use processor; see the CreateDeviceFromClipper command.
'I have used Direct X files to create the walls instead of surfaces, following an old
'Patrice Scribe example.
'It looks like having a old Wolfstein 3D or linear Doom view style, but I think it
'gives a nice effect, however.
'The collision detection is not implemented yet, sorry.
'I know there is the Pick command but I am not well documentated about it.
'This is my first program with DX7 Direct3DRM and I don't know it well.
'I don 't find documentations about it, in my country; the only news I can find about
'D3D are in the net.
'By the way, if someone wish to send me some informations about it, please email me to:
'FABIOCALVI@ YAHOO.COM
'Thank you, very much.
'Happy coding and enjoy, Fabio.

Option Explicit

' direct x objects
Dim Dx As New DirectX7
Dim Dd As DirectDraw4
Dim clip As DirectDrawClipper
Dim d3drm As Direct3DRM3
Dim scene As Direct3DRMFrame3
Dim cam As Direct3DRMFrame3
Dim dev As Direct3DRMDevice3
Dim view As Direct3DRMViewport2
Dim mesh As Direct3DRMMeshBuilder3
Dim XFileTex As Direct3DRMTexture3

Dim LightFrame As Direct3DRMFrame3
Dim ViewFrame(13, 23) As Direct3DRMFrame3
Dim Wall(8) As Direct3DRMMeshBuilder3
Dim TCase As Integer

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

'Main sub
Private Sub Form_Load()
ShowCursor 0
TCase = 10
Const Sin5 = 8.715574E-02!  ' Sin(5°)
Const Cos5 = 0.9961947!     ' Cos(5°)
    
    ' init direct draw and clipper
    Set Dd = Dx.DirectDraw4Create("")
    Set clip = Dd.CreateClipper(0)
    clip.SetHWnd Me.hWnd
    
    ' screen mode
    Dd.SetDisplayMode 800, 600, 32, 0, DDSDM_DEFAULT
    
    ' init direct 3drm and main frames
    Set d3drm = Dx.Direct3DRMCreate()
    Set scene = d3drm.CreateFrame(Nothing)
    Set cam = d3drm.CreateFrame(scene)
    cam.SetPosition scene, 10, 5, 30
    
    ' add lights
    Dim light As Direct3DRMLight
    Set light = d3drm.CreateLightRGB(D3DRMLIGHT_POINT, 0.55, 0.55, 0.55)
    light.SetUmbra 0.8: light.SetPenumbra 1.1
    Set LightFrame = d3drm.CreateFrame(scene)
    LightFrame.SetPosition Nothing, 0, 10, 0
    LightFrame.AddLight light
    
    ' add a bit of ambient light to the scene
    scene.AddLight d3drm.CreateLightRGB(D3DRMLIGHT_AMBIENT, 0.65, 0.65, 0.65)
    
    ' make viewport and device (I use IID_IDirect3DRGBDevice because I don't have a 3d card, sigh!)
    Set dev = d3drm.CreateDeviceFromClipper(clip, "IID_IDirect3DRGBDevice", Me.ScaleWidth, Me.ScaleHeight)
    ' unrem the next line to use hardware rendered (3D Card)
'    Set dev = d3drm.CreateDeviceFromClipper(clip, "IID_IDirect3DHALDevice", Me.ScaleWidth, Me.ScaleHeight)
    
    dev.SetQuality D3DRMFILL_SOLID + D3DRMLIGHT_ON + D3DRMSHADE_GOURAUD
    dev.SetDither D_TRUE ' texture correction
    Set view = d3drm.CreateViewport(dev, cam, 0, 0, Me.ScaleWidth, Me.ScaleHeight)
    view.SetBack 3000!
    
    ' create the meshbuilder from which we will add faces(walls) to create our world
    Set mesh = d3drm.CreateMeshBuilder()
    mesh.SetPerspective D_TRUE ' texture correction
    scene.AddVisual mesh ' add mesh builder to scene
    
    Set Wall(0) = Charge_X("mur.x")
    Set Wall(1) = Charge_X("sol.x")
    Set Wall(2) = Charge_X("red.x")
    Set Wall(3) = Charge_X("green.x")
    Set Wall(4) = Charge_X("blue.x")
    Set Wall(5) = Charge_X("mur2.x")
    Set Wall(6) = Charge_X("sol2.x")
    Set Wall(7) = Charge_X("mur3.x")
    Set Wall(8) = Charge_X("sol3.x")
    
    ' Create all wall & light
    Dim i%, j%
    For i% = 0 To 13
        For j% = 0 To 23
            ' Create view
            Set ViewFrame(i%, j%) = d3drm.CreateFrame(scene)
             ViewFrame(i%, j%).SetPosition scene, i% * TCase, 0, j% * TCase
        Next j%
    Next i%
    ' Load the map (this method is taken from an old Patrice Scribe example, thanks to him)
    Dim Carte$(13)
    Carte$(13) = "XXXXXXXXXXXXXXYYYYYYYYYY"
    Carte$(12) = "X............XY********Y"
    Carte$(11) = "X..X...X....RXY***Y****Y"
    Carte$(10) = "X..X...X......*********Y"
     Carte$(9) = "X...X.X.....RXY***Y****Y"
     Carte$(8) = "X...X.XX.....XY********Y"
     Carte$(7) = "X...X.XX..X..XYYYYYYYYYY"
     Carte$(6) = "X..GX.XX..X..XZZZZZZZZZZ"
     Carte$(5) = "X..XX.X......XZ°°°°°°°°Z"
     Carte$(4) = "X........X...XZ°°Z°°Z°°Z"
     Carte$(3) = "X..XXXX.....BXZ°°°°°°°°Z"
     Carte$(2) = "X....GXX......°°°°°°°°°Z"
     Carte$(1) = "X......XX...BXZ°°°°°°°°Z"
     Carte$(0) = "XXXXXXXXXXXXXXZZZZZZZZZZ"
    For i% = 0 To 13
        For j% = 0 To 23
            With ViewFrame(i%, j%)
                Select Case Mid$(Carte$(i%), j% + 1, 1)
                Case "X"
                    Call .AddVisual(Wall(0))
                Case "."
                    Call .AddVisual(Wall(1))
                Case "R"
                    Call .AddVisual(Wall(1))
                    Call .AddVisual(Wall(2))
                Case "G"
                    Call .AddVisual(Wall(1))
                    Call .AddVisual(Wall(3))
                Case "B"
                    Call .AddVisual(Wall(1))
                    Call .AddVisual(Wall(4))
                 Case "Y"
                    Call .AddVisual(Wall(5))
                 Case "*"
                    Call .AddVisual(Wall(6))
                 Case "Z"
                    Call .AddVisual(Wall(7))
                 Case "°"
                    Call .AddVisual(Wall(8))
               
                End Select
            End With
        Next j%
    Next i%
    ' show form before end of load
    Me.Show
    Me.Refresh
    DoEvents
    
    ' start main app loop
    Do While DoEvents()
        
       
        'Move forward
        If GetKeyState(vbKeyUp) < -1 Then cam.SetPosition cam, 0, 0, 2
   
        'Move back
        If GetKeyState(vbKeyDown) < -1 Then cam.SetPosition cam, 0, 0, -2
        
        'Rotate left
        If GetKeyState(vbKeyLeft) < -1 Then cam.SetOrientation cam, -Sin5, 0, Cos5, 0, 1, 0

        'Rotate right
        If GetKeyState(vbKeyRight) < -1 Then cam.SetOrientation cam, Sin5, 0, Cos5, 0, 1, 0
        
        'Funny view
        If GetKeyState(vbKeyPageUp) < -1 Then cam.SetPosition cam, 0, 1, 0
        If GetKeyState(vbKeyPageDown) < -1 Then cam.SetPosition cam, 0, -1, 0

'        Dim pos As D3DVECTOR
'        cam.GetPosition scene, pos 'Camera's position in scene
'
'        Dim wallpos As D3DVECTOR
'
'        For i% = 0 To 13
'           For j% = 0 To 23
'               ViewFrame(i%, j%).GetPosition scene, wallpos
'               If Mid$(Carte$(i%), j% + 1, 1) = "X" or _
'                  Mid$(Carte$(i%), j% + 1, 1) = "Y" or _
'                  Mid$(Carte$(i%), j% + 1, 1) = "Z" Then
'                  If Int(pos.x) > wallpos.x And Int(pos.x) < wallpos.x + 10 And _
'                     Int(pos.z) > wallpos.z And Int(pos.z) < wallpos.z + 10 Then
'                         -------------
'                  End If
'               End If
'           Next
'        Next
        
        ' render the scene
        view.Clear D3DRMCLEAR_ALL
        view.Render scene
        dev.Update
        
        ' check to exit
        If GetKeyState(vbKeyEscape) < -5 Then Unload Me
        
    Loop
    
End Sub
  
Private Function Charge_X(Nom$) As Direct3DRMMeshBuilder3
    Set Charge_X = d3drm.CreateMeshBuilder()
    Charge_X.LoadFromFile Nom$, 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
End Function

Private Sub Form_Unload(Cancel As Integer)
    
Dd.RestoreDisplayMode
Dd.SetCooperativeLevel frmMain.hWnd, DDSCL_NORMAL

Dim i%, j%
For i% = 0 To 13
    For j% = 0 To 23
        Set ViewFrame(i%, j%) = Nothing
    Next j%
Next i%
For i% = 0 To 6
    Set Wall(i%) = Nothing
Next i%

Set dev = Nothing
Set clip = Nothing
Set d3drm = Nothing
Set Dd = Nothing
Set Dx = Nothing

ShowCursor 1

End
End Sub
