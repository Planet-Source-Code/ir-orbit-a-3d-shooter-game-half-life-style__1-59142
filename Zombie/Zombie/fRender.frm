VERSION 5.00
Begin VB.Form fRender 
   Caption         =   "Form1"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   ScaleHeight     =   436
   ScaleMode       =   3  'Ïèêñåëü
   ScaleWidth      =   525
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "fRender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------'
' Zombie
'-----------------------------------------------------------'
' Alex, AKiR (c) 2004-2005                                 '
'-----------------------------------------------------------'

Option Explicit
Private iR          As New iR_Engine
Private Control     As New iR_Control
Private TexFactory  As New iR_TextureFactory
Private Inface      As New iR_Interface
Private Maths       As New iR_Maths
Private iFont       As New iRObject_Font
Private iSound      As New iR_SoundEngine
Private Light       As New iR_LightEngine
'==================================================
Private Camera      As New iR_Camera
Private CamPos      As iR_Vector3D
Private CamRot      As iR_Vector3D
Private NewCamPos   As iR_Vector3D
'==================================================
Private Level       As New iR_BSPTree
'==================================================
Private Hero(1)     As New iR_ActorMDL
Private HeroPos     As iR_Vector3D
Private HeroRot     As iR_Vector3D
Private HeroDeath   As Boolean
'==================================================
Private Aid(3)      As New iR_ActorMDL
Private AidPos(3)   As iR_Vector3D
Private AidRot(3)   As iR_Vector3D
Private AidEnbl(3) As Boolean
'==================================================
Private BoxShell(3)    As New iR_ActorMDL
Private BoxShellPos(3) As iR_Vector3D
Private BoxShellRot(3) As iR_Vector3D
Private BoxShellEnbl(3) As Boolean
'==================================================
Private Zombie(9)       As New iR_ActorMD2
Private ZombiePos(9)    As iR_Vector3D
Private ZombieRot(9)    As iR_Vector3D
Private NewZombPos(9)   As iR_Vector3D
Private ZombieDeath(9)  As Boolean
Private ZombieAttack(9) As Boolean
Private ZombieSpr(9)    As iR_Sprite
Private zDir(9)         As iR_Vector3D
Private ZombieMat(9)    As iR_Material
Private ZombieEnbl(9)   As Boolean
Private Const AmountZombie As Integer = 9
Private BloodColor      As Integer
Private ZombieSprite    As iR_Sprite
'==================================================
Private Prizel      As iR_Sprite
Private Shell       As iR_Sprite
Private Cross       As iR_Sprite
Private ShellBox    As iR_Sprite
Private SmokeSpr    As iR_Sprite
Private SmokeSprEnbl  As Boolean
Private SmokeSprColor As Integer
'==================================================
Private Bullet      As Integer
Private Health      As Integer
Private Weapon      As Integer
Private SLBox       As Integer
Private Possible    As Boolean
Private i           As Integer
Private ii          As Integer
Private HrSoundDt   As Boolean
Private eDir             As iR_Vector3D
Private Vec              As iR_Vector3D
Private ldir             As iR_Vector3D
Private ColorLight       As iR_ColorValue
Private ColorDiffuse(9)  As iR_ColorValue
Private ColorEmissive(9) As iR_ColorValue
'==================================================
Private OldTick     As Long
Private TickNow     As Long
Private Movement    As Single
'==================================================
Private Const Pi = 3.14159265358979

Private Sub Form_Load()
    
    iR.SetLogging 1
    If Not iR.InitWithDialog(Me.hWnd) Then End
    Me.Show
    iR.SetDisplayFPS 1
    iR.SetViewFrustum 3000, 90
    iR.SetBackGroundColor RGB(55, 55, 55)
    iR.SetTextureFilter iR_Filter_Bilinear
    'iR.SwCur
'======================================================================
    ColorLight.Blue = 0.64
    ColorLight.Green = 1
    ColorLight.Red = 1
    ColorLight.alpha = 1
    ldir.X = 1
    ldir.y = 1
    ldir.z = 1
    Light.Light_EnableLighting 1
    Light.Light_CreateDirectional ldir, ColorLight
    
    
    For i = 0 To 9
       ColorDiffuse(i).Blue = 1
       ColorDiffuse(i).Green = 1
       ColorDiffuse(i).Red = 1
       ZombieMat(i).Diffuse = ColorDiffuse(i)
       ColorEmissive(i).Blue = 0.3
       ColorEmissive(i).Green = 0.3
       ColorEmissive(i).Red = 0.3
       ZombieMat(i).Emissive = ColorEmissive(i)
    Next i
    
   
 '======================================================================
  Prizel = iR.CreateSprite((CInt(iR.GetScreenResolution.X) - 32) * 0.5, (CInt(iR.GetScreenResolution.y) - 32) * 0.5, 32, 32, _
           TexFactory.LoadTextureEx("gfx/Prizel.bmp", "prizel", &HFF000000), RGB(255, 255, 255), 1, 1, 0, 1)
  Shell = iR.CreateSprite((CInt(iR.GetScreenResolution.X) - 32) * 0.015, (CInt(iR.GetScreenResolution.y) - 92) * 0.95, 32, 92, _
           TexFactory.LoadTextureEx("gfx/Shell.bmp", "shell", &HFF000000), RGB(255, 255, 255), 1, 1, 1, 0)
  Cross = iR.CreateSprite((CInt(iR.GetScreenResolution.X) - 40) * 0.85, (CInt(iR.GetScreenResolution.y) - 40) * 0.02, 40, 40, _
           TexFactory.LoadTextureEx("gfx/Cross.bmp", "Cross", &HFF000000), RGB(255, 255, 255), 1, 1, 0, 1)
  ShellBox = iR.CreateSprite((CInt(iR.GetScreenResolution.X) - 40) * 0.05, (CInt(iR.GetScreenResolution.y) - 40) * 0.02, 40, 40, _
           TexFactory.LoadTextureEx("gfx/Caps.bmp", "ShellBox", &HFF000000), RGB(255, 255, 255), 1, 1, 1, 0)
  SmokeSpr = iR.CreateSprite((CInt(iR.GetScreenResolution.X) - 64) * 0.5, (CInt(iR.GetScreenResolution.y) - 64) * 0.5, 64, 64, _
           TexFactory.LoadTextureEx("gfx/Smoke.bmp", "Smoke", &HFF000000), RGB(155, 155, 155), 1, 1, 0, 1)
  
'======================================================================
    Level.LoadBSP "maps/gad1.bsp", 8
'======================================================================
    Hero(0).LoadActor "Actor/v_knife.mdl"
    Hero(0).SetScale iR.CreateVec3(0.5, 0.5, 0.6)
    Hero(0).SetAnimation 0
    Hero(0).SetSpeed 0.8
'======================================================================
    Hero(1).LoadActor "Actor/v_m3.mdl"
    Hero(1).SetScale iR.CreateVec3(0.6, 1, 1.2)
    Hero(1).SetAnimation 0
    Hero(1).SetSpeed 0.8
'======================================================================
 For i = 0 To 3
    BoxShell(i).LoadActor "Actor/chainammo.mdl"
    BoxShell(i).SetScale iR.CreateVec3(0.7, 0.7, 0.7)
    BoxShellEnbl(i) = True
    
 Next i
    BoxShell(0).SetPosition iR.CreateVec3(-1600, -50, 327)
    BoxShell(1).SetPosition iR.CreateVec3(-1600, -90, -298)
    BoxShell(2).SetPosition iR.CreateVec3(-56, -90, 365)
    BoxShell(3).SetPosition iR.CreateVec3(-684, -90, 304)
'======================================================================
 For i = 0 To 3
    Aid(i).LoadActor "Actor/w_medkitl.mdl"
    Aid(i).SetScale iR.CreateVec3(0.5, 0.5, 0.5)
    AidEnbl(i) = True
 Next i
    Aid(0).SetPosition iR.CreateVec3(-1600, -90, -320)
    Aid(1).SetPosition iR.CreateVec3(-915, -90, -177)
    Aid(2).SetPosition iR.CreateVec3(-417, -90, 368)
    Aid(3).SetPosition iR.CreateVec3(-58, -90, -335)
'======================================================================
For i = 0 To AmountZombie
    ZombieSpr(i) = iR.CreateSprite(0, 0, 0, 0, iR.LoadTexture("gfx/DM_Base.bmp"), RGB(255, 255, 255))
    Zombie(i).LoadActor "Actor/Zombie4.md2"
    Zombie(i).SetTexture ZombieSpr(i)
    Zombie(i).SetScale iR.CreateVec3(1, 1.2, 1)
    Zombie(i).SetDrawType RightHanded
    Zombie(i).SetStartFrame 0
    Zombie(i).SetEndFrame 30
    Zombie(i).SetMaterial ZombieMat(i)
    Zombie(i).SetAmimationSpeed 0.015
    ZombieEnbl(i) = True
Next i
    Zombie(0).SetPosition iR.CreateVec3(-1584, -70, -367)
    Zombie(1).SetPosition iR.CreateVec3(-855, -70, 345)
    Zombie(2).SetPosition iR.CreateVec3(-741, -70, -183)
    Zombie(3).SetPosition iR.CreateVec3(-874, -70, -312)
    Zombie(4).SetPosition iR.CreateVec3(-919, -70, -183)
    Zombie(5).SetPosition iR.CreateVec3(-386, -70, 294)
    Zombie(6).SetPosition iR.CreateVec3(-140, -70, 333)
    Zombie(7).SetPosition iR.CreateVec3(986, -70, -358)
    Zombie(8).SetPosition iR.CreateVec3(1000, -70, -24)
    Zombie(9).SetPosition iR.CreateVec3(1019, -70, 357)
'======================================================================
    iFont.CreateFont "gfx/cut.bmp", RGB(0, 0, 255), 44, 64, 0, 1
'======================================================================
    iSound.LoadSound3D "sound/drown1.wav" '0
    iSound.LoadSound3D "sound/pain4.wav" '1
    iSound.LoadSound3D "sound/plyrjmp8.wav" '2
    iSound.LoadSound3D "sound/hum1.wav" '3
    iSound.LoadSound3D "sound/pump1.wav" '4
    iSound.LoadSound3D "sound/pump2.wav" '5
    iSound.LoadSound3D "sound/buzzer_mmanson.wav" '6
    iSound.LoadSound3D "sound/fall1.wav" '7
    
'======================================================================
    Camera.SetPosition iR.CreateVec3(-1720, 0, 362)
    Camera.SetRotation iR.CreateVec3(0, -Pi * 0.5, 0)
    CamRot = Camera.GetRotation
    Bullet = 5
    Health = 100
    SLBox = -1
    Weapon = 0
    SmokeSprColor = 155
'======================================================================
    
    Do
        DoEvents
        
        iSound.PlaySound3D 3, False
        iSound.PlaySound3D 6, False
        
        TickNow = iR.GetTickPassed
        Movement = 0.1 * (TickNow - OldTick)
        OldTick = TickNow
        
        If Control.CheckKBKeyPressed(iR_Key_Escape) Then
        iSound.StopSound3D 7
        iSound.StopSound3D 6
        iSound.StopSound3D 3
        Set iR = Nothing
        End
        End If
        iR.Clear
        iR.BeginScene
        
        If Health < 1 Then HeroDeath = True
             
        If HeroDeath = False Then
           UpdateParameter
        Else
          If HrSoundDt = False Then HrSoundDt = True: iSound.PlaySound3D 0, False
          For i = 0 To AmountZombie
           Zombie(i).SetAmimationSpeed 0.0005
          Next i
           CamPos = Camera.GetPosition
           If CamPos.y > Level.GetHeight(CamPos) + 10 Then CamPos.y = CamPos.y - 0.2 * Movement
           Camera.SetPosition CamPos
           CamRot = Camera.GetRotation
           If CamRot.X < Pi * 0.5 Then
              CamRot.X = CamRot.X + 0.005 * Movement
              Camera.SetRotation CamRot
           Else
            For i = 0 To AmountZombie
              Zombie(i).SetAmimationSpeed 0
            Next i
           End If
        End If
        
        Camera.Update
         
        Level.Render
        
        
        
        For i = 0 To AmountZombie
        If Maths.GetDistance(CamPos, ZombiePos(i)) < 1000 And ZombieEnbl(i) Then
           Light.Light_EnableLighting 1
           Zombie(i).Render
           End If
        Next i
       
        
        For i = 0 To 3
           If BoxShellEnbl(i) Then BoxShell(i).Render
           If AidEnbl(i) Then Aid(i).Render
        Next i

       
        If HeroDeath = False Then
           If Weapon = 0 Then Hero(0).Render
           If Weapon = 1 Then Hero(1).Render
           If SmokeSprEnbl Then
              If SmokeSprColor > 10 Then
                 SmokeSprColor = SmokeSprColor - 3 * Movement
                 SmokeSpr.Height = SmokeSpr.Height + 10 * Movement
                 SmokeSpr.Width = SmokeSpr.Width + 10 * Movement
                 SmokeSpr.Pos.y = (CInt(iR.GetScreenResolution.y) - SmokeSpr.Height) * 0.5
                 SmokeSpr.Pos.X = (CInt(iR.GetScreenResolution.X) - SmokeSpr.Width) * 0.5
                 If SmokeSprColor > 0 Then SmokeSpr.Color = RGB(SmokeSprColor, SmokeSprColor, SmokeSprColor)
                 Inface.DrawSprite SmokeSpr
              Else
                 SmokeSprColor = 155
                 SmokeSprEnbl = False
              End If
           Else
             SmokeSpr.Height = 64
             SmokeSpr.Width = 64
           End If
           Inface.DrawSprite Prizel
           Inface.DrawSprite Cross
           If Weapon = 1 Then
           Inface.DrawSprite ShellBox
           iFont.DrawText Str(SLBox), Int(ShellBox.Pos.X + 8), Int(ShellBox.Pos.y - 10), 28
           For i = 0 To Bullet - 1
               Shell.Pos.X = i * 32
               Inface.DrawSprite Shell
           Next i
           End If
           iFont.DrawText Str(Health), Int(Cross.Pos.X + 8), Int(Cross.Pos.y - 10), 28
        End If
        
        
        
        'Inface.DrawText Str(ZombieRot.y + CamRot.y), 20, 210
        'Inface.DrawText Str(NewCamPos.y), 20, 230
        'Inface.DrawText Str(CamRot.y), 20, 250
        'Inface.DrawText Str(CamPos.X) & "   " & Str(CamPos.z), 20, 270
        'Inface.DrawText Str(SmokePart.Position.X) & "   " & Str(SmokePart.Position.z), 20, 290
        
       
      
        
        
       For i = 0 To AmountZombie
        If ZombieAttack(i) = True And ZombieDeath(i) = False Then
            If BloodColor > 5 And HeroDeath = False Then
               BloodColor = BloodColor - 3 * Movement
            Else
               BloodColor = 155
               If HeroDeath = False Then iSound.PlaySound3D 1, False
               If Health > 0 Then Health = Health - 20
            End If
            If BloodColor > 0 Then Inface.DrawBox2D iR.CreateVec2(0, 0), _
                           iR.CreateVec2(iR.GetScreenResolution.X, iR.GetScreenResolution.y), _
                           RGB(0, 0, BloodColor), True
            Exit For
        End If
      Next i
      
      
      
        If HeroDeath = True Then
        Inface.DrawBox2D iR.CreateVec2(0, 0), _
                         iR.CreateVec2(iR.GetScreenResolution.X, iR.GetScreenResolution.y * 0.15), _
                         RGB(0, 0, 0) ', True
        Inface.DrawBox2D iR.CreateVec2(0, iR.GetScreenResolution.y - iR.GetScreenResolution.y * 0.15), _
                         iR.CreateVec2(iR.GetScreenResolution.X, iR.GetScreenResolution.y), _
                         RGB(0, 0, 0) ', True
        End If
        
        
        
        iR.EndScene
        iR.Present
        
    Loop
    
End Sub
Public Sub CheckInput()
Control.Mouse_RotateCamera Camera
 
    If Control.CheckKBKeyPressed(iR_Key_Left) Then Camera.MoveSide Movement
    If Control.CheckKBKeyPressed(iR_Key_Right) Then Camera.MoveSide -Movement
    If Control.CheckKBKeyPressed(iR_Key_Up) Then Camera.MoveForward Movement
    If Control.CheckKBKeyPressed(iR_Key_Down) Then Camera.MoveForward -Movement
End Sub


Private Sub UpdateParameter()
'=====CAMERA===========================================================================
    CamPos = Camera.GetPosition
    CheckInput
    NewCamPos = Camera.GetPosition
    NewCamPos.y = CamPos.y
    
    
    eDir = Maths.Vec3Subtract(NewCamPos, CamPos)
    Level.SetPlayerBoundingBox iR.CreateVec3(-15, -30, -15), iR.CreateVec3(15, 10, 15)
    NewCamPos = Level.BoundingBoxCollision(CamPos, eDir)
    NewCamPos.y = Level.GetHeight(NewCamPos) + 60
    Camera.SetPosition NewCamPos
    
    For i = 0 To AmountZombie
       If ZombieDeath(i) = False Then
          ZombiePos(i) = Zombie(i).GetPosition
          If Maths.GetDistance(NewCamPos, ZombiePos(i)) < 50 Then
             Camera.SetPosition CamPos
          End If
       End If
    Next i
    
    

'=====HERO=============================================================================
    CamPos = Camera.GetPosition
    Hero(1).SetPosition CamPos
    Hero(0).SetPosition CamPos
    CamRot = Camera.GetRotation
    Camera.SetRotation CamRot
    HeroRot.y = CamRot.y
    HeroRot.z = CamRot.X
    Hero(1).SetRotation HeroRot
    Hero(0).SetRotation HeroRot
    
    If Control.CheckKBKeyPressed(iR_Key_1) And Weapon = 1 Then
       Weapon = 0
       Hero(0).SetAnimation 3
    End If
    If Hero(0).GetAnimation = 3 And Hero(0).GetFrame > 35 Then Hero(0).SetAnimation 0
    
    If Control.CheckKBKeyPressed(iR_Key_2) And Weapon = 0 Then
       Weapon = 1
       Hero(1).SetAnimation 6
    End If
    If Hero(1).GetAnimation = 6 And Hero(1).GetFrame > 29 Then Hero(1).SetAnimation 0
    
If Weapon = 0 Then
    If Control.Mouse_CheckButton(iR_MouseButton1) And Hero(0).GetAnimation = 0 Then
       Hero(0).SetAnimation 1
       iSound.StopSound3D 2
       iSound.PlaySound3D 2, False
     For i = 0 To AmountZombie
      If Maths.GetDistance(CamPos, ZombiePos(i)) <= 55 Then
        If Zombie(i).MousePick(iR.GetScreenResolution.X * 0.5, iR.GetScreenResolution.y * 0.5) And ZombieDeath(i) = False Then
           ZombieDeath(i) = True
           iSound.StopSound3D 7
           iSound.PlaySound3D 7, False
           Zombie(i).SetCurrentFrame 39
           Zombie(i).SetStartFrame 39
           Zombie(i).SetEndFrame 58
         Exit For
        End If
      End If
     Next i
    End If
    If Hero(0).GetAnimation = 1 And Hero(0).GetFrame > 23 Then Hero(0).SetAnimation 0
End If
    
If Weapon = 1 Then
    If Control.Mouse_CheckButton(iR_MouseButton1) And Hero(1).GetAnimation = 0 Then
      If Bullet > 0 Then
          SmokeSprEnbl = False
          SmokeSprEnbl = True
          Bullet = Bullet - 1
          iSound.StopSound3D 4
          iSound.PlaySound3D 4, False
          Hero(1).SetAnimation 1
        For i = 0 To AmountZombie
         If Zombie(i).MousePick(iR.GetScreenResolution.X * 0.5, iR.GetScreenResolution.y * 0.5) And ZombieDeath(i) = False Then
            ZombieDeath(i) = True
            iSound.StopSound3D 7
            iSound.PlaySound3D 7, False
            Zombie(i).SetCurrentFrame 39
            Zombie(i).SetStartFrame 39
            Zombie(i).SetEndFrame 58
          Exit For
         End If
        Next i
       End If
       If Bullet = 0 And SLBox >= 0 Then SLBox = SLBox - 1
    End If
    
   If SLBox >= 0 Then
      If Bullet < 5 Then
         If Hero(1).GetAnimation = 3 And Hero(1).GetFrame > 26 Then Possible = True
         If Possible = True And Hero(1).GetFrame < 26 Then
            Bullet = Bullet + 1
            Possible = False
         End If
      Else
         If Hero(1).GetAnimation = 3 And Hero(1).GetFrame < 26 Then Hero(1).SetAnimation 4
      End If
      If Bullet = 0 And Hero(1).GetFrame > 36 Then Hero(1).SetAnimation 3
      If Hero(1).GetAnimation = 1 And Hero(1).GetFrame > 36 Then Hero(1).SetAnimation 0
      If Hero(1).GetAnimation = 4 And Hero(1).GetFrame > 26 Then Hero(1).SetAnimation 0
      If Hero(1).GetAnimation = 1 And Hero(1).GetFrame > 20 And Hero(1).GetFrame < 21 Then
         iSound.StopSound3D 5
         iSound.PlaySound3D 5, False
      End If
      If Hero(1).GetAnimation = 4 And Hero(1).GetFrame > 12 And Hero(1).GetFrame < 13 Then
         iSound.StopSound3D 5
         iSound.PlaySound3D 5, False
      End If
      If Hero(1).GetAnimation = 3 And Hero(1).GetFrame > 12 And Hero(1).GetFrame < 13 Then
         iSound.StopSound3D 5
         iSound.PlaySound3D 5, False
      End If
      
    Else
      Hero(1).SetAnimation 3
      Hero(0).SetAnimation 3
      Weapon = 0
    End If
 End If
'=====Zombie=============================================================================
   For i = 0 To AmountZombie
   
    If ZombieDeath(i) = False And ZombieEnbl(i) Then
    
       ZombiePos(i) = Zombie(i).GetPosition
       ZombieRot(i) = Zombie(i).GetRotation
       
      Vec = Maths.Vec3Subtract(CamPos, ZombiePos(i))
      If Vec.X > 0 Then _
         ZombieRot(i).y = Atn(Vec.z / Vec.X)
      If Vec.X < 0 Then _
         ZombieRot(i).y = Atn(Vec.z / Vec.X) + Pi
       
      NewZombPos(i) = ZombiePos(i)
    
      If Maths.GetDistance(CamPos, ZombiePos(i)) > 50 Then
         Zombie(i).SetStartFrame 0
         Zombie(i).SetEndFrame 30
        If Maths.GetDistance(CamPos, ZombiePos(i)) < 900 Then
           NewZombPos(i).X = NewZombPos(i).X + Cos(ZombieRot(i).y) * 0.5 * Movement
           NewZombPos(i).z = NewZombPos(i).z + Sin(ZombieRot(i).y) * 0.5 * Movement
        End If
      Else
         If Zombie(i).GetCurrentFrame < 31 Then Zombie(i).SetCurrentFrame 31
         Zombie(i).SetStartFrame 31
         Zombie(i).SetEndFrame 39
      End If
      If Zombie(i).GetCurrentFrame >= 31 And Zombie(i).GetCurrentFrame <= 39 Then ZombieAttack(i) = True Else ZombieAttack(i) = False
    End If
    
       zDir(i) = Maths.Vec3Subtract(NewZombPos(i), ZombiePos(i))
       Level.SetPlayerBoundingBox iR.CreateVec3(-30, -15, -25), iR.CreateVec3(30, 15, 25)
       NewZombPos(i) = Level.BoundingBoxCollision(ZombiePos(i), zDir(i))
       Zombie(i).SetRotation ZombieRot(i)
       Zombie(i).SetPosition NewZombPos(i)
    
    For ii = 0 To AmountZombie
        If i <> ii Then
           If Maths.GetDistance(NewZombPos(ii), ZombiePos(i)) < 40 Then
            If Maths.GetDistance(CamPos, ZombiePos(ii)) < Maths.GetDistance(CamPos, NewZombPos(i)) Then
              If ZombieDeath(ii) = False Then Zombie(i).SetPosition ZombiePos(i)
            End If
           End If
        End If
    Next ii
    
    If CInt(Zombie(i).GetCurrentFrame) >= 57 And ZombieEnbl(i) And ZombieDeath(i) Then
         Zombie(i).SetStartFrame 58
         Zombie(i).SetEndFrame 58
         Zombie(i).SetAmimationSpeed 0
         ZombieSpr(i).Transparensy = 1
         Zombie(i).SetTexture ZombieSpr(i)
         If ColorDiffuse(i).Blue > 0 Then
            ColorDiffuse(i).Blue = ColorDiffuse(i).Blue - 0.003 * Movement
            ColorDiffuse(i).Red = ColorDiffuse(i).Blue
            ColorDiffuse(i).Green = ColorDiffuse(i).Blue
            ZombieMat(i).Diffuse = ColorDiffuse(i)
            ZombieMat(i).Emissive = ColorDiffuse(i)
            Zombie(i).SetMaterial ZombieMat(i)
         Else
           ZombieEnbl(i) = False
           Exit For
         End If
      End If
     
  Next i
'=====ShellBox=============================================================================
    For i = 0 To 3
    BoxShellPos(i) = BoxShell(i).GetPosition
    If Maths.GetDistance(CamPos, BoxShellPos(i)) <= 60 Then
       BoxShellEnbl(i) = False
       BoxShellPos(i).y = -200
       BoxShell(i).SetPosition BoxShellPos(i)
       SLBox = SLBox + 1
       Weapon = 1
       Hero(1).SetAnimation 3
    End If
    Next i
'=====Aid=============================================================================
          If Health < 100 Then
        For i = 0 To 3
         AidPos(i) = Aid(i).GetPosition
         If Maths.GetDistance(CamPos, AidPos(i)) <= 60 Then
            AidEnbl(i) = False
            AidPos(i).y = -200
            Aid(i).SetPosition AidPos(i)
            Health = Health + 40
            If Health > 100 Then Health = 100
         End If
        Next i
       End If
    
End Sub

