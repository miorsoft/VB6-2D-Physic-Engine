VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Physic Engine"
   ClientHeight    =   7965
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   531
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command5 
      Caption         =   "ADD Mini Chain"
      Height          =   615
      Left            =   13440
      TabIndex        =   8
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ADD Regular Poly"
      Height          =   615
      Left            =   13440
      TabIndex        =   7
      Top             =   4200
      Width           =   975
   End
   Begin VB.CheckBox chkJPG 
      Caption         =   "Save Jpg Frames"
      Height          =   495
      Left            =   13560
      TabIndex        =   6
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ADD BOX"
      Height          =   615
      Left            =   13440
      TabIndex        =   4
      Top             =   3360
      Width           =   975
   End
   Begin VB.ComboBox cmbScene 
      Height          =   315
      Left            =   13440
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ADD BALL"
      Height          =   615
      Left            =   13440
      TabIndex        =   2
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "(RE) START"
      Height          =   615
      Left            =   13440
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   6135
      Left            =   120
      ScaleHeight     =   409
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   641
      TabIndex        =   0
      Top             =   120
      Width           =   9615
   End
   Begin VB.Label Label1 
      Caption         =   "SCENE"
      Height          =   255
      Left            =   13440
      TabIndex        =   5
      Top             =   1320
      Width           =   1095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public WithEvents ENGINE As cls2DPhysic
Attribute ENGINE.VB_VarHelpID = -1



Private MouseDownX As Double
Private MouseDownY As Double



Private Sub chkJPG_Click()
    SaveFrames = (chkJPG.Value = vbChecked)
End Sub

Private Sub cmbScene_Change()
    CreateScene frmMain.cmbScene.ListIndex
End Sub

Private Sub cmbScene_Click()
    CreateScene frmMain.cmbScene.ListIndex
End Sub

Private Sub Command1_Click()

    ENGINE.BiggerGroup = 0
    CreateScene frmMain.cmbScene.ListIndex


End Sub

Private Sub Command2_Click()
    ENGINE.BodyCREATECircle Vec2(PicW * 0.5, 0), 5 + Rnd * 20, DefDensity

    ENGINE.BodySetGroup ENGINE.NofBodies, 1
    ENGINE.BodySetCollideWith ENGINE.NofBodies, ALL

End Sub

Private Sub Command3_Click()


'    BodyCREATERandomPoly Vec2(PicW \ 2, 0), DefDensity
    ENGINE.BodyCREATEBox Vec2(PicW \ 2, 0), 60, 30
    ENGINE.BodySetGroup ENGINE.NofBodies, 1
    ENGINE.BodySetCollideWith ENGINE.NofBodies, ALL
End Sub

Private Sub Command4_Click()
    ENGINE.BodyCREATERegularPoly Vec2(PicW \ 2, 0), 12 + Rnd * 30, 12 + Rnd * 30, 3 + Int(Rnd * 10), -Int(Rnd * 2), DefDensity
    ENGINE.BodySetGroup ENGINE.NofBodies, 1
    ENGINE.BodySetCollideWith ENGINE.NofBodies, ALL
End Sub

Private Sub Command5_Click()
    ENGINE.BodyCREATEBox Vec2(PicW * 0.5, 5), 80, 15
    ENGINE.BodyCREATEBox Vec2(PicW * 0.5 + 70, 5), 80, 15
    ENGINE.JointAdd2PinsJ ENGINE.NofBodies - 1, Vec2(35, 0), ENGINE.NofBodies, Vec2(-35, 0), 0, 1, 1


    'Make last 2 bodies collide with All but each other
    ENGINE.BodySetGroup ENGINE.NofBodies - 1, ENGINE.BiggerGroup * 2
    ENGINE.BodySetGroup ENGINE.NofBodies, ENGINE.BiggerGroup * 2
    ENGINE.BodySetCollideWith ENGINE.NofBodies - 1, ALL - ENGINE.BiggerGroup
    ENGINE.BodySetCollideWith ENGINE.NofBodies, ALL - ENGINE.BiggerGroup \ 2



End Sub


Private Sub ENGINE_CollisionEvent(bA As Long, bB As Long, posAX As Double, PosAY As Double, posBX As Double, PosBY As Double, Nx As Double, Ny As Double, ContactVelo As Double)
'Me.Caption = bA & " " & bB & "   " & ContactVelo

End Sub

Private Sub Form_Activate()
    MAINLOOP
End Sub

Private Sub Form_Load()
    Randomize Timer

    If Dir(App.Path & "\Frames", vbDirectory) = vbNullString Then MkDir App.Path & "\Frames"
    If Dir(App.Path & "\Frames\*.*") <> vbNullString Then Kill App.Path & "\Frames\*.*"

    PIC.Height = 360    '360
    PIC.Width = Int(PIC.Height * 16 / 9)



    pHDC = PIC.hDC
    PicW = PIC.Width
    PicH = PIC.Height




    Set ENGINE = New cls2DPhysic


    ENGINE.EngineINIT PIC



    cmbScene.AddItem "First"
    cmbScene.AddItem "Distance Joints"
    cmbScene.AddItem "1 Pin"
    cmbScene.AddItem "2 Pins Joints"
    cmbScene.AddItem "2 Pins Joints II"
    cmbScene.AddItem "Slope"
    cmbScene.AddItem "Gum Bridge"
    cmbScene.AddItem "Car(Rotor)"
    cmbScene.AddItem "Newton Cardle"
    cmbScene.AddItem "(Rotor2)"
    cmbScene.AddItem "CAR2"
    
    cmbScene.ListIndex = 1

    ENGINE.RenderCreateIntroFrames



    'ENGINE.RenderINITRC

    CreateScene cmbScene.ListIndex




End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ENGINE.RenderCreateOuttroFrames

End Sub

Private Sub Form_Unload(Cancel As Integer)

    ENGINE.UnLoad


    End

End Sub

Private Sub PIC_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim EO  As Long
    Dim rX  As Double
    Dim rY  As Double


    MouseDownX = x
    MouseDownY = y


    ENGINE.BodyGetNearest x * 1, y * 1, EO, rX, rY
    ENGINE.MouseSelectedObj = EO
    ENGINE.MouseDownRelX = rX
    ENGINE.MouseDownRelY = rY



End Sub

Private Sub PIC_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ENGINE.MouseMoveX = x
    ENGINE.MouseMoveY = y

End Sub

Private Sub PIC_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Dx  As Double
    Dim DY  As Double
    Dx = x - ENGINE.MouseDownRelX
    DY = y - ENGINE.MouseDownRelY
    

    'ENGINE.BodyApplyImpulse SelectedObj, _
     Vec2MUL(Vec2(Dx, Dy), 1), _
     Vec2ADD(ENGINE.BodyGetPOS(SelectedObj), Vec2(orX, orY))


    ENGINE.BodyApplyImpulse ENGINE.MouseSelectedObj, _
                            Vec2MUL(Vec2(Dx, DY), ENGINE.BodyGetMass(ENGINE.MouseSelectedObj) * 0.0085), _
                            Vec2(ENGINE.MouseDownRelX, ENGINE.MouseDownRelY)




    ENGINE.MouseSelectedObj = 0


End Sub
