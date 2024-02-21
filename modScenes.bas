Attribute VB_Name = "modScenes"
Option Explicit

'feel free to create your own scene and share it at VBForums.com (VB6 Phyisc Engine)

Public Sub CreateScene(Scene As Long)

    Dim I      As Long
    Dim Vertices() As tVec2


    With frmMain.ENGINE

        .NofBodies = 0
        .NJ = 0

        Select Case Scene
        Case 0
            For I = 1 To 20
                .BodyCREATECircle Vec2(I * 55, 50), 5 + Rnd * (20)
            Next

            ' .JointAddDistanceJ 2, 3, 50
            ' .JointAddDistanceJ 4, 5, 50
            ' .JointAddDistanceJ 6, 7, 50
            ' .JointAddDistanceJ 8, 9, 50
            ' .JointAddDistanceJ 10, 11, 50

            .BodyCREATERandomPoly Vec2(300, 150)
            .BodyCREATERandomPoly Vec2(350, 150)

            .JointAdd2PinsJ .NofBodies - 1, Vec2(30, 0), .NofBodies, Vec2(-30, 0), 80, 0.5


            For I = 20 + 1 To 20 + 9

                .BodyCREATECircle Vec2((I - 20 - 1) * 75, PicH + 40), 65

                .BodySetStatic .NofBodies
            Next

            .JointAddDistanceJ 20 + 6, 5, 200



            '-----------ROPE
            .BodyCREATECircle Vec2(100, 50), 10
            .BodySetStatic .NofBodies
            .BodyCREATECircle Vec2(100, 100), 10
            .BodyCREATECircle Vec2(100, 150), 10
            .BodyCREATECircle Vec2(100, 200), 10
            .BodyCREATECircle Vec2(100, 250), 10
            .JointAddDistanceJ .NofBodies, .NofBodies - 1, 50, 1, 0
            .JointAddDistanceJ .NofBodies - 1, .NofBodies - 2, 50, 1, 0
            .JointAddDistanceJ .NofBodies - 2, .NofBodies - 3, 50, 1, 0
            .JointAddDistanceJ .NofBodies - 3, .NofBodies - 4, 50, 1, 0


            .BodyCREATERandomPoly Vec2(500, 150)

            .JointAdd1PinJ .NofBodies, Vec2(30, 0), 50, 0.1, 0.1
            For I = 1 To .NofBodies
                .BodySetGroup I, 1
                .BodySetCollideWith I, ALL
            Next

        Case 1

            .BodyCREATEBox Vec2(PicW * 0.5, PicH - 15), PicW * 0.9, 25
            .BodySetStatic 1

            '-----------ROPE
            .BodyCREATECircle Vec2(100, 50), 7
            .BodySetStatic .NofBodies
            .BodyCREATECircle Vec2(100, 100), 7
            .BodyCREATECircle Vec2(100, 150), 7
            .BodyCREATECircle Vec2(100, 200), 7
            .BodyCREATECircle Vec2(100, 250), 7
            .JointAddDistanceJ .NofBodies, .NofBodies - 1, 50, 1, 0
            .JointAddDistanceJ .NofBodies - 1, .NofBodies - 2, 50, 1, 0
            .JointAddDistanceJ .NofBodies - 2, .NofBodies - 3, 50, 1, 0
            .JointAddDistanceJ .NofBodies - 3, .NofBodies - 4, 50, 1, 0

            .BodyCREATECircle Vec2(PicW * 0.75, PicH * 0.1), 7
            .BodySetStatic .NofBodies
            .BodyCREATEBox Vec2(PicW * 0.75, PicH * 0.2), 50, 20
            .JointAddDistanceJ .NofBodies, .NofBodies - 1, 100, 1, 0
            For I = 1 To .NofBodies
                .BodySetGroup I, 1
                .BodySetCollideWith I, ALL
            Next
        Case 2    '1PIN '


            .BodyCREATEBox Vec2(PicW * 0.5, PicH - 15), PicW * 0.9, 25
            .BodySetStatic 1

            .BodyCREATEBox Vec2(PicW * 0.5, PicH * 0.5), 100, 20
            .JointAdd1PinJ .NofBodies, Vec2(0, 0), 0

            .BodyCREATEBox Vec2(PicW * 0.25, PicH * 0.5), 100, 20
            .JointAdd1PinJ .NofBodies, Vec2(-25, 0), 0

            .BodyCREATEBox Vec2(PicW * 0.75, PicH * 0.4), 100, 20
            .JointAdd1PinJ .NofBodies, Vec2(-25, 0), 50


            .BodyCREATEBox Vec2(PicW * 0.9, PicH * 0.1), 100, 20
            .JointAdd1PinJ .NofBodies, Vec2(-25, 0), 50, 0.005, 0.005
            For I = 1 To .NofBodies
                .BodySetGroup I, 1
                .BodySetCollideWith I, ALL
            Next

        Case 3    '"2 Pins Joints"

            'Floor
            .BodyCREATEBox Vec2(PicW * 0.5, PicH - 15), PicW * 0.9, 25
            .BodySetStatic 1

            .BodyCREATEBox Vec2(PicW * 0.1 + 20, PicH * 0.5), 50, 20
            .JointAdd1PinJ .NofBodies, Vec2(-20, 0), 40, 0.01, 0

            For I = 1 To 5
                .BodyCREATEBox Vec2(PicW * 0.1 + 20 + 70 * I, PicH * 0.5), 50, 20
                ' .JointAdd1PinJ  .NofBodies, Vec2(-20, 0), 40, 0.01, 0
                .JointAdd2PinsJ .NofBodies - 1, Vec2(20, 0), _
                                .NofBodies, Vec2(-20, 0), 30, 1, 0
            Next

            For I = 1 To .NofBodies
                .BodySetGroup I, 1
                .BodySetCollideWith I, ALL
            Next



        Case 4    '"2 Pins Joints II

            'Floor
            .BodyCREATEBox Vec2(PicW * 0.5, PicH - 15), PicW * 0.9, 25
            .BodySetStatic 1

            ' .BodyCREATEBox 50, 20, Vec2(PicW * 0.1 + 20, PicH * 0.4)
            ' .JointAdd1PinJ  .NofBodies, Vec2(-20, 0), 40, 0.5, 0
            .BodyCREATEBox Vec2(PicW * 0.05 + 70 * 0, PicH * 0.4), 50, 20
            .BodySetStatic .NofBodies

            For I = 1 To 8
                .BodyCREATEBox Vec2(PicW * 0.05 + 70 * I, PicH * 0.4), 50, 20
                .JointAdd2PinsJ .NofBodies - 1, Vec2(20, 0), _
                                .NofBodies, Vec2(-20, 0), 30, 0.125, 0
            Next

            .BodySetStatic .NofBodies
            For I = 1 To .NofBodies
                .BodySetGroup I, 1
                .BodySetCollideWith I, ALL
            Next
        Case 5    'Slope

            .BodyCREATEBox Vec2(PicW * 0.2, PicH * 0.45), PicW * 0.5, 25, PI * 0.25
            .BodySetStatic 1
            .BodyCREATEBox Vec2(PicW * 0.8, PicH * 0.45), PicW * 0.5, 25, PI * 0.75
            .BodySetStatic .NofBodies

            .BodyCREATEBox Vec2(PicW * 0.5, PicH * 0.75), 58, 22
            .JointAdd1PinJ .NofBodies, Vec2(-25, 0), 0, 0.006, 0.006
            .JointAdd1PinJ .NofBodies, Vec2(25, 0), 0, 0.006, 0.006

            For I = 1 To .NofBodies
                .BodySetGroup I, 1
                .BodySetCollideWith I, ALL
            Next
        Case 6    'Gum Bridge

            For I = 1 To 8
                .BodyCREATEBox Vec2((I - 0.5) * 82, PicH * 0.7), 58, 22
                .JointAdd1PinJ .NofBodies, Vec2(-25, 0), 0, 0.006, 0.006
                .JointAdd1PinJ .NofBodies, Vec2(25, 0), 0, 0.006, 0.006
            Next
            For I = 1 To .NofBodies
                .BodySetGroup I, 1
                .BodySetCollideWith I, ALL
            Next
        Case 7    '''' CAR

            .BodyCREATEBox Vec2(PicW * 0.5, PicH - 15), PicW * 1, 25
            .BodySetStatic 1

            For I = 1 To .NofBodies
                .BodySetGroup I, 1
                .BodySetCollideWith I, ALL
            Next

            .BodySetGroup .NofBodies - 1, .BiggerGroup * 2      '=2
            .BodySetGroup .NofBodies, .BiggerGroup * 2      '=4
            .BodySetCollideWith .NofBodies - 1, ALL - .BiggerGroup      '=4
            .BodySetCollideWith .NofBodies, ALL - .BiggerGroup \ 2      '=2

            .BodyCREATEBox Vec2(100, 200), 150, 20 '40
            .BodyCREATECircle Vec2(100 - 50, 230), 20   'WHEEL
            .BodySetGroup .NofBodies - 1, .BiggerGroup * 2
            .BodySetGroup .NofBodies, .BiggerGroup * 2
            .BodySetCollideWith .NofBodies - 1, ALL - .BiggerGroup
            .BodySetCollideWith .NofBodies, ALL - .BiggerGroup \ 2
            .BodyCREATECircle Vec2(100 + 50, 230), 20    'WHEEL
            .BodySetGroup .NofBodies, .BiggerGroup
            .BodySetCollideWith .NofBodies, ALL - .BiggerGroup \ 4


            .JointAdd2PinsJ .NofBodies - 1, Vec2(0, 0), .NofBodies - 2, Vec2(0, 0), Sqr(50 * 50 + 30 * 30)
            .JointAdd2PinsJ .NofBodies, Vec2(0, 0), .NofBodies - 2, Vec2(0, 0), Sqr(50 * 50 + 30 * 30)

            .JointAdd2PinsJ .NofBodies - 1, Vec2(0, 0), .NofBodies - 2, Vec2(-50, 0), 30, 0.02, 0.02
            .JointAdd2PinsJ .NofBodies, Vec2(0, 0), .NofBodies - 2, Vec2(50, 0), 30, 0.02, 0.02



            .JoinAddRotorJ .NofBodies - 1, Vec2(20, 0), 0.07


        Case 8    'newton cardle


            'Floor
            .BodyCREATEBox Vec2(PicW * 0.5, PicH - 15), PicW * 0.9, 25
            .BodySetStatic 1


            For I = 1 To 5

                .BodyCREATECircle Vec2(200 + I * 50, 50), 25
                .JointAdd1PinJ .NofBodies, Vec2(0, 0), 140, 0.1, 0.1

            Next


            For I = 1 To .NofBodies
                .BodySetGroup I, 1
                .BodySetCollideWith I, ALL
            Next

        Case 9    '''' Rotor2

            .BodyCREATEBox Vec2(PicW * 0.5, PicH - 15), PicW * 1, 25
            .BodySetStatic 1

            .BodyCREATEBox Vec2(100, 200), 80, 33
            .BodyCREATEBox Vec2(140, 200), 80, 33
            .JoinAddRotor2J .NofBodies - 1, Vec2(35, 0), .NofBodies, Vec2(-35, 0), 0.1, 0.1

            .BodyCREATEBox Vec2(400, 200), 85, 22
            .BodyCREATEBox Vec2(440, 200), 85, 22
            .JoinAddRotor2J .NofBodies - 1, Vec2(12, 0), .NofBodies, Vec2(-12, 0), 0.1, 0.1



            For I = 1 To .NofBodies
                .BodySetGroup I, 1
                .BodySetCollideWith I, ALL
            Next

            .BodySetGroup .NofBodies - 1, .BiggerGroup * 2      '=2
            .BodySetGroup .NofBodies, .BiggerGroup * 2      '=4
            .BodySetCollideWith .NofBodies - 1, ALL - .BiggerGroup      '=4
            .BodySetCollideWith .NofBodies, ALL - .BiggerGroup \ 2      '=2


            .BodySetGroup .NofBodies - 3, .BiggerGroup * 2      '=2
            .BodySetGroup .NofBodies - 2, .BiggerGroup * 2    '=4
            .BodySetCollideWith .NofBodies - 3, ALL - .BiggerGroup      '=4
            .BodySetCollideWith .NofBodies - 2, ALL - .BiggerGroup \ 2    '=2



        Case 10
            ' CAR 2   Vertices test
            Dim CarL As Double
            Dim CarH As Double
            Dim WR As Double
            Dim WDX As Double
            Dim WDY As Double


            CarL = 111 '80
            CarH = 15 '25 '30
            WDX = 30
            WDY = 20
            WR = 15

            .BodyCREATEBox Vec2(PicW * 0.5, PicH - 15), PicW * 1, 25
            .BodySetStatic 1


            For I = 1 To .NofBodies
                .BodySetGroup I, 1
                .BodySetCollideWith I, ALL
            Next

            .BodySetGroup .NofBodies - 1, .BiggerGroup * 2      '=2
            .BodySetGroup .NofBodies, .BiggerGroup * 2      '=4
            .BodySetCollideWith .NofBodies - 1, ALL - .BiggerGroup      '=4
            .BodySetCollideWith .NofBodies, ALL - .BiggerGroup \ 2      '=2


            .BodyCREATEBox Vec2(50, 240), CarL, CarH, , True
            

            .BodyCREATECircle Vec2(50 - WDX, 240 + WDY), WR    'WHEEL
            .BodySetFriction .NofBodies, 0.5, 0.4

            .BodySetGroup .NofBodies - 1, .BiggerGroup * 2
            .BodySetGroup .NofBodies, .BiggerGroup * 2
            .BodySetCollideWith .NofBodies - 1, ALL - .BiggerGroup
            .BodySetCollideWith .NofBodies, ALL - .BiggerGroup \ 2

            .BodyCREATECircle Vec2(50 + WDX, 240 + WDY), WR  'WHEEL
            .BodySetFriction .NofBodies, 0.5, 0.4

            .BodySetGroup .NofBodies, .BiggerGroup
            .BodySetCollideWith .NofBodies, ALL - .BiggerGroup \ 4

            ''Diagonals
            .JointAdd2PinsJ .NofBodies - 1, Vec2(0, 0), .NofBodies - 2, Vec2(0, 0), Sqr(WDX * WDX + WDY * WDY)
            .JointAdd2PinsJ .NofBodies, Vec2(0, 0), .NofBodies - 2, Vec2(0, 0), Sqr(WDX * WDX + WDY * WDY)

            'Vericals
            .JointAdd2PinsJ .NofBodies - 1, Vec2(0, 0), .NofBodies - 2, Vec2(-WDX, 0), WDY, 0.015, 0.015
            .JointAdd2PinsJ .NofBodies, Vec2(0, 0), .NofBodies - 2, Vec2(WDX, 0), WDY, 0.015, 0.015



            .JoinAddRotorJ .NofBodies - 1, Vec2(WR, 0), 0.07


            ReDim Vertices(3)

            Vertices(1) = Vec2(PicW * 0.25, PicH - 28)
            Vertices(2) = Vec2(PicW * 0.25 + 140, PicH - 28 - 40)
            Vertices(3) = Vec2(PicW * 0.25 + 140, PicH - 28)

            .BodyCREATEPolygon Vertices
           ' .BodySetStatic .NofBodies
            .BodySetGroup .NofBodies, 1
            .BodySetCollideWith .NofBodies, ALL

            For I = 1 To UBound(Vertices)
                Vertices(I) = Vec2ADD(Vertices(I), Vec2(255, 0))
            Next
'            Vertices(4) = Vec2ADD(Vertices(4), Vec2(-50, 0))
            .BodyCREATEPolygon Vertices
         ''  .BodySetStatic .NofBodies
            .BodySetGroup .NofBodies, 1
            .BodySetCollideWith .NofBodies, ALL





        Case 11    '"2 Pins Joints"

            'Floor
            .BodyCREATEBox Vec2(PicW * 0.5, PicH - 15), PicW * 0.9, 25
            .BodySetStatic 1

            .BodyCREATEBox Vec2(PicW * 0.5 + 20, PicH * 0.25), 50, 20
            .JointAdd1PinJ .NofBodies, Vec2(-20, 0), 40, 0.01, 0

            For I = 1 To 2
                .BodyCREATEBox Vec2(PicW * 0.5 + 20 + 70 * I, PicH * 0.25), 50, 20
                ' .JointAdd1PinJ  .NofBodies, Vec2(-20, 0), 40, 0.01, 0
              
            Next

  .JointAdd2PinsJ .NofBodies - 2, Vec2(20, 0), _
                                .NofBodies - 1, Vec2(-20, 0), 30, 1, 0


 .JointAdd2PinsAlignedJ .NofBodies, .NofBodies - 1, 60, Vec2(1, 0), 0.01, 0.01

            For I = 1 To .NofBodies
                .BodySetGroup I, 1
                .BodySetCollideWith I, ALL
            Next


        End Select


    End With

End Sub

