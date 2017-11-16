Attribute VB_Name = "mod2DPhysic"
''# VB6 2D Physic Engine
''
''VB6 port of 2D Impulse Engine
''    by Randy Gaul:
''        http://www.randygaul.net/projects-open-sources/impulse-engine/
''    and Philip Diffenderfer:
''        https://github.com/ClickerMonkey/ImpulseEngine
''
''   + Joints by the Author
''
''
''   Author: Roberto Mior (aka reexre,miorsoft)
''   Contibutors: yet none.
''
''Requires:
''  * vbRichClient (for Render) http://vbrichclient.com/#/en/About/
''
''
''LICENSE: BSD. This allows you to use its source code in any application, commercial or otherwise,
''if you supply proper attribution. Proper attribution includes a notice of copyright and disclaimer
''of warranty.  (https://opensource.org/licenses/BSD-2-Clause)
''
''
''   Copyright © 2017 by Roberto Mior (Aka reexre,miorsoft)
''
''Redistribution and use in source and binary forms, with or without modification, are permitted provided
''that the following conditions are met:
''
''1. Redistributions of source code must retain the above copyright notice, this list of conditions and
''   the following disclaimer.
''2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and
''   the following disclaimer in the documentation and/or other materials provided with the distribution.
''
''THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED
''WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A
''PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
''FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING,
''BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS;
''OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT,
''STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
''SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
''



Option Explicit

'*************************************************************************
'************************* V E C T O R S & MATHS  ************************
'*************************************************************************
Public Type tVec2
    x          As Double
    y          As Double
End Type

Public Type tMAT2
    m00        As Double
    m01        As Double
    m10        As Double
    m11        As Double
End Type





Public Const DT As Double = 1 / 4    '4    ' 1 / 20 '20    '1 / 24  '1/20  '1 / 10   '1 / 60
Public Const Iterations As Long = 1   ' 2    ' 5 ' 20 '5   '10  '2    ' 4
Public Const DefDensity As Double = 1

Public Const PI As Double = 3.14159265358979
Public Const PI2 As Double = 6.28318530717959
Public Const PIh As Double = 1.5707963267949

Public Const EPSILON As Double = 0.00001          '0.0001
Public Const EPSILON_SQ As Double = EPSILON * EPSILON
Public Const BIAS_RELATIVE As Double = 0.98 ' '0.95    '0.9    '0.95
Public Const BIAS_ABSOLUTE As Double = 0.02 '0.01    '0.02    '0.01


Public Const PENETRATION_ALLOWANCE As Double = 0.05    '0.001    '0.01    ' 0.05    '0.1   ' 0.05
Public Const PENETRATION_CORRETION As Double = 0.25    '0.4    '.4 '0.9    '0.4   '0.125   '0.4

Public Const MAX_VALUE As Double = 1E+32


Public Const GlobalSTATICFRICTION As Double = 0.3   '0.5
Public Const GlobalDYNAMICFRICTION As Double = 0.2   '0.3
Public Const GlobalRestitution As Double = 0.8    '0.8


Public GRAVITY As tVec2
Public RESTING As Double


'********** ENGINE ***************

Public pHDC    As Long
Public PicW    As Long
Public PicH    As Long
Public Frame   As Long
Public SaveFrames As Long
Public Const ALL As Long = &HFFFFFFFF

Public INVdt   As Double
Public INVdt2  As Double

Public DisplayRefreshPeriod As Long
Public CNT     As Long
Public pCNT    As Long
Public FPS     As Long
'


Public New_c As cConstructor
Public Cairo As cCairo    '<- global defs of the two Main-"EntryPoints" into the RC5




'Public ENGINE   As cls2DPhysic
'Public WithEvents ENGINE As cls2DPhysic


'************************ TIMING ********************************
''Private Declare Function QueryPerformanceCounter Lib "kernel32" (x As Currency) As Boolean
''Private Declare Function QueryPerformanceFrequency Lib "kernel32" (x As Currency) As Boolean
''Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
''Private m_Time As Double    'DDOUBLE
''Private m_TimeFreq As Double    'DDOUBLE
''Private m_TimeStart As Currency


Public FPSTICK As clsTick

Public tComputed As Long
Public tDraw As Long
Public t1Sec As Long




'*************************************************************************
'************************* V E C T O R S & MATHS  ************************
'*************************************************************************


Public Function Vec2(x As Double, y As Double) As tVec2

    Vec2.x = x
    Vec2.y = y

End Function

Public Function Vec2Negative(V As tVec2) As tVec2
    Vec2Negative.x = -V.x
    Vec2Negative.y = -V.y
End Function



Public Function Vec2ADD(v1 As tVec2, v2 As tVec2) As tVec2
    Vec2ADD.x = v1.x + v2.x
    Vec2ADD.y = v1.y + v2.y
End Function

Public Function Vec2SUB(v1 As tVec2, v2 As tVec2) As tVec2
    Vec2SUB.x = v1.x - v2.x
    Vec2SUB.y = v1.y - v2.y
End Function

Public Function Vec2MULV(v1 As tVec2, v2 As tVec2) As tVec2
    Vec2MULV.x = v1.x * v2.x
    Vec2MULV.y = v1.y * v2.y
End Function
Public Function Vec2MUL(V As tVec2, S As Double) As tVec2
    Vec2MUL.x = V.x * S
    Vec2MUL.y = V.y * S
End Function

Public Function Vec2ADDScaled(v1 As tVec2, v2 As tVec2, S As Double) As tVec2
    Vec2ADDScaled.x = v1.x + v2.x * S
    Vec2ADDScaled.y = v1.y + v2.y * S
End Function

Public Function Vec2LengthSq(V As tVec2) As Double
    Vec2LengthSq = V.x * V.x + V.y * V.y
End Function

Public Function Vec2Length(V As tVec2) As Double
'   Vec2Length = FASTsqr(V.X * V.X + V.Y * V.Y)
    Vec2Length = Sqr(V.x * V.x + V.y * V.y)

End Function


Public Function Vec2Rotate(V As tVec2, radians As Double) As tVec2
'real c = std::cos( radians );
'real s = std::sin( radians );

'real xp = x * c - y * s;
'real yp = x * s + y * c;

    Dim S      As Double
    Dim c      As Double
    c = Cos(radians)
    S = Sin(radians)

    Vec2Rotate.x = V.x * c - V.y * S
    Vec2Rotate.y = V.x * S + V.y * c
End Function

Public Function Vec2Normalize(V As tVec2) As tVec2
    Dim D      As Double
    D = Vec2Length(V)
    If D Then
        D = 1# / D
        Vec2Normalize.x = V.x * D
        Vec2Normalize.y = V.y * D
    End If

End Function

Public Function Vec2MIN(A As tVec2, B As tVec2) As tVec2
    Vec2MIN.x = IIf(A.x < B.x, A.x, B.x)
    Vec2MIN.y = IIf(A.y < B.y, A.y, B.y)
End Function

Public Function Vec2MAX(A As tVec2, B As tVec2) As tVec2
    Vec2MAX.x = IIf(A.x > B.x, A.x, B.x)
    Vec2MAX.y = IIf(A.y > B.y, A.y, B.y)
End Function
'  return a.x * b.x + a.y * b.y;
Public Function Vec2DOT(A As tVec2, B As tVec2) As Double
    Vec2DOT = A.x * B.x + A.y * B.y
End Function
'inline Vec2 Cross( const Vec2& v, real a )
'{
'  return Vec2( a * v.y, -a * v.x );
'}
Public Function Vec2CROSSva(V As tVec2, A As Double) As tVec2
    Vec2CROSSva.x = A * V.y
    Vec2CROSSva.y = -A * V.x
End Function
'inline Vec2 Cross( real a, const Vec2& v )
'{
'  return Vec2( -a * v.y, a * v.x );
'}
Public Function Vec2CROSSav(A As Double, V As tVec2) As tVec2
    Vec2CROSSav.x = -A * V.y
    Vec2CROSSav.y = A * V.x
End Function
'inline real Cross( const Vec2& a, const Vec2& b )
'{
'  return a.x * b.y - a.y * b.x;
'}
Public Function Vec2CROSS(A As tVec2, B As tVec2) As Double
    Vec2CROSS = A.x * B.y - A.y * B.x
End Function


Public Function Vec2DISTANCEsq(A As tVec2, B As tVec2) As Double
    Dim Dx     As Double
    Dim DY     As Double
    Dx = A.x - B.x
    DY = A.y - B.y
    Vec2DISTANCEsq = Dx * Dx + DY * DY
End Function


'************************************************************************************



Public Function matTranspose(M As tMAT2) As tMAT2
    With M
        matTranspose.m00 = .m00
        matTranspose.m01 = .m10    '
        matTranspose.m10 = .m01    '
        matTranspose.m11 = .m11
    End With
End Function

Public Function matMULv(M As tMAT2, V As tVec2) As tVec2

'return Vec2( m00 * rhs.x + m01 * rhs.y, m10 * rhs.x + m11 * rhs.y );
    With M
        matMULv.x = .m00 * V.x + .m01 * V.y
        matMULv.y = .m10 * V.x + .m11 * V.y
    End With

End Function

Public Function SetOrient(radians As Double) As tMAT2
'    real c = std::cos( radians );
'    real s = std::sin( radians );
'
'    m00 = c; m01 = -s;
'    m10 = s; m11 =  c;

    Dim c      As Double
    Dim S      As Double

    c = Cos(radians)
    S = Sin(radians)

With SetOrient
    .m00 = c
    .m01 = -S
    .m10 = S
    .m11 = c
End With

End Function


Public Function VectorProject(ByRef V As tVec2, ByRef Vto As tVec2) As tVec2
'Poject Vector V to vector Vto
    Dim K      As Double
    Dim D      As Double



    D = Vto.x * Vto.x + Vto.y * Vto.y
    If D = 0 Then Exit Function

    D = 1 / Sqr(D)

    K = (V.x * Vto.x + V.y * Vto.y) * D

    VectorProject.x = (Vto.x * D) * K
    VectorProject.y = (Vto.y * D) * K

End Function

Public Function VectorReflect(ByRef V As tVec2, ByRef wall As tVec2) As tVec2
'Function returning the reflection of one vector around another.
'it's used to calculate the rebound of a Vector on another Vector
'Vector "V" represents current velocity of a point.
'Vector "Wall" represent the angle of a wall where the point Bounces.
'Returns the vector velocity that the point takes after the rebound

    Dim vDot   As Double
    Dim D      As Double
    Dim NwX    As Double
    Dim NwY    As Double

    D = (wall.x * wall.x + wall.y * wall.y)
    If D = 0 Then Exit Function

    D = 1 / Sqr(D)

    NwX = wall.x * D
    NwY = wall.y * D
    '    'Vect2 = Vect1 - 2 * WallN * (WallN DOT Vect1)
    'vDot = N.DotV(V)
    vDot = V.x * NwX + V.y * NwY

    NwX = NwX * vDot * 2
    NwY = NwY * vDot * 2

    VectorReflect.x = -V.x + NwX
    VectorReflect.y = -V.y + NwY


End Function


Public Function ACOS(x As Double) As Double

    ACOS = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)

End Function

Public Function AngleDIFF(A1 As Double, A2 As Double) As Double

    AngleDIFF = A1 - A2
    While AngleDIFF < -PI
        AngleDIFF = AngleDIFF + PI2
    Wend
    While AngleDIFF > PI
        AngleDIFF = AngleDIFF - PI2
    Wend

End Function
















Public Sub MAINLOOP()


    Dim I As Long
    Dim A As Long
    Dim B As Long

    '    Dim Accumulator As Long
    '    Dim currTime As Long
    '    Dim frameStart As Long
    '    frameStart = GetTickCount


    Dim pTime As Double   ''DDOUBLE
    Dim pTime2 As Double    ''DDOUBLE



    '''    Timing = 0
    '''    pTime = Timing
    '''    pTime2 = Timing


    Do

        '        'currTime = GetTickCount
        '        'Accumulator = Accumulator + currTime - frameStart
        '        'frameStart = currTime
        '        'If Accumulator > 200 Then Accumulator = 200
        '        'While Accumulator > 10
        '        '    EngineDoSTEP Accumulator * 0.01
        '        '    Accumulator = Accumulator - 10
        '        'Wend
        '
        '
        '        If Timing - pTime2 >= 2 Then
        '            FPS = (CNT - pCNT) * 0.5
        '            pCNT = CNT
        '            pTime2 = Timing
        '            frmMain.Caption = "Physic Engine   computed FPS:" & FPS & " DrawnFPS:" & FPS \ DisplayRefreshPeriod
        '        End If
        '
        '
        '        If ((Timing - pTime) >= 0.001) Then
        '
        '            pTime = Timing


        Select Case FPSTICK.WaitForNext

        Case tComputed
            frmMain.ENGINE.EngineDoSTEP
        Case tDraw

            '   If CNT Mod DisplayRefreshPeriod = 0 Then
            frmMain.ENGINE.RenderDRAWRC

            ''                TotalNContacts = 0
            ''                For I = 1 To NofContactMainFolds
            ''                    TotalNContacts = TotalNContacts + Contacts(I).contactCount
            ''                Next

            If SaveFrames Then
                frmMain.ENGINE.RenderSaveJPG App.Path & "\Frames\" & Format(Frame, "00000") & ".jpg"
                Frame = Frame + 1
            End If


        Case t1Sec
            frmMain.Caption = "Physic Engine   computed FPS:" & FPSTICK.Count(0) & " DrawnFPS:" & FPSTICK.Count(1)
            FPSTICK.ResetCount (0)
            FPSTICK.ResetCount (1)

            
        End Select
'            End If



        CNT = CNT + 1




        '        Else
        '            ' DoEvents
        '
        '
        '        End If



    Loop While True

End Sub



'*************************************************************
'************* IMPULSE MATH **********************************
'*************************************************************

Public Sub InitMATH()

    Set New_c = New cConstructor
    Set Cairo = New_c.Cairo

    

    GRAVITY.x = 0
    GRAVITY.y = 0.008 / DT


    RESTING = Vec2LengthSq(Vec2MUL(GRAVITY, DT)) + EPSILON

    INVdt = 1 / DT
    INVdt2 = 1 / (DT * DT)

    'DisplayRefreshPeriod = 2.5 / DT
    'DisplayRefreshPeriod = 3.5 / DT
    DisplayRefreshPeriod = 6 / DT

End Sub

Public Function Equal(A As Double, B As Double) As Boolean
    If Abs(A - B) <= EPSILON Then Equal = True
End Function

Public Function Clamp(F As Double, T As Double, A As Double) As Double
    Clamp = A
    If Clamp < F Then
        Clamp = F
    ElseIf Clamp > T Then
        Clamp = T
    End If
End Function

Public Function rndFT(F As Double, T As Double) As Double
    rndFT = (T - F) * Rnd + F
End Function

'inline bool BiasGreaterThan( real a, real b )
'{
'  const real k_biasRelative = 0.95f;
'  const real k_biasAbsolute = 0.01f;
'  return a >= b * k_biasRelative + a * k_biasAbsolute;
'}
Public Function BiasGreaterThan(A As Double, B As Double) As Boolean
    BiasGreaterThan = (A >= (B * BIAS_RELATIVE + A * BIAS_ABSOLUTE))
End Function

Public Function gt(A As Double, B As Double) As Boolean
'return a >= b * BIAS_RELATIVE + a * BIAS_ABSOLUTE;
    gt = (A >= (B * BIAS_RELATIVE + A * BIAS_ABSOLUTE))
End Function


'********************** MATHS: ********************************


Public Function Min(A As Double, B As Double) As Double
    If A < B Then
        Min = A
    Else
        Min = B
    End If
End Function
Public Function Max(A As Double, B As Double) As Double
    If A > B Then
        Max = A
    Else
        Max = B
    End If
End Function




'************************ TIMING ********************************

''Public Property Get Timing() As Double    ''DDOUBLE
''    Dim curTime As Currency
''    QueryPerformanceCounter curTime
''    Timing = (curTime - m_TimeStart) * m_TimeFreq + m_Time
''End Property
''
''Public Property Let Timing(ByVal NewValue As Double)  ''DDOUBLE
''    Dim curFreq As Currency, curOverhead As Currency
''    m_Time = NewValue
''    QueryPerformanceFrequency curFreq
''    m_TimeFreq = 1 / curFreq
''    QueryPerformanceCounter curOverhead
''    QueryPerformanceCounter m_TimeStart
''    m_TimeStart = m_TimeStart + (m_TimeStart - curOverhead)
''End Property

