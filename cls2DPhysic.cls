VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls2DPhysic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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






Option Explicit



'*************************************************************************
'************************* M A T E R I A L S *****************************
'*************************************************************************

Private Enum eMat
    mStone
    mMetal
    mGlass
    mWood
    mPlastic
    mRubber
    mFlesh
End Enum

Private Type tMaterial
    SF          As Double    'Static Friction
    DF          As Double    'Dynamic Friction
    Restitution As Double
    Density     As Double
End Type





Private Contacts() As tManifold
Public NofContactMainFolds As Long
Public MAXNofContactMainFolds As Long

Public BiggerGroup As Long




'*************************************************************************
'************************* B O D Y  *****************************
'*************************************************************************

Private Type tAABB
    pMin        As tVec2
    pMax        As tVec2
End Type

Private Enum eBodyType
    eCircle = 0    '&H0
    ePolygon = 1    '&H1
End Enum


Private Type tBody
    myType      As eBodyType
    mass        As Double
    invMass     As Double
    inertia     As Double
    invInertia  As Double

    ii          As Double


    Area        As Double


    COM         As tVec2

    '------------------
    'Circle
    Radius      As Double
    '------------------
    'Poligon
    Vertex()    As tVec2
    tVertex()   As tVec2

    normals()   As tVec2
    Nvertex     As Long
    '------------------


    Pos         As tVec2
    VEL         As tVec2
    FORCE       As tVec2
    angularVelocity As Double
    torque      As Double
    orient      As Double

    'Material
    staticFriction As Double
    dynamicFriction As Double
    Restitution As Double

    U           As tMAT2

    color       As Long

    AABB        As tAABB

    CollisionGroup As Long
    CollideWith As Long

    ImageExist  As Boolean
    ImageCenter As tVec2
    ImageSize   As tVec2

    MyIDX       As Long

End Type

Private ImageList As cImageList



''Private Type tSHAPE
''    myType  As eBodyType
''    '------------------
''    'Circle
''    Radius  As Double
''    '------------------
''    'Poligon
''    Vertex() As tVec2
''    tVertex() As tVec2
''
''    normals() As tVec2
''    Nvertex As Long
''    '------------------
''
''    Pos     As tVec2
''    orient  As Double
''
''    color   As Long
''
''    LToCT() As Long
''
''
''End Type


Private Body()  As tBody
'Private Shape() As tSHAPE

Private NBodies As Long


'***************** Mouse Interaction
Public MouseDownRelX As Double
Public MouseDownRelY As Double
Public MouseMoveX As Double
Public MouseMoveY As Double
Public MouseSelectedObj As Long



'*************************************************************************
'************************* ENGINE  *****************************
'*************************************************************************





Public TotalNContacts As Long

Public Version  As String




'**************************************************************
'********************JOINTS *********************************
'**************************************************************


Private Enum eJointType
    JointDistance
    Joint2Pins
    Joint1Pin
    Rotor1
    Rotor2
End Enum



Private Type tJoint
    JointType   As eJointType
    bA          As Long
    bB          As Long
    L           As Double
    AnchA       As tVec2
    AnchB       As tVec2
    tAnchA      As tVec2
    tAnchB      As tVec2

    StifPULL    As Double
    StifPUSH    As Double
End Type

Private Joints() As tJoint
Public NJ       As Long

Private Const KMASS As Double = 0.1
Private Const KDamp As Double = 0.125    '0.5 '0.125




'*********************************************************************
'******************* R E N D E R *************************************
'*********************************************************************

Private Srf As cCairoSurface, CC As cCairoContext    'Srf is similar to a DIB, the derived CC similar to a hDC
Attribute CC.VB_VarUserMemId = 1073938448
Public vbDrawCC As cCairoContext

Public vbDRAW   As cVBDraw
Attribute vbDRAW.VB_VarUserMemId = 1073938445
Private CONS    As cConstructor
Attribute CONS.VB_VarUserMemId = 1073938446

Public PicHDC   As Long
Attribute PicHDC.VB_VarUserMemId = 1073938447

Private Const JPGQuality As Long = 95


'***********************************************************************
'**************** COLLISIONS ******************************************
'**********************************************************************
Private Enum eCollision
    eCollisionCircleCircle = 0    '&H0
    eCollisionCirclePolygon = 1    '&H1
    eCollisionPolygonCircle = 2    '&H10
    eCollisionPolygonPolygon = 3    '&H11
End Enum


Private Type tManifold

    ContactType As eCollision

    bodyA       As Long
    bodyB       As Long

    Penetration As Double
    normal      As tVec2

    contactsPTS(1 To 2) As tVec2
    contactCount As Long
    invContactCount As Double

    MAXcontactCount As Long

    e           As Double
    DF          As Double
    SF          As Double

End Type





Public Event CollisionEvent(bA As Long, bB As Long, _
                            posAX As Double, PosAY As Double, _
                            posBX As Double, PosBY As Double, _
                            Nx As Double, Ny As Double, _
                            ContactVelo As Double)





'*************************************************************************
'************************* M A T E R I A L S *****************************
'*************************************************************************

Private MATERIAL() As tMaterial
Attribute MATERIAL.VB_VarUserMemId = 1073938448

Private Sub MaterialsINIT()

    ReDim MATERIAL(10)

    With MATERIAL(mStone)    'STONE
        .SF = 0.8
        .DF = 0.8
        .Restitution = 0.4
        .Density = 1
    End With

    With MATERIAL(mMetal)    'METAL
        .SF = 0.3
        .DF = 0.3
        .Restitution = 0.4
        .Density = 1
    End With

    With MATERIAL(mGlass)    'GLASS
        .SF = 0.6
        .DF = 0.6
        .Restitution = 0.5
        .Density = 1
    End With

    With MATERIAL(mWood)    'WOOD
        .SF = 0.6
        .DF = 0.6
        .Restitution = 0.5
    End With

    With MATERIAL(mFlesh)    'FLESH
        .SF = 0.9
        .DF = 0.9
        .Restitution = 0.3
        .Density = 1
    End With

    With MATERIAL(mPlastic)    'plastic
        .SF = 0.4
        .DF = 0.4
        .Restitution = 0.7
        .Density = 1
    End With

    With MATERIAL(mRubber)    'RUBBER
        .SF = 0.9
        .DF = 0.9
        .Restitution = 0.9
        .Density = 1
    End With



End Sub

'*************************************************************************

'*************************************************************************
'************************* COLLISIONS *****************************
'*************************************************************************




Private Sub collContactsINIT(wC As Long)
    Dim I       As Long
    Dim rA      As tVec2
    Dim rB      As tVec2
    Dim rV      As tVec2
    Dim A       As Long
    Dim B       As Long



    With Contacts(wC)
        A = .bodyA
        B = .bodyB


        '   // Calculate average restitution
        '   // e = std::min( A->restitution, B->restitution );
        '   e = StrictMath.min( A.restitution, B.restitution );
        .e = Min(Body(A).Restitution, Body(B).Restitution)

        '// Calculate static and dynamic friction
        '// sf = std::sqrt( A->staticFriction * A->staticFriction );
        '// df = std::sqrt( A->dynamicFriction * A->dynamicFriction );
        'sf = (float)StrictMath.sqrt( A.staticFriction * A.staticFriction );
        'df = (float)StrictMath.sqrt( A.dynamicFriction * A.dynamicFriction );

        'https://github.com/ClickerMonkey/ImpulseEngine/blob/master/src/org/magnos/impulse/Manifold.java:
        'sf = (float)StrictMath.sqrt( A.staticFriction * A.staticFriction + B.staticFriction * B.staticFriction);
        'df = (float)StrictMath.sqrt( A.dynamicFriction * A.dynamicFriction + B.dynamicFriction * B.dynamicFriction);

        .SF = Sqr(Body(A).staticFriction * Body(A).staticFriction + Body(B).staticFriction * Body(B).staticFriction)
        .DF = Sqr(Body(A).dynamicFriction * Body(A).dynamicFriction + Body(B).dynamicFriction * Body(B).dynamicFriction)


        For I = 1 To .contactCount

            '// Calculate radii from COM to contact
            '// Vec2 ra = contacts[i] - A->position;
            '// Vec2 rb = contacts[i] - B->position;
            'Vec2 ra = contacts[i].sub( A.position );
            'Vec2 rb = contacts[i].sub( B.position );

            rA = Vec2SUB(.contactsPTS(I), Body(A).Pos)
            rB = Vec2SUB(.contactsPTS(I), Body(B).Pos)

            ' // Vec2 rv = B->velocity + Cross( B->angularVelocity, rb ) -
            ' // A->velocity - Cross( A->angularVelocity, ra );
            ' Vec2 rv = B.velocity.add( Vec2.cross( B.angularVelocity, rb, new Vec2() ) ).subi( A.velocity ).subi( Vec2.cross( A.angularVelocity, ra, new Vec2() ) );

            rV = Vec2ADD(Body(B).VEL, Vec2CROSSav(Body(B).angularVelocity, rB))
            rV = Vec2SUB(rV, Body(A).VEL)
            rV = Vec2SUB(rV, Vec2CROSSav(Body(A).angularVelocity, rA))


            '            // Determine if we should perform a resting collision or not
            '            // The idea is if the only thing moving this object is gravity,
            '            // then the collision should be performed without any restitution
            '            // if(rv.LenSqr( ) < (dt * gravity).LenSqr( ) + EPSILON)
            '            if (rv.lengthSq() < ImpulseMath.RESTING)
            '            {
            '            e = 0.0f;
            '            }

            If Vec2LengthSq(rV) < RESTING Then .e = 0


        Next

    End With


End Sub
Private Sub CollinfiniteMassCorrection(wC As Long)
    Body(Contacts(wC).bodyA).VEL = Vec2(0, 0)
    Body(Contacts(wC).bodyB).VEL = Vec2(0, 0)
End Sub
Private Sub collContactsPositionalCorrection(wC As Long)
    Dim correction As Double

    '// const real k_slop = 0.05f; // Penetration allowance
    '// const real percent = 0.4f; // Penetration percentage to correct
    '// Vec2 correction = (std::max( penetration - k_slop, 0.0f ) / (A->im +
    '// B->im)) * normal * percent;
    '// A->position -= correction * A->im;
    '// B->position += correction * B->im;

    'float correction = StrictMath.max( penetration - ImpulseMath.PENETRATION_ALLOWANCE, 0.0f ) / (A.invMass + B.invMass) * ImpulseMath.PENETRATION_CORRETION;

    'A.position.addsi( normal, -A.invMass * correction );
    'B.position.addsi( normal, B.invMass * correction );


    With Contacts(wC)
        correction = (Max(.Penetration - PENETRATION_ALLOWANCE, 0) / (Body(.bodyA).invMass + Body(.bodyB).invMass)) _
                     * PENETRATION_CORRETION

        If correction Then
            Body(.bodyA).Pos = Vec2ADDScaled(Body(.bodyA).Pos, .normal, Body(.bodyA).invMass * -correction)
            Body(.bodyB).Pos = Vec2ADDScaled(Body(.bodyB).Pos, .normal, Body(.bodyB).invMass * correction)
        End If

    End With

End Sub
Private Sub collContactsApplyImpulse(wC As Long)
    Dim A       As Long
    Dim B       As Long

    Dim I       As Long

    Dim rA      As tVec2
    Dim rB      As tVec2
    Dim rV      As tVec2
    Dim contactVel As Double


    Dim rACrossN As Double
    Dim rBCrossN As Double
    Dim InvMassSUM As Double
    Dim MassSUM As Double

    Dim J       As Double
    Dim Jt      As Double


    Dim impulse As tVec2
    Dim T       As tVec2

    Dim tangentImpulse As tVec2


    With Contacts(wC)
        A = .bodyA
        B = .bodyB

        '// Early out and positional correct if both objects have infinite mass
        '// if(Equal( A->im + B->im, 0 ))
        'if (ImpulseMath.equal( A.invMass + B.invMass, 0 ))
        '{
        '    CollinfiniteMassCorrection();
        '    return;
        '}


        If Equal(Body(A).invMass + Body(B).invMass, 0) Then
            CollinfiniteMassCorrection wC
        Else

            For I = 1 To .contactCount

                '   // Calculate radii from COM to contact
                '   // Vec2 ra = contacts[i] - A->position;
                '   // Vec2 rb = contacts[i] - B->position;
                '   Vec2 ra = contacts[i].sub( A.position );
                '   Vec2 rb = contacts[i].sub( B.position );
                rA = Vec2SUB(.contactsPTS(I), Body(A).Pos)
                rB = Vec2SUB(.contactsPTS(I), Body(B).Pos)

                ' // Relative velocity
                ' // Vec2 rv = B->velocity + Cross( B->angularVelocity, rb ) -
                ' // A->velocity - Cross( A->angularVelocity, ra );
                ' Vec2 rv = B.velocity.add( Vec2.cross( B.angularVelocity, rb, new Vec2() ) ).subi( A.velocity ).subi( Vec2.cross( A.angularVelocity, ra, new Vec2() ) );

                rV = Vec2ADD(Body(B).VEL, Vec2CROSSav(Body(B).angularVelocity, rB))
                rV = Vec2SUB(rV, Body(A).VEL)
                rV = Vec2SUB(rV, Vec2CROSSav(Body(A).angularVelocity, rA))


                '// Relative velocity along the normal
                '// real contactVel = Dot( rv, normal );
                'float contactVel = Vec2.dot( rv, normal );
                contactVel = Vec2DOT(rV, .normal)

                '// Do not resolve if velocities are separating
                'if (contactVel > 0)
                '{
                'return;
                '}
                If contactVel <= 0 Then
                    '// real raCrossN = Cross( ra, normal );
                    '// real rbCrossN = Cross( rb, normal );
                    '// real invMassSum = A->im + B->im + FASTsqr( raCrossN ) * A->iI + FASTsqr(
                    '// rbCrossN ) * B->iI;
                    'float raCrossN = Vec2.cross( ra, normal );
                    'float rbCrossN = Vec2.cross( rb, normal );
                    'float invMassSum = A.invMass + B.invMass + (raCrossN * raCrossN) * A.invInertia + (rbCrossN * rbCrossN) * B.invInertia;
                    rACrossN = Vec2CROSS(rA, .normal)
                    rBCrossN = Vec2CROSS(rB, .normal)

                    'InvMassSUM = Body(A).invMass + Body(B).invMass + _
                     (rACrossN * rACrossN) * Body(A).invInertia + _
                     (rBCrossN * rBCrossN) * Body(B).invInertia
                    MassSUM = 1# / (Body(A).invMass + Body(B).invMass + _
                                    (rACrossN * rACrossN) * Body(A).invInertia + _
                                    (rBCrossN * rBCrossN) * Body(B).invInertia)



                    ' // Calculate impulse scalar
                    ' float j = -(1.0f + e) * contactVel;
                    ' j /= invMassSum;
                    ' j /= contactCount;
                    J = -(1 + .e) * contactVel
                    'J = J / InvMassSUM
                    J = J * MassSUM
                    J = J * .invContactCount    '/ .contactcount


                    ' // Apply impulse
                    ' Vec2 impulse = normal.mul( j );
                    ' A.applyImpulse( impulse.neg(), ra );
                    ' B.applyImpulse( impulse, rb );
                    impulse = Vec2MUL(.normal, J)
                    BodyApplyImpulse A, Vec2Negative(impulse), rA
                    BodyApplyImpulse B, impulse, rB






                    If contactVel * .invContactCount < -0.001 Then
                        RaiseEvent CollisionEvent(A, B, _
                                                  rA.X, rA.Y, _
                                                  rB.X, rB.Y, _
                                                  .normal.X, .normal.Y, _
                                                  -contactVel)
                    End If







                    '          // Friction impulse
                    '          // rv = B->velocity + Cross( B->angularVelocity, rb ) -
                    '          // A->velocity - Cross( A->angularVelocity, ra );
                    '          rv = B.velocity.add( Vec2.cross( B.angularVelocity, rb, new Vec2() ) ).subi( A.velocity ).subi( Vec2.cross( A.angularVelocity, ra, new Vec2() ) );
                    rV = Vec2ADD(Body(B).VEL, Vec2CROSSav(Body(B).angularVelocity, rB))
                    rV = Vec2SUB(rV, Body(A).VEL)
                    rV = Vec2SUB(rV, Vec2CROSSav(Body(A).angularVelocity, rA))

                    '// Vec2 t = rv - (normal * Dot( rv, normal ));
                    '// t.Normalize( );
                    'Vec2 t = new Vec2( rv );
                    't.addsi( normal, -Vec2.dot( rv, normal ) );
                    't.normalize();
                    T = Vec2SUB(rV, Vec2MUL(.normal, Vec2DOT(rV, .normal)))
                    T = Vec2Normalize(T)

                    '// j tangent magnitude
                    'float jt = -Vec2.dot( rv, t );
                    'jt /= invMassSum;
                    'jt /= contactCount;
                    Jt = -Vec2DOT(rV, T)
                    'Jt = Jt / InvMassSUM
                    Jt = Jt * MassSUM

                    Jt = Jt * .invContactCount    '/ .contactcount


                    '// Don    't apply tiny friction impulses
                    'if (ImpulseMath.equal( jt, 0.0f ))
                    '{
                    'return;
                    '}

                    If Not (Equal(Jt, 0)) Then

                        '// Coulumb    's law
                        'Vec2 tangentImpulse;
                        '// if(std::abs( jt ) < j * sf)
                        'if (StrictMath.abs( jt ) < j * sf)
                        '{
                        '// tangentImpulse = t * jt;
                        'tangentImpulse = t.mul( jt );
                        '}
                        'Else
                        '{
                        '// tangentImpulse = t * -j * df;
                        'tangentImpulse = t.mul( j ).muli( -df );
                        '}


                        If Abs(Jt) < J * .SF Then
                            tangentImpulse = Vec2MUL(T, Jt)
                        Else
                            tangentImpulse = Vec2MUL(T, -J * .DF)
                        End If

                        '// Apply friction impulse
                        '// A->ApplyImpulse( -tangentImpulse, ra );
                        '// B->ApplyImpulse( tangentImpulse, rb );
                        'A.applyImpulse( tangentImpulse.neg(), ra );
                        'B.applyImpulse( tangentImpulse, rb );
                        BodyApplyImpulse A, Vec2Negative(tangentImpulse), rA
                        BodyApplyImpulse B, tangentImpulse, rB






                    Else
                        Exit Sub    'Don't apply tiny friction impulses
                    End If
                Else
                    Exit Sub    ' Do not resolve if velocities are separating
                End If
            Next

        End If


    End With

End Sub

'Public Function CollisionCirclePolygon(wbA As Long, wBB As Long) As tManifold
'
'    Dim A      As tBody
'    Dim B      As tBody
'
'    Dim normal As tVec2
'    Dim DistDist As double
'    Dim radius As double
'    Dim distance As double
'
'    Dim center As tVec2
'
'
'    Dim separation As double
'    Dim faceNormal As Long
'
'    Dim I      As Long
'    Dim I2     As Long
'
'    Dim S      As double
'
'    Dim v1     As tVec2
'    Dim v2     As tVec2
'
'    Dim Dot1   As double
'    Dim Dot2   As double
'
'    Dim N      As tVec2
'
'
'
'    CollisionCirclePolygon.bodyA = wbA
'    CollisionCirclePolygon.bodyB = wBB
'    A = Body(wbA)
'    B = Body(wBB)
'
'    '*******************************************************************************
'    '*******************************************************************************
'    CollisionCirclePolygon.contactCount = 0
'
'    '// Transform circle center to Polygon model space
'    '// Vec2 center = a->position;
'    '// center = B->u.Transpose( ) * (center - b->position);
'    'Vec2 center = B.u.transpose().muli( a.position.sub( b.position ) );
'    center = matMULv(matTranspose(B.U), Vec2SUB(A.Pos, B.Pos))
'
'
'    '        // Find edge with minimum penetration
'    '        // Exact concept as using support points in Polygon vs Polygon
'    '        float separation = -Float.MAX_VALUE;
'    '        int faceNormal = 0;
'    separation = -MAX_VALUE
'    faceNormal = 0
'
'
'
'
'    For I = 1 To B.Nvertex
'
'        '// real s = Dot( B->m_normals[i], center - B->m_vertices[i] );
'        'float s = Vec2.dot( B.normals[i], center.sub( B.vertices[i] ) );
'
'        S = Vec2DOT(B.normals(I), Vec2SUB(center, B.Vertex(I)))
'
'        If (S < A.radius) Then
'            If (S > separation) Then
'                separation = S
'                faceNormal = I
'            End If
'        End If
'
'    Next
'
'
'
'    '// Grab face's vertices
'    'Vec2 v1 = B.vertices[faceNormal];
'    'int i2 = faceNormal + 1 < B.vertexCount ? faceNormal + 1 : 0;
'    'Vec2 v2 = B.vertices[i2];
'
'    v1 = B.Vertex(faceNormal)
'    I2 = faceNormal + 1
'    If I2 > B.Nvertex Then I2 = 1
'    v2 = B.Vertex(I2)
'
'
'
'    '// Check to see if center is within polygon
'    'if (separation < ImpulseMath.EPSILON)
'    '{
'    '    // m->contact_count = 1;
'    '    // m->normal = -(B->u * B->m_normals[faceNormal]);
'    '    // m->contacts[0] = m->normal * A->radius + a->position;
'    '    // m->penetration = A->radius;
'
'    '    m.contactCount = 1;
'    '    B.u.mul( B.normals[faceNormal], m.normal ).negi();
'    '    m.contacts[0].set( m.normal ).muli( A.radius ).addi( a.position );
'    '    m.penetration = A.radius;
'    '    return;
'    '}
'    If separation < EPSILON Then
'        CollisionCirclePolygon.contactCount = 1
'        ReDim CollisionCirclePolygon.contactsPTS(1)
'        CollisionCirclePolygon.normal = Vec2Negative(matMULv(B.U, B.normals(faceNormal)))
'        CollisionCirclePolygon.contactsPTS(1) = Vec2ADD(Vec2MUL(CollisionCirclePolygon.normal, A.radius), A.Pos)
'        CollisionCirclePolygon.penetration = A.radius
'    Else    'because of return
'
'
'        '// Determine which voronoi region of the edge center of circle lies within
'        '// real dot1 = Dot( center - v1, v2 - v1 );
'        '// real dot2 = Dot( center - v2, v1 - v2 );
'        '// m->penetration = A->radius - separation;
'        'float dot1 = Vec2.dot( center.sub( v1 ), v2.sub( v1 ) );
'        'float dot2 = Vec2.dot( center.sub( v2 ), v1.sub( v2 ) );
'        'm.penetration = A.radius - separation;
'
'        Dot1 = Vec2DOT(Vec2SUB(center, v1), Vec2SUB(v2, v1))
'        Dot2 = Vec2DOT(Vec2SUB(center, v2), Vec2SUB(v1, v2))
'        CollisionCirclePolygon.penetration = A.radius - separation
'
'
'        '        // Closest to v1
'        '        if (dot1 <= 0.0f)
'        '        {
'        '            if (Vec2.distanceSq( center, v1 ) > A.radius * A.radius)
'        '            {
'        '                return;
'        '            }
'        '
'        '            // m->contact_count = 1;
'        '            // Vec2 n = v1 - center;
'        '            // n = B->u * n;
'        '            // n.Normalize( );
'        '            // m->normal = n;
'        '            // v1 = B->u * v1 + b->position;
'        '            // m->contacts[0] = v1;
'        '
'        '            m.contactCount = 1;
'        '            B.u.muli( m.normal.set( v1 ).subi( center ) ).normalize();
'        '            B.u.mul( v1, m.contacts[0] ).addi( b.position );
'        '        }
'        If Dot1 <= 0 Then
'
'            If Vec2DISTANCEsq(center, v1) < A.radius * A.radius Then
'
'                CollisionCirclePolygon.contactCount = 1
'                ReDim CollisionCirclePolygon.contactsPTS(1)
'                N = Vec2SUB(v1, center)
'                CollisionCirclePolygon.normal = Vec2Normalize(matMULv(B.U, N))
'                v1 = Vec2ADD(matMULv(B.U, v1), B.Pos)
'                CollisionCirclePolygon.contactsPTS(1) = v1
'
'            End If
'            '        else if (dot2 <= 0.0f)
'        ElseIf Dot2 <= 0 Then
'
'            '                {
'            '                if (Vec2.distanceSq( center, v2 ) > A.radius * A.radius)
'            '                {
'            '                return;
'            '                }
'            '
'            '                // m->contact_count = 1;
'            '                // Vec2 n = v2 - center;
'            '                // v2 = B->u * v2 + b->position;
'            '                // m->contacts[0] = v2;
'            '                // n = B->u * n;
'            '                // n.Normalize( );
'            '                // m->normal = n;
'            '
'            '                m.contactCount = 1;
'            '                B.u.muli( m.normal.set( v2 ).subi( center ) ).normalize();
'            '                B.u.mul( v2, m.contacts[0] ).addi( b.position );
'            '                }
'            If Vec2DISTANCEsq(center, v2) < A.radius * A.radius Then
'
'                CollisionCirclePolygon.contactCount = 1
'                ReDim CollisionCirclePolygon.contactsPTS(1)
'                N = Vec2SUB(v2, center)
'                v2 = Vec2ADD(matMULv(B.U, v2), B.Pos)
'                CollisionCirclePolygon.contactsPTS(1) = v2
'                CollisionCirclePolygon.normal = Vec2Normalize(matMULv(B.U, N))
'            End If
'
'
'
'            '// Closest to face
'        Else
'            '                Vec2 n = B.normals[faceNormal];
'            '
'            '                if (Vec2.dot( center.sub( v1 ), n ) > A.radius)
'            '                {
'            '                return;
'            '                }
'            '
'            '                // n = B->u * n;
'            '                // m->normal = -n;
'            '                // m->contacts[0] = m->normal * A->radius + a->position;
'            '                // m->contact_count = 1;
'            '
'            '                m.contactCount = 1;
'            '                B.u.mul( n, m.normal ).negi();
'            '                m.contacts[0].set( a.position ).addsi( m.normal, A.radius );
'
'
'            N = B.normals(faceNormal)
'            If Vec2DOT(Vec2SUB(center, v1), N) < A.radius Then
'
'
'                N = matMULv(B.U, N)
'                CollisionCirclePolygon.normal = Vec2Negative(N)
'                CollisionCirclePolygon.contactCount = 1
'                ReDim CollisionCirclePolygon.contactsPTS(1)
'                CollisionCirclePolygon.contactsPTS(1) = Vec2ADD(Vec2MUL(CollisionCirclePolygon.normal, A.radius), A.Pos)
'
'            End If
'
'        End If
'
'    End If
'
'End Function


Private Function collCollisionSOLVE(wbA As Long, wbB As Long) As tManifold
    Dim A       As tBody
    Dim B       As tBody

    Dim normal  As tVec2
    Dim DistDist As Double
    Dim Radius  As Double
    Dim distance As Double

    Dim Center  As tVec2

    Dim separation As Double
    Dim faceNormal As Long

    Dim I       As Long
    Dim I2      As Long

    Dim S       As Double

    Dim V1      As tVec2
    Dim V2      As tVec2

    Dim Dot1    As Double
    Dim Dot2    As Double

    Dim N       As tVec2

    Dim ContactType As eCollision    'Long

    Dim faceA   As Long    '---------------polypoly
    Dim faceB   As Long
    Dim penetrationA As Double
    Dim penetrationB As Double
    Dim flip    As Boolean
    Dim RefPoly As tBody
    Dim IncPoly As tBody
    Dim incidentFace(0 To 1) As tVec2
    Dim referenceIndex As Long
    Dim sidePlaneNormal As tVec2
    Dim refFaceNormal As tVec2
    Dim refC    As Double
    Dim posSide As Double
    Dim negSide As Double
    Dim cp      As Double



    '---------------------------------------------------------------
    With collCollisionSOLVE

        ContactType = Body(wbA).myType * 2 + Body(wbB).myType

        Select Case ContactType

        Case eCollisionCircleCircle    '0

            '*******************************************************************************
            '*******************************************************************************
            .bodyA = wbA
            .bodyB = wbB
            A = Body(wbA)
            B = Body(wbB)

            '        // Calculate translational vector, which is normal
            '        // Vec2 normal = b->position - a->position;
            '        Vec2 normal = b.position.sub( a.position );
            normal = Vec2SUB(B.Pos, A.Pos)
            '        // real DistDist = normal.LenSqr( );
            '        // real radius = A->radius + B->radius;
            '        float DistDist = normal.lengthSq();
            '        float radius = A.radius + B.radius;
            DistDist = Vec2LengthSq(normal)
            Radius = A.Radius + B.Radius

            '        // Not in contact
            If (DistDist >= Radius * Radius) Then
                .contactCount = 0
            Else

                distance = Sqr(DistDist)

                .contactCount = 1
                .invContactCount = 1

                If (distance = 0) Then
                    ' // m->penetration = A->radius;
                    ' // m->normal = Vec2( 1, 0 );
                    ' // m->contacts [0] = a->position;
                    ' m.penetration = A.radius;
                    ' m.normal.set( 1.0f, 0.0f );
                    ' m.contacts[0].set( a.position );
                    .Penetration = A.Radius
                    .normal.X = 1
                    .normal.Y = 0
                    'ReDim .contactsPTS(1)
                    .contactsPTS(1) = A.Pos

                Else
                    '// m->penetration = radius - distance;
                    '// m->normal = normal / distance; // Faster than using Normalized since
                    '// we already performed sqrt
                    '// m->contacts[0] = m->normal * A->radius + a->position;
                    'm.penetration = radius - distance;
                    'm.normal.set( normal ).divi( distance );
                    'm.contacts[0].set( m.normal ).muli( A.radius ).addi( a.position );
                    .Penetration = Radius - distance
                    .normal = Vec2MUL(normal, 1# / distance)
                    'ReDim .contactsPTS(1)
                    .contactsPTS(1) = Vec2ADD(A.Pos, Vec2MUL(.normal, A.Radius))
                End If

            End If


        Case eCollisionCirclePolygon    '1    ' CirclePolygon
            .bodyA = wbA
            .bodyB = wbB
            A = Body(wbA)
            B = Body(wbB)
            GoSub LABELCirclePolygon


        Case eCollisionPolygonCircle    '2    'PolygonCircle

            .bodyA = wbA
            .bodyB = wbB
            A = Body(wbB)    '*
            B = Body(wbA)    '*
            GoSub LABELCirclePolygon
            .normal = Vec2Negative(.normal)



        Case eCollisionPolygonPolygon    '3    ' PolygonPolygon
            '*******************************************************************************
            '*******************************************************************************
            .bodyA = wbA
            .bodyB = wbB
            A = Body(wbA)
            B = Body(wbB)
            GoSub LabelPolygonPolygon

        End Select





        Exit Function

LABELCirclePolygon:

        '*******************************************************************************
        '*******************************************************************************
        .contactCount = 0

        '// Transform circle center to Polygon model space
        '// Vec2 center = a->position;
        '// center = B->u.Transpose( ) * (center - b->position);
        'Vec2 center = B.u.transpose().muli( a.position.sub( b.position ) );

        Center = matMULv(matTranspose(B.U), Vec2SUB(A.Pos, B.Pos))



        '        // Find edge with minimum penetration
        '        // Exact concept as using support points in Polygon vs Polygon
        '        float separation = -Float.MAX_VALUE;
        '        int faceNormal = 0;
        separation = -MAX_VALUE
        faceNormal = 0

        For I = 1 To B.Nvertex

            '// real s = Dot( B->m_normals[i], center - B->m_vertices[i] );
            'float s = Vec2.dot( B.normals[i], center.sub( B.vertices[i] ) );

            S = Vec2DOT(B.normals(I), Vec2SUB(Center, B.Vertex(I)))
            '    if(s > A->radius)
            '      return;
            '    if(s > separation)
            '    {
            '      separation = s;
            '      faceNormal = i;
            '    }
            If (S <= A.Radius) Then
                If (S > separation) Then
                    separation = S
                    faceNormal = I
                End If
            Else

                Return

            End If
        Next




        '// Grab face's vertices
        'Vec2 v1 = B.vertices[faceNormal];
        'int i2 = faceNormal + 1 < B.vertexCount ? faceNormal + 1 : 0;
        'Vec2 v2 = B.vertices[i2];

        V1 = B.Vertex(faceNormal)
        I2 = faceNormal + 1: If I2 > B.Nvertex Then I2 = 1
        V2 = B.Vertex(I2)



        '// Check to see if center is within polygon
        'if (separation < ImpulseMath.EPSILON)
        '{
        '    // m->contact_count = 1;
        '    // m->normal = -(B->u * B->m_normals[faceNormal]);
        '    // m->contacts[0] = m->normal * A->radius + a->position;
        '    // m->penetration = A->radius;

        '    m.contactCount = 1;
        '    B.u.mul( B.normals[faceNormal], m.normal ).negi();
        '    m.contacts[0].set( m.normal ).muli( A.radius ).addi( a.position );
        '    m.penetration = A.radius;
        '    return;
        '}
        If separation < EPSILON Then
            .contactCount = 1
            .invContactCount = 1
            'ReDim .contactsPTS(1)
            .normal = Vec2Negative(matMULv(B.U, B.normals(faceNormal)))
            .contactsPTS(1) = Vec2ADD(Vec2MUL(.normal, A.Radius), A.Pos)
            .Penetration = A.Radius
            Return
            'Else    'because of return
        End If

        '// Determine which voronoi region of the edge center of circle lies within
        '// real dot1 = Dot( center - v1, v2 - v1 );
        '// real dot2 = Dot( center - v2, v1 - v2 );
        '// m->penetration = A->radius - separation;
        'float dot1 = Vec2.dot( center.sub( v1 ), v2.sub( v1 ) );
        'float dot2 = Vec2.dot( center.sub( v2 ), v1.sub( v2 ) );
        'm.penetration = A.radius - separation;

        Dot1 = Vec2DOT(Vec2SUB(Center, V1), Vec2SUB(V2, V1))
        Dot2 = Vec2DOT(Vec2SUB(Center, V2), Vec2SUB(V1, V2))
        .Penetration = A.Radius - separation


        '        // Closest to v1
        '        if (dot1 <= 0.0f)
        '        {
        '            if (Vec2.distanceSq( center, v1 ) > A.radius * A.radius)
        '            {
        '                return;
        '            }
        '
        '            // m->contact_count = 1;
        '            // Vec2 n = v1 - center;
        '            // n = B->u * n;
        '            // n.Normalize( );
        '            // m->normal = n;
        '            // v1 = B->u * v1 + b->position;
        '            // m->contacts[0] = v1;
        '
        '            m.contactCount = 1;
        '            B.u.muli( m.normal.set( v1 ).subi( center ) ).normalize();
        '            B.u.mul( v1, m.contacts[0] ).addi( b.position );
        '        }
        If Dot1 <= 0 Then

            If Vec2DISTANCEsq(Center, V1) <= A.Radius * A.Radius Then

                .contactCount = 1
                .invContactCount = 1
                'ReDim .contactsPTS(1)
                N = Vec2SUB(V1, Center)
                .normal = Vec2Normalize(matMULv(B.U, N))
                V1 = Vec2ADD(matMULv(B.U, V1), B.Pos)
                .contactsPTS(1) = V1
            Else
                Return
            End If
            '        else if (dot2 <= 0.0f)
        ElseIf Dot2 <= 0 Then

            '                {
            '                if (Vec2.distanceSq( center, v2 ) > A.radius * A.radius)
            '                {
            '                return;
            '                }
            '
            '                // m->contact_count = 1;
            '                // Vec2 n = v2 - center;
            '                // v2 = B->u * v2 + b->position;
            '                // m->contacts[0] = v2;
            '                // n = B->u * n;
            '                // n.Normalize( );
            '                // m->normal = n;
            '
            '                m.contactCount = 1;
            '                B.u.muli( m.normal.set( v2 ).subi( center ) ).normalize();
            '                B.u.mul( v2, m.contacts[0] ).addi( b.position );
            '                }
            If Vec2DISTANCEsq(Center, V2) <= A.Radius * A.Radius Then

                .contactCount = 1
                .invContactCount = 1
                ' ReDim .contactsPTS(1)
                N = Vec2SUB(V2, Center)
                V2 = Vec2ADD(matMULv(B.U, V2), B.Pos)
                .contactsPTS(1) = V2
                .normal = Vec2Normalize(matMULv(B.U, N))
            Else
                Return
            End If



            '// Closest to face
        Else

            '                Vec2 n = B.normals[faceNormal];
            '
            '                if (Vec2.dot( center.sub( v1 ), n ) > A.radius)
            '                {
            '                return;
            '                }
            '
            '                // n = B->u * n;
            '                // m->normal = -n;
            '                // m->contacts[0] = m->normal * A->radius + a->position;
            '                // m->contact_count = 1;
            '
            '                m.contactCount = 1;
            '                B.u.mul( n, m.normal ).negi();
            '                m.contacts[0].set( a.position ).addsi( m.normal, A.radius );


            N = B.normals(faceNormal)
            If Vec2DOT(Vec2SUB(Center, V1), N) <= A.Radius Then

                N = matMULv(B.U, N)
                .normal = Vec2Negative(N)
                .contactCount = 1
                .invContactCount = 1
                ' ReDim .contactsPTS(1)
                .contactsPTS(1) = Vec2ADD(Vec2MUL(.normal, A.Radius), A.Pos)
            Else
                Return

            End If

        End If

        'End If ''''(Return up)
        Return




LabelPolygonPolygon:
        '*******************************************************************************
        '*******************************************************************************
        .contactCount = 0


        'penetrationA = collFindAxisLeastPenetration( &faceA, A, B );
        penetrationA = collFindAxisLeastPenetration(faceA, A, B)
        If penetrationA >= 0 Then Return
        penetrationB = collFindAxisLeastPenetration(faceB, B, A)
        If penetrationB >= 0 Then Return

        '// Determine which shape contains reference face
        If (BiasGreaterThan(penetrationA, penetrationB)) Then
            RefPoly = A
            IncPoly = B
            referenceIndex = faceA
            flip = False
        Else
            RefPoly = B
            IncPoly = A
            referenceIndex = faceB
            flip = True
        End If

        '  // World space incident face


        collFindIncidentFace incidentFace(), RefPoly, IncPoly, referenceIndex

        ''  //        y
        ''  //        ^  ->n       ^
        ''  //      +---c ------posPlane--
        ''  //  x < | i |\
        ''  //      +---+ c-----negPlane--
        ''  //             \       v
        ''  //              r
        ''  //
        ''  //  r : reference face
        ''  //  i : incident poly
        ''  //  c : clipped point
        ''  //  n : incident normal



        '// Setup reference face vertices
        'Vec2 v1 = RefPoly->m_vertices[referenceIndex];
        'referenceIndex = referenceIndex + 1 == RefPoly->m_vertexCount ? 0 : referenceIndex + 1;
        'Vec2 v2 = RefPoly->m_vertices[referenceIndex];
        '// Transform vertices to world space
        'v1 = RefPoly->u * v1 + RefPoly->body->position;
        'v2 = RefPoly->u * v2 + RefPoly->body->position;
        '// Setup reference face vertices
        V1 = RefPoly.Vertex(referenceIndex)
        referenceIndex = referenceIndex + 1: If referenceIndex > RefPoly.Nvertex Then referenceIndex = 1
        V2 = RefPoly.Vertex(referenceIndex)

        '// Transform vertices to world space
        'v1 = RefPoly->u * v1 + RefPoly->body->position;
        'v2 = RefPoly->u * v2 + RefPoly->body->position;
        V1 = Vec2ADD(matMULv(RefPoly.U, V1), RefPoly.Pos)
        V2 = Vec2ADD(matMULv(RefPoly.U, V2), RefPoly.Pos)

        '  // Calculate reference face side normal in world space
        '  Vec2 sidePlaneNormal = (v2 - v1);
        '  sidePlaneNormal.Normalize( );
        '  // Orthogonalize
        '  Vec2 refFaceNormal( sidePlaneNormal.y, -sidePlaneNormal.x );
        sidePlaneNormal = Vec2Normalize(Vec2SUB(V2, V1))
        refFaceNormal.X = sidePlaneNormal.Y
        refFaceNormal.Y = -sidePlaneNormal.X



        '  // ax + by = c
        '  // c is distance from origin
        '  real refC = Dot( refFaceNormal, v1 );
        '  real negSide = -Dot( sidePlaneNormal, v1 );
        '  real posSide =  Dot( sidePlaneNormal, v2 );
        refC = Vec2DOT(refFaceNormal, V1)
        negSide = -Vec2DOT(sidePlaneNormal, V1)
        posSide = Vec2DOT(sidePlaneNormal, V2)



        '  // Clip2 incident face to reference face side planes
        '  if(Clip2( -sidePlaneNormal, negSide, incidentFace ) < 2)
        '    return; // Due to floating point error, possible to not have required points
        '  if(Clip2(  sidePlaneNormal, posSide, incidentFace ) < 2)
        '    return; // Due to floating point error, possible to not have required points
        '  // Flip
        '  m->normal = flip ? -refFaceNormal : refFaceNormal;


        '************** HERE
        If Clip2(Vec2Negative(sidePlaneNormal), negSide, incidentFace) < 2 Then
            Return
        End If
        If Clip2(sidePlaneNormal, posSide, incidentFace) < 2 Then
            Return
        End If

        If flip Then
            .normal = Vec2Negative(refFaceNormal)
        Else
            .normal = refFaceNormal
        End If





        '// Keep points behind reference face
        '  uint32 cp = 0; // clipped points behind reference face
        '  real separation = Dot( refFaceNormal, incidentFace[0] ) - refC;
        '  if(separation <= 0.0f)
        '  {
        '    m->contacts[cp] = incidentFace[0];
        '    m->penetration = -separation;
        '    ++cp;
        '  }
        '  Else
        '    m->penetration = 0;
        '
        '  separation = Dot( refFaceNormal, incidentFace[1] ) - refC;
        '  if(separation <= 0.0f)
        '  {
        '    m->contacts[cp] = incidentFace[1];
        '
        '    m->penetration += -separation;
        '    ++cp;
        '
        '    // Average penetration
        '    m->penetration /= (real)cp;
        '  }
        '
        '  m->contact_count = cp;



        'ReDim .contactsPTS(2)

        '// Keep points behind reference face
        cp = 0
        separation = Vec2DOT(refFaceNormal, incidentFace(0)) - refC
        If (separation < 0#) Then
            cp = cp + 1
            .contactsPTS(cp) = incidentFace(0)
            .Penetration = -separation
            .contactCount = cp
        Else
            .Penetration = 0
        End If


        'separation = Dot( refFaceNormal, incidentFace[1] ) - refC;
        separation = Vec2DOT(refFaceNormal, incidentFace(1)) - refC
        If (separation < 0) Then
            cp = cp + 1
            .contactsPTS(cp) = incidentFace(1)
            .Penetration = .Penetration - separation
            .Penetration = .Penetration / cp
            .contactCount = cp
        End If

        .contactCount = cp
        If cp Then .invContactCount = 1# / cp

        Return


    End With

End Function









'Real collFindAxisLeastPenetration(uint32 * faceIndex, PolygonShape * A, PolygonShape * B)
'{
'  real bestDistance = -MAX_VALUE;
'  uint32 bestIndex;
'
'  for(uint32 i = 0; i < A->m_vertexCount; ++i)
'  {
'    // Retrieve a face normal from A
'    Vec2 n = A->m_normals[i];
'    Vec2 nw = A->u * n;
'
'    // Transform face normal into B's model space
'    Mat2 buT = B->u.Transpose( );
'    n = buT * nw;
'
'    // Retrieve support point from B along -n
'    Vec2 s = B->GetSupport( -n );
'
'    // Retrieve vertex on face from A, transform into
'    // B's model space
'    Vec2 v = A->m_vertices[i];
'    v = A->u * v + A->body->position;
'    v -= B->body->position;
'    v = buT * v;
'
'    // Compute penetration distance (in B's model space)
'    real d = Dot( n, s - v );
'
'    // Store greatest distance
'    if(d > bestDistance)
'    {
'      bestDistance = d;
'      bestIndex = i;
'    }
'  }
Private Function collFindAxisLeastPenetration(faceIndex As Long, A As tBody, B As tBody) As Double
    Dim I       As Long

    Dim bestIndex As Long
    Dim bestDistance As Double


    Dim N       As tVec2
    Dim nw      As tVec2
    Dim buT     As tMAT2
    Dim S       As tVec2
    Dim V       As tVec2
    Dim D       As Double


    bestDistance = -MAX_VALUE

    For I = 1 To A.Nvertex

        '// Retrieve a face normal from A
        ' Vec2 n = A->m_normals[i];
        ' nw = A->u * n;
        N = A.normals(I)
        nw = matMULv(A.U, N)

        '// Transform face normal into B    's model space
        'Mat2 buT = B->u.Transpose( );
        'n = buT * nw;
        buT = matTranspose(B.U)
        N = matMULv(buT, nw)


        '// Retrieve support point from B along -n
        'Vec2 s = B->GetSupport( -n );
        S = CollGetSupport(B, Vec2Negative(N))

        '// Retrieve vertex on face from A, transform into
        '// B    's model space
        'Vec2 v = A->m_vertices[i];
        'v = A->u * v + A->body->position;
        'v -= B->body->position;
        'v = buT * v;
        V = A.Vertex(I)
        V = Vec2ADD(matMULv(A.U, V), A.Pos)
        V = Vec2SUB(V, B.Pos)
        V = matMULv(buT, V)


        '// Compute penetration distance (in B    's model space)
        'real d = Dot( n, s - v );
        D = Vec2DOT(N, Vec2SUB(S, V))

        '// Store greatest distance
        If (D > bestDistance) Then
            bestDistance = D
            bestIndex = I
        End If
    Next

    collFindAxisLeastPenetration = bestDistance
    faceIndex = bestIndex

End Function


'void collFindIncidentFace( Vec2 *v, PolygonShape *RefPoly, PolygonShape *IncPoly, uint32 referenceIndex )
'{
'  Vec2 referenceNormal = RefPoly->m_normals[referenceIndex];
'
'  // Calculate normal in incident's frame of reference
'  referenceNormal = RefPoly->u * referenceNormal; // To world space
'  referenceNormal = IncPoly->u.Transpose( ) * referenceNormal; // To incident's model space
'
'  // Find most anti-normal face on incident polygon
'  int32 incidentFace = 0;
'  real minDot = MAX_VALUE;
'  for(uint32 i = 0; i < IncPoly->m_vertexCount; ++i)
'  {
'    real dot = Dot( referenceNormal, IncPoly->m_normals[i] );
'    if(dot < minDot)
'    {
'      minDot = dot;
'      incidentFace = i;
'    }
'  }
'
'  // Assign face vertices for incidentFace
'  v[0] = IncPoly->u * IncPoly->m_vertices[incidentFace] + IncPoly->body->position;
'  incidentFace = incidentFace + 1 >= (int32)IncPoly->m_vertexCount ? 0 : incidentFace + 1;
'  v[1] = IncPoly->u * IncPoly->m_vertices[incidentFace] + IncPoly->body->position;
'}
Private Sub collFindIncidentFace(V() As tVec2, RefPoly As tBody, IncPoly As tBody, referenceIndex As Long)

    Dim referenceNormal As tVec2
    Dim incidentFace As Long
    Dim I       As Long
    Dim dot     As Double
    Dim minDot  As Double

    'Vec2 referenceNormal = RefPoly->m_normals[referenceIndex];
    referenceNormal = RefPoly.normals(referenceIndex)

    '// Calculate normal in incident's frame of reference
    'referenceNormal = RefPoly->u * referenceNormal; // To world space
    'referenceNormal = IncPoly->u.Transpose( ) * referenceNormal; // To incident's model space
    referenceNormal = matMULv(RefPoly.U, referenceNormal)
    referenceNormal = matMULv(matTranspose(IncPoly.U), referenceNormal)


    '// Find most anti-normal face on incident polygon
    '  for(uint32 i = 0; i < IncPoly->m_vertexCount; ++i)
    '  {
    '    real dot = Dot( referenceNormal, IncPoly->m_normals[i] );
    '    if(dot < minDot)
    '    {
    '      minDot = dot;
    '      incidentFace = i;
    '    }
    '  }
    incidentFace = 0
    minDot = MAX_VALUE
    For I = 1 To IncPoly.Nvertex
        dot = Vec2DOT(referenceNormal, IncPoly.normals(I))
        If (dot < minDot) Then
            minDot = dot
            incidentFace = I
        End If

    Next

    '// Assign face vertices for incidentFace
    'v[0] = IncPoly->u * IncPoly->m_vertices[incidentFace] + IncPoly->body->position;
    'incidentFace = incidentFace + 1 >= (int32)IncPoly->m_vertexCount ? 0 : incidentFace + 1;
    'v[1] = IncPoly->u * IncPoly->m_vertices[incidentFace] + IncPoly->body->position;

    V(0) = Vec2ADD(matMULv(IncPoly.U, IncPoly.Vertex(incidentFace)), IncPoly.Pos)
    incidentFace = incidentFace + 1: If incidentFace > IncPoly.Nvertex Then incidentFace = 1
    V(1) = Vec2ADD(matMULv(IncPoly.U, IncPoly.Vertex(incidentFace)), IncPoly.Pos)


End Sub



'int32 Clip2( Vec2 n, real c, Vec2 *face )
'{
'  uint32 sp = 0;
'  Vec2 out[2] = {
'    face[0],
'    face [1]
'  };
'
'  // Retrieve distances from each endpoint to the line
'  // d = ax + by - c
'  real d1 = Dot( n, face[0] ) - c;
'  real d2 = Dot( n, face[1] ) - c;
'
'  // If negative (behind plane) Clip2
'  if(d1 <= 0.0f) out[sp++] = face[0];
'  if(d2 <= 0.0f) out[sp++] = face[1];
'
'  // If the points are on different sides of the plane
'  if(d1 * d2 < 0.0f) // less than to ignore -0.0f
'  {
'    // Push interesection point
'    real alpha = d1 / (d1 - d2);
'    out[sp] = face[0] + alpha * (face[1] - face[0]);
'    ++sp;
'  }
'
'  // Assign our new converted values
'  face[0] = out[0];
'  face[1] = out[1];
'
'  assert( sp != 3 );
'
'  return sp;
'}



Private Function Clip2(N As tVec2, c As Double, face() As tVec2) As Long
    Dim sp      As Long
    Dim out(0 To 1) As tVec2
    Dim d1      As Double
    Dim d2      As Double
    Dim alpha   As Double

    '  uint32 sp = 0;
    '  Vec2 out[2] = {
    '    face[0],
    '    face [1]
    '  };
    out(0) = face(0)
    out(1) = face(1)

    ' // Retrieve distances from each endpoint to the line
    ' // d = ax + by - c
    ' real d1 = Dot( n, face[0] ) - c;
    ' real d2 = Dot( n, face[1] ) - c;
    d1 = Vec2DOT(N, face(0)) - c
    d2 = Vec2DOT(N, face(1)) - c


    '// If negative (behind plane) Clip2
    'if(d1 <= 0.0f) out[sp++] = face[0];
    'if(d2 <= 0.0f) out[sp++] = face[1];
    If d1 <= 0 Then out(sp) = face(0): sp = sp + 1
    If d2 <= 0 Then out(sp) = face(1): sp = sp + 1


    '// If the points are on different sides of the plane
    'if(d1 * d2 < 0.0f) // less than to ignore -0.0f
    '{
    '  // Push interesection point
    '  real alpha = d1 / (d1 - d2);
    '  out[sp] = face[0] + alpha * (face[1] - face[0]);
    '  ++sp;
    '}
    If d1 * d2 < 0 Then

        alpha = d1 / (d1 - d2)
        out(sp) = Vec2ADD(face(0), Vec2MUL(Vec2SUB(face(1), face(0)), alpha))
        sp = sp + 1
    End If

    '// Assign our new converted values
    face(0) = out(0)
    face(1) = out(1)

    '    assert( sp != 3 );
    ''    If sp = 3 Then MsgBox "sp=3"
    Clip2 = sp
End Function




















'*************************************************************************
'************************* B O D Y  *****************************
'*************************************************************************




Private Sub BodyCalcCentroid(wB As Long)


'// Calculate centroid and moment of inertia

    Dim triangleArea As Double


    Const k_inv3 As Double = 1 / 3

    Dim I       As Long
    Dim J       As Long

    Dim P1      As tVec2
    Dim p2      As tVec2
    Dim D       As Double

    Dim weight  As Double

    Dim intx2   As Double
    Dim inty2   As Double

    With Body(wB)

        .ii = 0
        .COM.X = 0
        .COM.Y = 0
        .Area = 0

        For I = 1 To .Nvertex
            J = I + 1
            If J > .Nvertex Then J = 1
            '            // Triangle vertices, third vertex implied as (0, 0)
            P1 = .Vertex(I)
            p2 = .Vertex(J)


            D = Vec2CROSS(P1, p2)
            triangleArea = 0.5 * D

            .Area = .Area + triangleArea

            '            // Use area to weight the centroid average, not just vertex position
            weight = triangleArea * k_inv3

            '            com.addsi( p1, weight );
            '            com.addsi( p2, weight );

            .COM = Vec2ADD(.COM, Vec2MUL(P1, weight))
            .COM = Vec2ADD(.COM, Vec2MUL(p2, weight))

            P1 = Vec2SUB(P1, .Pos)
            p2 = Vec2SUB(p2, .Pos)

            intx2 = P1.X * P1.X + p2.X * P1.X + p2.X * p2.X
            inty2 = P1.Y * P1.Y + p2.Y * P1.Y + p2.Y * p2.Y

            .ii = .ii + (0.25 * k_inv3 * D) * (intx2 + inty2)

        Next
        'com.muli( 1.0f / area );
        .COM = Vec2MUL(.COM, 1 / .Area)

        '       .ii = .ii / .Area ^ 0.75 '-----------<<<<<<<<<<<<<<<< main missing line causing polygons not to rotate .... But in original source there isnt !!!???
    End With

End Sub


Private Sub BodyComputeMass(wB As Long, Density As Double)
    Dim I       As Long

    With Body(wB)

        If .myType = eCircle Then
            .mass = PI * .Radius * .Radius * Density
            .invMass = IIf(.mass <> 0#, 1# / .mass, 0)
            .inertia = .mass * .Radius * .Radius
            .invInertia = IIf(.inertia <> 0#, 1# / .inertia, 0#)
        End If

        If .myType = ePolygon Then

            BodyCalcCentroid wB

            ' Translate vertices to centroid (make the centroid (0, 0)
            ' for the polygon in model space)
            ' Not really necessary, but I like doing this anyway
            For I = 1 To .Nvertex
                .Vertex(I) = Vec2SUB(.Vertex(I), .COM)
            Next

            .mass = Density * .Area
            .invMass = IIf(.mass <> 0#, 1# / .mass, 0)
            .inertia = .ii * Density
            .invInertia = IIf(.inertia <> 0#, 1# / .inertia, 0#)
        End If


    End With

End Sub




Friend Sub BodyApplyForce(wB As Long, F As tVec2)
    With Body(wB)
        .FORCE = Vec2ADD(.FORCE, F)
    End With
End Sub

Friend Sub BodyApplyImpulse(wB As Long, impulse As tVec2, contactVector As tVec2)
'      velocity.addsi( impulse, invMass );
'      angularVelocity += invInertia * Vec2.cross( contactVector, impulse );
    With Body(wB)
        .VEL = Vec2ADD(.VEL, Vec2MUL(impulse, .invMass))
        .angularVelocity = .angularVelocity + Vec2CROSS(contactVector, impulse) * .invInertia
    End With
End Sub

Public Sub BodySetStatic(wB As Long)
    With Body(wB)
        .inertia = MAX_VALUE
        .invInertia = 0
        .mass = 0
        .invMass = 0
    End With

End Sub

Public Sub BodySetGroup(wB As Long, G As Long)
    Body(wB).CollisionGroup = G
    If BiggerGroup < G Then BiggerGroup = G
End Sub
Public Sub BodySetCollideWith(wB As Long, M As Long)
    Body(wB).CollideWith = M
End Sub


Public Sub BodySetFriction(wB As Long, StaticF As Double, DynamicF As Double)
    Body(wB).staticFriction = StaticF
    Body(wB).dynamicFriction = DynamicF
End Sub
Public Sub BodySetRestitution(wB As Long, Restitution As Double)
    Body(wB).Restitution = Restitution
End Sub

Private Sub BodyPOLYGONComputeFaceNormals(wB As Long)
'
'        // Compute face normals
'        for (int i = 0; i < vertexCount; ++i)
'        {
'            Vec2 face = vertices[(i + 1) % vertexCount].sub( vertices[i] );
'
'            // Calculate normal with 2D cross product between vector and scalar
'            normals[i].set( face.y, -face.x );
'            normals[i].normalize();
'        }

    Dim I       As Long
    Dim J       As Long
    Dim face    As tVec2
    Dim N       As tVec2

    If Body(wB).myType <> ePolygon Then Exit Sub

    With Body(wB)
        ReDim .normals(.Nvertex)
        For I = 1 To .Nvertex
            J = I + 1
            If J > .Nvertex Then J = 1
            face = Vec2SUB(.Vertex(J), .Vertex(I))
            N.X = face.Y
            N.Y = -face.X
            .normals(I) = Vec2Normalize(N)
        Next
    End With


End Sub
Friend Sub BodyCREATECircle(Pos As tVec2, r As Double, Optional Density As Double = DefDensity)
    Dim B       As Long

    NBodies = NBodies + 1
    ReDim Preserve Body(NBodies)
    With Body(NBodies)
        .myType = eCircle
        .Pos = Pos
        .Radius = r
        .staticFriction = GlobalSTATICFRICTION   ' 0.15    '0.3    '0.5
        .dynamicFriction = GlobalDYNAMICFRICTION    ' 0.5 '0.07    ' 0.1    '0.3
        .Restitution = GlobalRestitution
        .orient = rndFT(-PI, PI)
        .color = RGB(100 + Rnd * 155, 100 + Rnd * 155, 100 + Rnd * 155)
        .U = SetOrient(0)
        .VEL = Vec2(0, 0)
        .angularVelocity = 0
    End With


    BodyComputeMass NBodies, Density


        B = Int(Rnd * 4)
        Select Case B
            Case 0
                BodySETSprite NBodies, 0, 0, r * 2, r * 2, App.Path & "\Res\" & "RainbowBall.png"
            Case 1
                BodySETSprite NBodies, 0, 0, r * 2, r * 2, App.Path & "\Res\" & "Ball_Yellow.svg"
            Case 2
                BodySETSprite NBodies, 0, 0, r * 2.1, r * 2.1, App.Path & "\Res\" & "Ball.png"
            Case 3
                BodySETSprite NBodies, 0, 0, r * 2, r * 2, App.Path & "\Res\" & "SoccerBall.png"
        End Select


    '    CopyBodyToShape
End Sub

Friend Sub BodyCREATERandomPoly(Pos As tVec2, Optional Density As Double = DefDensity)
    Dim I       As Long


    NBodies = NBodies + 1
    ReDim Preserve Body(NBodies)
    With Body(NBodies)
        .myType = ePolygon
        .Pos = Pos
        .staticFriction = GlobalSTATICFRICTION   '0.5
        .dynamicFriction = GlobalDYNAMICFRICTION    '0.3
        .Restitution = GlobalRestitution

        .orient = rndFT(-PI, PI)

        .color = RGB(100 + Rnd * 155, 100 + Rnd * 155, 100 + Rnd * 155)

        .Nvertex = 4    '+ Rnd * 2

        ReDim .Vertex(.Nvertex)
        ReDim .tVertex(.Nvertex)

        '        For I = 1 To .Nvertex
        '            '        For I = .Nvertex To 1 Step -1
        '            .Vertex(I) = Vec2ADD(Pos, _
                     '                                 Vec2((10 + Rnd * 30) * Cos(PI2 * (I - 1) / .Nvertex), _
                     '                                      (10 + Rnd * 30) * Sin(PI2 * (I - 1) / .Nvertex)))
        '        Next

        .Vertex(1) = Vec2(Pos.X - 20, Pos.Y - 15)
        .Vertex(2) = Vec2(Pos.X + 40, Pos.Y - 15)
        .Vertex(3) = Vec2(Pos.X + 40, Pos.Y + 15)
        .Vertex(4) = Vec2(Pos.X - 20, Pos.Y + 15)
    End With


    BodyPOLYGONComputeFaceNormals NBodies
    BodyComputeMass NBodies, Density
    Body(NBodies).Pos = Body(NBodies).COM



    '''    CopyBodyToShape

End Sub


Friend Sub BodyCREATEBox(Pos As tVec2, w As Double, h As Double, Optional Ang As Double = 0, Optional DoChamfer As Boolean = False, Optional Density As Double = DefDensity)
    Dim I         As Long



    NBodies = NBodies + 1
    ReDim Preserve Body(NBodies)
    With Body(NBodies)
        .myType = ePolygon
        .Pos = Pos
        .staticFriction = GlobalSTATICFRICTION   '0.5
        .dynamicFriction = GlobalDYNAMICFRICTION    '0.3
        .Restitution = GlobalRestitution

        .orient = Ang
        .color = RGB(100 + Rnd * 155, 100 + Rnd * 155, 100 + Rnd * 155)
        .Nvertex = 4                             '+ Rnd * 2

        .U = SetOrient(Ang)
        .VEL = Vec2(0, 0)
        .angularVelocity = 0

        ReDim .Vertex(.Nvertex)
        ReDim .tVertex(.Nvertex)

        .Vertex(1) = Vec2(Pos.X - w * 0.5, Pos.Y - h * 0.5)
        .Vertex(2) = Vec2(Pos.X + w * 0.5, Pos.Y - h * 0.5)
        .Vertex(3) = Vec2(Pos.X + w * 0.5, Pos.Y + h * 0.5)
        .Vertex(4) = Vec2(Pos.X - w * 0.5, Pos.Y + h * 0.5)


    End With

    BodyPOLYGONComputeFaceNormals NBodies
    BodyComputeMass NBodies, Density
    Body(NBodies).Pos = Body(NBodies).COM


    '    If Rnd < 0.5 Then BodyChamfer NBodies, 4 + Rnd * 5
    If DoChamfer Then
        BodyChamfer NBodies, 4 + Rnd * 5
    Else
        BodySETSprite NBodies, 0, 0, w, h, App.Path & "\Res\" & "Box.png"
    End If


    '  CopyBodyToShape

End Sub


Friend Sub BodyCREATERegularPoly(Pos As tVec2, Rw As Double, Rh As Double, N As Long, Flat As Long, Optional Density As Double = DefDensity)
    Dim I       As Long
    Dim A       As Double

    NBodies = NBodies + 1
    ReDim Preserve Body(NBodies)
    With Body(NBodies)
        .myType = ePolygon
        .Pos = Pos
        .staticFriction = GlobalSTATICFRICTION   '0.5
        .dynamicFriction = GlobalDYNAMICFRICTION    '0.3
        .Restitution = GlobalRestitution

        .orient = 0
        .color = RGB(100 + Rnd * 155, 100 + Rnd * 155, 100 + Rnd * 155)

        .U = SetOrient(.orient)
        .VEL = Vec2(0, 0)
        .angularVelocity = 0

        .Nvertex = N

        ReDim .Vertex(.Nvertex)
        ReDim .tVertex(.Nvertex)

        If Flat Then A = 0.5 * PI2 / N

        For I = 1 To .Nvertex
            '        For I = .Nvertex To 1 Step -1
            .Vertex(I) = Vec2ADD(Pos, _
                                 Vec2((Rw) * Cos(A + PI2 * (I - 1) / .Nvertex), _
                                      (Rh) * Sin(A + PI2 * (I - 1) / .Nvertex)))
        Next

    End With


    BodyPOLYGONComputeFaceNormals NBodies
    BodyComputeMass NBodies, Density
    Body(NBodies).Pos = Body(NBodies).COM

    '   If Rnd < 0.5 Then BodyChamfer NBodies, 9
    BodyChamfer NBodies, 8 + Rnd * 8, , 4


    '    CopyBodyToShape


End Sub


Friend Sub BodyCREATEPolygon(Vertex() As tVec2)

    Dim I       As Long



    NBodies = NBodies + 1
    ReDim Preserve Body(NBodies)
    With Body(NBodies)
        .myType = ePolygon

        .staticFriction = GlobalSTATICFRICTION   '0.5
        .dynamicFriction = GlobalDYNAMICFRICTION    '0.3
        .Restitution = GlobalRestitution

        .orient = 0
        .color = RGB(100 + Rnd * 155, 100 + Rnd * 155, 100 + Rnd * 155)

        .U = SetOrient(.orient)
        .VEL = Vec2(0, 0)
        .angularVelocity = 0

        .Nvertex = UBound(Vertex)

        ReDim .Vertex(.Nvertex)
        ReDim .tVertex(.Nvertex)

        .Vertex = Vertex


        '------------------------------------ IMPORTANT
        .Pos = Vec2(0, 0)
        For I = 1 To .Nvertex
            .Pos = Vec2ADD(.Pos, .Vertex(I))
        Next
        .Pos = Vec2MUL(.Pos, 1 / .Nvertex)

    End With


    BodyPOLYGONComputeFaceNormals NBodies
    BodyComputeMass NBodies, DefDensity
    Body(NBodies).Pos = Body(NBodies).COM

    ''    '   If Rnd < 0.5 Then BodyChamfer NBodies, 9
    ''    BodyChamfer NBodies, 8 + Rnd * 8, , 4


    '    CopyBodyToShape



End Sub

'Private Sub CopyBodyToShape()
'    Dim I   As Long
'
'    ReDim Preserve Shape(NBodies)
'    With Shape(NBodies)
'        .myType = Body(NBodies).myType
'        .Pos = Body(NBodies).Pos
'        .Nvertex = Body(NBodies).Nvertex
'        .Vertex = Body(NBodies).Vertex
'        .Radius = Body(NBodies).Radius
'        .color = Body(NBodies).color
'        .orient = Body(NBodies).orient
'        ReDim .tVertex(.Nvertex)
'        ReDim .LToCT(.Nvertex)
'
'        For I = 1 To .Nvertex
'            .LToCT(I) = 1
'        Next
'
'    End With


'End Sub




Public Sub BodySETSprite(wB As Long, _
                         cx As Double, cy As Double, _
                         w As Double, h As Double, FilePath As String)


    With Body(wB)
        .ImageExist = True
        .ImageCenter = Vec2(cx, cy)
        .ImageSize = Vec2(w, h)

    End With

    ImageList.AddImage CStr(wB), FilePath, w, h


End Sub






'  // The extreme point along a direction within a polygon
'  Vec2 CollGetSupport( const Vec2& dir )
'  {
'    real bestProjection = -MAX_VALUE;
'    Vec2 bestVertex;
'
'    for(uint32 i = 0; i < m_vertexCount; ++i)
'    {
'      Vec2 v = m_vertices[i];
'      real projection = Dot( v, dir );
'
'      if(projection > bestProjection)
'      {
'        bestVertex = v;
'        bestProjection = projection;
'      }
'    }
'
'    return bestVertex;
'  }
Private Function CollGetSupport(Body As tBody, dire As tVec2) As tVec2
'// The extreme point along a direction within a polygon
    Dim bestProjection As Double
    Dim bestVertex As tVec2
    Dim V       As tVec2
    Dim I       As Long
    Dim projection As Double

    bestProjection = -MAX_VALUE
    For I = 1 To Body.Nvertex
        V = Body.Vertex(I)
        projection = Vec2DOT(V, dire)

        If (projection > bestProjection) Then
            bestVertex = V
            bestProjection = projection
        End If
    Next

    CollGetSupport = bestVertex

End Function


'''
'''    /**
'''     * Chamfers a set of vertices by giving them rounded corners, returns a new set of vertices.
'''     * The radius parameter is a double number or an array to specify the radius for each vertex.
'''     * @method BodyChamfer
'''     * @param {vertices} vertices
'''     * @param {number[]} radius
'''     * @param {number} quality
'''     * @param {number} qualityMin
'''     * @param {number} qualityMax
'''     */
'''    Vertices.BodyChamfer = function(vertices, radius, quality, qualityMin, qualityMax) {
'''        radius = radius || [8];
'''
'''        if (!radius.length)
'''            radius = [radius];
'''
'''        // quality defaults to -1, which is auto
'''        quality = (typeof quality !== 'undefined') ? quality : -1;
'''        qualityMin = qualityMin || 2;
'''        qualityMax = qualityMax || 14;
'''
'''        var newVertices = [];
'''
'''        for (var i = 0; i < vertices.length; i++) {
'''            var prevVertex = vertices[i - 1 >= 0 ? i - 1 : vertices.length - 1],
'''                vertex = vertices[i],
'''                nextVertex = vertices[(i + 1) % vertices.length],
'''                currentRadius = radius[i < radius.length ? i : radius.length - 1];
'''
'''            if (currentRadius === 0) {
'''                newVertices.push(vertex);
'''                continue;
'''            }
'''
'''            var prevNormal = Vector.normalise({
'''                x: vertex.y - prevVertex.y,
'''y:                 prevVertex.x -Vertex.x
'''            });
'''
'''            var nextNormal = Vector.normalise({
'''                x: nextVertex.y - vertex.y,
'''y:                 Vertex.x -nextVertex.x
'''            });
'''
'''            var diagonalRadius = Math.sqrt(2 * Math.pow(currentRadius, 2)),
'''                radiusVector = Vector.mult(Common.clone(prevNormal), currentRadius),
'''                midNormal = Vector.normalise(Vector.mult(Vector.add(prevNormal, nextNormal), 0.5)),
'''                scaledVertex = Vector.sub(vertex, Vector.mult(midNormal, diagonalRadius));
'''
'''            var precision = quality;
'''
'''            if (quality === -1) {
'''                // automatically decide precision
'''                precision = Math.pow(currentRadius, 0.32) * 1.75;
'''            }
'''
'''            precision = Common.clamp(precision, qualityMin, qualityMax);
'''
'''            // use an even value for precision, more likely to reduce axes by using symmetry
'''            if (precision % 2 === 1)
'''                precision += 1;
'''
'''            var alpha = Math.acos(Vector.dot(prevNormal, nextNormal)),
'''                theta = alpha / precision;
'''
'''            for (var j = 0; j < precision; j++) {
'''                newVertices.push(Vector.add(Vector.rotate(radiusVector, theta * j), scaledVertex));
'''            }
'''        }
'''
'''        return newVertices;
'''    };

'
'    /**
'     * Chamfers a set of vertices by giving them rounded corners, returns a new set of vertices.
'     * The radius parameter is a double number or an array to specify the radius for each vertex.
'     * @method BodyChamfer
'     * @param {vertices} vertices
'     * @param {number[]} radius
'     * @param {number} quality
'     * @param {number} qualityMin
'     * @param {number} qualityMax
'     */
Public Sub BodyChamfer(wB As Long, Radius As Double, _
                       Optional Quality As Double = -1, Optional QualityMin As Double = 2, Optional QualityMax As Double = 14)
'        radius = radius || [8];
'
'        if (!radius.length)
'            radius = [radius];

'// quality defaults to -1, which is auto
'quality = (typeof quality !== 'undefined') ? quality : -1;
'qualityMin = qualityMin || 2;
'qualityMax = qualityMax || 14;
    Dim prevVertex As tVec2
    Dim nextVertex As tVec2
    Dim Vertex  As tVec2
    Dim newVertex() As tVec2
    Dim newVertexN As Long

    Dim I       As Long

    Dim Precision As Double
    Dim currentRadius As Double
    Dim prevNormal As tVec2
    Dim nextNormal As tVec2
    Dim diagonalRadius As Double
    Dim radiusVector As tVec2
    Dim midNormal As tVec2
    Dim scaledVertex As tVec2
    Dim alpha   As Double
    Dim theta   As Double
    Dim J       As Long


    With Body(wB)





        For I = 1 To .Nvertex
            Vertex = .Vertex(I)
            If I = .Nvertex Then nextVertex = .Vertex(1) Else: nextVertex = .Vertex(I + 1)
            If Vec2Length(Vec2SUB(Vertex, nextVertex)) <= Radius * 2 Then Exit Sub
        Next

        For I = 1 To .Nvertex
            Vertex = .Vertex(I)
            If I = 1 Then prevVertex = .Vertex(.Nvertex) Else: prevVertex = .Vertex(I - 1)
            If I = .Nvertex Then nextVertex = .Vertex(1) Else: nextVertex = .Vertex(I + 1)




            '        for (var i = 0; i < vertices.length; i++) {
            '            var prevVertex = vertices[i - 1 >= 0 ? i - 1 : vertices.length - 1],
            '                vertex = vertices[i],
            '                nextVertex = vertices[(i + 1) % vertices.length],
            '                currentRadius = radius[i < radius.length ? i : radius.length - 1];
            currentRadius = Radius

            '            if (currentRadius === 0) {
            '                newVertices.push(vertex);
            '                continue;
            '            }

            '            var prevNormal = Vector.normalise({
            '                x: vertex.y - prevVertex.y,
            'y:                 prevVertex.x -Vertex.x
            '            });
            '
            '            var nextNormal = Vector.normalise({
            '                x: nextVertex.y - vertex.y,
            'y:                 Vertex.x -nextVertex.x
            '            });
            prevNormal = Vec2Normalize(Vec2(Vertex.Y - prevVertex.Y, prevVertex.X - Vertex.X))
            nextNormal = Vec2Normalize(Vec2(nextVertex.Y - Vertex.Y, Vertex.X - nextVertex.X))

            '            var diagonalRadius = Math.sqrt(2 * Math.pow(currentRadius, 2)),
            '                radiusVector = Vector.mult(Common.clone(prevNormal), currentRadius),
            '                midNormal = Vector.normalise(Vector.mult(Vector.add(prevNormal, nextNormal), 0.5)),
            '                scaledVertex = Vector.sub(vertex, Vector.mult(midNormal, diagonalRadius));
            diagonalRadius = Sqr(2 * (currentRadius ^ 2))
            radiusVector = Vec2MUL(prevNormal, currentRadius)
            midNormal = Vec2Normalize(Vec2MUL(Vec2ADD(prevNormal, nextNormal), 0.5))
            scaledVertex = Vec2SUB(Vertex, Vec2MUL(midNormal, diagonalRadius))

            Precision = Quality
            '            var precision = quality;

            '            if (quality === -1) {
            '                // automatically decide precision
            '                precision = Math.pow(currentRadius, 0.32) * 1.75;
            '            }
            If Quality = -1 Then Precision = (currentRadius ^ 0.32) * 1.75



            '            precision = Common.clamp(precision, qualityMin, qualityMax);
            Precision = Clamp(Precision, QualityMin, QualityMax)

            '            // use an even value for precision, more likely to reduce axes by using symmetry
            '            if (precision % 2 === 1)
            '                precision += 1;
            Precision = Round(Precision)
            If Precision Mod 2 = 0 Then Precision = Precision + 1

            '            var alpha = Math.acos(Vector.dot(prevNormal, nextNormal)),
            '                theta = alpha / precision;
            alpha = ACOS(Vec2DOT(prevNormal, nextNormal))
            theta = alpha / Precision

            '  for (var j = 0; j < precision; j++) {
            '  newVertices.push(Vector.add(Vector.rotate(radiusVector, theta * j), scaledVertex));
            '  }

            For J = 0 To Precision
                newVertexN = newVertexN + 1
                ReDim Preserve newVertex(newVertexN)
                newVertex(newVertexN) = Vec2ADD(Vec2Rotate(radiusVector, theta * J), scaledVertex)
            Next


        Next


        .Vertex = newVertex
        .Nvertex = newVertexN
        ReDim .tVertex(newVertexN)

    End With

    BodyPOLYGONComputeFaceNormals wB
    '     BodyComputeMass wB, DefDensity  '''' Causes error to much inertia!!!
    '    Body(wB).Pos = Body(NBodies).COM


End Sub



















'*************************************************************************
'************************* E N G I N E *****************************
'*************************************************************************

Public Property Get NofBodies() As Long
    NofBodies = NBodies
End Property

Public Property Let NofBodies(V As Long)
    NBodies = V
    BiggerGroup = 0

End Property


Public Sub EngineINIT(PIC As PictureBox)

    InitMATH
    RenderINITRC PIC
    MaterialsINIT

    Set ImageList = Cairo.ImageList

    Version = App.Major & "." & App.Minor & "." & App.Revision

    Set FPSTICK = New clsTick

    tComputed = FPSTICK.Add(800)
    tDraw = FPSTICK.Add(30)
    t1Sec = FPSTICK.Add(1)




End Sub

'Public Sub EngineDoSTEP(DeltaTime As double)
Public Sub EngineDoSTEP()

    Dim I       As Long
    Dim J       As Long
    Dim A       As tBody

    Dim MinX    As Double
    Dim MinY    As Double
    Dim MaxX    As Double
    Dim MaxY    As Double

    'DT = DeltaTime
    'RESTING = Vec2LengthSq(Vec2MUL(GRAVITY, DT)) + EPSILON


    Dim tmpContacts As tManifold

    Dim ContactType As Long

    Dim V       As tVec2

    ''   // Generate new collision info
    ''       contacts.clear();
    NofContactMainFolds = 0

    For I = 1 To NBodies - 1
        '        A = Body(I)
        For J = I + 1 To NBodies
            'If (A.invMass <> 0 Or Body(J).invMass <> 0) Then
            If (Body(I).invMass <> 0 Or Body(J).invMass <> 0) Then

                'If I = 4 And J = 5 Then Stop

                If AABBvsAABB(I, J) Then

                    If ((Body(I).CollideWith And Body(J).CollisionGroup)) Then    'Or _
                                                                                  '              (Body(J).CollideWith And Body(I).CollisionGroup)) Then

                        '                Manifold m = new Manifold( A, B );
                        '                m.solve();
                        '
                        '                if (m.contactCount > 0)
                        '                {
                        '                contacts.add( m );
                        '                }
                        '                    If AABBvsAABB(I, J) Then
                        tmpContacts = collCollisionSOLVE(I, J)




                        If tmpContacts.contactCount > 0 Then
                            NofContactMainFolds = NofContactMainFolds + 1
                            If NofContactMainFolds > MAXNofContactMainFolds Then



                                MAXNofContactMainFolds = NofContactMainFolds + 20
                                ReDim Preserve Contacts(MAXNofContactMainFolds)
                            End If

                            Contacts(NofContactMainFolds) = tmpContacts


                        End If

                    End If


                End If
            End If

        Next
    Next



    '// Integrate forces
    'for (int i = 0; i < bodies.size(); ++i)
    '{
    'EngineIntegrateForces( bodies.get( i ), dt );
    '}

    ''    For I = 1 To NBodies'''<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    ''        EngineIntegrateForces I    ', DT
    ''    Next


    ''// Initialize collision
    'for (int i = 0; i < contacts.size(); ++i)
    '{
    'contacts.get( i ).initialize();
    '}

    For I = 1 To NofContactMainFolds
        collContactsINIT I
    Next


    '// Solve collisions
    'for (int j = 0; j < iterations; ++j)
    '{
    'for (int i = 0; i < contacts.size(); ++i)
    '{
    'contacts.get( i ).applyImpulse();
    '}
    '}

    For J = 1 To Iterations
        For I = 1 To NofContactMainFolds
            collContactsApplyImpulse I
        Next
    Next

    JointsResolve    'Ver 1.0.147

    '// Integrate velocities
    'for (int i = 0; i < bodies.size(); ++i)
    '{
    'EngineIntegrateVelocity( bodies.get( i ), dt );
    '}
    For I = 1 To NBodies
        EngineIntegrateVelocity I    ', DT
    Next

    '// Correct positions
    'for (int i = 0; i < contacts.size(); ++i)
    '{
    'contacts.get( i ).positionalCorrection();
    '}

    For I = 1 To NofContactMainFolds
        collContactsPositionalCorrection I
    Next


    '// Clear all forces
    'for (int i = 0; i < bodies.size(); ++i)
    '{
    'Body b = bodies.get( i );
    'b.force.set( 0, 0 );
    'b.torque = 0;
    '}



    For I = 1 To NBodies
        With Body(I)
            .FORCE.X = 0
            .FORCE.Y = 0
            .torque = 0


            'SET AABB
            If .myType = eCircle Then

                .AABB.pMin = Vec2ADD(.Pos, Vec2(-.Radius, -.Radius))
                .AABB.pMax = Vec2ADD(.Pos, Vec2(.Radius, .Radius))

            Else

                MinX = MAX_VALUE
                MinY = MAX_VALUE
                MaxX = -MAX_VALUE
                MaxY = -MAX_VALUE

                For J = 1 To .Nvertex
                    V = matMULv(.U, .Vertex(J))
                    If V.X < MinX Then MinX = V.X
                    If V.Y < MinY Then MinY = V.Y
                    If V.X > MaxX Then MaxX = V.X
                    If V.Y > MaxY Then MaxY = V.Y
                    .tVertex(J) = V
                Next

                .AABB.pMin = Vec2ADD(Vec2(MinX, MinY), .Pos)
                .AABB.pMax = Vec2ADD(Vec2(MaxX, MaxY), .Pos)

            End If


            '-------------------------------


        End With


    Next

    '''    UpdateShapes


    ' JointsResolve moved after collContactsApplyImpulse '--- Ver 1.0.147



End Sub
'Private Sub UpdateShapes()
'    Dim I   As Long
'    Dim J   As Long
'    Dim V   As tVec2
'    Dim U   As tMAT2
'
'
'    For I = 1 To NBodies
'        With Shape(I)
'            U = Body(I).U
'            For J = 1 To .Nvertex
'                V = matMULv(U, .Vertex(J))
'                .tVertex(J) = V
'            Next
'            .Pos = Body(I).Pos
'            .orient = Body(I).orient
'        End With
'    Next
'
'
'End Sub
Private Function AABBvsAABB(wA As Long, wB As Long) As Boolean
    Dim ab1     As tAABB
    Dim ab2     As tAABB

    ab1 = Body(wA).AABB
    ab2 = Body(wB).AABB

    If ab1.pMin.Y > ab2.pMax.Y Then Exit Function
    If ab2.pMin.Y > ab1.pMax.Y Then Exit Function
    If ab1.pMin.X > ab2.pMax.X Then Exit Function
    If ab2.pMin.X > ab1.pMax.X Then Exit Function

    AABBvsAABB = True

End Function





Private Sub EngineIntegrateForces(wB As Long)    ', DT As double)
'''    // see http://www.niksula.hut.fi/~hkankaan/Homepages/gravity.html
'''    public void EngineIntegrateForces( Body b, float dt )
'''    {
'''//      if(b->im == 0.0f)
'''//          return;
'''//      b->velocity += (b->force * b->im + gravity) * (dt / 2.0f);
'''//      b->angularVelocity += b->torque * b->iI * (dt / 2.0f);
'''
'''        if (b.invMass == 0.0f)
'''        {
'''            return;
'''        }
'''
'''        float dts = dt * 0.5f;
'''
'''        b.velocity.addsi( b.force, b.invMass * dts );
'''        b.velocity.addsi( ImpulseMath.GRAVITY, dts );
'''        b.angularVelocity += b.torque * b.invInertia * dts;
'''    }
    Dim dts     As Double
    dts = DT * 0.5

    Dim Dx      As Double


    With Body(wB)
        If .invMass <> 0 Then


            .VEL = Vec2ADD(.VEL, Vec2MUL(.FORCE, .invMass * dts))
            '.VEL = Vec2ADD(.VEL, Vec2MUL(GRAVITY, .invMass * dts))
            .VEL = Vec2ADD(.VEL, Vec2MUL(GRAVITY, dts))

            .angularVelocity = .angularVelocity + .torque * .invInertia * dts

            .angularVelocity = .angularVelocity * 0.9999    'Air Resistence
            .VEL = Vec2MUL(.VEL, 0.9999)


            If .Pos.Y > PicH + 500 Then
                Dx = .Pos.X
                While .Pos.X > PicW
                    .Pos.X = .Pos.X - PicW
                Wend
                While .Pos.X < 0
                    .Pos.X = .Pos.X + PicW
                Wend
                .Pos.Y = 0
                .VEL.Y = 0
                .VEL.X = .VEL.X * 0.5
                '.angularVelocity = 0
                EngineRepositionConnected wB, .Pos.X - Dx
            End If




            '    If .Pos.X + .radius > PicW Then
            '
            '    BodyApplyImpulse wB, Vec2(-.VEL.X * 2, 0), Vec2(.Pos.X + .radius, .Pos.Y)
            '    End If


            '            If .Pos.X + .radius > PicW Then
            '                .VEL.X = -.VEL.X * .restitution
            '                .Pos.X = PicW - .radius
            '            End If

            '   If .Pos.Y + .radius > PicH Then
            '                .VEL.Y = -.VEL.Y * .restitution
            '                .Pos.Y = PicH - .radius
            '            End If

        End If

    End With
End Sub

Private Sub EngineIntegrateVelocity(wB As Long)    ', DT As double)
'        if (b.invMass == 0.0f)
'        {
'            return;
'        }
'
'        b.position.addsi( b.velocity, dt );
'        b.orient += b.angularVelocity * dt;
'        b.setOrient( b.orient );
'
'        EngineIntegrateForces( b, dt );
    With Body(wB)
        If .invMass <> 0 Then

            .Pos = Vec2ADD(.Pos, Vec2MUL(.VEL, DT))
            .orient = .orient + .angularVelocity * DT

            'If .myType = ePolygon Then
            .U = SetOrient(.orient)
            'End If

            EngineIntegrateForces wB

        End If

    End With
End Sub





Private Sub EngineRepositionConnected(wB As Long, Dx As Double)
    Dim I       As Long
    Dim J       As Long
    Dim ConnB   As Long

    For J = 1 To NJ
        ConnB = 0
        If Joints(J).bA = wB Then ConnB = Joints(J).bB
        If Joints(J).bB = wB Then ConnB = Joints(J).bA
        If ConnB Then
            Debug.Print wB & "   " & ConnB

            With Body(ConnB)
                .Pos.X = .Pos.X + Dx
                While .Pos.X < 0
                    .Pos.X = .Pos.X + PicW
                Wend
                While .Pos.X > PicW
                    .Pos.X = .Pos.X - PicW
                Wend

                .Pos.Y = 0
                .VEL.Y = 0
                .VEL.X = .VEL.X * 0.5
                '.angularVelocity = 0
            End With
        End If
    Next

End Sub














'**************************************************************
'********************JOINTS *********************************
'**************************************************************

Public Sub JointAddDistanceJ(bA As Long, bB As Long, D As Double, Optional StiffPull As Double = 1, Optional stiffPush As Double = 1)
    NJ = NJ + 1
    ReDim Preserve Joints(NJ)
    With Joints(NJ)
        .JointType = JointDistance
        .bA = bA
        .bB = bB
        .L = D
        .StifPULL = StiffPull
        .StifPUSH = stiffPush
    End With
End Sub

Friend Sub JointAdd2PinsJ(bA As Long, AnchA As tVec2, bB As Long, AnchB As tVec2, D As Double, _
                          Optional StiffPull As Double = 1, Optional stiffPush As Double = 1)
    NJ = NJ + 1
    ReDim Preserve Joints(NJ)
    With Joints(NJ)
        .JointType = Joint2Pins
        .bA = bA
        .bB = bB
        .AnchA = AnchA
        .AnchB = AnchB
        .L = D
        .StifPULL = StiffPull
        .StifPUSH = stiffPush
    End With
End Sub

Friend Sub JointAdd1PinJ(wBody As Long, Anch As tVec2, Optional D As Double = 0, _
                         Optional StiffPull As Double = 1, Optional stiffPush As Double = 1)
    NJ = NJ + 1
    ReDim Preserve Joints(NJ)
    With Joints(NJ)
        .JointType = Joint1Pin
        .bA = wBody
        .bB = 0
        .AnchA = Anch
        .AnchB = Vec2ADD(Body(wBody).Pos, Anch)
        .L = D
        .StifPULL = StiffPull
        .StifPUSH = stiffPush
    End With
End Sub

Friend Sub JoinAddRotorJ(wBody As Long, Leva As tVec2, Speed As Double)
    NJ = NJ + 1
    ReDim Preserve Joints(NJ)
    With Joints(NJ)
        .JointType = Rotor1
        .bA = wBody
        .bB = 0
        .AnchA = Leva
        .L = Speed
    End With
End Sub

Friend Sub JoinAddRotor2J(bA As Long, AnchA As tVec2, bB As Long, AnchB As tVec2, _
                          Optional StiffPull As Double = 1, Optional stiffPush As Double = 1)
    NJ = NJ + 1
    ReDim Preserve Joints(NJ)
    With Joints(NJ)
        .JointType = Rotor2
        .bA = bA
        .bB = bB
        .AnchA = AnchA
        .AnchB = AnchB
        .L = 0
        '        .StifPULL = StiffPull
        '        .StifPUSH = stiffPush
    End With
End Sub

Private Sub JointsResolve()

'By Roberto Mior (aka reexre,miorsoft)

    Dim I       As Long
    Dim J       As Long

    Dim D       As Double
    Dim N       As tVec2
    Dim V       As tVec2

    Dim Center  As tVec2
    Dim FIXa    As tVec2
    Dim FIXb    As tVec2

    Dim vAprj   As tVec2
    Dim vBprj   As tVec2
    Dim RVA     As tVec2
    Dim RVB     As tVec2
    Dim Axe     As tVec2

    Dim Tmass   As Double
    Dim ORI     As tMAT2
    Dim RotDiff As Double


    For I = 1 To NJ

        With Joints(I)
            Select Case .JointType

            Case JointDistance


                N = Vec2SUB(Body(.bA).Pos, Body(.bB).Pos)

                'N = Vec2ADD(N, Vec2MUL(Vec2SUB(Body(.bA).VEL, Body(.bB).VEL), invDT))

                V = Vec2MUL(Vec2SUB(Body(.bA).VEL, Body(.bB).VEL), KDamp * INVdt)
                V = VectorProject(V, N)

                N = Vec2ADD(N, V)

                D = Vec2Length(N)
                N = Vec2MUL(N, 1 / D)
                D = (D - .L) * (Body(.bA).mass + Body(.bB).mass) * KMASS
                If D > 0 Then
                    D = D * .StifPULL
                Else
                    D = D * .StifPUSH
                End If

                '                  Body(.bA).FORCE = Vec2ADD(Body(.bA).FORCE, Vec2MUL(N, -D))
                '                  Body(.bB).FORCE = Vec2ADD(Body(.bB).FORCE, Vec2MUL(N, D))


                'For J = 1 To Iterations
                BodyApplyImpulse .bA, Vec2MUL(N, -D), Vec2(0, 0)
                BodyApplyImpulse .bB, Vec2MUL(N, D), Vec2(0, 0)
                'Next

            Case Joint2Pins



                ORI = SetOrient(Body(.bA).orient + Body(.bA).angularVelocity * KDamp * INVdt)
                '    .tAnchA = matMULv(Body(.bA).U, .AnchA)
                .tAnchA = matMULv(ORI, .AnchA)
                FIXa = Vec2ADD(Body(.bA).Pos, .tAnchA)

                V = Vec2MUL(Body(.bA).VEL, KDamp * INVdt)
                'V = VectorProject(V, .tAnchA) ' -------------------------removed Ver 1.0.147
                FIXa = Vec2ADD(FIXa, V)   '.------------------<<<<<<<


                ORI = SetOrient(Body(.bB).orient + Body(.bB).angularVelocity * KDamp * INVdt)
                '.tAnchB = matMULv(Body(.bB).U, .AnchB)
                .tAnchB = matMULv(ORI, .AnchB)

                FIXb = Vec2ADD(Body(.bB).Pos, .tAnchB)
                V = Vec2MUL(Body(.bB).VEL, KDamp * INVdt)
                'V = VectorProject(V, .tAnchB) ' -------------------------removed Ver 1.0.147
                FIXb = Vec2ADD(FIXb, V)   '.------------------<<<<<<<

                'vAprj = VectorProject(Body(.bA).VEL, .tAnchA)
                'vBprj = VectorProject(Body(.bB).VEL, .tAnchB)
                'Axe = Vec2SUB(.tAnchB, .tAnchB)



                '''''    Axe = Vec2SUB(FIXa, FIXb) ' -------------------------removed Ver 1.0.147


                'V = Vec2MUL(Body(.bA).VEL, KDamp * INVdt)
                'V = Vec2ADD(V, VectorProject(V, Vec2MUL(Body(.bB).VEL, KDamp * INVdt)))
                'V = VectorProject(V, Axe)


                '                                    vAprj = VectorProject(Body(.bA).VEL, Axe)
                '                                    vBprj = VectorProject(Body(.bB).VEL, Axe)
                ''''vAprj = VectorProject(Vec2MUL(Body(.bA).VEL, 1 * 1), Axe)'---TOLTO
                ''''vBprj = VectorProject(Vec2MUL(Body(.bB).VEL, 1 * 1), Axe)'---TOLTO

                N = Vec2SUB(FIXa, FIXb)
                ''''N = Vec2ADD(N, Vec2SUB(vAprj, vBprj))'---TOLTO

                'N = Vec2ADD(N, V)

                D = Vec2Length(N)
                If D Then N = Vec2MUL(N, 1 / D)
                D = (D - .L) * (Body(.bA).mass + Body(.bB).mass) * KMASS
                If D > 0 Then
                    D = D * .StifPULL
                Else
                    D = D * .StifPUSH
                End If


                BodyApplyImpulse .bA, Vec2MUL(N, -D), .tAnchA
                BodyApplyImpulse .bB, Vec2MUL(N, D), .tAnchB


            Case Joint1Pin


                ORI = SetOrient(Body(.bA).orient + Body(.bA).angularVelocity * KDamp * INVdt)

                '                    .tAnchA = matMULv(Body(.bA).U + Body(.bA).angularVelocity, .AnchA)
                .tAnchA = matMULv(ORI, .AnchA)

                Axe = .tAnchA    ' -------------------------removed Ver 1.0.147


                ''''vAprj = VectorProject(Vec2MUL(Body(.bA).VEL, 1 * 1), Axe)'---TOLTO

                FIXa = Vec2ADD(Body(.bA).Pos, .tAnchA)


                V = Vec2MUL(Body(.bA).VEL, KDamp * INVdt)
                'V = VectorProject(V, .tAnchA)' -------------------------removed Ver 1.0.147

                FIXa = Vec2ADD(FIXa, V)   '.------------------<<<<<<<
                'FIXa = Vec2ADD(Body(.bA).Pos, Vec2ADD(.tAnchA, RVA))
                N = Vec2SUB(FIXa, .AnchB)

                ''''N = Vec2ADD(N, Vec2MUL(vAprj, 1))'---TOLTO


                D = Vec2Length(N)

                'If D Then N = Vec2MUL(N, 1 / D) '*******************************

                D = (D - .L) * (Body(.bA).mass) * KMASS    ' * 2

                If D > 0 Then
                    D = D * .StifPULL
                Else
                    D = D * .StifPUSH
                End If

                N = Vec2Normalize(N)    '*************************************
                BodyApplyImpulse .bA, Vec2MUL(N, -D), .tAnchA



            Case Rotor1


                'ORI = SetOrient(Body(.bA).orient + Body(.bA).angularVelocity)
                ORI = SetOrient(Body(.bA).orient + Body(.bA).angularVelocity * KDamp * INVdt)
                .tAnchA = matMULv(ORI, .AnchA)
                'FIXa = Vec2ADD(Body(.bA).Pos, .tAnchA)
                N.X = -.tAnchA.Y
                N.Y = .tAnchA.X
                RotDiff = .L - Body(.bA).angularVelocity


                RotDiff = RotDiff * 20
                If Sgn(RotDiff) = Sgn(.L) Then
                    N = Vec2MUL(N, RotDiff)
                    BodyApplyImpulse .bA, N, Vec2ADD(.tAnchA, Body(.bA).VEL)
                End If



            Case Rotor2
                '''' Like Joint2pins

                ORI = SetOrient(Body(.bA).orient + Body(.bA).angularVelocity * KDamp * INVdt)
                '    .tAnchA = matMULv(Body(.bA).U, .AnchA)


                .tAnchA = matMULv(ORI, .AnchA)
                FIXa = Vec2ADD(Body(.bA).Pos, .tAnchA)
                V = Vec2MUL(Body(.bA).VEL, KDamp * INVdt)
                'V = VectorProject(V, .tAnchA)
                FIXa = Vec2ADD(FIXa, V)   '.------------------<<<<<<<

                ORI = SetOrient(Body(.bB).orient + Body(.bB).angularVelocity * KDamp * INVdt)
                '.tAnchB = matMULv(Body(.bB).U, .AnchB)
                .tAnchB = matMULv(ORI, .AnchB)

                FIXb = Vec2ADD(Body(.bB).Pos, .tAnchB)
                V = Vec2MUL(Body(.bB).VEL, KDamp * INVdt)
                'V = VectorProject(V, .tAnchB)
                FIXb = Vec2ADD(FIXb, V)   '.------------------<<<<<<<

                Axe = Vec2SUB(FIXb, FIXa)

                'vAprj = VectorProject(Vec2MUL(Body(.bA).VEL, 1 * 1), Axe)'---TOLTO
                'vBprj = VectorProject(Vec2MUL(Body(.bB).VEL, 1 * 1), Axe)'---TOLTO

                N = Vec2SUB(FIXa, FIXb)
                'N = Vec2ADD(N, Vec2SUB(vAprj, vBprj))''---TOLTO

                D = Vec2Length(N)
                If D Then N = Vec2MUL(N, 1 / D)
                D = (D - .L) * (Body(.bA).mass + Body(.bB).mass) * KMASS
                If D > 0 Then
                    D = D * 0.5
                Else
                    D = D * 0.5
                End If

                BodyApplyImpulse .bA, Vec2MUL(N, -D), .tAnchA
                BodyApplyImpulse .bB, Vec2MUL(N, D), .tAnchB
                '---------------------------------

                RotDiff = AngleDIFF(PI * 0.8 * Cos(CNT * 0.0008) + _
                                    Body(.bB).orient + Body(.bB).angularVelocity * KDamp * INVdt, _
                                    Body(.bA).orient + Body(.bA).angularVelocity * KDamp * INVdt)

                If Abs(RotDiff) > 1 Then RotDiff = Sgn(RotDiff) * 1

                ORI = SetOrient(Body(.bA).orient + Body(.bA).angularVelocity * KDamp * INVdt)
                .tAnchA = matMULv(ORI, .AnchA)
                .tAnchA.X = -.tAnchA.X
                .tAnchA.Y = -.tAnchA.Y
                N.X = -.tAnchA.Y
                N.Y = .tAnchA.X
                N = Vec2MUL(N, RotDiff)
                BodyApplyImpulse .bA, N, Vec2ADD(.tAnchA, Body(.bA).VEL)


                ORI = SetOrient(Body(.bB).orient + Body(.bB).angularVelocity * KDamp * INVdt)
                .tAnchB = matMULv(ORI, .AnchB)
                .tAnchB.X = -.tAnchB.X
                .tAnchB.Y = -.tAnchB.Y
                N.X = -.tAnchB.Y
                N.Y = .tAnchB.X
                N = Vec2MUL(N, -RotDiff)
                BodyApplyImpulse .bB, N, Vec2ADD(.tAnchB, Body(.bB).VEL)


            End Select

        End With
    Next

End Sub



'*********************************************************************
'******************* R E N D E R *************************************
'*********************************************************************



Private Sub RenderINITRC(PIC As PictureBox)
' Set Srf = Cairo.CreateSurface(400, 400)    'size of our rendering-area in Pixels
' Set CC = Srf.CreateContext    'create a Drawing-Context from the PixelSurface above



    Set vbDRAW = Cairo.CreateVBDrawingObject
    '    Set vbDRAW.Srf = Cairo.CreateSurface(400, 400)    'size of our rendering-area in Pixels
    Set vbDRAW.Srf = Cairo.CreateSurface(PIC.Width, PIC.Height, ImageSurface)           'size of our rendering-area in Pixels

    Set vbDRAW.CC = vbDRAW.Srf.CreateContext    'create a Drawing-Context from the PixelSurface above

    'vbDRAW.BindTo  PIC

    With vbDRAW

        .CC.AntiAlias = CAIRO_ANTIALIAS_GRAY

        '.CC.SetSourceSurface Srf
        .CC.SetLineCap CAIRO_LINE_CAP_ROUND
        .CC.SetLineJoin CAIRO_LINE_JOIN_ROUND


        .CC.SetLineWidth 1, True


        .CC.SelectFont "Courier New", 9, vbGreen

    End With

    PicHDC = PIC.hDC

    '     PIC.Cls
    '     PIC.Height = 640    '480    '360    ' 480
    '     PIC.Width = Int( PIC.Height * 4 / 3)

    PicW = PIC.Width
    PicH = PIC.Height


    Set vbDrawCC = vbDRAW.CC    ''''''''<<<<<<<<<<<< to speed up (Maybe)


End Sub

Private Sub RenderUnloadRC()
    Set CC = Nothing
    Set Srf = Nothing
    Set vbDRAW = Nothing

    Set CONS = New cConstructor
    CONS.CleanupRichClientDll
End Sub

Public Sub RenderCreateIntroFrames()
    vbDRAW.CC.SetSourceColor 0
    vbDRAW.CC.Paint

    RenderStringOut "                               "
    RenderStringOut "2D-Impulse-Engine:   (V" & Version & ")"
    RenderStringOut "                               "
    RenderStringOut "VB6 Port of Randy Gaul Impulse Engine.  "
    RenderStringOut "                               "
    RenderStringOut "Parameters:"
    RenderStringOut "Delta time:  " & DT
    RenderStringOut "Iterations:  " & Iterations
    RenderStringOut "G. Static  Friction:   " & GlobalSTATICFRICTION
    RenderStringOut "G. Dynamic Friction:   " & GlobalDYNAMICFRICTION
    RenderStringOut "Restitution: " & GlobalRestitution
    RenderStringOut "Penetration Allowance: " & PENETRATION_ALLOWANCE
    RenderStringOut "Penetration Corretion: " & PENETRATION_CORRETION
    RenderStringOut "Epsilon:     " & EPSILON
    RenderStringOut "                               "
    RenderStringOut "port to VB6 & Joints by MiorSoft"
    RenderStringOut "                                                                     "
    RenderStringOut "                                                                     "

End Sub
Public Sub RenderCreateOuttroFrames()
    vbDRAW.CC.SetSourceColor 0
    vbDRAW.CC.Paint

    RenderStringOut "                               "
    RenderStringOut "                                Thanks for watching !"
    RenderStringOut "                               "
    RenderStringOut "                               "
    RenderStringOut "                               "
    RenderStringOut "                               "


End Sub
Private Sub RenderStringOut(S As String)
    Dim I       As Double
    Const sStep As Double = 40    ' 10 ' 30
    Static Y    As Double
    Dim S2      As String

    Do
        S = S & " "
    Loop While (Len(S) - 1) Mod sStep <> 0

    S2 = S

    For I = 1 To Len(S2) Step sStep
        vbDRAW.CC.TextOut 10 + I * 7, Y, Mid$(S, I, sStep)
        vbDRAW.Srf.WriteContentToJpgFile App.Path & "\Frames\" & Format(Frame, "00000") & ".jpg", JPGQuality
        Frame = Frame + 1
    Next
    Y = Y + 16

    If Y > PicH Then Y = 5




End Sub
Public Sub RenderDRAWRC(Optional RefreshPicNow As Boolean = True)
    Dim x1      As Double
    Dim y1      As Double
    Dim x2      As Double
    Dim y2      As Double
    Dim x3      As Double
    Dim y3      As Double

    Dim x1d     As Double
    Dim y1d     As Double
    Dim x2d     As Double
    Dim y2d     As Double


    Dim I       As Long
    Dim J       As Long
    Dim JJ      As Long
    Dim S       As String


    vbDrawCC.SetSourceColor 2631720    '0
    vbDrawCC.Paint
    vbDrawCC.SetLineWidth 1.25


    For I = 1 To NBodies

        '        With Shape(I)
        With Body(I)

            If .myType = eCircle Then
                '                x1 = .Pos.x
                '                y1 = .Pos.y
                '
                '                vbDrawCC.SetSourceColor .color
                '                ' vbDrawCC.Ellipse x1, y1, .Radius * 2, .Radius * 2
                '                vbDrawCC.ARC x1, y1, .Radius
                '                vbDrawCC.Fill
                '                'vbDrawCC.Stroke
                '
                '                x2 = x1 + .Radius * Cos(.orient)
                '                y2 = y1 + .Radius * Sin(.orient)
                '                vbDrawCC.DrawLine x1, y1, x2, y2, , , 0, 0.25

                vbDrawCC.Save
                vbDrawCC.TranslateDrawings .Pos.X + .ImageCenter.Y, _
                                           .Pos.Y + .ImageCenter.Y
                vbDrawCC.RotateDrawings .orient
                vbDrawCC.RenderSurfaceContent ImageList(CStr(I)), -.ImageSize.X * 0.5, -.ImageSize.Y * 0.5
                vbDrawCC.Restore



            Else

                If .ImageExist Then
                    vbDrawCC.Save

                    vbDrawCC.TranslateDrawings .Pos.X + .ImageCenter.Y, _
                                               .Pos.Y + .ImageCenter.Y
                    vbDrawCC.RotateDrawings .orient
                    vbDrawCC.RenderSurfaceContent ImageList(CStr(I)), -.ImageSize.X * 0.5, -.ImageSize.Y * 0.5
                    vbDrawCC.Restore

                Else




                    '''' FILL
                    vbDrawCC.SetSourceColor .color

                    x1 = .tVertex(1).X + .Pos.X
                    y1 = .tVertex(1).Y + .Pos.Y
                    vbDrawCC.MoveTo x1, y1
                    '                For J = 2 To .Nvertex 'STADARD
                    '                    X1 = .tVertex(J).X + .Pos.X
                    '                    y1 = .tVertex(J).Y + .Pos.Y
                    '                    vbDrawCC.LineTo X1, y1
                    '                Next
                    J = 2
                    Do
                        'If .LToCT(J) = 1 Then
                        x1 = .tVertex(J).X + .Pos.X
                        y1 = .tVertex(J).Y + .Pos.Y
                        vbDrawCC.LineTo x1, y1
                        J = J + 1
                        '                    Else
                        '                        x1 = .tVertex(J).x + .Pos.x
                        '                        y1 = .tVertex(J).y + .Pos.y
                        '                        J = J + 1
                        '                        x2 = .tVertex(J).x + .Pos.x
                        '                        y2 = .tVertex(J).y + .Pos.y
                        '                        J = J + 1
                        '                        x3 = .tVertex(J).x + .Pos.x
                        '                        y3 = .tVertex(J).y + .Pos.y
                        '                        vbDrawCC.CurveTo x1, y1, x2, y2, x3, y3
                        '                        J = J + 1
                        '                    End If
                    Loop While J <= .Nvertex


                    vbDrawCC.Fill
                    'vbDrawCC.Stroke


                    vbDrawCC.SetSourceColor 0, 0.25
                    'vbDrawCC.Ellipse .Pos.x, .Pos.y, 3, 3
                    vbDrawCC.ARC .Pos.X, .Pos.Y, 1.5
                    vbDrawCC.Fill

                    If .Nvertex > 4 Then vbDrawCC.DrawLine x1, y1, .Pos.X, .Pos.Y, , , 0, 0.25



                End If




            End If

        End With

    Next
















    '    ' DRAW Contact Points----------------------------------------------------------------------
    '    For I = 1 To NofContactMainFolds
    '        With Contacts(I)
    '            For J = 1 To .contactCount
    '                x1 = .contactsPTS(J).X
    '                y1 = .contactsPTS(J).Y
    '
    '
    '                x2 = x1 + .normal.X * (1 + .penetration * 25)
    '                y2 = y1 + .normal.Y * (1 + .penetration * 25)
    '
    '                vbDrawCC.DrawLine x1, y1, x2, y2, , 2, vbBlue, 0.5
    '            Next
    '        End With
    '    Next
    '    -------------------------------------------------------------------------------------------



    '    For I = 1 To NJ
    '        With Joints(I)
    '            x1 = shape(.bA).Pos.x
    '            y1 = shape(.bA).Pos.y
    '            x2 = shape(.bB).Pos.x
    '            y2 = shape(.bB).Pos.y
    '            FastLine pHDC, x1, y1, x2, y2, 1, vbWhite
    '        End With
    '    Next
    For I = 1 To NJ

        With Joints(I)
            Select Case .JointType

            Case JointDistance
                x1 = Body(.bA).Pos.X
                y1 = Body(.bA).Pos.Y
                x2 = Body(.bB).Pos.X
                y2 = Body(.bB).Pos.Y
                vbDrawCC.DrawLine x1, y1, x2, y2, , 5, vbBlue, 0.5


            Case Joint2Pins
                x1 = Body(.bA).Pos.X + .tAnchA.X
                y1 = Body(.bA).Pos.Y + .tAnchA.Y
                x2 = Body(.bB).Pos.X + .tAnchB.X
                y2 = Body(.bB).Pos.Y + .tAnchB.Y
                vbDrawCC.DrawLine x1, y1, x2, y2, , 5, vbBlue, 0.5

            Case Joint1Pin
                x1 = .AnchB.X
                y1 = .AnchB.Y
                x2 = Body(.bA).Pos.X + .tAnchA.X
                y2 = Body(.bA).Pos.Y + .tAnchA.Y
                vbDrawCC.DrawLine x1, y1, x2, y2, , 5, vbBlue, 0.5
                vbDrawCC.SetSourceColor vbRed, 0.5
                'vbDrawCC.Ellipse x1, y1, 5, 5
                vbDrawCC.ARC x1, y1, 2.5
                vbDrawCC.Fill

            Case Rotor1

                x2 = Body(.bA).Pos.X - .tAnchA.Y
                y2 = Body(.bA).Pos.Y + .tAnchA.X
                x1 = x2 - .tAnchA.X
                y1 = y2 - .tAnchA.Y
                vbDrawCC.DrawLine x1, y1, x2, y2, , 4, vbRed, 0.5

            Case Rotor2
                x1 = Body(.bA).Pos.X - .tAnchA.X
                y1 = Body(.bA).Pos.Y - .tAnchA.Y
                x2 = Body(.bB).Pos.X - .tAnchB.X
                y2 = Body(.bB).Pos.Y - .tAnchB.Y
                '                vbDrawCC.DrawLine x1, y1, x2, y2, , 3, vbBlue, 0.5
                vbDrawCC.SetSourceColor vbRed, 0.5
                vbDrawCC.ARC (x1 + x2) * 0.5, (y1 + y2) * 0.5, 6

                vbDrawCC.Fill

            End Select

        End With
    Next


    If MouseSelectedObj Then
        vbDrawCC.SetSourceColor vbRed
        vbDrawCC.MoveTo Body(MouseSelectedObj).Pos.X + MouseDownRelX, Body(MouseSelectedObj).Pos.Y + MouseDownRelY
        vbDrawCC.LineTo MouseMoveX, MouseMoveY
        vbDrawCC.Stroke
    End If


    '    vbDrawCC.TextOut 5, 5, S
    If RefreshPicNow Then vbDRAW.Srf.DrawToDC PicHDC

    DoEvents


End Sub

Public Sub RenderRefreshPic()
    vbDRAW.Srf.DrawToDC PicHDC

    DoEvents
End Sub


Public Sub RenderSaveJPG(FN As String)
    vbDRAW.Srf.WriteContentToJpgFile FN, JPGQuality
End Sub
Public Sub UnLoad()
    RenderUnloadRC
    Erase Contacts
    Erase Body
End Sub



Public Sub BodyGetNearest(X As Double, Y As Double, retObj As Long, retRelX As Double, retRelY As Double)

    Dim I       As Long
    Dim J       As Long

    Dim Dmin    As Double
    Dim D       As Double
    Dim Dx      As Double
    Dim DY      As Double

    Dmin = MAX_VALUE

    For I = 1 To NBodies
        If Body(I).myType = eCircle Then
            Dx = X - Body(I).Pos.X
            DY = Y - Body(I).Pos.Y
            D = Dx * Dx + DY * DY
            If D < Dmin Then
                Dmin = D
                retObj = I
            End If

        Else
            For J = 1 To Body(I).Nvertex

                Dx = X - Body(I).tVertex(J).X
                DY = Y - Body(I).tVertex(J).Y
                Dx = Dx - Body(I).Pos.X
                DY = DY - Body(I).Pos.Y

                D = Dx * Dx + DY * DY
                If D < Dmin Then

                    retObj = I
                End If
            Next

            Dx = X - Body(I).Pos.X
            DY = Y - Body(I).Pos.Y
            D = Dx * Dx + DY * DY
            If D < Dmin Then
                Dmin = D
                retObj = I
            End If

        End If



    Next

    retRelX = X - Body(retObj).Pos.X
    retRelY = Y - Body(retObj).Pos.Y


End Sub

Friend Function BodyGetPOS(wB As Long) As tVec2
    BodyGetPOS = Body(wB).Pos
End Function

Public Function BodyGetMass(wB As Long) As Double
    BodyGetMass = Body(wB).mass
End Function


Friend Function BodyGetVEL(wB As Long) As tVec2
    BodyGetVEL = Body(wB).VEL
End Function





''Friend Sub RenderDRAWRCzoom(ZOOM As Double, Center As tVec2, Optional RefreshPicNow As Boolean = True)
''    Dim x1      As Double
''    Dim y1      As Double
''    Dim x2      As Double
''    Dim y2      As Double
''    Dim x3      As Double
''    Dim y3      As Double
''
''    Dim x1d     As Double
''    Dim y1d     As Double
''    Dim x2d     As Double
''    Dim y2d     As Double
''
''
''    Dim I       As Long
''    Dim J       As Long
''    Dim JJ      As Long
''    Dim S       As String
''
''
''    vbDrawCC.Save
''    vbDrawCC.TranslateDrawings CenX - Center.X * ZOOM, CenY - Center.Y * ZOOM
''    vbDrawCC.ScaleDrawings ZOOM, ZOOM
''
''    vbDrawCC.SetSourceColor 2631720    '0
''    vbDrawCC.Paint
''    vbDrawCC.SetLineWidth 1.25
''
''
''    For I = 1 To NBodies
''
'''''        If I <> gPlayer Then
''
''            With Body(I)
''
''                If .myType = eCircle Then
''
''                    If .ImageExist Then
''
''
''                        vbDrawCC.Save
''                        vbDrawCC.TranslateDrawings .Pos.X + .ImageCenter.Y, _
''                                                   .Pos.Y + .ImageCenter.Y
''                        vbDrawCC.RotateDrawings .orient
''                        vbDrawCC.RenderSurfaceContent ImageList(CStr(I)), -.ImageSize.X * 0.5, -.ImageSize.Y * 0.5
''                        vbDrawCC.Restore
''
''                    Else
''
''                        x1 = .Pos.X
''                        y1 = .Pos.Y
''
''                        vbDrawCC.SetSourceColor .color
''                        ' vbDrawCC.Ellipse x1, y1, .Radius * 2, .Radius * 2
''                        vbDrawCC.ARC x1, y1, .Radius
''                        vbDrawCC.Fill
''                        'vbDrawCC.Stroke
''
''                        x2 = x1 + .Radius * Cos(.orient)
''                        y2 = y1 + .Radius * Sin(.orient)
''                        vbDrawCC.DrawLine x1, y1, x2, y2, , , 0, 0.25
''
''                    End If
''
''
''
''                Else
''
''
''                    If .ImageExist Then
''                        vbDrawCC.Save
''                        vbDrawCC.TranslateDrawings .Pos.X + .ImageCenter.Y, _
''                                                   .Pos.Y + .ImageCenter.Y
''                        vbDrawCC.RotateDrawings .orient
''                        vbDrawCC.RenderSurfaceContent ImageList(CStr(I)), -.ImageSize.X * 0.5, -.ImageSize.Y * 0.5
''                        vbDrawCC.Restore
''
''                    Else
''
''
''
''
''                        '''' FILL
''                        vbDrawCC.SetSourceColor .color
''
''                        x1 = .tVertex(1).X + .Pos.X
''                        y1 = .tVertex(1).Y + .Pos.Y
''                        vbDrawCC.MoveTo x1, y1
''                        '                For J = 2 To .Nvertex 'STADARD
''                        '                    X1 = .tVertex(J).X + .Pos.X
''                        '                    y1 = .tVertex(J).Y + .Pos.Y
''                        '                    vbDrawCC.LineTo X1, y1
''                        '                Next
''                        J = 2
''                        Do
''                            'If .LToCT(J) = 1 Then
''                            x1 = .tVertex(J).X + .Pos.X
''                            y1 = .tVertex(J).Y + .Pos.Y
''                            vbDrawCC.LineTo x1, y1
''                            J = J + 1
''                            '                    Else
''                            '                        x1 = .tVertex(J).x + .Pos.x
''                            '                        y1 = .tVertex(J).y + .Pos.y
''                            '                        J = J + 1
''                            '                        x2 = .tVertex(J).x + .Pos.x
''                            '                        y2 = .tVertex(J).y + .Pos.y
''                            '                        J = J + 1
''                            '                        x3 = .tVertex(J).x + .Pos.x
''                            '                        y3 = .tVertex(J).y + .Pos.y
''                            '                        vbDrawCC.CurveTo x1, y1, x2, y2, x3, y3
''                            '                        J = J + 1
''                            '                    End If
''                        Loop While J <= .Nvertex
''
''
''                        vbDrawCC.Fill
''                        'vbDrawCC.Stroke
''
''
''                        vbDrawCC.SetSourceColor 0, 0.25
''                        'vbDrawCC.Ellipse .Pos.x, .Pos.y, 3, 3
''                        vbDrawCC.ARC .Pos.X, .Pos.Y, 1.5
''                        vbDrawCC.Fill
''
''                        If .Nvertex > 4 Then vbDrawCC.DrawLine x1, y1, .Pos.X, .Pos.Y, , , 0, 0.25
''
''
''
''                    End If
''
''
''
''
''                End If
''
''            End With
'''''        Else
'''''        '********************************************************
'''''        '********************************************************
'''''        '********************************************************
'''''
'''''
'''''
'''''        '********************************************************
'''''        '********************************************************
'''''        '********************************************************
'''''
'''''        End If
''
''    Next
''
''
''
''
''
''
''
''
''
''
''
''
''
''
''
''
''    '    ' DRAW Contact Points----------------------------------------------------------------------
''    '    For I = 1 To NofContactMainFolds
''    '        With Contacts(I)
''    '            For J = 1 To .contactCount
''    '                x1 = .contactsPTS(J).X
''    '                y1 = .contactsPTS(J).Y
''    '
''    '
''    '                x2 = x1 + .normal.X * (1 + .penetration * 25)
''    '                y2 = y1 + .normal.Y * (1 + .penetration * 25)
''    '
''    '                vbDrawCC.DrawLine x1, y1, x2, y2, , 2, vbBlue, 0.5
''    '            Next
''    '        End With
''    '    Next
''    '    -------------------------------------------------------------------------------------------
''
''
''
''    '    For I = 1 To NJ
''    '        With Joints(I)
''    '            x1 = shape(.bA).Pos.x
''    '            y1 = shape(.bA).Pos.y
''    '            x2 = shape(.bB).Pos.x
''    '            y2 = shape(.bB).Pos.y
''    '            FastLine pHDC, x1, y1, x2, y2, 1, vbWhite
''    '        End With
''    '    Next
''    For I = 1 To NJ
''
''        With Joints(I)
''            Select Case .JointType
''
''            Case JointDistance
''                x1 = Body(.bA).Pos.X
''                y1 = Body(.bA).Pos.Y
''                x2 = Body(.bB).Pos.X
''                y2 = Body(.bB).Pos.Y
''                vbDrawCC.DrawLine x1, y1, x2, y2, , 5, vbBlue, 0.5
''
''
''            Case Joint2Pins
''                x1 = Body(.bA).Pos.X + .tAnchA.X
''                y1 = Body(.bA).Pos.Y + .tAnchA.Y
''                x2 = Body(.bB).Pos.X + .tAnchB.X
''                y2 = Body(.bB).Pos.Y + .tAnchB.Y
''                vbDrawCC.DrawLine x1, y1, x2, y2, , 5, vbBlue, 0.5
''
''            Case Joint1Pin
''                x1 = .AnchB.X
''                y1 = .AnchB.Y
''                x2 = Body(.bA).Pos.X + .tAnchA.X
''                y2 = Body(.bA).Pos.Y + .tAnchA.Y
''                vbDrawCC.DrawLine x1, y1, x2, y2, , 5, vbBlue, 0.5
''                vbDrawCC.SetSourceColor vbRed, 0.5
''                'vbDrawCC.Ellipse x1, y1, 5, 5
''                vbDrawCC.ARC x1, y1, 2.5
''                vbDrawCC.Fill
''
''            Case Rotor1
''
''                x2 = Body(.bA).Pos.X - .tAnchA.Y
''                y2 = Body(.bA).Pos.Y + .tAnchA.X
''                x1 = x2 - .tAnchA.X
''                y1 = y2 - .tAnchA.Y
''                vbDrawCC.DrawLine x1, y1, x2, y2, , 4, vbRed, 0.5
''
''            Case Rotor2
''                x1 = Body(.bA).Pos.X - .tAnchA.X
''                y1 = Body(.bA).Pos.Y - .tAnchA.Y
''                x2 = Body(.bB).Pos.X - .tAnchB.X
''                y2 = Body(.bB).Pos.Y - .tAnchB.Y
''                '                vbDrawCC.DrawLine x1, y1, x2, y2, , 3, vbBlue, 0.5
''                vbDrawCC.SetSourceColor vbRed, 0.5
''                vbDrawCC.ARC (x1 + x2) * 0.5, (y1 + y2) * 0.5, 6
''
''                vbDrawCC.Fill
''
''            End Select
''
''        End With
''    Next
''
''
''
''    If MouseSelectedObj Then
''        vbDrawCC.SetSourceColor vbRed
''        vbDrawCC.MoveTo Body(MouseSelectedObj).Pos.X + MouseDownRelX, Body(MouseSelectedObj).Pos.Y + MouseDownRelY
''        vbDrawCC.LineTo MouseMoveX, MouseMoveY
''        vbDrawCC.Stroke
''    End If
''
''
''
''
''    vbDrawCC.Restore
''
''
''
''
''    '    vbDrawCC.TextOut 5, 5, S
''    If RefreshPicNow Then vbDRAW.Srf.DrawToDC PicHDC
''
''    DoEvents
''
''
''End Sub


