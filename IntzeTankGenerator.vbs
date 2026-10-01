' IntzeTankGenerator.vbs
' Writes a STAAD.Pro input file (.std) with the shell-element geometry of an
' Intze water tank. No STAAD scripting API is used: it only writes text, so it
' runs with plain Windows Script Host (double-click, or: cscript IntzeTankGenerator.vbs).
'
' Profile (revolved about the Y axis, Y is up):
'   bottom dome  -> conical wall -> cylindrical wall -> top dome
' Edit the parameters below, run the script, then open IntzeTank.std in STAAD.Pro.
' Units: metres, kN.

Option Explicit

' ---------------- PARAMETERS (edit these) ----------------
Const R_CYL      = 5.0    ' radius of cylindrical wall (m)
Const H_CYL      = 4.0    ' height of cylindrical wall (m)
Const R_RING     = 3.5    ' radius at the bottom of the conical wall = bottom dome base (m)
Const H_CONE     = 2.0    ' height of conical wall (m)
Const RISE_BOT   = 1.2    ' rise of the bottom dome (m, bulges downward)
Const RISE_TOP   = 1.0    ' rise of the top dome (m)
Const N_AROUND   = 24     ' divisions around the circumference
Const N_WALL     = 4      ' divisions along the cylindrical wall
Const N_CONE     = 3      ' divisions along the conical wall
Const N_DOME     = 4      ' divisions along each dome
Const T_WALL     = 0.25   ' shell thickness, cylindrical wall (m)
Const T_CONE     = 0.35   ' shell thickness, conical wall (m)
Const T_BOT      = 0.30   ' shell thickness, bottom dome (m)
Const T_TOP      = 0.10   ' shell thickness, top dome (m)
Const OUT_FILE   = "IntzeTank.std"
' ----------------------------------------------------------

Const PI = 3.14159265358979

Dim rr(), yy(), seg(), nPts
nPts = 0
ReDim rr(500): ReDim yy(500): ReDim seg(500)

Sub AddPt(r, y, s)
    rr(nPts) = r: yy(nPts) = y: seg(nPts) = s
    nPts = nPts + 1
End Sub

' Build the profile from the bottom-centre going outward and up.
Dim i, a, Rs, ang0, ang, yTop
' 1) bottom dome: sphere of radius Rs, apex at r=0, y=-RISE_BOT, base at (R_RING, 0)
Rs = (R_RING ^ 2 + RISE_BOT ^ 2) / (2 * RISE_BOT)
ang0 = ATan2(R_RING, Rs - RISE_BOT)
For i = 0 To N_DOME
    ang = ang0 * i / N_DOME
    AddPt Rs * Sin(ang), -RISE_BOT + Rs * (1 - Cos(ang)), 1
Next
' 2) conical wall from (R_RING, 0) to (R_CYL, H_CONE)
For i = 1 To N_CONE
    AddPt R_RING + (R_CYL - R_RING) * i / N_CONE, H_CONE * i / N_CONE, 2
Next
' 3) cylindrical wall up to y = H_CONE + H_CYL
For i = 1 To N_WALL
    AddPt R_CYL, H_CONE + H_CYL * i / N_WALL, 3
Next
' 4) top dome from (R_CYL, yTop) in to the apex at r=0
yTop = H_CONE + H_CYL
Rs = (R_CYL ^ 2 + RISE_TOP ^ 2) / (2 * RISE_TOP)
ang0 = ATan2(R_CYL, Rs - RISE_TOP)
For i = 1 To N_DOME
    ang = ang0 * (1 - i / N_DOME)
    AddPt Rs * Sin(ang), yTop + (Rs * Cos(ang) - (Rs - RISE_TOP)), 4
Next

Function ATan2(x, y)
    ' angle whose sine is x/hyp and cosine is y/hyp (both positive here)
    If y = 0 Then
        ATan2 = PI / 2
    Else
        ATan2 = Atn(x / y)
    End If
End Function

Function F(v)
    F = Replace(FormatNumber(v, 4, True, False, False), ",", ".")
End Function

' ---------------- joints ----------------
' Apex points (r = 0) get a single joint; every other ring has N_AROUND joints.
Dim nodeId(), nNodes, j, theta, k
ReDim nodeId(nPts - 1, N_AROUND - 1)
nNodes = 0
Dim fso, ts
Set fso = CreateObject("Scripting.FileSystemObject")
Set ts = fso.CreateTextFile(OUT_FILE, True)

ts.WriteLine "STAAD SPACE INTZE WATER TANK (generated)"
ts.WriteLine "START JOB INFORMATION"
ts.WriteLine "ENGINEER DATE " & Date
ts.WriteLine "END JOB INFORMATION"
ts.WriteLine "INPUT WIDTH 79"
ts.WriteLine "UNIT METER KN"
ts.WriteLine "JOINT COORDINATES"
For i = 0 To nPts - 1
    If rr(i) < 0.0001 Then
        nNodes = nNodes + 1
        For j = 0 To N_AROUND - 1: nodeId(i, j) = nNodes: Next
        ts.WriteLine nNodes & " 0 " & F(yy(i)) & " 0;"
    Else
        For j = 0 To N_AROUND - 1
            theta = 2 * PI * j / N_AROUND
            nNodes = nNodes + 1
            nodeId(i, j) = nNodes
            ts.WriteLine nNodes & " " & F(rr(i) * Cos(theta)) & " " & F(yy(i)) & " " & F(rr(i) * Sin(theta)) & ";"
        Next
    End If
Next

' ---------------- elements ----------------
Dim nEl, e1, e2, e3, e4, jn
Dim firstEl(4), lastEl(4)
For k = 1 To 4: firstEl(k) = 0: lastEl(k) = -1: Next
nEl = 0
ts.WriteLine "ELEMENT INCIDENCES SHELL"
For i = 0 To nPts - 2
    For j = 0 To N_AROUND - 1
        jn = (j + 1) Mod N_AROUND
        e1 = nodeId(i, j): e2 = nodeId(i, jn)
        e3 = nodeId(i + 1, jn): e4 = nodeId(i + 1, j)
        nEl = nEl + 1
        k = seg(i + 1)
        If firstEl(k) = 0 Then firstEl(k) = nEl
        lastEl(k) = nEl
        If e1 = e2 Then
            ts.WriteLine nEl & " " & e1 & " " & e3 & " " & e4 & ";"      ' triangle at bottom apex
        ElseIf e3 = e4 Then
            ts.WriteLine nEl & " " & e1 & " " & e2 & " " & e3 & ";"      ' triangle at top apex
        Else
            ts.WriteLine nEl & " " & e1 & " " & e2 & " " & e3 & " " & e4 & ";"
        End If
    Next
Next

' ---------------- properties ----------------
ts.WriteLine "ELEMENT PROPERTY"
Dim th: th = Array(0, T_BOT, T_CONE, T_WALL, T_TOP)
For k = 1 To 4
    If firstEl(k) > 0 Then ts.WriteLine firstEl(k) & " TO " & lastEl(k) & " THICKNESS " & F(th(k))
Next
ts.WriteLine "CONSTANTS"
ts.WriteLine "E 2.17185e+007 ALL"
ts.WriteLine "POISSON 0.17 ALL"
ts.WriteLine "DENSITY 25 ALL"

' Temporary supports on the ring where the bottom dome meets the cone.
' Replace these with your staging / columns.
ts.WriteLine "SUPPORTS"
Dim sup: sup = ""
For j = 0 To N_AROUND - 1
    sup = sup & nodeId(N_DOME, j) & " "
Next
ts.WriteLine sup & "PINNED"
ts.WriteLine "FINISH"
ts.Close

MsgBox "Created " & fso.GetAbsolutePathName(OUT_FILE) & vbCrLf & _
       nNodes & " joints, " & nEl & " plates." & vbCrLf & _
       "Open it in STAAD.Pro (File > Open).", vbInformation, "Intze Tank Generator"
