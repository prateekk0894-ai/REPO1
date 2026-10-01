' IntzeTankModel.vbs
' Builds the complete STAAD.Pro model of the Intze tank described in
' INTZE_TANK_DESIGN.xlsm: shell geometry, ring beams, columns, bracing, raft girder,
' platform, section properties, supports, every load case and the load combinations.
'
' The input parameters are READ FROM THE EXCEL SHEET (DATA SHEET and TANK DESIGN),
' so change the sheet, run this script again, and the model follows.
'
' How to run:  double-click, or   cscript IntzeTankModel.vbs ["path\to\INTZE_TANK_DESIGN.xlsm"]
' Result:      INTZE_TANK_MODEL.std next to this script -> STAAD.Pro, File > Open.
'
' Excel is used only to read cell values (macros are disabled while reading). If Excel
' or the file cannot be found, the values the sheet had when this script was written
' are used, so the script still produces a model.
'
' Units: metre, kN.  Y is up.  Y = 0 is the top of the footing / raft.
' The shell is modelled on the WATER FACE (inside surface), so the water volume equals
' the capacity computed in the sheet.

Option Explicit

Const DEFAULT_XLSM    = "D:\CLAUDE AI\CIRCULAR UGT AND OVERHEAD TANK\INTZE_TANK_DESIGN.xlsm"
Const OUT_NAME        = "INTZE_TANK_MODEL.std"
Const MESH_AROUND     = 24     ' minimum divisions round the circumference (raised to a multiple of the column count)
Const N_DOME          = 4      ' divisions along each dome
Const N_CONE          = 3      ' divisions along the conical wall
Const N_WALL          = 5      ' divisions along the wetted cylindrical wall
' The sheet takes capacity = V1 + V2 - V3, i.e. the bottom dome RISES into the tank.
' Set to False for the classical Intze tank whose bottom dome hangs down (capacity + V3).
Const BOTTOM_DOME_UP  = True

Const PI   = 3.14159265358979
Const GRAV = 9.81

' ---------------------------------------------------------------- input values
' Defaults = values in the sheet when this script was written.
Dim D, H, fb, hd, td, b1w, b1d, tw, b2w, b2d, hc, tc, Ds, bottomText, hb, tb
Dim b3w, b3d, nCol, colShape, Dcol, Dside, hCol, nBr, brw, brd, rfB, rfT, b4w, b4d
Dim gw, gc, LLroof, FIN, platOn, pw, pt, platLL, rail, windOn, eqOn, Ec, capV
Dim windP, wa, wbc, wc, wd, we, wcol
Dim Ahi, Ahc, mi, mc, msm, hiS, hcS, hs, AhE, hcg, VE, ME
D = 8: H = 4.9: fb = 0.55: hd = 1.35: td = 0.1
b1w = 0.3: b1d = 0.6: tw = 0.2: b2w = 0.5: b2d = 0.3: hc = 1.35: tc = 0.2: Ds = 5.5
bottomText = "Spherical dome": hb = 1.1: tb = 0.2: b3w = 0.4: b3d = 0.6
nCol = 6: colShape = "Circular": Dcol = 0.5: Dside = 0.75: hCol = 16: nBr = 3
brw = 0.4: brd = 0.5: rfB = 3: rfT = 0.4: b4w = 0.4: b4d = 0.6
gw = 10: gc = 25: LLroof = 1.5: FIN = 1: platOn = True: pw = 1: pt = 0.15
platLL = 3: rail = 5: windOn = True: eqOn = True: Ec = 27386127.9: capV = 250
windP = 1.0833: wa = 36.9437: wbc = 6.9099: wc = 4.2897: wd = 62.3961: we = 10.7243: wcol = 1.2
Ahi = 0.0729: Ahc = 0.0439: mi = 165.7202: mc = 92.484: msm = 178.4476
hiS = 3.3088: hcS = 3.6974: hs = 16.6
AhE = 0.0938: VE = 164.1161: ME = 3222.3748

Dim fso, xl, wbk, wsD, wsT, srcNote
Set fso = CreateObject("Scripting.FileSystemObject")
Set wsD = Nothing: Set wsT = Nothing: Set xl = Nothing: Set wbk = Nothing
srcNote = "built-in defaults (Excel file not read)"

Function CellVal(ws, addr, def)
    Dim v
    CellVal = def
    If ws Is Nothing Then Exit Function
    v = ws.Range(addr).Value2
    If IsEmpty(v) Then Exit Function
    If IsError(v) Then Exit Function
    If VarType(def) = vbString Then
        CellVal = CStr(v)
    ElseIf VarType(def) = vbBoolean Then
        CellVal = (InStr(1, CStr(v), "YES", vbTextCompare) > 0)
    ElseIf IsNumeric(v) Then
        CellVal = CDbl(v)
    End If
End Function

Sub OpenExcel()
    Dim path
    If WScript.Arguments.Count > 0 Then
        path = WScript.Arguments(0)
    Else
        path = fso.GetParentFolderName(WScript.ScriptFullName) & "\INTZE_TANK_DESIGN.xlsm"
        If Not fso.FileExists(path) Then path = DEFAULT_XLSM
    End If
    If Not fso.FileExists(path) Then Exit Sub
    On Error Resume Next
    Set xl = CreateObject("Excel.Application")
    If xl Is Nothing Then
        Exit Sub
    End If
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AutomationSecurity = 3          ' msoAutomationSecurityForceDisable: never run the macros
    Set wbk = xl.Workbooks.Open(path, 0, True)
    If Err.Number <> 0 Or wbk Is Nothing Then
        Err.Clear
        xl.Quit
        Set xl = Nothing
        Exit Sub
    End If
    Set wsD = wbk.Worksheets("DATA SHEET")
    Set wsT = wbk.Worksheets("TANK DESIGN")
    If Err.Number <> 0 Then
        Err.Clear
        Set wsD = Nothing: Set wsT = Nothing
    Else
        srcNote = path
    End If
    On Error GoTo 0
End Sub

Sub ReadInputs()
    D = CellVal(wsD, "B10", D):      H = CellVal(wsD, "B11", H):    fb = CellVal(wsD, "B12", fb)
    hd = CellVal(wsD, "B13", hd):    td = CellVal(wsD, "B14", td)
    b1w = CellVal(wsD, "B16", b1w):  b1d = CellVal(wsD, "C16", b1d): tw = CellVal(wsD, "B17", tw)
    b2w = CellVal(wsD, "B19", b2w):  b2d = CellVal(wsD, "C19", b2d)
    hc = CellVal(wsD, "B20", hc):    tc = CellVal(wsD, "B21", tc):  Ds = CellVal(wsD, "B22", Ds)
    bottomText = CellVal(wsD, "B23", bottomText)
    hb = CellVal(wsD, "B24", hb):    tb = CellVal(wsD, "B25", tb)
    b3w = CellVal(wsD, "B27", b3w):  b3d = CellVal(wsD, "C27", b3d)
    nCol = CellVal(wsD, "B28", nCol): colShape = CellVal(wsD, "B29", colShape)
    Dcol = CellVal(wsD, "B30", Dcol): Dside = CellVal(wsD, "B31", Dside)
    hCol = CellVal(wsD, "B32", hCol): nBr = CellVal(wsD, "B33", nBr)
    brw = CellVal(wsD, "B35", brw):  brd = CellVal(wsD, "C35", brd)
    rfB = CellVal(wsD, "B38", rfB):  rfT = CellVal(wsD, "C38", rfT)
    b4w = CellVal(wsD, "B40", b4w):  b4d = CellVal(wsD, "C40", b4d)
    LLroof = CellVal(wsD, "B50", LLroof): FIN = CellVal(wsD, "B51", FIN)
    gw = CellVal(wsD, "B52", gw):    gc = CellVal(wsD, "B53", gc)
    platOn = CellVal(wsD, "B60", platOn)
    pw = CellVal(wsD, "B62", pw):    pt = CellVal(wsD, "C62", pt)
    platLL = CellVal(wsD, "B63", platLL): rail = CellVal(wsD, "B64", rail)
    windOn = CellVal(wsD, "B66", windOn): eqOn = CellVal(wsD, "B75", eqOn)
    capV = CellVal(wsD, "B9", capV * 1000) / 1000
    ' derived values computed by the sheet
    Ec = CellVal(wsT, "C11", Ec / 1000) * 1000
    windP = CellVal(wsT, "C301", windP)
    wa = CellVal(wsT, "C302", wa):   wbc = CellVal(wsT, "C303", wbc): wc = CellVal(wsT, "C304", wc)
    wd = CellVal(wsT, "C306", wd):   we = CellVal(wsT, "C307", we):   wcol = CellVal(wsT, "C305", wcol)
    mi = CellVal(wsT, "C313", mi):   mc = CellVal(wsT, "C314", mc):   msm = CellVal(wsT, "C315", msm)
    Ahi = CellVal(wsT, "C330", Ahi): Ahc = CellVal(wsT, "C331", Ahc)
    hiS = CellVal(wsT, "C332", hiS): hcS = CellVal(wsT, "C333", hcS): hs = CellVal(wsT, "C334", hs)
    AhE = CellVal(wsT, "C338", AhE): VE = CellVal(wsT, "C339", VE):   ME = CellVal(wsT, "C340", ME)
    If VE > 0 Then hcg = ME / VE Else hcg = hs + 3
End Sub

OpenExcel
ReadInputs
If Not (xl Is Nothing) Then
    On Error Resume Next
    wbk.Close False
    xl.Quit
    On Error GoTo 0
End If
If hcg = 0 Then hcg = hs + 3

' ---------------------------------------------------------------- geometry levels
Dim rD, rS, hp, yB, yC, yT, yW, sgn, isDome
rD = D / 2: rS = Ds / 2
hp = hCol / (nBr + 1)            ' panel height of the staging
yB = hCol + b3d                  ' container base: top of girder B3
yC = yB + hc                     ' top of the cone (ring beam B2)
yT = yC + H                      ' top of the wall (ring beam B1)
yW = yT - fb                     ' water level
If BOTTOM_DOME_UP Then sgn = 1 Else sgn = -1
isDome = (InStr(1, bottomText, "dome", vbTextCompare) > 0) And (hb > 0)

' ---------------------------------------------------------------- profile (centre of bottom -> apex of top dome)
Dim pr(200), py(200), pg(200), npts
npts = 0
Sub AddP(r, y, g)
    pr(npts) = r: py(npts) = y: pg(npts) = g
    npts = npts + 1
End Sub

Dim i, j, k, ang, Rb, a0, Rt, a1
If isDome Then
    Rb = (rS ^ 2 + hb ^ 2) / (2 * hb)
    a0 = Atn(rS / (Rb - hb))
    For i = 0 To N_DOME
        ang = a0 * i / N_DOME
        AddP Rb * Sin(ang), yB + sgn * hb - sgn * Rb * (1 - Cos(ang)), 1
    Next
Else
    For i = 0 To N_DOME
        AddP rS * i / N_DOME, yB, 1
    Next
End If
Dim idxB, idxC, idxT
idxB = npts - 1
For i = 1 To N_CONE
    AddP rS + (rD - rS) * i / N_CONE, yB + hc * i / N_CONE, 2
Next
idxC = npts - 1
If fb > 0.01 Then
    For i = 1 To N_WALL
        AddP rD, yC + (yW - yC) * i / N_WALL, 3
    Next
    AddP rD, yT, 3
Else
    For i = 1 To N_WALL
        AddP rD, yC + H * i / N_WALL, 3
    Next
End If
idxT = npts - 1
Rt = (rD ^ 2 + hd ^ 2) / (2 * hd)
a1 = Atn(rD / (Rt - hd))
For i = 1 To N_DOME
    ang = a1 * (1 - i / N_DOME)
    AddP Rt * Sin(ang), yT + Rt * Cos(ang) - (Rt - hd), 4
Next

' ---------------------------------------------------------------- nodes
Dim NA, nIn, cap
NA = nCol * (-Int(-MESH_AROUND / nCol))
cap = npts * NA + NA + nCol * (nBr + 2) + 50
Dim nx(), ny(), nz(), nid(), nNodes
ReDim nx(cap): ReDim ny(cap): ReDim nz(cap)
ReDim nid(npts, NA)
nNodes = 0

Function NewNode(x, y, z)
    nNodes = nNodes + 1
    nx(nNodes) = x: ny(nNodes) = y: nz(nNodes) = z
    NewNode = nNodes
End Function

Dim th, nn
For i = 0 To npts - 1
    If pr(i) < 0.000001 Then
        nn = NewNode(0, py(i), 0)
        For j = 0 To NA - 1: nid(i, j) = nn: Next
    Else
        For j = 0 To NA - 1
            th = 2 * PI * j / NA
            nid(i, j) = NewNode(pr(i) * Cos(th), py(i), pr(i) * Sin(th))
        Next
    End If
Next
' platform outer ring (at the level of B2)
Dim pOut(), rPo
ReDim pOut(NA)
rPo = rD + tw + pw
If platOn Then
    For j = 0 To NA - 1
        th = 2 * PI * j / NA
        pOut(j) = NewNode(rPo * Cos(th), yC, rPo * Sin(th))
    Next
End If
' column and raft-girder nodes: k = 0 is the base (y = 0), k = nBr the top brace level
Dim cn(), jc
ReDim cn(nCol, nBr + 1)
For i = 0 To nCol - 1
    jc = i * (NA \ nCol)
    th = 2 * PI * jc / NA
    For k = 0 To nBr
        cn(i, k) = NewNode(rS * Cos(th), hp * k, rS * Sin(th))
    Next
    cn(i, nBr + 1) = nid(idxB, jc)         ' the node on girder B3
Next

' ---------------------------------------------------------------- plates
Dim ne, pe1(), pe2(), pe3(), pe4(), pgrp()
cap = (npts - 1) * NA + NA + 10
ReDim pe1(cap): ReDim pe2(cap): ReDim pe3(cap): ReDim pe4(cap): ReDim pgrp(cap)
ne = 0
Dim jn, g1(6), g2(6), grpFirst(6), grpLast(6)
For k = 1 To 5: grpFirst(k) = 0: grpLast(k) = -1: Next
Sub AddEl(a, b, c, d, g)
    ne = ne + 1
    pe1(ne) = a: pe2(ne) = b: pe3(ne) = c: pe4(ne) = d: pgrp(ne) = g
    If grpFirst(g) = 0 Then grpFirst(g) = ne
    grpLast(g) = ne
End Sub
Dim ea, eb, ec2, ed
For i = 0 To npts - 2
    For j = 0 To NA - 1
        jn = (j + 1) Mod NA
        ea = nid(i, j): eb = nid(i, jn): ec2 = nid(i + 1, jn): ed = nid(i + 1, j)
        If ea = eb Then
            AddEl ea, ec2, ed, 0, pg(i + 1)
        ElseIf ec2 = ed Then
            AddEl ea, eb, ec2, 0, pg(i + 1)
        Else
            AddEl ea, eb, ec2, ed, pg(i + 1)
        End If
    Next
Next
If platOn Then
    For j = 0 To NA - 1
        jn = (j + 1) Mod NA
        AddEl nid(idxC, j), nid(idxC, jn), pOut(jn), pOut(j), 5
    Next
End If

' ---------------------------------------------------------------- members
Dim nm, ma(), mb(), mgrp()
cap = 3 * NA + nCol * (nBr + 1) + nCol * nBr + nCol + 10
ReDim ma(cap): ReDim mb(cap): ReDim mgrp(cap)
Dim mFirst(8), mLast(8)
For k = 1 To 7: mFirst(k) = 0: mLast(k) = -1: Next
nm = 0
Sub AddMb(a, b, g)
    nm = nm + 1
    ma(nm) = a: mb(nm) = b: mgrp(nm) = g
    If mFirst(g) = 0 Then mFirst(g) = nm
    mLast(g) = nm
End Sub
' groups: 1 B1 (top ring), 2 B2, 3 B3 (girder), 4 columns, 5 braces.  The raft (B4) is designed in the sheet; the columns are fixed at the top of the footing.
For j = 0 To NA - 1: AddMb nid(idxT, j), nid(idxT, (j + 1) Mod NA), 1: Next
For j = 0 To NA - 1: AddMb nid(idxC, j), nid(idxC, (j + 1) Mod NA), 2: Next
For j = 0 To NA - 1: AddMb nid(idxB, j), nid(idxB, (j + 1) Mod NA), 3: Next
For i = 0 To nCol - 1
    For k = 0 To nBr
        AddMb cn(i, k), cn(i, k + 1), 4
    Next
Next
For k = 1 To nBr
    For i = 0 To nCol - 1
        AddMb cn(i, k), cn((i + 1) Mod nCol, k), 5
    Next
Next

' ---------------------------------------------------------------- loads
Const NCASE = 12
Dim lf()
ReDim lf(NCASE, nNodes, 2)
Sub AddF(c, n, fx, fy, fz)
    lf(c, n, 0) = lf(c, n, 0) + fx
    lf(c, n, 1) = lf(c, n, 1) + fy
    lf(c, n, 2) = lf(c, n, 2) + fz
End Sub

Dim gAx, gAy, gAz, gCx, gCy, gCz, gN
Sub PlateGeom(e)
    Dim n1, n2, n3, n4, d1x, d1y, d1z, d2x, d2y, d2z
    n1 = pe1(e): n2 = pe2(e): n3 = pe3(e): n4 = pe4(e)
    If n4 = 0 Then
        d1x = nx(n2) - nx(n1): d1y = ny(n2) - ny(n1): d1z = nz(n2) - nz(n1)
        d2x = nx(n3) - nx(n1): d2y = ny(n3) - ny(n1): d2z = nz(n3) - nz(n1)
        gCx = (nx(n1) + nx(n2) + nx(n3)) / 3
        gCy = (ny(n1) + ny(n2) + ny(n3)) / 3
        gCz = (nz(n1) + nz(n2) + nz(n3)) / 3
        gN = 3
    Else
        d1x = nx(n3) - nx(n1): d1y = ny(n3) - ny(n1): d1z = nz(n3) - nz(n1)
        d2x = nx(n4) - nx(n2): d2y = ny(n4) - ny(n2): d2z = nz(n4) - nz(n2)
        gCx = (nx(n1) + nx(n2) + nx(n3) + nx(n4)) / 4
        gCy = (ny(n1) + ny(n2) + ny(n3) + ny(n4)) / 4
        gCz = (nz(n1) + nz(n2) + nz(n3) + nz(n4)) / 4
        gN = 4
    End If
    gAx = 0.5 * (d1y * d2z - d1z * d2y)
    gAy = 0.5 * (d1z * d2x - d1x * d2z)
    gAz = 0.5 * (d1x * d2y - d1y * d2x)
End Sub

Sub SpreadEl(c, e, fx, fy, fz)
    Dim q, nq
    nq = gN
    For q = 1 To nq
        Select Case q
            Case 1: AddF c, pe1(e), fx / nq, fy / nq, fz / nq
            Case 2: AddF c, pe2(e), fx / nq, fy / nq, fz / nq
            Case 3: AddF c, pe3(e), fx / nq, fy / nq, fz / nq
            Case 4: AddF c, pe4(e), fx / nq, fy / nq, fz / nq
        End Select
    Next
End Sub

' Case ids
Const C_DL = 1, C_WAT = 2, C_SIDL = 3, C_LL = 4, C_WX = 5, C_WZ = 6
Const C_EIX = 7, C_ECX = 8, C_EIZ = 9, C_ECZ = 10, C_EEX = 11, C_EEZ = 12

Dim e, p, dotv, wTot, wVert, domeArea, grp
wTot = 0: wVert = 0
For e = 1 To ne
    PlateGeom e
    grp = pgrp(e)
    ' --- C_WAT: hydrostatic pressure on bottom, cone and wall (pushes away from the water)
    If grp <= 3 Then
        p = gw * (yW - gCy)
        If p > 0 Then
            dotv = gAx * gCx + gAy * (gCy - yW) + gAz * gCz
            If dotv < 0 Then gAx = -gAx: gAy = -gAy: gAz = -gAz
            SpreadEl C_WAT, e, p * gAx, p * gAy, p * gAz
            wVert = wVert + p * gAy
        End If
    End If
    ' --- finishes and live load on the top dome (per plan area); live on the platform
    If grp = 4 Then
        SpreadEl C_SIDL, e, 0, -FIN * Abs(gAy), 0
        SpreadEl C_LL, e, 0, -LLroof * Abs(gAy), 0
    ElseIf grp = 5 Then
        SpreadEl C_LL, e, 0, -platLL * Abs(gAy), 0
    End If
Next
' railing at the platform edge
If platOn Then
    For j = 0 To NA - 1
        AddF C_SIDL, pOut(j), 0, -rail * (2 * PI * rPo / NA), 0
    Next
End If

' --- wind: forces from the sheet spread over the nodes they act on
Function RowOf(y, r0, r1)
    Dim q, best, bd
    best = r0: bd = 1E+30
    For q = r0 To r1
        If Abs(py(q) - y) < bd Then bd = Abs(py(q) - y): best = q
    Next
    RowOf = best
End Function

Sub SpreadRows(c, r0, r1, f, dirx)
    Dim q, jj, cnt, nd
    cnt = 0
    For q = r0 To r1: cnt = cnt + NA: Next
    For q = r0 To r1
        For jj = 0 To NA - 1
            nd = nid(q, jj)
            If dirx Then AddF c, nd, f / cnt, 0, 0 Else AddF c, nd, 0, 0, f / cnt
        Next
    Next
End Sub

Sub SpreadRow(c, q, f, dirx)
    Dim jj
    For jj = 0 To NA - 1
        If dirx Then AddF c, nid(q, jj), f / NA, 0, 0 Else AddF c, nid(q, jj), 0, 0, f / NA
    Next
End Sub

Dim wcolW, ib, cW, dx
If windOn Then
    For ib = 0 To 1
        If ib = 0 Then cW = C_WX: dx = True Else cW = C_WZ: dx = False
        SpreadRows cW, idxC, idxT, wa, dx          ' (a) dome and wall
        SpreadRows cW, idxB, idxC - 1, wbc, dx     ' (b) cone
        SpreadRow cW, idxB, wc, dx                 ' (c) girder
        For i = 0 To nCol - 1                      ' (e) braces, at the brace levels of the columns
            For k = 1 To nBr
                If dx Then AddF cW, cn(i, k), we / (nCol * nBr), 0, 0 Else AddF cW, cn(i, k), 0, 0, we / (nCol * nBr)
            Next
        Next
    Next
    wcolW = wd / (nCol * hCol)                 ' (d) columns, kN/m on each column, as in the sheet
End If

' --- earthquake, tank full: impulsive + structure, convective   (IS 1893 Part 2, two mass model)
If eqOn Then
    Dim rI, rCv, rG, Fi, Fst, Fc
    rI = RowOf(hs + hiS, idxB, idxT)
    rCv = RowOf(hs + hcS, idxB, idxT)
    rG = RowOf(hcg, idxB, idxT)
    Fi = Ahi * mi * GRAV
    Fst = Ahi * msm * GRAV
    Fc = Ahc * mc * GRAV
    SpreadRow C_EIX, rI, Fi, True:  SpreadRow C_EIX, rG, Fst, True:  SpreadRow C_ECX, rCv, Fc, True
    SpreadRow C_EIZ, rI, Fi, False: SpreadRow C_EIZ, rG, Fst, False: SpreadRow C_ECZ, rCv, Fc, False
    ' tank empty: only the structural mass
    SpreadRow C_EEX, rG, AhE * msm * GRAV, True
    SpreadRow C_EEZ, rG, AhE * msm * GRAV, False
End If

' ---------------------------------------------------------------- write the STAAD file
Function F3(v)
    F3 = Replace(FormatNumber(v, 4, True, False, False), ",", ".")
End Function
Function FS(v)
    FS = Replace(CStr(CDbl(v)), ",", ".")
End Function

Dim outPath, ts
outPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\" & OUT_NAME
Set ts = fso.CreateTextFile(outPath, True)

ts.WriteLine "STAAD SPACE INTZE WATER TANK - generated by IntzeTankModel.vbs"
ts.WriteLine "START JOB INFORMATION"
ts.WriteLine "ENGINEER DATE " & Day(Now) & "-" & Month(Now) & "-" & Year(Now)
ts.WriteLine "END JOB INFORMATION"
ts.WriteLine "INPUT WIDTH 79"
ts.WriteLine "UNIT METER KN"
ts.WriteLine "* Parameters read from: " & srcNote
ts.WriteLine "* D=" & D & " H=" & H & " freeboard=" & fb & " Ds=" & Ds & " columns=" & nCol & " h=" & hCol & " braces=" & nBr
ts.WriteLine "JOINT COORDINATES"
For i = 1 To nNodes
    ts.WriteLine i & " " & F3(nx(i)) & " " & F3(ny(i)) & " " & F3(nz(i)) & ";"
Next
ts.WriteLine "MEMBER INCIDENCES"
For i = 1 To nm
    ts.WriteLine i & " " & ma(i) & " " & mb(i) & ";"
Next
ts.WriteLine "ELEMENT INCIDENCES SHELL"
For i = 1 To ne
    If pe4(i) = 0 Then
        ts.WriteLine i & " " & pe1(i) & " " & pe2(i) & " " & pe3(i) & ";"
    Else
        ts.WriteLine i & " " & pe1(i) & " " & pe2(i) & " " & pe3(i) & " " & pe4(i) & ";"
    End If
Next

' member sections: YD = depth, ZD = width; a circular column is given by YD only
ts.WriteLine "MEMBER PROPERTY"
Dim sz(6)
sz(1) = "PRIS YD " & F3(b1d) & " ZD " & F3(b1w)
sz(2) = "PRIS YD " & F3(b2d) & " ZD " & F3(b2w)
sz(3) = "PRIS YD " & F3(b3d) & " ZD " & F3(b3w)
If colShape = "Circular" Then sz(4) = "PRIS YD " & F3(Dcol) Else sz(4) = "PRIS YD " & F3(Dside) & " ZD " & F3(Dcol)
sz(5) = "PRIS YD " & F3(brd) & " ZD " & F3(brw)
sz(6) = "PRIS YD " & F3(b4d) & " ZD " & F3(b4w)
For k = 1 To 6
    If mFirst(k) > 0 Then ts.WriteLine mFirst(k) & " TO " & mLast(k) & " " & sz(k)
Next
ts.WriteLine "ELEMENT PROPERTY"
Dim eth(5)
eth(1) = tb: eth(2) = tc: eth(3) = tw: eth(4) = td: eth(5) = pt
For k = 1 To 5
    If grpFirst(k) > 0 Then ts.WriteLine grpFirst(k) & " TO " & grpLast(k) & " THICKNESS " & F3(eth(k))
Next
ts.WriteLine "CONSTANTS"
ts.WriteLine "E " & FS(Ec) & " ALL"
ts.WriteLine "POISSON 0.2 ALL"
ts.WriteLine "DENSITY " & FS(gc) & " ALL"
ts.WriteLine "SUPPORTS"
Dim sup: sup = ""
For i = 0 To nCol - 1
    sup = sup & cn(i, 0) & " "
    If Len(sup) > 60 Then ts.WriteLine sup & "-": sup = ""
Next
ts.WriteLine sup & "FIXED"

' ---- load cases
Dim cs, hasLoad, line, nd
Sub WriteJointLoads(c)
    Dim q
    ts.WriteLine "JOINT LOAD"
    For q = 1 To nNodes
        If Abs(lf(c, q, 0)) + Abs(lf(c, q, 1)) + Abs(lf(c, q, 2)) > 0.0000001 Then
            line = q
            If Abs(lf(c, q, 0)) > 0.0000001 Then line = line & " FX " & F3(lf(c, q, 0))
            If Abs(lf(c, q, 1)) > 0.0000001 Then line = line & " FY " & F3(lf(c, q, 1))
            If Abs(lf(c, q, 2)) > 0.0000001 Then line = line & " FZ " & F3(lf(c, q, 2))
            ts.WriteLine line
        End If
    Next
End Sub

ts.WriteLine "LOAD 1 LOADTYPE Dead TITLE DL SELF WEIGHT"
ts.WriteLine "SELFWEIGHT Y -1"
ts.WriteLine "LOAD 2 LOADTYPE Dead TITLE WATER FULL TANK"
WriteJointLoads C_WAT
ts.WriteLine "LOAD 3 LOADTYPE Dead TITLE SIDL FINISHES AND RAILING"
WriteJointLoads C_SIDL
ts.WriteLine "LOAD 4 LOADTYPE Live TITLE LL ROOF AND PLATFORM"
WriteJointLoads C_LL
If windOn Then
    ts.WriteLine "LOAD 5 LOADTYPE Wind TITLE WL +X"
    ts.WriteLine "MEMBER LOAD"
    ts.WriteLine mFirst(4) & " TO " & mLast(4) & " UNI GX " & F3(wcolW)
    WriteJointLoads C_WX
    ts.WriteLine "LOAD 6 LOADTYPE Wind TITLE WL +Z"
    ts.WriteLine "MEMBER LOAD"
    ts.WriteLine mFirst(4) & " TO " & mLast(4) & " UNI GZ " & F3(wcolW)
    WriteJointLoads C_WZ
End If
If eqOn Then
    ts.WriteLine "LOAD 7 LOADTYPE Seismic TITLE EQ +X FULL IMPULSIVE+STRUCTURE"
    WriteJointLoads C_EIX
    ts.WriteLine "LOAD 8 LOADTYPE Seismic TITLE EQ +X FULL CONVECTIVE"
    WriteJointLoads C_ECX
    ts.WriteLine "LOAD 9 LOADTYPE Seismic TITLE EQ +Z FULL IMPULSIVE+STRUCTURE"
    WriteJointLoads C_EIZ
    ts.WriteLine "LOAD 10 LOADTYPE Seismic TITLE EQ +Z FULL CONVECTIVE"
    WriteJointLoads C_ECZ
    ts.WriteLine "LOAD 11 LOADTYPE Seismic TITLE EQ +X EMPTY"
    WriteJointLoads C_EEX
    ts.WriteLine "LOAD 12 LOADTYPE Seismic TITLE EQ +Z EMPTY"
    WriteJointLoads C_EEZ
End If

' ---- load combinations  (IS 456 Table 18 for the limit state; IS 875 / IS 1893 factors)
Dim cmb: cmb = 20
Sub Comb(title, spec)
    cmb = cmb + 1
    ts.WriteLine "LOAD COMB " & cmb & " " & title
    ts.WriteLine spec
End Sub
Comb "SERVICE DL+SIDL+WATER+LL", "1 1.0 3 1.0 2 1.0 4 1.0"
Comb "SERVICE EMPTY DL+SIDL+LL", "1 1.0 3 1.0 4 1.0"
Comb "ULS 1.5(DL+SIDL+WATER+LL)", "1 1.5 3 1.5 2 1.5 4 1.5"
Comb "ULS 1.5(DL+SIDL+LL) EMPTY", "1 1.5 3 1.5 4 1.5"
If windOn Then
    Comb "ULS 1.5(DL+SIDL+WATER+WLX)", "1 1.5 3 1.5 2 1.5 5 1.5"
    Comb "ULS 1.5(DL+SIDL+WATER+WLZ)", "1 1.5 3 1.5 2 1.5 6 1.5"
    Comb "ULS 1.5(DL+SIDL+WLX) EMPTY", "1 1.5 3 1.5 5 1.5"
    Comb "ULS 1.5(DL+SIDL+WLZ) EMPTY", "1 1.5 3 1.5 6 1.5"
    Comb "ULS 1.2(DL+SIDL+WATER+LL+WLX)", "1 1.2 3 1.2 2 1.2 4 1.2 5 1.2"
    Comb "ULS 1.2(DL+SIDL+WATER+LL+WLZ)", "1 1.2 3 1.2 2 1.2 4 1.2 6 1.2"
    Comb "ULS 0.9(DL+SIDL)+1.5WLX EMPTY", "1 0.9 3 0.9 5 1.5"
    Comb "ULS 0.9(DL+SIDL)+1.5WLZ EMPTY", "1 0.9 3 0.9 6 1.5"
    Comb "SERVICE DL+SIDL+WATER+WLX", "1 1.0 3 1.0 2 1.0 5 1.0"
    Comb "SERVICE DL+SIDL+WATER+WLZ", "1 1.0 3 1.0 2 1.0 6 1.0"
End If
If eqOn Then
    Comb "ULS 1.5(DL+SIDL+WATER+EQX)", "1 1.5 3 1.5 2 1.5 7 1.5 8 1.5"
    Comb "ULS 1.5(DL+SIDL+WATER+EQZ)", "1 1.5 3 1.5 2 1.5 9 1.5 10 1.5"
    Comb "ULS 1.2(DL+SIDL+WATER+LL+EQX)", "1 1.2 3 1.2 2 1.2 4 1.2 7 1.2 8 1.2"
    Comb "ULS 1.2(DL+SIDL+WATER+LL+EQZ)", "1 1.2 3 1.2 2 1.2 4 1.2 9 1.2 10 1.2"
    Comb "ULS 1.5(DL+SIDL+EQX) EMPTY", "1 1.5 3 1.5 11 1.5"
    Comb "ULS 1.5(DL+SIDL+EQZ) EMPTY", "1 1.5 3 1.5 12 1.5"
    Comb "ULS 0.9(DL+SIDL)+1.5EQX EMPTY", "1 0.9 3 0.9 11 1.5"
    Comb "ULS 0.9(DL+SIDL)+1.5EQZ EMPTY", "1 0.9 3 0.9 12 1.5"
End If
ts.WriteLine "PERFORM ANALYSIS"
ts.WriteLine "FINISH"
ts.Close

Dim msg
msg = "Model written to:" & vbCrLf & outPath & vbCrLf & vbCrLf & _
      "Inputs from: " & srcNote & vbCrLf & vbCrLf & _
      nNodes & " joints, " & ne & " plates, " & nm & " members, " & (cmb - 20) & " load combinations" & vbCrLf & _
      "Water on the shell (vertical): " & FormatNumber(-wVert, 1) & " kN" & vbCrLf & _
      "Weight of the water in the sheet: " & FormatNumber(gw * capV, 1) & " kN" & vbCrLf & vbCrLf & _
      "Open it in STAAD.Pro: File > Open."
MsgBox msg, vbInformation, "Intze Tank Model"
