REM  *****  BASIC  *****
Option Explicit

Private Function NumCell(oTable As Object, sCell As String) As Double
    Dim s As String
    s = Trim(oTable.getCellByName(sCell).getString())
    s = Replace(s, ",", ".")
    If s = "" Then
        NumCell = 0
    Else
        NumCell = CDbl(s)
    End If
End Function

Private Function Fmt(x As Double, Optional n As Integer = 3) As String
    Fmt = Format(x, "0." & String(n, "0"))
End Function

Private Sub Put(oTable As Object, sCell As String, sValue As String)
    oTable.getCellByName(sCell).setString(sValue)
End Sub

Sub ПересчитатьРасчет()
    On Error GoTo ErrHandler
    Dim oDoc As Object, oTables As Object
    Dim inp As Object, steel As Object, conc As Object, total As Object
    Dim m As Double, g As Double, a As Double, ds As Double, Rs As Double
    Dim Nb As Double, b As Double, Rb As Double
    Dim P As Double, M As Double, W As Double, sigS As Double, etaS As Double
    Dim Ab As Double, sigB As Double, etaB As Double
    Dim steelOK As Boolean, concOK As Boolean

    oDoc = ThisComponent
    oTables = oDoc.TextTables
    inp = oTables.getByName("ИсходныеДанные")
    steel = oTables.getByName("РезультатыСталь")
    conc = oTables.getByName("РезультатыБетон")
    total = oTables.getByName("ОбщийВывод")

    ' Входные значения: столбец B, строки 2...9
    m = NumCell(inp, "B2")
    g = NumCell(inp, "B3")
    a = NumCell(inp, "B4")
    ds = NumCell(inp, "B5")
    Rs = NumCell(inp, "B6")
    Nb = NumCell(inp, "B7")
    b = NumCell(inp, "B8")
    Rb = NumCell(inp, "B9")

    If m <= 0 Or g <= 0 Or a <= 0 Or ds <= 0 Or Rs <= 0 Or Nb <= 0 Or b <= 0 Or Rb <= 0 Then
        Put(steel, "B2", "Недостаточно данных: проверьте m, g, a, ds, Rs")
        Put(steel, "C2", "—")
        Put(steel, "B3", "—"): Put(steel, "C3", "—")
        Put(steel, "B4", "—"): Put(steel, "C4", "—")
        Put(steel, "B5", "—"): Put(steel, "C5", "—")
        Put(steel, "B6", "—"): Put(steel, "C6", "—")
        Put(steel, "B7", "Не рассчитано")
        Put(conc, "B2", "Недостаточно данных: проверьте Nb, b, Rb")
        Put(conc, "C2", "—")
        Put(conc, "B3", "—"): Put(conc, "C3", "—")
        Put(conc, "B4", "—"): Put(conc, "C4", "—")
        Put(conc, "B5", "Не рассчитано")
        Put(total, "B2", "НЕДОСТАТОЧНО ДАННЫХ — заполните исходные значения и повторите пересчёт")
        Exit Sub
    End If

    P = m * g
    M = P * a
    W = 3.14159265358979 * ds ^ 3 / 32
    sigS = M / W
    etaS = sigS / Rs
    steelOK = (sigS <= Rs)

    Put(steel, "B2", "P = m·g = " & Fmt(m, 2) & "·" & Fmt(g, 2) & " = " & Fmt(P, 2) & " Н")
    Put(steel, "C2", Fmt(P, 2) & " Н")
    Put(steel, "B3", "M = P·a = " & Fmt(P, 2) & "·" & Fmt(a, 1) & " = " & Fmt(M, 2) & " Н·мм")
    Put(steel, "C3", Fmt(M / 1000, 4) & " кН·м")
    Put(steel, "B4", "W = π·ds³/32 = π·" & Fmt(ds, 2) & "³/32")
    Put(steel, "C4", Fmt(W, 2) & " мм³")
    Put(steel, "B5", "σs = M/W = " & Fmt(M, 2) & "/" & Fmt(W, 2))
    Put(steel, "C5", Fmt(sigS, 2) & " МПа")
    Put(steel, "B6", "ηs = σs/Rs = " & Fmt(sigS, 2) & "/" & Fmt(Rs, 2))
    Put(steel, "C6", Fmt(etaS, 3))
    If steelOK Then
        Put(steel, "B7", "ВЫПОЛНЕНО: σs ≤ Rs")
    Else
        Put(steel, "B7", "НЕ ВЫПОЛНЕНО: σs > Rs")
    End If

    Ab = b ^ 2
    sigB = Nb * 1000 / Ab
    etaB = sigB / Rb
    concOK = (sigB <= Rb)

    Put(conc, "B2", "Ab = b² = " & Fmt(b, 2) & "²")
    Put(conc, "C2", Fmt(Ab, 2) & " мм²")
    Put(conc, "B3", "σb = Nb·1000/Ab = " & Fmt(Nb, 3) & "·1000/" & Fmt(Ab, 2))
    Put(conc, "C3", Fmt(sigB, 3) & " МПа")
    Put(conc, "B4", "ηb = σb/Rb = " & Fmt(sigB, 3) & "/" & Fmt(Rb, 2))
    Put(conc, "C4", Fmt(etaB, 3))
    If concOK Then
        Put(conc, "B5", "ВЫПОЛНЕНО: σb ≤ Rb")
    Else
        Put(conc, "B5", "НЕ ВЫПОЛНЕНО: σb > Rb")
    End If

    If steelOK And concOK Then
        Put(total, "B2", "ОБЕ ПРОВЕРКИ ВЫПОЛНЕНЫ. Проверка скалывания бетона не выполнялась.")
    Else
        Put(total, "B2", "ОТРИЦАТЕЛЬНЫЙ ИТОГ: как минимум одна проверка не выполнена. Проверка скалывания бетона не выполнялась.")
    End If
    oDoc.calculateAll()
    Exit Sub

ErrHandler:
    MsgBox "Не удалось выполнить расчёт. Проверьте имена таблиц и формат исходных данных. Подробности: " & Error$, 16, "Расчёт шпильки"
End Sub
