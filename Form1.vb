﻿Imports System.Numerics
Imports System.Reflection

Public Class Form1
    Dim d(100, 4), idx, err
    Sub rdata()
        FileOpen(6, "C:\Users\dora0\Desktop\6.txt", OpenMode.Input)
        idx = 0
        Do While Not EOF(6)
            idx += 1
            For i = 1 To 3
                Input(6, d(idx, i))
            Next
        Loop
        FileClose(6)
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call rdata()
        For i = 1 To idx
            err = ""
            If err = "" Then Call sp1(i)
            If err = "" Then Call sp2(i)
            If err = "" Then Call sp3(i)
            d(i, 4) = err
        Next
        Call wdata()
    End Sub
    Sub sp1(i)
        Dim idno = d(i, 1)
        Dim m1 = Len(idno)
        If m1 <> 10 Then err = "FORMAT ERROR"
        Dim m2 = Mid(idno, 1, 1)
        If m2 < "A" Or m2 > "Z" Then err = "FORMAT ERROR"
        For j = 2 To 8
            Dim m3 = Mid(idno, j, 1)
            If m3 < "0" Or m3 > "9" Then err = "FORMAT ERROR"
        Next
    End Sub
    Sub sp2(i)
        Dim idno = d(i, 1)
        Dim sex_12 = Mid(idno, 2, 1)
        Dim sex_MF = d(i, 3)
        Dim sex = sex_12 & sex_MF
        If sex <> "1M" And sex <> "2F" Then
            err = "SEX ERROR"
        End If
    End Sub
    Sub sp3(i)
        Dim idno = d(i, 1)
        Dim L1 = Mid(idno, 1, 1)
        Dim s26 = "ABCDEFGHJKLMNPQRSEUVWXYZIO"
        Dim m1 = InStr(s26, L1) + 9
        Dim x1 = m1 \ 10
        Dim x2 = m1 Mod 10
        Dim a(9)
        For j = 2 To 10
            a(j - 1) = Mid(idno, j, 1)
        Next
        Dim y
        y = x1 + 9 * x2
        For j = 1 To 8
            y = y + (9 - j) * a(j)
        Next
        y = y + a(9)
        If y Mod 10 <> 0 Then err = "CHECK SUM ERROR"
    End Sub
    Sub wdata()
        Dim table As New DataTable
        table.Columns.Add("ID_NO")
        table.Columns.Add("NAME")
        table.Columns.Add("SEX")
        table.Columns.Add("ERROR")
        For i = 1 To idx
            Dim tr As DataRow = table.NewRow
            tr(0) = d(i, 1)
            tr(1) = d(i, 2)
            tr(2) = d(i, 3)
            tr(3) = d(i, 4)
            table.Rows.Add(tr)
        Next
        dgv.DataSource = table
        dgv.Sort(dgv.Columns(0), 0)
        dgv.Columns(2).Width = 50
        dgv.Columns(3).Width = 150

        dgv.Columns(2).Width = 50

    End Sub
End Class
