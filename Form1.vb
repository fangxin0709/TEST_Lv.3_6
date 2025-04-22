Public Class Form1
    'Dim d的二維陣列,資料筆數計數器(列),錯誤訊息
    Dim d(100, 4), idx, err
    'rdata函數(取測資並輸入)
    Sub rdata()
        FileOpen(6, "C:\Users\dora0\Desktop\NTVS\丙級\軟設丙\測資\6.txt", OpenMode.Input)
        '計數器歸零
        idx = 0
        '跑do while迴圈直到檔案6跑完(EOF) 這裡一定要用do while
        Do While Not EOF(6)
            '計數器(列)+1
            idx = idx + 1
            'i為計數器(行) 三行所以跑三次(第四行為錯誤訊息 不是測資有的 所以不需跑到四次)
            For i = 1 To 3
                '輸入到陣列 Ex:第一次do while迴圈--d(1,1) d(1,2) d(1,3)--跑完第一列  第二次--d(2,1) d(2,2) d(2,3)--
                Input(6, d(idx, i))
            Next
        Loop
        '一定要記得關檔
        FileClose(6)
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '呼叫rdata()函式
        Call rdata()
        '跑計數器(列)直到全列數跑完
        For i = 1 To idx
            '先清空錯誤訊息
            err = ""
            '開始逐一檢查三個項目
            If err = "" Then Call sp1(i)
            If err = "" Then Call sp2(i)
            If err = "" Then Call sp3(i)
            '第四行為存放錯誤訊息的地方
            d(i, 4) = err
        Next
        '呼叫wdata()函式
        Call wdata()
    End Sub
    '第一個檢查英文字母和數字函式(要記得i)
    Sub sp1(i)
        'Dim 每列第一行的資料
        Dim idno = d(i, 1)
        'Dim 資料總字數(Len)
        Dim m1 = Len(idno)
        '字數不等於10就輸出錯誤
        If m1 <> 10 Then err = "FORMAT ERROR"
        'Dim 資料第一個字 = Mid(要讀取的字串(idno),第幾(1)個字開始讀取,讀取幾(1)個字)
        '注意這裡取出的是長度為1的字串 不是字元!!!
        Dim m2 = Mid(idno, 1, 1)
        'vb語言會自動將字轉成電腦看得懂的編碼(ASCII) 所以可以用<>來比較
        'm2不是大寫A-Z就輸出錯誤
        If m2 < "A" Or m2 > "Z" Then err = "FORMAT ERROR"
        '讀取資料後面2-10位的字
        For j = 2 To 10
            Dim m3 = Mid(idno, j, 1)
            'm3不是數字0-9就輸出錯誤
            If m3 < "0" Or m3 > "9" Then err = "FORMAT ERROR"
        Next
    End Sub
    '第二個檢查男女函式(要記得i)
    Sub sp2(i)
        'Dim 每列第一行的資料
        Dim idno = d(i, 1)
        'Dim 資料的第二個字
        Dim sex_12 = Mid(idno, 2, 1)
        'Dim 每列第三行的MF
        Dim sex_MF = d(i, 3)
        'Dim 將上面兩個字串結合成一個字串變數
        Dim msex = sex_12 & sex_MF
        'msex不是1M(男生)和2F(女生)就輸出錯誤
        If msex <> "1M" And msex <> "2F" Then err = "SEX CODE ERROR"
    End Sub
    '第三個檢查公式函式(要記得i)
    Sub sp3(i)
        'Dim 每列第一行的資料
        Dim idno = d(i, 1)
        'Dim 資料的第一個字
        Dim L1 = Mid(idno, 1, 1)
        'Dim 英文字母字串(注意IO)
        Dim s26 = "ABCDEFGHJKLMNPQRSTUVXYWZIO"
        'Dim 代號(題目上表格)
        '--InStr(要搜尋的目標字串(s26),尋找的關鍵字(L1))
        'Ex:假設我在資料的第一個字(L1)是B,程式會回去找B在目標字串(s26)第幾(2)位,題目上的代碼顯示B是11,所以最後要+9(很重要!!!)
        Dim m1 = InStr(s26, L1) + 9
        'Dim 取十位數 (\是取商數)
        Dim x1 = m1 \ 10
        'Dim 取個位數 (Mod是取餘數)
        Dim x2 = m1 Mod 10
        'Dim a陣列(題目用d 但上面已經宣告過全域陣列d了 改用a)
        Dim a(9)
        '將資料後面2-10位的字存入a陣列
        For j = 2 To 10
            a(j - 1) = Mid(idno, j, 1)
        Next
        '題目有這個公式 照抄就好
        Dim y = x1 + 9 * x2 + 8 * a(1) + 7 * a(2) + 6 * a(3) + 5 * a(4) + 4 * a(5) + 3 * a(6) + 2 * a(7) + a(8) + a(9)
        'y不能被10整除(餘數不為0)就輸出錯誤
        If y Mod 10 <> 0 Then err = "CHECK SUM ERROR"
    End Sub
    'wdata函數(輸出到畫面)
    Sub wdata()
        'Dim 創建一個新的DataTable
        Dim table As New DataTable
        '新增欄位(行)
        table.Columns.Add("id_no")
        table.Columns.Add("name")
        table.Columns.Add("sex")
        table.Columns.Add("error")
        '有多少筆資料(列)就跑幾次
        For i = 1 To idx
            'Dim 創建一列
            Dim tr As DataRow = table.NewRow
            '將每列中的1,2,3,4欄填上資料
            tr(0) = d(i, 1) 'id_no
            tr(1) = d(i, 2) 'name
            tr(2) = d(i, 3) 'sex
            tr(3) = d(i, 4) 'error
            '新增列
            table.Rows.Add(tr)
        Next
        '將table指定給dgv顯示(DataGridView) 記得去屬性那邊 將name改成dgv方便這邊輸入(想直接打DataGridView也可)
        dgv.DataSource = table
        '排列改成升冪 這行可加可不加 不加的話預設降冪 要點一下id_no的欄位手動改成升冪(考試一定要記得 重新執行就會重製喔)
        dgv.Sort(dgv.Columns(0), 0)
        '這兩行要加 改欄位寬度的
        dgv.Columns(2).Width = 50
        dgv.Columns(3).Width = 150
    End Sub

End Class

