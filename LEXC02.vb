'┌─────┬────────────────────────────────────────────────────────────┐
'│ 類別名稱 │LEXC02 查詢                                                                                                             │
'├─────┼──────┬───────┬─────────────────────────────────────────────┤
'│ 日期     │撰寫人      │版本號        │撰寫內容                                                                                  │
'├─────┼──────┼───────┼─────────────────────────────────────────────┤
'│2011/05/25│tianen      │2011.05.25.01 │LEXC02查詢                                                                                │
'└─────┴──────┴───────┴─────────────────────────────────────────────┘
'2012.10.1.1 在2011.5.27.1基礎上修改，增加繁簡體轉換支持


Public Class LEXC02

    Public Sqlcommand As String = ""
    Dim k As Integer
    Dim filename As String = ""
    Dim issearch As Boolean = False



    Private Sub ToolBar1_ButtonClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles ToolBar1.ButtonClick

        Select Case e.Button.Text
            Case StrConv("查詢", SpiChangeLanguage.Value, 9)
                If issearch = True Then
                    check()
                    barcode_check()

                    issearch = False
                End If


            Case StrConv("匯入", SpiChangeLanguage.Value, 9)
                updata()
                If issearch = False Then
                    count.Text = ""
                    reel_listview.Items.Clear()
                    Me.OpenFileDialog1.InitialDirectory = "d:\"
                    Me.OpenFileDialog1.Filter = "EXECLE files (*.xls)|*.xls|All files (*.*)|*.*"
                    Me.OpenFileDialog1.FileName = ""
                    Me.OpenFileDialog1.ShowDialog()
                End If


            Case StrConv("結束", SpiChangeLanguage.Value, 9)
                Me.Close()
            Case StrConv("匯出", SpiChangeLanguage.Value, 9)

                If barcode_listview.Items.Count > 0 Then
                    Me.SaveFileDialog1.InitialDirectory = "d:\"
                    Me.SaveFileDialog1.Filter = "EXECLE files (*.xls)|*.xls|All files (*.*)|*.*"
                    Me.SaveFileDialog1.FileName = ""
                    Me.SaveFileDialog1.ShowDialog()


                Else
                    StatusBar1.Panels(0).Text = StrConv("無資料匯出", SpiChangeLanguage.Value, 9)
                End If

            Case StrConv("重整", SpiChangeLanguage.Value, 9)

                updata()


        End Select


    End Sub
    Sub check()
        StatusBar1.Panels(0).Text = ""
        work_listview.Items.Add("")
        work_listview.Items(0).SubItems.Add("")
        Try

            If Me.ClsDBInfo.datalink(Me.ClsCommonInfo.uLocalIP) Then
                For k = 0 To reel_listview.Items.Count - 1
                    Sqlcommand = "SELECT se17.f003,wk_ord_id, se08.f002,se17.f005 ,se11.f001" & _
                                  " FROM sajet.g_part_led_issue led, mse0017 se17, mse0008 se08,mse0011 se11" & _
                                   " WHERE led.part_sn = '" & reel_listview.Items(k).SubItems(0).Text & "'" & _
                                  " AND led.wk_ord_id = se17.f001" & _
                                  " AND se17.f005 = se08.f001" & _
                                  " and se17.f005 =se11.f002" & _
                                  " and se11.f004=6"
                    Dim dt As DataTable = New DataTable("dt")
                    Me.ClsDBInfo.ExecuteSQL(Sqlcommand, dt)

                    '*********************************
                    '**以上查出工單ID，回焊ID,機種ID**
                    '*********************************
                    If dt.Rows.Count > 0 Then
                        Dim workid As Integer
                        workid = dt.Rows(0).Item(1)
                        If work_listview.Items(0).SubItems(1).Text = "" Then
                            work_listview.Items(0).SubItems(0).Text = dt.Rows(0).Item(0).ToString
                            work_listview.Items(0).SubItems(1).Text = dt.Rows(0).Item(1).ToString
                            work_listview.Items(0).SubItems.Add(dt.Rows(0).Item(2))
                            work_listview.Items(0).SubItems.Add(dt.Rows(0).Item(3))
                            work_listview.Items(0).SubItems.Add(dt.Rows(0).Item(4))
                            'work_listview.Items(0).SubItems.Add(dt.Rows(0).Item(5))



                            Sqlcommand = "SELECT TO_CHAR (in_time, 'yyyy/mm/dd hh24:mi:ss') AS in_time," & _
                                          " TO_CHAR (out_time, 'yyyy/mm/dd hh24:mi:ss') AS out_time, msl_no " & _
                                          "  FROM smt.g_smt_travel " & _
                                           " WHERE reel_no='" & reel_listview.Items(k).SubItems(0).Text & "'"

                            Dim dt3 As DataTable = New DataTable("dt3")
                            Me.ClsDBInfo.ExecuteSQL(Sqlcommand, dt3)
                            If dt3.Rows.Count > 0 Then
                                work_listview.Items(0).SubItems.Add(dt3.Rows(0).Item("in_time"))
                                work_listview.Items(0).SubItems.Add(dt3.Rows(0).Item("out_time"))
                            Else
                                StatusBar1.Panels(0).Text = StrConv("此lot上下料時間無法確定", SpiChangeLanguage.Value, 9) & reel_listview.Items(k).SubItems(0).Text
                            End If
                            '**************************
                            '以上第一次查出列出表格里**
                            '**************************

                            '********************************************************************************
                            '**以下判斷是否有重複的工單，如果有則不顯示，繼續執行下個，如果沒有則列出表格里**
                            '********************************************************************************                  
                        Else
                            Dim j As Integer
                            j = work_listview.Items.Count
                            If workid <> work_listview.Items(j - 1).SubItems(1).Text Then
                                work_listview.Items.Add("")
                                work_listview.Items(j).SubItems(0).Text = dt.Rows(0).Item(0).ToString
                                work_listview.Items(j).SubItems.Add(dt.Rows(0).Item(1))
                                work_listview.Items(j).SubItems.Add(dt.Rows(0).Item(2))
                                work_listview.Items(j).SubItems.Add(dt.Rows(0).Item(3))
                                work_listview.Items(j).SubItems.Add(dt.Rows(0).Item(4))
                                'work_listview.Items(j).SubItems.Add(dt.Rows(0).Item(5))

                                Sqlcommand = "SELECT TO_CHAR (in_time, 'yyyy/mm/dd hh24:mi:ss') AS in_time," & _
                                                                          " TO_CHAR (out_time, 'yyyy/mm/dd hh24:mi:ss') AS out_time, msl_no " & _
                                                                          "  FROM smt.g_smt_travel " & _
                                                                           " WHERE reel_no='" & reel_listview.Items(k).SubItems(0).Text & "'"

                                Dim dt4 As DataTable = New DataTable("dt4")
                                Me.ClsDBInfo.ExecuteSQL(Sqlcommand, dt4)
                                If dt4.Rows.Count > 0 Then
                                    work_listview.Items(j).SubItems.Add(dt4.Rows(0).Item("in_time"))
                                    work_listview.Items(j).SubItems.Add(dt4.Rows(0).Item("out_time"))
                                End If



                            Else

                                '**********************************************
                                '**以下是更新時間，找到最早的時間和最晚的時間**
                                '**********************************************                            
                                Sqlcommand = "SELECT TO_CHAR (in_time, 'yyyy/mm/dd hh24:mi:ss') AS in_time," & _
                                              " TO_CHAR (out_time, 'yyyy/mm/dd hh24:mi:ss') AS out_time, msl_no " & _
                                              "  FROM smt.g_smt_travel " & _
                                              " WHERE reel_no='" & reel_listview.Items(k).SubItems(0).Text & "'"

                                Dim dt1 As DataTable = New DataTable("dt1")
                                Me.ClsDBInfo.ExecuteSQL(Sqlcommand, dt1)
                                If dt1.Rows.Count > 0 Then
                                    Dim in_time, out_time As String
                                    in_time = dt1.Rows(0).Item(0).ToString
                                    out_time = dt1.Rows(0).Item(1).ToString

                                    'If work_listview.Items(j - 1).SubItems(5).Text = "" Then
                                    'work_listview.Items(j - 1).SubItems(5).Text = dt1.Rows(0).Item(0).ToString
                                    'work_listview.Items(j - 1).SubItems.Add(dt1.Rows(0).Item(1))
                                    'work_listview.Items(j - 1).SubItems.Add(dt1.Rows(0).Item(2))
                                    'Else
                                    If in_time < work_listview.Items(j - 1).SubItems(5).Text Then
                                        work_listview.Items(j - 1).SubItems(5).Text = in_time
                                    End If
                                    If out_time > work_listview.Items(j - 1).SubItems(6).Text Then
                                        work_listview.Items(j - 1).SubItems(6).Text = out_time
                                    End If
                                    'End If
                                Else
                                    StatusBar1.Panels(0).Text = StrConv("此lot上下料時間無法確定", SpiChangeLanguage.Value, 9) & reel_listview.Items(k).SubItems(0).Text
                                End If
                            End If
                        End If

                    Else

                        StatusBar1.Panels(0).Text = StrConv("請確定此LOT是否為LED料，無法確定工單或回焊站", SpiChangeLanguage.Value, 9) & reel_listview.Items(k).SubItems(0).Text
                    End If

                Next


            End If
        Catch ex As Exception
            ErrorInfo.SysErrMessageInfomation(StrConv(ex.Message, SpiChangeLanguage.Value, 9))
        End Try
    End Sub
    Private Sub barcode_check()
        '**********************
        '**根據時間查過帳記錄**
        '**********************
        Try
            If work_listview.Items.Count <> 0 And work_listview.Items(0).SubItems(6).Text <> "" Then
                Dim a As Integer = 0
                For a = 0 To work_listview.Items.Count - 1
                    Sqlcommand = " SELECT   f003, f005    FROM mse0060" & _
                                 " WHERE f005 BETWEEN TO_DATE ('" & work_listview.Items(a).SubItems(5).Text & "', 'yyyy/mm/dd hh24:mi:ss')" & _
                                 " AND TO_DATE ('" & work_listview.Items(a).SubItems(6).Text & "', 'yyyy/mm/dd hh24:mi:ss')" & _
                                 " AND f004 = " & work_listview.Items(a).SubItems(4).Text & " " & _
                                 " AND f002 = " & work_listview.Items(a).SubItems(3).Text & "  ORDER BY f005"
                    Dim dt2 As DataTable = New DataTable("dt2")
                    Me.ClsDBInfo.ExecuteSQL(Sqlcommand, dt2)
                    If dt2.Rows.Count > 0 Then

                        Dim i As Integer = 0
                        Dim j As Integer
                        j = barcode_listview.Items.Count
                        For i = 0 To dt2.Rows.Count - 1
                            barcode_listview.Items.Add(i + j)
                            barcode_listview.Items(i + j).SubItems.Add(work_listview.Items(a).SubItems(0))
                            barcode_listview.Items(i + j).SubItems.Add(dt2.Rows(i).Item("f003"))
                            barcode_listview.Items(i + j).SubItems.Add(dt2.Rows(i).Item("f005"))
                        Next
                    Else
                        StatusBar1.Panels(0).Text = work_listview.Items(a).SubItems(0).Text & StrConv("工單無資料", SpiChangeLanguage.Value, 9)
                    End If
                Next
            End If
        Catch ex As Exception
            ErrorInfo.SysErrMessageInfomation(StrConv(ex.Message, SpiChangeLanguage.Value, 9))
        End Try

    End Sub
  



    Private Sub OpenFileDialog1_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk
        filename = Me.OpenFileDialog1.FileName
        Me.OpenFileDialog1.Dispose()

        Dim app As Object = CreateObject("Excel.Application") '定義excel實例
        Dim xlbook As Object = app.WorkBooks.Open(filename) '打開已經存在的工作薄
        Dim xlsheet As Object = xlbook.worksheets(1) '設置xlsheet為工作表1

        xlsheet.Activate() '設置xlsheet為當前工作表
        Dim numR As Integer = xlsheet.usedrange.rows.count
        Dim numC As Integer = xlsheet.usedrange.columns.count
        Dim num As Integer = 0
        If numC <> 1 Then
            MsgBox(StrConv("文檔格式有誤，請檢查", SpiChangeLanguage.Value, 9))

        End If

        Dim i As Integer = 0
        Dim j As Integer = 0
        For i = 1 To numR
            num = reel_listview.Items.Count
            reel_listview.Items.Insert(num, Trim(xlsheet.cells(i, 1).value))
        Next
        xlbook.close()
        app.Quit()

        app = Nothing
        GC.Collect()

        count.Text = reel_listview.Items.Count
        issearch = True
    End Sub


    Private Sub SaveFileDialog1_FileOk_1(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles SaveFileDialog1.FileOk
        filename = Me.SaveFileDialog1.FileName
        Me.SaveFileDialog1.Dispose()

        Dim app As Object = CreateObject("Excel.Application") '定義excel實例
        Dim xlbook As Object = app.workbooks.add() '增加新工作薄
        Dim xlsheet As Object = xlbook.worksheets(1) '設置xlsheet為工作表1
        Dim xlsheet2 As Object = xlbook.worksheets(2)
        Dim xlsheet3 As Object = xlbook.worksheets(2)
        xlsheet.Activate() '設置xlsheet為當前工作表
        Dim i As Integer
        Dim temp As String = ""
        If barcode_listview.Items.Count > 0 Then
            xlsheet.cells(1, 1) = "序號"
            xlsheet.cells(1, 2) = "工單"
            xlsheet.cells(1, 3) = "生產條碼"
            xlsheet.cells(1, 4) = "過帳時間"
            xlsheet.cells(1, 5) = "脆盤條碼"
            xlsheet.cells(1, 6) = "外箱條碼"

            For i = 0 To barcode_listview.Items.Count - 1
                xlsheet.cells(i + 2, 1) = barcode_listview.Items(i).SubItems(0).Text
                xlsheet.cells(i + 2, 2) = barcode_listview.Items(i).SubItems(1).Text
                xlsheet.cells(i + 2, 3) = barcode_listview.Items(i).SubItems(2).Text
                xlsheet.cells(i + 2, 4) = barcode_listview.Items(i).SubItems(3).Text



            Next
        End If
        app.Visible = False 'excel程序是否可見

        xlbook.SaveAs(filename) '存文檔
        xlbook.close()
        app.Quit()

        app = Nothing
        GC.Collect()

        updata()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
       
        If work_listview.Items.Count > 0 Or barcode_listview.Items.Count > 0 Then
            Dim yes_no As String
            yes_no = MsgBox(StrConv("是否清除資料", SpiChangeLanguage.Value, 9), MsgBoxStyle.OkCancel, StrConv("警告", SpiChangeLanguage.Value, 9))
            If yes_no = vbOK Then
                updata()
            End If
        End If

        If TextBox1.Text <> "" Then

            reel_listview.Items.Add(TextBox1.Text)
            count.Text = reel_listview.Items.Count
            TextBox1.Text = ""
            TextBox1.Focus()

        End If
        TextBox1.Focus()
        issearch = True



    End Sub
    Private Sub updata()
        count.Text = ""
        reel_listview.Items.Clear()
        work_listview.Items.Clear()
        barcode_listview.Items.Clear()
        TextBox1.Text = ""

        StatusBar1.Panels(0).Text = ""
        issearch = False


    End Sub

    Private Sub LEXC02_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        updata()
    End Sub
End Class