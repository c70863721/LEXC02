'�z�w�w�w�w�w�s�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�{
'�x ���O�W�� �xLEXC02 �d��                                                                                                             �x
'�u�w�w�w�w�w�q�w�w�w�w�w�w�s�w�w�w�w�w�w�w�s�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�t
'�x ���     �x���g�H      �x������        �x���g���e                                                                                  �x
'�u�w�w�w�w�w�q�w�w�w�w�w�w�q�w�w�w�w�w�w�w�q�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�t
'�x2011/05/08�xtianen      �x2011.05.08.01 �xLEXC02�d��                                                                                �x
'�|�w�w�w�w�w�r�w�w�w�w�w�w�r�w�w�w�w�w�w�w�r�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�}

Public Class LEXC02

    Public Sqlcommand As String = ""
    Dim k As Integer
    Dim filename As String = ""
    Dim part_id As Integer
    Dim liucheng As Integer

    Private Sub ToolBar1_ButtonClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles ToolBar1.ButtonClick

        Select Case e.Button.Text
            Case "�d��"
                check()

            Case "�פJ"
                count.Text = ""
                ' If reel_listview.Items.Count = 0 And Pass_listview.Items.Count = 0 And rtno_listview.Items.Count = 0 And barcode_listview.Items.Count = 0 Then
                Me.OpenFileDialog1.InitialDirectory = "d:\"
                Me.OpenFileDialog1.Filter = "EXECLE files (*.xls)|*.xls|All files (*.*)|*.*"
                Me.OpenFileDialog1.FileName = ""
                Me.OpenFileDialog1.ShowDialog()

            Case "����"
                Me.Close()
            Case "�ץX"
                'If lv1.Items.Count > 0 Then
                '    Me.SaveFileDialog1.InitialDirectory = "d:\"
                '    Me.SaveFileDialog1.Filter = "EXECLE files (*.xls)|*.xls|All files (*.*)|*.*"
                '    Me.SaveFileDialog1.FileName = ""
                '    Me.SaveFileDialog1.ShowDialog()
                'End If
        End Select


    End Sub
    Sub check()
        Try

            If Me.ClsDBInfo.datalink(Me.ClsCommonInfo.uLocalIP) Then
                For k = 0 To reel_listview.Items.Count - 1
                    Sqlcommand = "SELECT se17.f003, wk_ord_id, se08.f002,se17.f005 ,se11.f001" & _
                                  " FROM sajet.g_part_led_issue led, mse0017 se17, mse0008 se08,mse0011 se11" & _
                                   " WHERE led.part_sn = '" & reel_listview.Items(k).SubItems(0).Text & "'" & _
                                  " AND led.wk_ord_id = se17.f001" & _
                                  " AND se17.f005 = se08.f001" & _
                                  " and se17.f005 =se11.f002" & _
                                  " and se11.f004=6"
                    Dim dt As DataTable = New DataTable("dt")
                    Me.ClsDBInfo.ExecuteSQL(Sqlcommand, dt)
                    If dt.Rows.Count > 0 Then
                        Dim i, J As Integer
                        For i = 0 To dt.Rows.Count - 1
                            work_listview.Items.Add("")
                            work_listview.Items(0).SubItems.Add(dt.Rows(0).Item(0))
                            work_listview.Items(0).SubItems.Add(dt.Rows(0).Item(1))
                            work_listview.Items(0).SubItems.Add(dt.Rows(0).Item(2))
                            work_listview.Items(0).SubItems.Add(dt.Rows(0).Item(3))
                            work_listview.Items(0).SubItems.Add(dt.Rows(0).Item(4))
                            For J = 0 To work_listview.Items.Count - 1
                                If dt.Rows(i).Item(1).ToString = work_listview.Items(J).SubItems(3).Text Then
                                    Exit For
                                End If
                                work_listview.Items(i).SubItems.Add(dt.Rows(i).Item(0))
                                work_listview.Items(i).SubItems.Add(dt.Rows(i).Item(1))
                                work_listview.Items(i).SubItems.Add(dt.Rows(i).Item(2))
                                work_listview.Items(i).SubItems.Add(dt.Rows(i).Item(3))
                                work_listview.Items(i).SubItems.Add(dt.Rows(i).Item(4))
                                'work_listview.Items(i).SubItems.Add(reel)

                                'work_listview.Items(i).SubItems.Add(dt.Rows(i).Item("ends"))
                                'work_listview.Items(i).SubItems.Add(dt.Rows(i).Item("f005"))
                            Next
                        Next
                        Dim w As Integer = 0
                        Dim time(w) As String
                        For w = 0 To reel_listview.Items.Count


                            time(w) = reel_listview.Items(w).SubItems(0).Text




                        Next
                        Sqlcommand = "SELECT   TO_CHAR (in_time, 'yyyy/mm/dd hh24:mi:ss') AS in_time," & _
                                                                   " TO_CHAR (out_time, 'yyyy/mm/dd hh24:mi:ss') AS out_time, msl_no" & _
                                                                    "FROM smt.g_smt_travel" & _
                                                                     "WHERE reel_no IN ('" & time(w) & "')"


                        Dim dt1 As DataTable = New DataTable("dt1")
                        Me.ClsDBInfo.ExecuteSQL(Sqlcommand, dt1)
                        If dt1.Rows.Count > 0 Then

                        End If

                        'Sqlcommand = "SELECT TO_CHAR (in_time, 'yyyy/mm/dd hh24:mi:ss') AS in_time," & _
                        '              " TO_CHAR (out_time, 'yyyy/mm/dd hh24:mi:ss') AS out_time, msl_no " & _
                        '              "  FROM(smt.g_smt_travel) " & _
                        '              " WHERE reel_no='" & reel_listview.Items(k).SubItems(0).Text & "'"
                        'Dim dt1 As DataTable = New DataTable("dt1")
                        'Me.ClsDBInfo.ExecuteSQL(Sqlcommand, dt1)
                        'If dt1.Rows.Count > 0 Then
                        '    'For i = 0 To dt1.Rows.Count - 1
                        '    '    'For J = 0 To dt1.Rows.Count - 1
                        '    '    '    'If dt1.Rows(i).Item(0) > dt1.Rows(J).Item(0) Then

                        '    '    '    'End If

                        '    '    'Next
                        '    'Next
                        '    ' max(dt1.Rows.Item(0))





                        'Else
                        '    StatusBar1.Panels(0).Text = reel_listview.Items(k).SubItems(0).Text & "�ӮƸ��L�k�T�ߤW�Ʈɶ�"
                        'End If
                    Else
                        StatusBar1.Panels(0).Text = reel_listview.Items(k).SubItems(0).Text & "�L�k�T�ߤu��M�^�k���A�нT�w���L�o��"
                    End If
                Next
            End If

            '    Sqlcommand = "SELECT TO_CHAR (in_time, 'yyyy/mm/dd hh24:mi:ss') AS in_time," & _
            '                        "TO_CHAR (out_time, 'yyyy/mm/dd hh24:mi:ss') AS out_time" & _
            '                " FROM smt.g_smt_travel " & _
            '                " WHERE reel_no = '" & reel & "'"
            '    Dim dt1 As DataTable = New DataTable("dt1")
            '    Me.ClsDBInfo.ExecuteSQL(Sqlcommand, dt1)
            '    If dt1.Rows.Count > 0 Then
            '        Dim in_time As String
            '        Dim out_time As String
            '        in_time = dt1.Rows(0).Item(0)
            '        out_time = dt1.Rows(0).Item(1)
            '        Dim w As Integer
            '        For w = 0 To work_listview.Items.Count - 1


            '            Sqlcommand = "SELECT f003,f005  FROM mse0060 " & _
            '                         " WHERE f003 BETWEEN '" & work_listview.Items(k).SubItems(2).Text & "' AND '" & work_listview.Items(k).SubItems(3).Text & "'" & _
            '                            " AND f004 = " & liucheng & "   AND f002 = " & part_id & " " & _
            '                            " and f009 < to_date ('" & out_time & "','yyyy/mm/dd hh24:mi:ss') " & _
            '                            " AND f009 > TO_DATE ('" & in_time & "', 'yyyy/mm/dd hh24:mi:ss')" & _
            '                            " order by f005"
            '            Dim dt2 As DataTable = New DataTable("dt2")
            '            Me.ClsDBInfo.ExecuteSQL(Sqlcommand, dt2)
            '            If dt2.Rows.Count > 0 Then
            '                Dim z As Integer
            '                z = barcode_listview.Items.Count
            '                Dim j As Integer = 0
            '                For j = 0 To dt2.Rows.Count - 1
            '                    barcode_listview.Items.Add(j)
            '                    barcode_listview.Items(j + z).SubItems.Add(reel)
            '                    barcode_listview.Items(j + z).SubItems.Add(dt2.Rows(j).Item(0))
            '                    barcode_listview.Items(j + z).SubItems.Add(dt2.Rows(j).Item(1))
            '                Next

            '            Else
            '                StatusBar1.Panels(0).Text = "�L��ơA"
            '            End If







            '        Next

            '    Else
            '        StatusBar1.Panels(0).Text = "�нT�w" & reel & "�ƨ����L�W��"
            '    End If
            'Else
            '    StatusBar1.Panels(0).Text = reel & "�L�k�T�ߤu��A�нT�w���L�o�ƩΩw�q���X�d��"
            'End If
            '    Next
            'End If
        Catch ex As Exception
            ErrorInfo.SysErrMessageInfomation(ex.Message)
        End Try
    End Sub



    Private Sub OpenFileDialog1_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk
        filename = Me.OpenFileDialog1.FileName
        Me.OpenFileDialog1.Dispose()

        Dim app As Object = CreateObject("Excel.Application") '�w�qexcel���
        Dim xlbook As Object = app.WorkBooks.Open(filename) '���}�w�g�s�b���u�@��
        Dim xlsheet As Object = xlbook.worksheets(1) '�]�mxlsheet���u�@��1

        xlsheet.Activate() '�]�mxlsheet����e�u�@��
        Dim numR As Integer = xlsheet.usedrange.rows.count
        Dim numC As Integer = xlsheet.usedrange.columns.count
        Dim num As Integer = 0
        If numC <> 1 Then
            MsgBox("���ɮ榡���~�A���ˬd")

        End If

        Dim i As Integer = 0
        Dim j As Integer = 0


        For i = 1 To numR
            num = reel_listview.Items.Count
            reel_listview.Items.Insert(num, Trim(xlsheet.cells(i, 1).value))
        Next
        count.Text = reel_listview.Items.Count
    End Sub


    Private Sub SaveFileDialog1_FileOk_1(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles SaveFileDialog1.FileOk
        filename = Me.SaveFileDialog1.FileName
        Me.SaveFileDialog1.Dispose()

        Dim app As Object = CreateObject("Excel.Application") '�w�qexcel���
        Dim xlbook As Object = app.workbooks.add() '�W�[�s�u�@��
        Dim xlsheet As Object = xlbook.worksheets(1) '�]�mxlsheet���u�@��1
        Dim xlsheet2 As Object = xlbook.worksheets(2)
        Dim xlsheet3 As Object = xlbook.worksheets(2)
        xlsheet.Activate() '�]�mxlsheet����e�u�@��
        Dim i As Integer
        Dim temp As String = ""
        If reel_listview.Items.Count > 0 Then
            xlsheet.cells(1, 1) = "�Ͳ����X"
            xlsheet.cells(1, 2) = "�ƥe��"
            xlsheet.cells(1, 3) = "�Ƹ�"
            xlsheet.cells(1, 4) = "�ƨ��s��"
            xlsheet.cells(1, 5) = "BIN"
            xlsheet.cells(1, 6) = "���A"
            xlsheet.cells(1, 7) = "����"
            xlsheet.cells(1, 8) = "OP"
            xlsheet.cells(1, 9) = "�W�Ʈɶ�"
            xlsheet.cells(1, 10) = "�U�Ʈɶ�"
            xlsheet.cells(1, 11) = "�����ӧ帹"
            xlsheet.cells(1, 12) = "�ƶq"
            For i = 0 To reel_listview.Items.Count - 1
                'xlsheet.cells(i + 2, 1) = reel_listview.Items(i).SubItems(0).Text
                'xlsheet.cells(i + 2, 2) = reel_listview.Items(i).SubItems(1).Text
                'xlsheet.cells(i + 2, 3) = reel_listview.Items(i).SubItems(2).Text
                'xlsheet.cells(i + 2, 4) = reel_listview.Items(i).SubItems(3).Text
                'xlsheet.cells(i + 2, 5) = reel_listview.Items(i).SubItems(4).Text
                'xlsheet.cells(i + 2, 6) = reel_listview.Items(i).SubItems(5).Text
                'xlsheet.cells(i + 2, 7) = reel_listview.Items(i).SubItems(6).Text
                'xlsheet.cells(i + 2, 8) = reel_listview.Items(i).SubItems(7).Text
                'xlsheet.cells(i + 2, 9) = reel_listview.Items(i).SubItems(8).Text
                'xlsheet.cells(i + 2, 10) = reel_listview.Items(i).SubItems(9).Text
                'xlsheet.cells(i + 2, 11) = reel_listview.Items(i).SubItems(10).Text
                'xlsheet.cells(i + 2, 12) = reel_listview.Items(i).SubItems(11).Text
            Next
        End If
        app.Visible = False 'excel�{�ǬO�_�i��

        xlbook.SaveAs(filename) '�s����
        xlbook.close()
        app.Quit()

        app = Nothing
        GC.Collect()
        'issearch = True
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        reel_listview.Items.Add(TextBox1.Text)
        count.Text = reel_listview.Items.Count


    End Sub
End Class