'�z�w�w�w�w�w�s�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�{
'�x ���O�W�� �xLEXC02 �d��                                                                                                             �x
'�u�w�w�w�w�w�q�w�w�w�w�w�w�s�w�w�w�w�w�w�w�s�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�t
'�x ���     �x���g�H      �x������        �x���g���e                                                                                  �x
'�u�w�w�w�w�w�q�w�w�w�w�w�w�q�w�w�w�w�w�w�w�q�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�t
'�x2011/05/25�xtianen      �x2011.05.25.01 �xLEXC02�d��                                                                                �x
'�|�w�w�w�w�w�r�w�w�w�w�w�w�r�w�w�w�w�w�w�w�r�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�}
'2012.10.1.1 �b2011.5.27.1��¦�W�ק�A�W�[�c²���ഫ���


Public Class LEXC02

    Public Sqlcommand As String = ""
    Dim k As Integer
    Dim filename As String = ""
    Dim issearch As Boolean = False



    Private Sub ToolBar1_ButtonClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles ToolBar1.ButtonClick

        Select Case e.Button.Text
            Case StrConv("�d��", SpiChangeLanguage.Value, 9)
                If issearch = True Then
                    check()
                    barcode_check()

                    issearch = False
                End If


            Case StrConv("�פJ", SpiChangeLanguage.Value, 9)
                updata()
                If issearch = False Then
                    count.Text = ""
                    reel_listview.Items.Clear()
                    Me.OpenFileDialog1.InitialDirectory = "d:\"
                    Me.OpenFileDialog1.Filter = "EXECLE files (*.xls)|*.xls|All files (*.*)|*.*"
                    Me.OpenFileDialog1.FileName = ""
                    Me.OpenFileDialog1.ShowDialog()
                End If


            Case StrConv("����", SpiChangeLanguage.Value, 9)
                Me.Close()
            Case StrConv("�ץX", SpiChangeLanguage.Value, 9)

                If barcode_listview.Items.Count > 0 Then
                    Me.SaveFileDialog1.InitialDirectory = "d:\"
                    Me.SaveFileDialog1.Filter = "EXECLE files (*.xls)|*.xls|All files (*.*)|*.*"
                    Me.SaveFileDialog1.FileName = ""
                    Me.SaveFileDialog1.ShowDialog()


                Else
                    StatusBar1.Panels(0).Text = StrConv("�L��ƶץX", SpiChangeLanguage.Value, 9)
                End If

            Case StrConv("����", SpiChangeLanguage.Value, 9)

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
                    '**�H�W�d�X�u��ID�A�^�kID,����ID**
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
                                StatusBar1.Panels(0).Text = StrConv("��lot�W�U�Ʈɶ��L�k�T�w", SpiChangeLanguage.Value, 9) & reel_listview.Items(k).SubItems(0).Text
                            End If
                            '**************************
                            '�H�W�Ĥ@���d�X�C�X��樽**
                            '**************************

                            '********************************************************************************
                            '**�H�U�P�_�O�_�����ƪ��u��A�p�G���h����ܡA�~�����U�ӡA�p�G�S���h�C�X��樽**
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
                                '**�H�U�O��s�ɶ��A���̦����ɶ��M�̱ߪ��ɶ�**
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
                                    StatusBar1.Panels(0).Text = StrConv("��lot�W�U�Ʈɶ��L�k�T�w", SpiChangeLanguage.Value, 9) & reel_listview.Items(k).SubItems(0).Text
                                End If
                            End If
                        End If

                    Else

                        StatusBar1.Panels(0).Text = StrConv("�нT�w��LOT�O�_��LED�ơA�L�k�T�w�u��Φ^�k��", SpiChangeLanguage.Value, 9) & reel_listview.Items(k).SubItems(0).Text
                    End If

                Next


            End If
        Catch ex As Exception
            ErrorInfo.SysErrMessageInfomation(StrConv(ex.Message, SpiChangeLanguage.Value, 9))
        End Try
    End Sub
    Private Sub barcode_check()
        '**********************
        '**�ھڮɶ��d�L�b�O��**
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
                        StatusBar1.Panels(0).Text = work_listview.Items(a).SubItems(0).Text & StrConv("�u��L���", SpiChangeLanguage.Value, 9)
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

        Dim app As Object = CreateObject("Excel.Application") '�w�qexcel���
        Dim xlbook As Object = app.WorkBooks.Open(filename) '���}�w�g�s�b���u�@��
        Dim xlsheet As Object = xlbook.worksheets(1) '�]�mxlsheet���u�@��1

        xlsheet.Activate() '�]�mxlsheet����e�u�@��
        Dim numR As Integer = xlsheet.usedrange.rows.count
        Dim numC As Integer = xlsheet.usedrange.columns.count
        Dim num As Integer = 0
        If numC <> 1 Then
            MsgBox(StrConv("���ɮ榡���~�A���ˬd", SpiChangeLanguage.Value, 9))

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

        Dim app As Object = CreateObject("Excel.Application") '�w�qexcel���
        Dim xlbook As Object = app.workbooks.add() '�W�[�s�u�@��
        Dim xlsheet As Object = xlbook.worksheets(1) '�]�mxlsheet���u�@��1
        Dim xlsheet2 As Object = xlbook.worksheets(2)
        Dim xlsheet3 As Object = xlbook.worksheets(2)
        xlsheet.Activate() '�]�mxlsheet����e�u�@��
        Dim i As Integer
        Dim temp As String = ""
        If barcode_listview.Items.Count > 0 Then
            xlsheet.cells(1, 1) = "�Ǹ�"
            xlsheet.cells(1, 2) = "�u��"
            xlsheet.cells(1, 3) = "�Ͳ����X"
            xlsheet.cells(1, 4) = "�L�b�ɶ�"
            xlsheet.cells(1, 5) = "�ܽL���X"
            xlsheet.cells(1, 6) = "�~�c���X"

            For i = 0 To barcode_listview.Items.Count - 1
                xlsheet.cells(i + 2, 1) = barcode_listview.Items(i).SubItems(0).Text
                xlsheet.cells(i + 2, 2) = barcode_listview.Items(i).SubItems(1).Text
                xlsheet.cells(i + 2, 3) = barcode_listview.Items(i).SubItems(2).Text
                xlsheet.cells(i + 2, 4) = barcode_listview.Items(i).SubItems(3).Text



            Next
        End If
        app.Visible = False 'excel�{�ǬO�_�i��

        xlbook.SaveAs(filename) '�s����
        xlbook.close()
        app.Quit()

        app = Nothing
        GC.Collect()

        updata()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
       
        If work_listview.Items.Count > 0 Or barcode_listview.Items.Count > 0 Then
            Dim yes_no As String
            yes_no = MsgBox(StrConv("�O�_�M�����", SpiChangeLanguage.Value, 9), MsgBoxStyle.OkCancel, StrConv("ĵ�i", SpiChangeLanguage.Value, 9))
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