<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class LEXC02
    'Inherits System.Windows.Forms.Form
    Inherits CommonModule.BaseForm

    'Form 覆寫 Dispose 以清除元件清單。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    '為 Windows Form 設計工具的必要項
    Private components As System.ComponentModel.IContainer

    '注意: 以下為 Windows Form 設計工具所需的程序
    '可以使用 Windows Form 設計工具進行修改。
    '請不要使用程式碼編輯器進行修改。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(LEXC02))
        Me.lb2 = New System.Windows.Forms.Label
        Me.StatusBarPanel1 = New System.Windows.Forms.StatusBarPanel
        Me.StatusBar1 = New System.Windows.Forms.StatusBar
        Me.ToolBarButton14 = New System.Windows.Forms.ToolBarButton
        Me.ToolBarButton13 = New System.Windows.Forms.ToolBarButton
        Me.ToolBarButton12 = New System.Windows.Forms.ToolBarButton
        Me.ToolBarButton11 = New System.Windows.Forms.ToolBarButton
        Me.ToolBarButton10 = New System.Windows.Forms.ToolBarButton
        Me.ToolBarButton9 = New System.Windows.Forms.ToolBarButton
        Me.ToolBarButton8 = New System.Windows.Forms.ToolBarButton
        Me.ToolBarButton7 = New System.Windows.Forms.ToolBarButton
        Me.ToolBarButton6 = New System.Windows.Forms.ToolBarButton
        Me.ToolBarButton5 = New System.Windows.Forms.ToolBarButton
        Me.ToolBarButton4 = New System.Windows.Forms.ToolBarButton
        Me.ToolBarButton3 = New System.Windows.Forms.ToolBarButton
        Me.ToolBarButton2 = New System.Windows.Forms.ToolBarButton
        Me.ToolBarButton1 = New System.Windows.Forms.ToolBarButton
        Me.ToolBar1 = New System.Windows.Forms.ToolBar
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.reel_listview = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.barcode_listview = New System.Windows.Forms.ListView
        Me.ColumnHeader5 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader8 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader6 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader7 = New System.Windows.Forms.ColumnHeader
        Me.work_listview = New System.Windows.Forms.ListView
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader9 = New System.Windows.Forms.ColumnHeader
        Me.Label3 = New System.Windows.Forms.Label
        Me.count = New System.Windows.Forms.Label
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
        Me.Button1 = New System.Windows.Forms.Button
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.ColumnHeader11 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader10 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader12 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader13 = New System.Windows.Forms.ColumnHeader
        CType(Me.StatusBarPanel1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lb2
        '
        Me.lb2.AutoSize = True
        Me.lb2.ForeColor = System.Drawing.Color.Blue
        Me.lb2.Location = New System.Drawing.Point(365, 79)
        Me.lb2.Name = "lb2"
        Me.lb2.Size = New System.Drawing.Size(0, 12)
        Me.lb2.TabIndex = 87
        '
        'StatusBarPanel1
        '
        Me.StatusBarPanel1.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Spring
        Me.StatusBarPanel1.Icon = CType(resources.GetObject("StatusBarPanel1.Icon"), System.Drawing.Icon)
        Me.StatusBarPanel1.Name = "StatusBarPanel1"
        Me.StatusBarPanel1.Width = 671
        '
        'StatusBar1
        '
        Me.StatusBar1.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.StatusBar1.Location = New System.Drawing.Point(0, 504)
        Me.StatusBar1.Name = "StatusBar1"
        Me.StatusBar1.Panels.AddRange(New System.Windows.Forms.StatusBarPanel() {Me.StatusBarPanel1})
        Me.StatusBar1.ShowPanels = True
        Me.StatusBar1.Size = New System.Drawing.Size(671, 22)
        Me.StatusBar1.SizingGrip = False
        Me.StatusBar1.TabIndex = 23
        '
        'ToolBarButton14
        '
        Me.ToolBarButton14.ImageIndex = 13
        Me.ToolBarButton14.Name = "ToolBarButton14"
        Me.ToolBarButton14.Text = "結束"
        '
        'ToolBarButton13
        '
        Me.ToolBarButton13.ImageIndex = 12
        Me.ToolBarButton13.Name = "ToolBarButton13"
        Me.ToolBarButton13.Text = "重整"
        '
        'ToolBarButton12
        '
        Me.ToolBarButton12.Enabled = False
        Me.ToolBarButton12.ImageIndex = 11
        Me.ToolBarButton12.Name = "ToolBarButton12"
        Me.ToolBarButton12.Text = "存檔"
        '
        'ToolBarButton11
        '
        Me.ToolBarButton11.Enabled = False
        Me.ToolBarButton11.ImageIndex = 10
        Me.ToolBarButton11.Name = "ToolBarButton11"
        Me.ToolBarButton11.Text = "說明"
        '
        'ToolBarButton10
        '
        Me.ToolBarButton10.Enabled = False
        Me.ToolBarButton10.ImageIndex = 9
        Me.ToolBarButton10.Name = "ToolBarButton10"
        Me.ToolBarButton10.Text = "計算"
        '
        'ToolBarButton9
        '
        Me.ToolBarButton9.Enabled = False
        Me.ToolBarButton9.ImageIndex = 8
        Me.ToolBarButton9.Name = "ToolBarButton9"
        Me.ToolBarButton9.Text = "還原"
        '
        'ToolBarButton8
        '
        Me.ToolBarButton8.Enabled = False
        Me.ToolBarButton8.ImageIndex = 7
        Me.ToolBarButton8.Name = "ToolBarButton8"
        Me.ToolBarButton8.Text = "過帳"
        '
        'ToolBarButton7
        '
        Me.ToolBarButton7.ImageIndex = 6
        Me.ToolBarButton7.Name = "ToolBarButton7"
        Me.ToolBarButton7.Text = "匯出"
        '
        'ToolBarButton6
        '
        Me.ToolBarButton6.ImageIndex = 5
        Me.ToolBarButton6.Name = "ToolBarButton6"
        Me.ToolBarButton6.Text = "匯入"
        '
        'ToolBarButton5
        '
        Me.ToolBarButton5.Enabled = False
        Me.ToolBarButton5.ImageIndex = 4
        Me.ToolBarButton5.Name = "ToolBarButton5"
        Me.ToolBarButton5.Text = "列印"
        '
        'ToolBarButton4
        '
        Me.ToolBarButton4.ImageIndex = 3
        Me.ToolBarButton4.Name = "ToolBarButton4"
        Me.ToolBarButton4.Text = "查詢"
        '
        'ToolBarButton3
        '
        Me.ToolBarButton3.Enabled = False
        Me.ToolBarButton3.ImageIndex = 2
        Me.ToolBarButton3.Name = "ToolBarButton3"
        Me.ToolBarButton3.Text = "刪除"
        '
        'ToolBarButton2
        '
        Me.ToolBarButton2.Enabled = False
        Me.ToolBarButton2.ImageIndex = 1
        Me.ToolBarButton2.Name = "ToolBarButton2"
        Me.ToolBarButton2.Text = "修改"
        '
        'ToolBarButton1
        '
        Me.ToolBarButton1.Enabled = False
        Me.ToolBarButton1.ImageIndex = 0
        Me.ToolBarButton1.Name = "ToolBarButton1"
        Me.ToolBarButton1.Text = "新增"
        '
        'ToolBar1
        '
        Me.ToolBar1.AutoSize = False
        Me.ToolBar1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ToolBar1.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.ToolBarButton1, Me.ToolBarButton2, Me.ToolBarButton3, Me.ToolBarButton4, Me.ToolBarButton5, Me.ToolBarButton6, Me.ToolBarButton7, Me.ToolBarButton8, Me.ToolBarButton9, Me.ToolBarButton10, Me.ToolBarButton11, Me.ToolBarButton12, Me.ToolBarButton13, Me.ToolBarButton14})
        Me.ToolBar1.ButtonSize = New System.Drawing.Size(40, 35)
        Me.ToolBar1.Divider = False
        Me.ToolBar1.DropDownArrows = True
        Me.ToolBar1.Font = New System.Drawing.Font("新細明體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.ToolBar1.ImageList = Me.ImageList1
        Me.ToolBar1.Location = New System.Drawing.Point(0, 0)
        Me.ToolBar1.Name = "ToolBar1"
        Me.ToolBar1.ShowToolTips = True
        Me.ToolBar1.Size = New System.Drawing.Size(671, 42)
        Me.ToolBar1.TabIndex = 21
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList1.Images.SetKeyName(0, "")
        Me.ImageList1.Images.SetKeyName(1, "")
        Me.ImageList1.Images.SetKeyName(2, "")
        Me.ImageList1.Images.SetKeyName(3, "")
        Me.ImageList1.Images.SetKeyName(4, "")
        Me.ImageList1.Images.SetKeyName(5, "")
        Me.ImageList1.Images.SetKeyName(6, "")
        Me.ImageList1.Images.SetKeyName(7, "")
        Me.ImageList1.Images.SetKeyName(8, "")
        Me.ImageList1.Images.SetKeyName(9, "")
        Me.ImageList1.Images.SetKeyName(10, "")
        Me.ImageList1.Images.SetKeyName(11, "")
        Me.ImageList1.Images.SetKeyName(12, "")
        Me.ImageList1.Images.SetKeyName(13, "")
        '
        'reel_listview
        '
        Me.reel_listview.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1})
        Me.reel_listview.FullRowSelect = True
        Me.reel_listview.GridLines = True
        Me.reel_listview.Location = New System.Drawing.Point(2, 107)
        Me.reel_listview.Name = "reel_listview"
        Me.reel_listview.Size = New System.Drawing.Size(132, 397)
        Me.reel_listview.TabIndex = 103
        Me.reel_listview.UseCompatibleStateImageBehavior = False
        Me.reel_listview.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "料卷編號"
        Me.ColumnHeader1.Width = 125
        '
        'barcode_listview
        '
        Me.barcode_listview.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader5, Me.ColumnHeader8, Me.ColumnHeader6, Me.ColumnHeader7})
        Me.barcode_listview.FullRowSelect = True
        Me.barcode_listview.GridLines = True
        Me.barcode_listview.Location = New System.Drawing.Point(140, 218)
        Me.barcode_listview.Name = "barcode_listview"
        Me.barcode_listview.Size = New System.Drawing.Size(530, 282)
        Me.barcode_listview.TabIndex = 104
        Me.barcode_listview.UseCompatibleStateImageBehavior = False
        Me.barcode_listview.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "序號"
        Me.ColumnHeader5.Width = 42
        '
        'ColumnHeader8
        '
        Me.ColumnHeader8.Text = "料卷編號"
        Me.ColumnHeader8.Width = 147
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "生產條碼"
        Me.ColumnHeader6.Width = 148
        '
        'ColumnHeader7
        '
        Me.ColumnHeader7.Text = "時間"
        Me.ColumnHeader7.Width = 70
        '
        'work_listview
        '
        Me.work_listview.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4, Me.ColumnHeader9, Me.ColumnHeader11, Me.ColumnHeader10, Me.ColumnHeader12, Me.ColumnHeader13})
        Me.work_listview.FullRowSelect = True
        Me.work_listview.GridLines = True
        Me.work_listview.Location = New System.Drawing.Point(140, 90)
        Me.work_listview.Name = "work_listview"
        Me.work_listview.Size = New System.Drawing.Size(530, 122)
        Me.work_listview.TabIndex = 105
        Me.work_listview.UseCompatibleStateImageBehavior = False
        Me.work_listview.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "工單"
        Me.ColumnHeader2.Width = 87
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "工單ID"
        Me.ColumnHeader3.Width = 80
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "機種"
        Me.ColumnHeader4.Width = 109
        '
        'ColumnHeader9
        '
        Me.ColumnHeader9.Text = "機種ID"
        Me.ColumnHeader9.Width = 92
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label3.Font = New System.Drawing.Font("新細明體", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label3.Location = New System.Drawing.Point(-1, 90)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(86, 15)
        Me.Label3.TabIndex = 106
        Me.Label3.Text = "導入的數量:"
        '
        'count
        '
        Me.count.AutoSize = True
        Me.count.Location = New System.Drawing.Point(90, 92)
        Me.count.Name = "count"
        Me.count.Size = New System.Drawing.Size(11, 12)
        Me.count.TabIndex = 107
        Me.count.Text = "0"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'SaveFileDialog1
        '
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("新細明體", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Button1.Location = New System.Drawing.Point(299, 49)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(109, 33)
        Me.Button1.TabIndex = 102
        Me.Button1.Text = "增加"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(95, 54)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(190, 22)
        Me.TextBox1.TabIndex = 103
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("新細明體", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label2.Location = New System.Drawing.Point(9, 56)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(76, 16)
        Me.Label2.TabIndex = 101
        Me.Label2.Text = "料卷編號:"
        '
        'TextBox2
        '
        Me.TextBox2.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.TextBox2.Enabled = False
        Me.TextBox2.Location = New System.Drawing.Point(2, 44)
        Me.TextBox2.Multiline = True
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(668, 43)
        Me.TextBox2.TabIndex = 109
        '
        'ColumnHeader11
        '
        Me.ColumnHeader11.Text = "回焊ID"
        Me.ColumnHeader11.Width = 81
        '
        'ColumnHeader10
        '
        Me.ColumnHeader10.Text = "最早時間"
        Me.ColumnHeader10.Width = 90
        '
        'ColumnHeader12
        '
        Me.ColumnHeader12.Text = "最晚時間"
        Me.ColumnHeader12.Width = 83
        '
        'ColumnHeader13
        '
        Me.ColumnHeader13.Text = "料占表"
        '
        'LEXC02
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(671, 526)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.count)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.work_listview)
        Me.Controls.Add(Me.barcode_listview)
        Me.Controls.Add(Me.reel_listview)
        Me.Controls.Add(Me.lb2)
        Me.Controls.Add(Me.StatusBar1)
        Me.Controls.Add(Me.ToolBar1)
        Me.MaximizeBox = False
        Me.Name = "LEXC02"
        Me.Text = "隆達料卷查詢barcode"
        CType(Me.StatusBarPanel1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lb2 As System.Windows.Forms.Label
    Friend WithEvents StatusBarPanel1 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents StatusBar1 As System.Windows.Forms.StatusBar
    Friend WithEvents ToolBarButton14 As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButton13 As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButton12 As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButton11 As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButton10 As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButton9 As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButton8 As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButton7 As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButton6 As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButton5 As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButton4 As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButton3 As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButton2 As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButton1 As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBar1 As System.Windows.Forms.ToolBar
    Friend WithEvents reel_listview As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents barcode_listview As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader7 As System.Windows.Forms.ColumnHeader
    Friend WithEvents work_listview As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader9 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents count As System.Windows.Forms.Label
    Friend WithEvents ColumnHeader8 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents ColumnHeader11 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader10 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader12 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader13 As System.Windows.Forms.ColumnHeader
End Class
