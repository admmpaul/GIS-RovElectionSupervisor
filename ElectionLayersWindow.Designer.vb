<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ElectionLayersWindow
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ElectionLayersWindow))
        Me.CloseBtn = New System.Windows.Forms.Button()
        Me.StatusStrip = New System.Windows.Forms.StatusStrip()
        Me.TSSLabel = New System.Windows.Forms.ToolStripStatusLabel()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.tbElectionName = New System.Windows.Forms.TextBox()
        Me.rbOverwrite = New System.Windows.Forms.RadioButton()
        Me.rbAppend = New System.Windows.Forms.RadioButton()
        Me.Label7003 = New System.Windows.Forms.Label()
        Me.Label7011 = New System.Windows.Forms.Label()
        Me.UploadBtn = New System.Windows.Forms.Button()
        Me.R700_03TB = New System.Windows.Forms.TextBox()
        Me.R701_01TB = New System.Windows.Forms.TextBox()
        Me.Desc1Lbl = New System.Windows.Forms.Label()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.lblPollingPlace = New System.Windows.Forms.Label()
        Me.btnUploadPollingPlaceFile = New System.Windows.Forms.Button()
        Me.tbPollingPlaceFile = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.StatusStrip.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'CloseBtn
        '
        Me.CloseBtn.Location = New System.Drawing.Point(312, 356)
        Me.CloseBtn.Name = "CloseBtn"
        Me.CloseBtn.Size = New System.Drawing.Size(75, 23)
        Me.CloseBtn.TabIndex = 1
        Me.CloseBtn.Text = "Close"
        Me.CloseBtn.UseVisualStyleBackColor = True
        '
        'StatusStrip
        '
        Me.StatusStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TSSLabel})
        Me.StatusStrip.Location = New System.Drawing.Point(0, 384)
        Me.StatusStrip.Name = "StatusStrip"
        Me.StatusStrip.Size = New System.Drawing.Size(409, 22)
        Me.StatusStrip.Stretch = False
        Me.StatusStrip.TabIndex = 3
        Me.StatusStrip.Text = "Ready"
        '
        'TSSLabel
        '
        Me.TSSLabel.Name = "TSSLabel"
        Me.TSSLabel.Size = New System.Drawing.Size(39, 17)
        Me.TSSLabel.Text = "Ready"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.tbElectionName)
        Me.GroupBox1.Controls.Add(Me.rbOverwrite)
        Me.GroupBox1.Controls.Add(Me.rbAppend)
        Me.GroupBox1.Controls.Add(Me.Label7003)
        Me.GroupBox1.Controls.Add(Me.Label7011)
        Me.GroupBox1.Controls.Add(Me.UploadBtn)
        Me.GroupBox1.Controls.Add(Me.R700_03TB)
        Me.GroupBox1.Controls.Add(Me.R701_01TB)
        Me.GroupBox1.Controls.Add(Me.Desc1Lbl)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(384, 222)
        Me.GroupBox1.TabIndex = 6
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Upload of Precint Data"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(15, 135)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(77, 13)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "Election name:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tbElectionName
        '
        Me.tbElectionName.Location = New System.Drawing.Point(16, 151)
        Me.tbElectionName.Name = "tbElectionName"
        Me.tbElectionName.Size = New System.Drawing.Size(272, 20)
        Me.tbElectionName.TabIndex = 11
        '
        'rbOverwrite
        '
        Me.rbOverwrite.AutoSize = True
        Me.rbOverwrite.Location = New System.Drawing.Point(14, 178)
        Me.rbOverwrite.Name = "rbOverwrite"
        Me.rbOverwrite.Size = New System.Drawing.Size(132, 17)
        Me.rbOverwrite.TabIndex = 10
        Me.rbOverwrite.Text = "Overwrite existing data"
        Me.rbOverwrite.UseVisualStyleBackColor = True
        '
        'rbAppend
        '
        Me.rbAppend.AutoSize = True
        Me.rbAppend.Location = New System.Drawing.Point(14, 201)
        Me.rbAppend.Name = "rbAppend"
        Me.rbAppend.Size = New System.Drawing.Size(136, 17)
        Me.rbAppend.TabIndex = 9
        Me.rbAppend.Text = "Append to existing data"
        Me.rbAppend.UseVisualStyleBackColor = True
        '
        'Label7003
        '
        Me.Label7003.AutoSize = True
        Me.Label7003.Location = New System.Drawing.Point(15, 59)
        Me.Label7003.Name = "Label7003"
        Me.Label7003.Size = New System.Drawing.Size(109, 13)
        Me.Label7003.TabIndex = 8
        Me.Label7003.Text = "r700.09 Spreadsheet:"
        Me.Label7003.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label7011
        '
        Me.Label7011.AutoSize = True
        Me.Label7011.Location = New System.Drawing.Point(13, 97)
        Me.Label7011.Name = "Label7011"
        Me.Label7011.Size = New System.Drawing.Size(109, 13)
        Me.Label7011.TabIndex = 7
        Me.Label7011.Text = "r701.01 Spreadsheet:"
        Me.Label7011.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'UploadBtn
        '
        Me.UploadBtn.Location = New System.Drawing.Point(215, 187)
        Me.UploadBtn.Name = "UploadBtn"
        Me.UploadBtn.Size = New System.Drawing.Size(75, 23)
        Me.UploadBtn.TabIndex = 6
        Me.UploadBtn.Text = "Upload"
        Me.UploadBtn.UseVisualStyleBackColor = True
        '
        'R700_03TB
        '
        Me.R700_03TB.Location = New System.Drawing.Point(18, 74)
        Me.R700_03TB.Name = "R700_03TB"
        Me.R700_03TB.Size = New System.Drawing.Size(272, 20)
        Me.R700_03TB.TabIndex = 5
        '
        'R701_01TB
        '
        Me.R701_01TB.Location = New System.Drawing.Point(16, 112)
        Me.R701_01TB.Name = "R701_01TB"
        Me.R701_01TB.Size = New System.Drawing.Size(272, 20)
        Me.R701_01TB.TabIndex = 4
        '
        'Desc1Lbl
        '
        Me.Desc1Lbl.Location = New System.Drawing.Point(11, 16)
        Me.Desc1Lbl.Name = "Desc1Lbl"
        Me.Desc1Lbl.Size = New System.Drawing.Size(364, 45)
        Me.Desc1Lbl.TabIndex = 2
        Me.Desc1Lbl.Text = resources.GetString("Desc1Lbl.Text")
        Me.Desc1Lbl.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.lblPollingPlace)
        Me.GroupBox2.Controls.Add(Me.btnUploadPollingPlaceFile)
        Me.GroupBox2.Controls.Add(Me.tbPollingPlaceFile)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 250)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(384, 95)
        Me.GroupBox2.TabIndex = 7
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Load Polling Place layer"
        '
        'lblPollingPlace
        '
        Me.lblPollingPlace.AutoSize = True
        Me.lblPollingPlace.Location = New System.Drawing.Point(11, 49)
        Me.lblPollingPlace.Name = "lblPollingPlace"
        Me.lblPollingPlace.Size = New System.Drawing.Size(120, 13)
        Me.lblPollingPlace.TabIndex = 10
        Me.lblPollingPlace.Text = "Polling Place Export file:"
        Me.lblPollingPlace.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnUploadPollingPlaceFile
        '
        Me.btnUploadPollingPlaceFile.Location = New System.Drawing.Point(298, 63)
        Me.btnUploadPollingPlaceFile.Name = "btnUploadPollingPlaceFile"
        Me.btnUploadPollingPlaceFile.Size = New System.Drawing.Size(75, 23)
        Me.btnUploadPollingPlaceFile.TabIndex = 9
        Me.btnUploadPollingPlaceFile.Text = "Upload"
        Me.btnUploadPollingPlaceFile.UseVisualStyleBackColor = True
        '
        'tbPollingPlaceFile
        '
        Me.tbPollingPlaceFile.Location = New System.Drawing.Point(14, 64)
        Me.tbPollingPlaceFile.Name = "tbPollingPlaceFile"
        Me.tbPollingPlaceFile.Size = New System.Drawing.Size(272, 20)
        Me.tbPollingPlaceFile.TabIndex = 8
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(6, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(364, 40)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Load the Polling Place export file and create and new Polling Place layer.  Note " & _
    "that this will replace the existing Polling Place layer."
        '
        'ElectionLayersWindow
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(409, 406)
        Me.ControlBox = False
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.StatusStrip)
        Me.Controls.Add(Me.CloseBtn)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ElectionLayersWindow"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Election Layers Export Process"
        Me.StatusStrip.ResumeLayout(False)
        Me.StatusStrip.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents CloseBtn As System.Windows.Forms.Button
    Friend WithEvents StatusStrip As System.Windows.Forms.StatusStrip
    Friend WithEvents TSSLabel As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents R701_01TB As System.Windows.Forms.TextBox
    Friend WithEvents Desc1Lbl As System.Windows.Forms.Label
    Friend WithEvents R700_03TB As System.Windows.Forms.TextBox
    Friend WithEvents UploadBtn As System.Windows.Forms.Button
    Friend WithEvents Label7003 As System.Windows.Forms.Label
    Friend WithEvents Label7011 As System.Windows.Forms.Label
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblPollingPlace As System.Windows.Forms.Label
    Friend WithEvents btnUploadPollingPlaceFile As System.Windows.Forms.Button
    Friend WithEvents tbPollingPlaceFile As System.Windows.Forms.TextBox
    Friend WithEvents rbOverwrite As System.Windows.Forms.RadioButton
    Friend WithEvents rbAppend As System.Windows.Forms.RadioButton
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents tbElectionName As System.Windows.Forms.TextBox

End Class
