<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ADV
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
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ADV))
        Me.Txt = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.ans = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.tx1 = New System.Windows.Forms.TextBox
        Me.rest = New System.Windows.Forms.TextBox
        Me.fr = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.bsp = New System.Windows.Forms.TextBox
        Me.hlp = New System.Windows.Forms.TextBox
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.src = New System.Windows.Forms.TextBox
        Me.Timer5 = New System.Windows.Forms.Timer(Me.components)
        Me.SuspendLayout()
        '
        'Txt
        '
        Me.Txt.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.Txt.Enabled = False
        Me.Txt.Location = New System.Drawing.Point(166, 94)
        Me.Txt.Name = "Txt"
        Me.Txt.Size = New System.Drawing.Size(817, 20)
        Me.Txt.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label1.Location = New System.Drawing.Point(12, 45)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(73, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "SAY the Input"
        '
        'ans
        '
        Me.ans.Enabled = False
        Me.ans.Location = New System.Drawing.Point(583, 57)
        Me.ans.Name = "ans"
        Me.ans.Size = New System.Drawing.Size(400, 20)
        Me.ans.TabIndex = 2
        Me.ans.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label2.Location = New System.Drawing.Point(746, 277)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(44, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "READY"
        '
        'tx1
        '
        Me.tx1.BackColor = System.Drawing.SystemColors.Menu
        Me.tx1.Enabled = False
        Me.tx1.Location = New System.Drawing.Point(583, 193)
        Me.tx1.Name = "tx1"
        Me.tx1.Size = New System.Drawing.Size(400, 20)
        Me.tx1.TabIndex = 4
        '
        'rest
        '
        Me.rest.BackColor = System.Drawing.SystemColors.Menu
        Me.rest.Enabled = False
        Me.rest.Location = New System.Drawing.Point(583, 139)
        Me.rest.Name = "rest"
        Me.rest.Size = New System.Drawing.Size(400, 20)
        Me.rest.TabIndex = 5
        '
        'fr
        '
        Me.fr.BackColor = System.Drawing.SystemColors.Menu
        Me.fr.Enabled = False
        Me.fr.Location = New System.Drawing.Point(166, 251)
        Me.fr.Name = "fr"
        Me.fr.Size = New System.Drawing.Size(305, 20)
        Me.fr.TabIndex = 6
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label3.Location = New System.Drawing.Point(12, 258)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(37, 13)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Result"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label4.Location = New System.Drawing.Point(445, 139)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(40, 13)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "Input a"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label5.Location = New System.Drawing.Point(448, 196)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(39, 13)
        Me.Label5.TabIndex = 10
        Me.Label5.Text = "input b"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label6.Location = New System.Drawing.Point(12, 101)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(40, 13)
        Me.Label6.TabIndex = 11
        Me.Label6.Text = "Phrase"
        '
        'bsp
        '
        Me.bsp.Enabled = False
        Me.bsp.Location = New System.Drawing.Point(166, 12)
        Me.bsp.Name = "bsp"
        Me.bsp.Size = New System.Drawing.Size(586, 20)
        Me.bsp.TabIndex = 12
        Me.bsp.Visible = False
        '
        'hlp
        '
        Me.hlp.Enabled = False
        Me.hlp.Location = New System.Drawing.Point(786, 12)
        Me.hlp.Name = "hlp"
        Me.hlp.Size = New System.Drawing.Size(197, 20)
        Me.hlp.TabIndex = 13
        Me.hlp.Visible = False
        '
        'TextBox1
        '
        Me.TextBox1.Enabled = False
        Me.TextBox1.ForeColor = System.Drawing.Color.Red
        Me.TextBox1.Location = New System.Drawing.Point(166, 57)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(400, 20)
        Me.TextBox1.TabIndex = 14
        '
        'src
        '
        Me.src.Location = New System.Drawing.Point(583, 228)
        Me.src.Name = "src"
        Me.src.Size = New System.Drawing.Size(400, 20)
        Me.src.TabIndex = 15
        Me.src.Visible = False
        '
        'Timer5
        '
        Me.Timer5.Interval = 1000
        '
        'ADV
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.MintCream
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(990, 299)
        Me.Controls.Add(Me.src)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.hlp)
        Me.Controls.Add(Me.bsp)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.fr)
        Me.Controls.Add(Me.rest)
        Me.Controls.Add(Me.tx1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.ans)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Txt)
        Me.DoubleBuffered = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.HelpButton = True
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "ADV"
        Me.Text = "Advanced Voice Calc"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Txt As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ans As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents tx1 As System.Windows.Forms.TextBox
    Friend WithEvents rest As System.Windows.Forms.TextBox
    Friend WithEvents fr As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents bsp As System.Windows.Forms.TextBox
    Friend WithEvents hlp As System.Windows.Forms.TextBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents src As System.Windows.Forms.TextBox
    Friend WithEvents Timer5 As System.Windows.Forms.Timer
End Class
