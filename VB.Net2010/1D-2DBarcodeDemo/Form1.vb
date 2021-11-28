Option Strict Off
Option Explicit On
Friend Class Form1
	Inherits System.Windows.Forms.Form
#Region "Windows Form Designer generated code "
	Public Sub New()
		MyBase.New()
		If m_vb6FormDefInstance Is Nothing Then
			If m_InitializingDefInstance Then
				m_vb6FormDefInstance = Me
			Else
				Try 
					'For the start-up form, the first instance created is the default instance.
					If System.Reflection.Assembly.GetExecutingAssembly.EntryPoint.DeclaringType Is Me.GetType Then
						m_vb6FormDefInstance = Me
					End If
				Catch
				End Try
			End If
		End If
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents Command12 As System.Windows.Forms.Button
    Public WithEvents Command11 As System.Windows.Forms.Button
    Public WithEvents Command7 As System.Windows.Forms.Button
    Public WithEvents Command6 As System.Windows.Forms.Button
    Public WithEvents Command8 As System.Windows.Forms.Button
    Public WithEvents Command9 As System.Windows.Forms.Button
    Public WithEvents Command10 As System.Windows.Forms.Button
    Public WithEvents Command5 As System.Windows.Forms.Button
    Public WithEvents Command20 As System.Windows.Forms.Button
    Public WithEvents Command3 As System.Windows.Forms.Button
    Public WithEvents Command2 As System.Windows.Forms.Button
    Public WithEvents Command1 As System.Windows.Forms.Button
    Public WithEvents Scanner1 As AxSCANNERLib.AxScanner
    'Public WithEvents Frame1 As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    Friend WithEvents ColorDialog1 As System.Windows.Forms.ColorDialog
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cbopdfimagecount As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cbopdfpage As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents chkscanmore As System.Windows.Forms.CheckBox
    Friend WithEvents txtcustomrotation As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents chkScan90Rotation As System.Windows.Forms.CheckBox
    Friend WithEvents chkScan45CounterRotation As System.Windows.Forms.CheckBox
    Friend WithEvents chkScan45Rotation As System.Windows.Forms.CheckBox
    Friend WithEvents chkscannorotation As System.Windows.Forms.CheckBox
    Friend WithEvents chkScanMultiple As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents Button7 As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button6 As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Command12 = New System.Windows.Forms.Button()
        Me.Command11 = New System.Windows.Forms.Button()
        Me.Command7 = New System.Windows.Forms.Button()
        Me.Command6 = New System.Windows.Forms.Button()
        Me.Command8 = New System.Windows.Forms.Button()
        Me.Command9 = New System.Windows.Forms.Button()
        Me.Command10 = New System.Windows.Forms.Button()
        Me.Command5 = New System.Windows.Forms.Button()
        Me.Command20 = New System.Windows.Forms.Button()
        Me.Command3 = New System.Windows.Forms.Button()
        Me.Command2 = New System.Windows.Forms.Button()
        Me.Command1 = New System.Windows.Forms.Button()
        Me.Scanner1 = New AxSCANNERLib.AxScanner()
        Me.ColorDialog1 = New System.Windows.Forms.ColorDialog()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.cbopdfimagecount = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cbopdfpage = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.chkscanmore = New System.Windows.Forms.CheckBox()
        Me.txtcustomrotation = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.chkScan90Rotation = New System.Windows.Forms.CheckBox()
        Me.chkScan45CounterRotation = New System.Windows.Forms.CheckBox()
        Me.chkScan45Rotation = New System.Windows.Forms.CheckBox()
        Me.chkscannorotation = New System.Windows.Forms.CheckBox()
        Me.chkScanMultiple = New System.Windows.Forms.CheckBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button7 = New System.Windows.Forms.Button()
        CType(Me.Scanner1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'Command12
        '
        Me.Command12.BackColor = System.Drawing.SystemColors.Control
        Me.Command12.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command12.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command12.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command12.Location = New System.Drawing.Point(8, 185)
        Me.Command12.Name = "Command12"
        Me.Command12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command12.Size = New System.Drawing.Size(89, 25)
        Me.Command12.TabIndex = 12
        Me.Command12.Text = "25% of Original"
        Me.Command12.UseVisualStyleBackColor = False
        '
        'Command11
        '
        Me.Command11.BackColor = System.Drawing.SystemColors.Control
        Me.Command11.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command11.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command11.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command11.Location = New System.Drawing.Point(8, 217)
        Me.Command11.Name = "Command11"
        Me.Command11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command11.Size = New System.Drawing.Size(89, 25)
        Me.Command11.TabIndex = 11
        Me.Command11.Text = "33% of Original"
        Me.Command11.UseVisualStyleBackColor = False
        '
        'Command7
        '
        Me.Command7.BackColor = System.Drawing.SystemColors.Control
        Me.Command7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command7.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command7.Location = New System.Drawing.Point(8, 313)
        Me.Command7.Name = "Command7"
        Me.Command7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command7.Size = New System.Drawing.Size(89, 25)
        Me.Command7.TabIndex = 10
        Me.Command7.Text = "100% of Original"
        Me.Command7.UseVisualStyleBackColor = False
        '
        'Command6
        '
        Me.Command6.BackColor = System.Drawing.SystemColors.Control
        Me.Command6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command6.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command6.Location = New System.Drawing.Point(8, 345)
        Me.Command6.Name = "Command6"
        Me.Command6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command6.Size = New System.Drawing.Size(89, 25)
        Me.Command6.TabIndex = 9
        Me.Command6.Text = "150% of Original"
        Me.Command6.UseVisualStyleBackColor = False
        '
        'Command8
        '
        Me.Command8.BackColor = System.Drawing.SystemColors.Control
        Me.Command8.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command8.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command8.Location = New System.Drawing.Point(8, 249)
        Me.Command8.Name = "Command8"
        Me.Command8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command8.Size = New System.Drawing.Size(89, 25)
        Me.Command8.TabIndex = 8
        Me.Command8.Text = "50% of Original"
        Me.Command8.UseVisualStyleBackColor = False
        '
        'Command9
        '
        Me.Command9.BackColor = System.Drawing.SystemColors.Control
        Me.Command9.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command9.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command9.Location = New System.Drawing.Point(8, 281)
        Me.Command9.Name = "Command9"
        Me.Command9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command9.Size = New System.Drawing.Size(89, 25)
        Me.Command9.TabIndex = 7
        Me.Command9.Text = "75% of Original"
        Me.Command9.UseVisualStyleBackColor = False
        '
        'Command10
        '
        Me.Command10.BackColor = System.Drawing.SystemColors.Control
        Me.Command10.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command10.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command10.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command10.Location = New System.Drawing.Point(8, 377)
        Me.Command10.Name = "Command10"
        Me.Command10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command10.Size = New System.Drawing.Size(89, 25)
        Me.Command10.TabIndex = 6
        Me.Command10.Text = "200% of Original"
        Me.Command10.UseVisualStyleBackColor = False
        '
        'Command5
        '
        Me.Command5.BackColor = System.Drawing.SystemColors.Control
        Me.Command5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command5.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command5.Location = New System.Drawing.Point(8, 409)
        Me.Command5.Name = "Command5"
        Me.Command5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command5.Size = New System.Drawing.Size(89, 25)
        Me.Command5.TabIndex = 5
        Me.Command5.Text = "Fit to Window"
        Me.Command5.UseVisualStyleBackColor = False
        '
        'Command20
        '
        Me.Command20.BackColor = System.Drawing.SystemColors.Control
        Me.Command20.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command20.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command20.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command20.Location = New System.Drawing.Point(8, 440)
        Me.Command20.Name = "Command20"
        Me.Command20.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command20.Size = New System.Drawing.Size(89, 25)
        Me.Command20.TabIndex = 4
        Me.Command20.Text = "Aspect Ratio"
        Me.Command20.UseVisualStyleBackColor = False
        '
        'Command3
        '
        Me.Command3.BackColor = System.Drawing.SystemColors.Control
        Me.Command3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command3.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command3.Location = New System.Drawing.Point(8, 153)
        Me.Command3.Name = "Command3"
        Me.Command3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command3.Size = New System.Drawing.Size(89, 25)
        Me.Command3.TabIndex = 3
        Me.Command3.Text = "Rotate 270"
        Me.Command3.UseVisualStyleBackColor = False
        '
        'Command2
        '
        Me.Command2.BackColor = System.Drawing.SystemColors.Control
        Me.Command2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command2.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command2.Location = New System.Drawing.Point(8, 121)
        Me.Command2.Name = "Command2"
        Me.Command2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command2.Size = New System.Drawing.Size(89, 25)
        Me.Command2.TabIndex = 2
        Me.Command2.Text = "Rotate 180"
        Me.Command2.UseVisualStyleBackColor = False
        '
        'Command1
        '
        Me.Command1.BackColor = System.Drawing.SystemColors.Control
        Me.Command1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command1.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command1.Location = New System.Drawing.Point(8, 89)
        Me.Command1.Name = "Command1"
        Me.Command1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command1.Size = New System.Drawing.Size(89, 25)
        Me.Command1.TabIndex = 1
        Me.Command1.Text = "Rotate 90"
        Me.Command1.UseVisualStyleBackColor = False
        '
        'Scanner1
        '
        Me.Scanner1.Enabled = True
        Me.Scanner1.Location = New System.Drawing.Point(112, 78)
        Me.Scanner1.Name = "Scanner1"
        Me.Scanner1.OcxState = CType(resources.GetObject("Scanner1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.Scanner1.Size = New System.Drawing.Size(425, 498)
        Me.Scanner1.TabIndex = 0
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(15, 12)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(129, 60)
        Me.Button2.TabIndex = 117
        Me.Button2.Text = "Read QR Code Demo"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(169, 12)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(154, 60)
        Me.Button3.TabIndex = 118
        Me.Button3.Text = "Read 1D Barcode Demo"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(342, 12)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(154, 60)
        Me.Button4.TabIndex = 119
        Me.Button4.Text = "Read Multiple Barcode Demo"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(522, 12)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(154, 60)
        Me.Button5.TabIndex = 120
        Me.Button5.Text = "Read DataMatrix Demo"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Button6
        '
        Me.Button6.Location = New System.Drawing.Point(696, 12)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(154, 60)
        Me.Button6.TabIndex = 121
        Me.Button6.Text = "Read barcode at any orientation"
        Me.Button6.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.cbopdfimagecount)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.cbopdfpage)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(543, 250)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(364, 88)
        Me.GroupBox1.TabIndex = 122
        Me.GroupBox1.TabStop = False
        '
        'cbopdfimagecount
        '
        Me.cbopdfimagecount.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbopdfimagecount.FormattingEnabled = True
        Me.cbopdfimagecount.Location = New System.Drawing.Point(266, 34)
        Me.cbopdfimagecount.Name = "cbopdfimagecount"
        Me.cbopdfimagecount.Size = New System.Drawing.Size(92, 22)
        Me.cbopdfimagecount.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(191, 38)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(68, 14)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Select Image"
        '
        'cbopdfpage
        '
        Me.cbopdfpage.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbopdfpage.FormattingEnabled = True
        Me.cbopdfpage.Location = New System.Drawing.Point(82, 33)
        Me.cbopdfpage.Name = "cbopdfpage"
        Me.cbopdfpage.Size = New System.Drawing.Size(92, 22)
        Me.cbopdfpage.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(7, 37)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 14)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Select Page"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.chkscanmore)
        Me.GroupBox2.Controls.Add(Me.txtcustomrotation)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.chkScan90Rotation)
        Me.GroupBox2.Controls.Add(Me.chkScan45CounterRotation)
        Me.GroupBox2.Controls.Add(Me.chkScan45Rotation)
        Me.GroupBox2.Controls.Add(Me.chkscannorotation)
        Me.GroupBox2.Controls.Add(Me.chkScanMultiple)
        Me.GroupBox2.Location = New System.Drawing.Point(543, 89)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(576, 147)
        Me.GroupBox2.TabIndex = 125
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "UnSelected some option, it will scan more faster"
        '
        'chkscanmore
        '
        Me.chkscanmore.AutoSize = True
        Me.chkscanmore.Checked = True
        Me.chkscanmore.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkscanmore.Location = New System.Drawing.Point(264, 121)
        Me.chkscanmore.Name = "chkscanmore"
        Me.chkscanmore.Size = New System.Drawing.Size(128, 18)
        Me.chkscanmore.TabIndex = 7
        Me.chkscanmore.Text = "Scan More Accuracy"
        Me.chkscanmore.UseVisualStyleBackColor = True
        '
        'txtcustomrotation
        '
        Me.txtcustomrotation.Location = New System.Drawing.Point(386, 93)
        Me.txtcustomrotation.Name = "txtcustomrotation"
        Me.txtcustomrotation.Size = New System.Drawing.Size(128, 20)
        Me.txtcustomrotation.TabIndex = 6
        Me.txtcustomrotation.Text = "0"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(270, 97)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(85, 14)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Custom Rotation"
        '
        'chkScan90Rotation
        '
        Me.chkScan90Rotation.AutoSize = True
        Me.chkScan90Rotation.Checked = True
        Me.chkScan90Rotation.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkScan90Rotation.Location = New System.Drawing.Point(264, 62)
        Me.chkScan90Rotation.Name = "chkScan90Rotation"
        Me.chkScan90Rotation.Size = New System.Drawing.Size(170, 18)
        Me.chkScan90Rotation.TabIndex = 4
        Me.chkScan90Rotation.Text = "Scan With 90 Degree Rotation"
        Me.chkScan90Rotation.UseVisualStyleBackColor = True
        '
        'chkScan45CounterRotation
        '
        Me.chkScan45CounterRotation.AutoSize = True
        Me.chkScan45CounterRotation.Checked = True
        Me.chkScan45CounterRotation.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkScan45CounterRotation.Location = New System.Drawing.Point(264, 29)
        Me.chkScan45CounterRotation.Name = "chkScan45CounterRotation"
        Me.chkScan45CounterRotation.Size = New System.Drawing.Size(264, 18)
        Me.chkScan45CounterRotation.TabIndex = 3
        Me.chkScan45CounterRotation.Text = "Scan With 45 Degree Counter Clockwise Rotation"
        Me.chkScan45CounterRotation.UseVisualStyleBackColor = True
        '
        'chkScan45Rotation
        '
        Me.chkScan45Rotation.AutoSize = True
        Me.chkScan45Rotation.Checked = True
        Me.chkScan45Rotation.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkScan45Rotation.Location = New System.Drawing.Point(25, 97)
        Me.chkScan45Rotation.Name = "chkScan45Rotation"
        Me.chkScan45Rotation.Size = New System.Drawing.Size(170, 18)
        Me.chkScan45Rotation.TabIndex = 2
        Me.chkScan45Rotation.Text = "Scan With 45 Degree Rotation"
        Me.chkScan45Rotation.UseVisualStyleBackColor = True
        '
        'chkscannorotation
        '
        Me.chkscannorotation.AutoSize = True
        Me.chkscannorotation.Checked = True
        Me.chkscannorotation.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkscannorotation.Location = New System.Drawing.Point(25, 62)
        Me.chkscannorotation.Name = "chkscannorotation"
        Me.chkscannorotation.Size = New System.Drawing.Size(132, 18)
        Me.chkscannorotation.TabIndex = 1
        Me.chkscannorotation.Text = "Scan Without Rotation"
        Me.chkscannorotation.UseVisualStyleBackColor = True
        '
        'chkScanMultiple
        '
        Me.chkScanMultiple.AutoSize = True
        Me.chkScanMultiple.Checked = True
        Me.chkScanMultiple.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkScanMultiple.Location = New System.Drawing.Point(25, 29)
        Me.chkScanMultiple.Name = "chkScanMultiple"
        Me.chkScanMultiple.Size = New System.Drawing.Size(133, 18)
        Me.chkScanMultiple.TabIndex = 0
        Me.chkScanMultiple.Text = "Scan Multiple Barcode"
        Me.chkScanMultiple.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Button7)
        Me.GroupBox3.Controls.Add(Me.Button1)
        Me.GroupBox3.Location = New System.Drawing.Point(546, 352)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(389, 128)
        Me.GroupBox3.TabIndex = 126
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "If BarCodeReadFullPage cannot detect the barcode, try use for Advanced Barcode re" & _
            "ader"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(7, 46)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(170, 60)
        Me.Button1.TabIndex = 124
        Me.Button1.Text = "Read PDF Demo"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button7
        '
        Me.Button7.Location = New System.Drawing.Point(191, 46)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(192, 59)
        Me.Button7.TabIndex = 125
        Me.Button7.Text = "Scan for Advanced Barcode reader  "
        Me.Button7.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(1148, 682)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Command12)
        Me.Controls.Add(Me.Command11)
        Me.Controls.Add(Me.Command7)
        Me.Controls.Add(Me.Command6)
        Me.Controls.Add(Me.Command8)
        Me.Controls.Add(Me.Command9)
        Me.Controls.Add(Me.Command10)
        Me.Controls.Add(Me.Command5)
        Me.Controls.Add(Me.Command20)
        Me.Controls.Add(Me.Command3)
        Me.Controls.Add(Me.Command2)
        Me.Controls.Add(Me.Command1)
        Me.Controls.Add(Me.Scanner1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(4, 30)
        Me.Name = "Form1"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "Read 1D and 2D Barcode"
        CType(Me.Scanner1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
#End Region
#Region "Upgrade Support "
    Private Shared m_vb6FormDefInstance As Form1
    Private Shared m_InitializingDefInstance As Boolean
    Public Shared Property DefInstance() As Form1
        Get
            If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
                m_InitializingDefInstance = True
                m_vb6FormDefInstance = New Form1()
                m_InitializingDefInstance = False
            End If
            DefInstance = m_vb6FormDefInstance
        End Get
        Set(ByVal value As Form1)
            m_vb6FormDefInstance = value
        End Set
    End Property
#End Region
    Dim iX As Short
    Dim iY As Short

    Dim clrHashForeColor As Object
    Dim clrHashBackColor As Object

    Private Structure PictDesc
        Dim cbSizeofStruct As Integer
        Dim picType As Integer
        Dim hImage As Integer
        Dim xExt As Integer
        Dim yExt As Integer
    End Structure

    Private Structure Guid
        Dim Data1 As Integer
        Dim Data2 As Short
        Dim Data3 As Short
        <VBFixedArray(7)> Dim Data4() As Byte

        Public Sub Initialize()
            ReDim Data4(7)
        End Sub
    End Structure

    Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (ByRef lpPictDesc As PictDesc, ByRef riid As Guid, ByVal fPictureOwnsHandle As Integer, ByRef ipic As System.Drawing.Image) As Integer






    Private Sub chkshowtext_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
    End Sub

    Private Sub cmdoutlineback_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        If ColorDialog1.ShowDialog() Then
            Scanner1.SetOutlineTextBackColor(Color2Uint32(ColorDialog1.Color))
        End If

    End Sub

    Private Sub cmdoutlineborder_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        If ColorDialog1.ShowDialog() Then
            Scanner1.SetOutlineTextBorderColor(Color2Uint32(ColorDialog1.Color))
        End If

    End Sub

    Private Sub cmdtextcolor_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        If ColorDialog1.ShowDialog() Then
            Me.Scanner1.TextColor = ColorDialog1.Color
        End If


    End Sub


    Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
        Me.Scanner1.Rotate90()
    End Sub

    Private Sub Command10_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command10.Click

        Scanner1.View = 7
        Scanner1.Focus()

    End Sub

    Private Sub Command11_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command11.Click
        Scanner1.View = 2
        Scanner1.Focus()
    End Sub

    Private Sub Command12_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command12.Click

        Scanner1.View = 1
        Scanner1.Focus()

    End Sub

    Private Sub Command13_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        If ColorDialog1.ShowDialog() Then

            Me.Scanner1.SetBackgroundColor(Color2Uint32(ColorDialog1.Color))

        End If

    End Sub


    Private Sub Command17_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)


    End Sub


    Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
        Me.Scanner1.Rotate180()
    End Sub

    Private Sub Command20_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command20.Click

        Scanner1.View = 10
        Scanner1.Focus()

    End Sub



    Private Sub Command23_Click()



    End Sub





    Private Sub Command27_Click()



    End Sub

    Private Sub Command28_Click()
        Me.Scanner1.SetScrollBarPos(iX, iY)


        Me.Scanner1.Focus()

    End Sub

    Private Sub Command29_Click()

        iX = Me.Scanner1.GetHorzScrollBarPos
        iY = Me.Scanner1.GetVertScrollBarPos

        Me.Scanner1.Focus()

    End Sub

    Private Sub Command3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command3.Click
        Me.Scanner1.Rotate270()
    End Sub


    Private Sub Command31_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)

    End Sub


    Private Sub Command33_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)

    End Sub

    Private Sub Command4_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)



    End Sub


    Private Sub Command5_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command5.Click

        Scanner1.View = 9
        Scanner1.Focus()

    End Sub

    Private Sub Command6_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command6.Click

        Scanner1.View = 6
        Scanner1.Focus()

    End Sub

    Private Sub Command7_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command7.Click
        Scanner1.View = 5
        Scanner1.Focus()


    End Sub

    Private Sub Command8_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command8.Click
        Scanner1.View = 3
        Scanner1.Focus()


    End Sub

    Private Sub Command9_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command9.Click

        Scanner1.View = 4
        Scanner1.Focus()

    End Sub
    Public strApp As String
    Private Sub Form1_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Dim iFindIndex As Integer
        strApp = Application.ExecutablePath
        iFindIndex = strApp.IndexOf("Examples")

        strApp = strApp.Substring(0, iFindIndex)

        strApp = strApp + "barcodeimage"





    End Sub


    Private Sub Scanner1_EndScan(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Scanner1.EndScan



    End Sub



    Private Sub Text1_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs)

    End Sub

    Private Sub txtLeft_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs)

    End Sub

    Private Sub txtTop_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs)

    End Sub







    Public Function Color2Uint32(ByVal clr As Color) As UInt32
        Dim t As Int32
        Dim a() As Byte

        t = ColorTranslator.ToOle(clr)


        a = BitConverter.GetBytes(t)

        Return BitConverter.ToUInt32(a, 0)


    End Function

    Private Sub DisplayBarCode(ByVal ibarcodeCount As Integer)
        Dim str1 As String
        Dim strTmp As String
        Dim i As Integer

        If ibarcodeCount < 1 Then

            MessageBox.Show("No barcode found, try scan to black and white color and set the specific area for detect the barcode")
            Exit Sub
        End If


        str1 = "Total " + Str(ibarcodeCount) + " BarCode detected" + Chr(13) + Chr(10)

        For i = 0 To ibarcodeCount - 1

            strTmp = Me.Scanner1.BarCodeGetType(i) + " score:" + Str(Me.Scanner1.BarCodeGetScore(i)) + " value:" + Me.Scanner1.BarCodeGetValue(i)

            str1 = str1 + Chr(13) + Chr(10) + strTmp
        Next

        MessageBox.Show(str1)
    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim ibarcodeCount As Integer

        Scanner1.LoadImage(strApp + "\barcodetest6.pdf")
        AdvBarcodeReader_OpenPDF(strApp + "\barcodetest6.pdf")
    End Sub




    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim ibarcodeCount As Integer

        chkscanmore.Checked = True
        chkScanMultiple.Checked = False
        chkScan90Rotation.Checked = False
        chkScan45Rotation.Checked = False
        chkScan45CounterRotation.Checked = False
        chkscannorotation.Checked = True

        Scanner1.LoadImage(strApp + "\barcodetest1.jpg")
        BarcodeEngineSetting()
        ibarcodeCount = Me.Scanner1.BarCodeReadFullPage

        DisplayBarCode(ibarcodeCount)
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim ibarcodeCount As Integer


        chkscanmore.Checked = True
        chkScanMultiple.Checked = False
        chkScan90Rotation.Checked = False
        chkScan45Rotation.Checked = False
        chkScan45CounterRotation.Checked = False
        chkscannorotation.Checked = True

        Scanner1.LoadImage(strApp + "\barcodetest2.jpg")
        BarcodeEngineSetting()
        ibarcodeCount = Me.Scanner1.BarCodeReadFullPage

        DisplayBarCode(ibarcodeCount)
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim ibarcodeCount As Integer

        chkscanmore.Checked = True
        chkScanMultiple.Checked = True
        chkScan90Rotation.Checked = False
        chkScan45Rotation.Checked = False
        chkScan45CounterRotation.Checked = False
        chkscannorotation.Checked = True

        Scanner1.LoadImage(strApp + "\barcodetest3.jpg")
        BarcodeEngineSetting()
        ibarcodeCount = Me.Scanner1.BarCodeReadFullPage

        DisplayBarCode(ibarcodeCount)
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim ibarcodeCount As Integer

        chkscanmore.Checked = True
        chkScanMultiple.Checked = False
        chkScan90Rotation.Checked = False
        chkScan45Rotation.Checked = False
        chkScan45CounterRotation.Checked = False
        chkscannorotation.Checked = True


        Scanner1.LoadImage(strApp + "\barcodetest4.png")
        BarcodeEngineSetting()
        ibarcodeCount = Me.Scanner1.BarCodeReadFullPage

        DisplayBarCode(ibarcodeCount)
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim ibarcodeCount As Integer

        chkscanmore.Checked = True
        chkScanMultiple.Checked = False
        chkScan90Rotation.Checked = True
        chkScan45Rotation.Checked = True
        chkScan45CounterRotation.Checked = False
        chkscannorotation.Checked = True


        Scanner1.LoadImage(strApp + "\barcodetest5.jpg")
        BarcodeEngineSetting()
        ibarcodeCount = Me.Scanner1.BarCodeReadFullPage

        DisplayBarCode(ibarcodeCount)
    End Sub

   
    Private Sub BarcodeEngineSetting()
        Me.Scanner1.BarCodeReadScanMultiple = chkScanMultiple.Checked
        Me.Scanner1.BarCodeReadScanWithoutRotation = chkscannorotation.Checked
        Me.Scanner1.BarCodeReadScan45Rotation = chkScan45Rotation.Checked
        Me.Scanner1.BarCodeReadScan45CouterRotation = chkScan45CounterRotation.Checked
        Me.Scanner1.BarCodeReadScan90Rotation = chkScan90Rotation.Checked
        Me.Scanner1.BarCodeReadCustomRotation = txtcustomrotation.Text
        Me.Scanner1.BarCodeReadScanAccuracy = chkscanmore.Checked
    End Sub
    Private Sub AdvBarcodeReader_OpenPDF(ByVal strFileName As String)
        Dim iPDFPageCount As Integer
        Dim i As Integer
        cbopdfpage.Items.Clear()
        cbopdfimagecount.Items.Clear()

        iPDFPageCount = Scanner1.BarCodeReadPDFOpen(strFileName)

        cbopdfpage.Items.Clear()
        For i = 1 To iPDFPageCount
            cbopdfpage.Items.Add("Page " + Trim(Str(i)))
        Next

        If cbopdfpage.Items.Count > 0 Then
            cbopdfpage.SelectedIndex = 0

        End If

        chkScanMultiple.Checked = False
        chkScan45CounterRotation.Checked = False
        chkScan90Rotation.Checked = False
        chkScan45Rotation.Checked = False
    End Sub

    Private Sub cbopdfpage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbopdfpage.Click

    End Sub

    Private Sub cbopdfpage_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbopdfpage.SelectedIndexChanged
        Dim iImageCountPerPage As Integer
        Dim i As Integer
        Me.cbopdfimagecount.Items.Clear()
        'first page is 1
        iImageCountPerPage = Me.Scanner1.BarCodeReadPDFImageCount(cbopdfpage.SelectedIndex + 1)
        For i = 1 To iImageCountPerPage
            cbopdfimagecount.Items.Add("Image " + Trim(Str(i)))
        Next

        If cbopdfimagecount.Items.Count > 0 Then
            cbopdfimagecount.SelectedIndex = 0

        End If


    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim ibarcodeCount As Integer
        BarcodeEngineSetting()
        ibarcodeCount = Scanner1.BarCodeReadPDFScan(Me.cbopdfpage.SelectedIndex + 1, cbopdfimagecount.SelectedIndex + 1)
        DisplayBarCode(ibarcodeCount)
    End Sub
End Class