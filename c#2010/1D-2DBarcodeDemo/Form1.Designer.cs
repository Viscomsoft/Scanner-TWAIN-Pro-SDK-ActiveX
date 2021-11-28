namespace WindowsFormsApplication1
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.GroupBox2 = new System.Windows.Forms.GroupBox();
            this.chkscanmore = new System.Windows.Forms.CheckBox();
            this.txtcustomrotation = new System.Windows.Forms.TextBox();
            this.Label3 = new System.Windows.Forms.Label();
            this.chkScan90Rotation = new System.Windows.Forms.CheckBox();
            this.chkScan45CounterRotation = new System.Windows.Forms.CheckBox();
            this.chkScan45Rotation = new System.Windows.Forms.CheckBox();
            this.chkscannorotation = new System.Windows.Forms.CheckBox();
            this.chkScanMultiple = new System.Windows.Forms.CheckBox();
            this.cbopdfimagecount = new System.Windows.Forms.ComboBox();
            this.GroupBox1 = new System.Windows.Forms.GroupBox();
            this.Label2 = new System.Windows.Forms.Label();
            this.cbopdfpage = new System.Windows.Forms.ComboBox();
            this.Label1 = new System.Windows.Forms.Label();
            this.Button6 = new System.Windows.Forms.Button();
            this.Button5 = new System.Windows.Forms.Button();
            this.Button4 = new System.Windows.Forms.Button();
            this.Button3 = new System.Windows.Forms.Button();
            this.Button2 = new System.Windows.Forms.Button();
            this.Command12 = new System.Windows.Forms.Button();
            this.Command11 = new System.Windows.Forms.Button();
            this.Command7 = new System.Windows.Forms.Button();
            this.Command6 = new System.Windows.Forms.Button();
            this.Command8 = new System.Windows.Forms.Button();
            this.Command9 = new System.Windows.Forms.Button();
            this.Command10 = new System.Windows.Forms.Button();
            this.ToolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.Command5 = new System.Windows.Forms.Button();
            this.ColorDialog1 = new System.Windows.Forms.ColorDialog();
            this.Command20 = new System.Windows.Forms.Button();
            this.Command3 = new System.Windows.Forms.Button();
            this.OpenFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.Command2 = new System.Windows.Forms.Button();
            this.Command1 = new System.Windows.Forms.Button();
            this.button8 = new System.Windows.Forms.Button();
            this.axScanner1 = new AxSCANNERLib.AxScanner();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.Button1 = new System.Windows.Forms.Button();
            this.Button7 = new System.Windows.Forms.Button();
            this.GroupBox2.SuspendLayout();
            this.GroupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.axScanner1)).BeginInit();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // GroupBox2
            // 
            this.GroupBox2.Controls.Add(this.chkscanmore);
            this.GroupBox2.Controls.Add(this.txtcustomrotation);
            this.GroupBox2.Controls.Add(this.Label3);
            this.GroupBox2.Controls.Add(this.chkScan90Rotation);
            this.GroupBox2.Controls.Add(this.chkScan45CounterRotation);
            this.GroupBox2.Controls.Add(this.chkScan45Rotation);
            this.GroupBox2.Controls.Add(this.chkscannorotation);
            this.GroupBox2.Controls.Add(this.chkScanMultiple);
            this.GroupBox2.Location = new System.Drawing.Point(378, 68);
            this.GroupBox2.Name = "GroupBox2";
            this.GroupBox2.Size = new System.Drawing.Size(576, 147);
            this.GroupBox2.TabIndex = 146;
            this.GroupBox2.TabStop = false;
            this.GroupBox2.Text = "UnSelected some option, it will scan more faster";
            // 
            // chkscanmore
            // 
            this.chkscanmore.AutoSize = true;
            this.chkscanmore.Checked = true;
            this.chkscanmore.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkscanmore.Location = new System.Drawing.Point(264, 121);
            this.chkscanmore.Name = "chkscanmore";
            this.chkscanmore.Size = new System.Drawing.Size(132, 16);
            this.chkscanmore.TabIndex = 7;
            this.chkscanmore.Text = "Scan More Accuracy";
            this.chkscanmore.UseVisualStyleBackColor = true;
            // 
            // txtcustomrotation
            // 
            this.txtcustomrotation.Location = new System.Drawing.Point(386, 93);
            this.txtcustomrotation.Name = "txtcustomrotation";
            this.txtcustomrotation.Size = new System.Drawing.Size(128, 21);
            this.txtcustomrotation.TabIndex = 6;
            this.txtcustomrotation.Text = "0";
            // 
            // Label3
            // 
            this.Label3.AutoSize = true;
            this.Label3.Location = new System.Drawing.Point(270, 97);
            this.Label3.Name = "Label3";
            this.Label3.Size = new System.Drawing.Size(95, 12);
            this.Label3.TabIndex = 5;
            this.Label3.Text = "Custom Rotation";
            // 
            // chkScan90Rotation
            // 
            this.chkScan90Rotation.AutoSize = true;
            this.chkScan90Rotation.Checked = true;
            this.chkScan90Rotation.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkScan90Rotation.Location = new System.Drawing.Point(264, 62);
            this.chkScan90Rotation.Name = "chkScan90Rotation";
            this.chkScan90Rotation.Size = new System.Drawing.Size(192, 16);
            this.chkScan90Rotation.TabIndex = 4;
            this.chkScan90Rotation.Text = "Scan With 90 Degree Rotation";
            this.chkScan90Rotation.UseVisualStyleBackColor = true;
            // 
            // chkScan45CounterRotation
            // 
            this.chkScan45CounterRotation.AutoSize = true;
            this.chkScan45CounterRotation.Checked = true;
            this.chkScan45CounterRotation.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkScan45CounterRotation.Location = new System.Drawing.Point(264, 29);
            this.chkScan45CounterRotation.Name = "chkScan45CounterRotation";
            this.chkScan45CounterRotation.Size = new System.Drawing.Size(300, 16);
            this.chkScan45CounterRotation.TabIndex = 3;
            this.chkScan45CounterRotation.Text = "Scan With 45 Degree Counter Clockwise Rotation";
            this.chkScan45CounterRotation.UseVisualStyleBackColor = true;
            // 
            // chkScan45Rotation
            // 
            this.chkScan45Rotation.AutoSize = true;
            this.chkScan45Rotation.Checked = true;
            this.chkScan45Rotation.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkScan45Rotation.Location = new System.Drawing.Point(25, 97);
            this.chkScan45Rotation.Name = "chkScan45Rotation";
            this.chkScan45Rotation.Size = new System.Drawing.Size(192, 16);
            this.chkScan45Rotation.TabIndex = 2;
            this.chkScan45Rotation.Text = "Scan With 45 Degree Rotation";
            this.chkScan45Rotation.UseVisualStyleBackColor = true;
            // 
            // chkscannorotation
            // 
            this.chkscannorotation.AutoSize = true;
            this.chkscannorotation.Checked = true;
            this.chkscannorotation.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkscannorotation.Location = new System.Drawing.Point(25, 62);
            this.chkscannorotation.Name = "chkscannorotation";
            this.chkscannorotation.Size = new System.Drawing.Size(150, 16);
            this.chkscannorotation.TabIndex = 1;
            this.chkscannorotation.Text = "Scan Without Rotation";
            this.chkscannorotation.UseVisualStyleBackColor = true;
            // 
            // chkScanMultiple
            // 
            this.chkScanMultiple.AutoSize = true;
            this.chkScanMultiple.Checked = true;
            this.chkScanMultiple.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkScanMultiple.Location = new System.Drawing.Point(25, 29);
            this.chkScanMultiple.Name = "chkScanMultiple";
            this.chkScanMultiple.Size = new System.Drawing.Size(150, 16);
            this.chkScanMultiple.TabIndex = 0;
            this.chkScanMultiple.Text = "Scan Multiple Barcode";
            this.chkScanMultiple.UseVisualStyleBackColor = true;
            // 
            // cbopdfimagecount
            // 
            this.cbopdfimagecount.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbopdfimagecount.FormattingEnabled = true;
            this.cbopdfimagecount.Location = new System.Drawing.Point(266, 34);
            this.cbopdfimagecount.Name = "cbopdfimagecount";
            this.cbopdfimagecount.Size = new System.Drawing.Size(92, 20);
            this.cbopdfimagecount.TabIndex = 3;
            // 
            // GroupBox1
            // 
            this.GroupBox1.Controls.Add(this.cbopdfimagecount);
            this.GroupBox1.Controls.Add(this.Label2);
            this.GroupBox1.Controls.Add(this.cbopdfpage);
            this.GroupBox1.Controls.Add(this.Label1);
            this.GroupBox1.Location = new System.Drawing.Point(378, 229);
            this.GroupBox1.Name = "GroupBox1";
            this.GroupBox1.Size = new System.Drawing.Size(364, 88);
            this.GroupBox1.TabIndex = 143;
            this.GroupBox1.TabStop = false;
            // 
            // Label2
            // 
            this.Label2.AutoSize = true;
            this.Label2.Location = new System.Drawing.Point(191, 38);
            this.Label2.Name = "Label2";
            this.Label2.Size = new System.Drawing.Size(77, 12);
            this.Label2.TabIndex = 2;
            this.Label2.Text = "Select Image";
            // 
            // cbopdfpage
            // 
            this.cbopdfpage.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbopdfpage.FormattingEnabled = true;
            this.cbopdfpage.Location = new System.Drawing.Point(82, 33);
            this.cbopdfpage.Name = "cbopdfpage";
            this.cbopdfpage.Size = new System.Drawing.Size(92, 20);
            this.cbopdfpage.TabIndex = 1;
            this.cbopdfpage.SelectedIndexChanged += new System.EventHandler(this.cbopdfpage_SelectedIndexChanged);
            // 
            // Label1
            // 
            this.Label1.AutoSize = true;
            this.Label1.Location = new System.Drawing.Point(7, 37);
            this.Label1.Name = "Label1";
            this.Label1.Size = new System.Drawing.Size(71, 12);
            this.Label1.TabIndex = 0;
            this.Label1.Text = "Select Page";
            // 
            // Button6
            // 
            this.Button6.Location = new System.Drawing.Point(691, 2);
            this.Button6.Name = "Button6";
            this.Button6.Size = new System.Drawing.Size(154, 60);
            this.Button6.TabIndex = 142;
            this.Button6.Text = "Read barcode at any orientation";
            this.Button6.UseVisualStyleBackColor = true;
            this.Button6.Click += new System.EventHandler(this.Button6_Click);
            // 
            // Button5
            // 
            this.Button5.Location = new System.Drawing.Point(517, 2);
            this.Button5.Name = "Button5";
            this.Button5.Size = new System.Drawing.Size(154, 60);
            this.Button5.TabIndex = 141;
            this.Button5.Text = "Read DataMatrix Demo";
            this.Button5.UseVisualStyleBackColor = true;
            this.Button5.Click += new System.EventHandler(this.Button5_Click);
            // 
            // Button4
            // 
            this.Button4.Location = new System.Drawing.Point(337, 2);
            this.Button4.Name = "Button4";
            this.Button4.Size = new System.Drawing.Size(154, 60);
            this.Button4.TabIndex = 140;
            this.Button4.Text = "Read Multiple Barcode Demo";
            this.Button4.UseVisualStyleBackColor = true;
            this.Button4.Click += new System.EventHandler(this.Button4_Click);
            // 
            // Button3
            // 
            this.Button3.Location = new System.Drawing.Point(164, 2);
            this.Button3.Name = "Button3";
            this.Button3.Size = new System.Drawing.Size(154, 60);
            this.Button3.TabIndex = 139;
            this.Button3.Text = "Read 1D Barcode Demo";
            this.Button3.UseVisualStyleBackColor = true;
            this.Button3.Click += new System.EventHandler(this.Button3_Click);
            // 
            // Button2
            // 
            this.Button2.Location = new System.Drawing.Point(-150, -9);
            this.Button2.Name = "Button2";
            this.Button2.Size = new System.Drawing.Size(129, 60);
            this.Button2.TabIndex = 138;
            this.Button2.Text = "Read QR Code Demo";
            this.Button2.UseVisualStyleBackColor = true;
            // 
            // Command12
            // 
            this.Command12.BackColor = System.Drawing.SystemColors.Control;
            this.Command12.Cursor = System.Windows.Forms.Cursors.Default;
            this.Command12.Font = new System.Drawing.Font("Arial", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Command12.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Command12.Location = new System.Drawing.Point(-157, 164);
            this.Command12.Name = "Command12";
            this.Command12.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Command12.Size = new System.Drawing.Size(89, 25);
            this.Command12.TabIndex = 137;
            this.Command12.Text = "25% of Original";
            this.Command12.UseVisualStyleBackColor = false;
            // 
            // Command11
            // 
            this.Command11.BackColor = System.Drawing.SystemColors.Control;
            this.Command11.Cursor = System.Windows.Forms.Cursors.Default;
            this.Command11.Font = new System.Drawing.Font("Arial", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Command11.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Command11.Location = new System.Drawing.Point(-157, 196);
            this.Command11.Name = "Command11";
            this.Command11.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Command11.Size = new System.Drawing.Size(89, 25);
            this.Command11.TabIndex = 136;
            this.Command11.Text = "33% of Original";
            this.Command11.UseVisualStyleBackColor = false;
            // 
            // Command7
            // 
            this.Command7.BackColor = System.Drawing.SystemColors.Control;
            this.Command7.Cursor = System.Windows.Forms.Cursors.Default;
            this.Command7.Font = new System.Drawing.Font("Arial", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Command7.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Command7.Location = new System.Drawing.Point(-157, 292);
            this.Command7.Name = "Command7";
            this.Command7.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Command7.Size = new System.Drawing.Size(89, 25);
            this.Command7.TabIndex = 135;
            this.Command7.Text = "100% of Original";
            this.Command7.UseVisualStyleBackColor = false;
            // 
            // Command6
            // 
            this.Command6.BackColor = System.Drawing.SystemColors.Control;
            this.Command6.Cursor = System.Windows.Forms.Cursors.Default;
            this.Command6.Font = new System.Drawing.Font("Arial", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Command6.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Command6.Location = new System.Drawing.Point(-157, 324);
            this.Command6.Name = "Command6";
            this.Command6.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Command6.Size = new System.Drawing.Size(89, 25);
            this.Command6.TabIndex = 134;
            this.Command6.Text = "150% of Original";
            this.Command6.UseVisualStyleBackColor = false;
            // 
            // Command8
            // 
            this.Command8.BackColor = System.Drawing.SystemColors.Control;
            this.Command8.Cursor = System.Windows.Forms.Cursors.Default;
            this.Command8.Font = new System.Drawing.Font("Arial", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Command8.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Command8.Location = new System.Drawing.Point(-157, 228);
            this.Command8.Name = "Command8";
            this.Command8.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Command8.Size = new System.Drawing.Size(89, 25);
            this.Command8.TabIndex = 133;
            this.Command8.Text = "50% of Original";
            this.Command8.UseVisualStyleBackColor = false;
            // 
            // Command9
            // 
            this.Command9.BackColor = System.Drawing.SystemColors.Control;
            this.Command9.Cursor = System.Windows.Forms.Cursors.Default;
            this.Command9.Font = new System.Drawing.Font("Arial", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Command9.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Command9.Location = new System.Drawing.Point(-157, 260);
            this.Command9.Name = "Command9";
            this.Command9.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Command9.Size = new System.Drawing.Size(89, 25);
            this.Command9.TabIndex = 132;
            this.Command9.Text = "75% of Original";
            this.Command9.UseVisualStyleBackColor = false;
            // 
            // Command10
            // 
            this.Command10.BackColor = System.Drawing.SystemColors.Control;
            this.Command10.Cursor = System.Windows.Forms.Cursors.Default;
            this.Command10.Font = new System.Drawing.Font("Arial", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Command10.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Command10.Location = new System.Drawing.Point(-157, 356);
            this.Command10.Name = "Command10";
            this.Command10.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Command10.Size = new System.Drawing.Size(89, 25);
            this.Command10.TabIndex = 131;
            this.Command10.Text = "200% of Original";
            this.Command10.UseVisualStyleBackColor = false;
            // 
            // Command5
            // 
            this.Command5.BackColor = System.Drawing.SystemColors.Control;
            this.Command5.Cursor = System.Windows.Forms.Cursors.Default;
            this.Command5.Font = new System.Drawing.Font("Arial", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Command5.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Command5.Location = new System.Drawing.Point(-157, 388);
            this.Command5.Name = "Command5";
            this.Command5.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Command5.Size = new System.Drawing.Size(89, 25);
            this.Command5.TabIndex = 130;
            this.Command5.Text = "Fit to Window";
            this.Command5.UseVisualStyleBackColor = false;
            // 
            // Command20
            // 
            this.Command20.BackColor = System.Drawing.SystemColors.Control;
            this.Command20.Cursor = System.Windows.Forms.Cursors.Default;
            this.Command20.Font = new System.Drawing.Font("Arial", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Command20.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Command20.Location = new System.Drawing.Point(-157, 419);
            this.Command20.Name = "Command20";
            this.Command20.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Command20.Size = new System.Drawing.Size(89, 25);
            this.Command20.TabIndex = 129;
            this.Command20.Text = "Aspect Ratio";
            this.Command20.UseVisualStyleBackColor = false;
            // 
            // Command3
            // 
            this.Command3.BackColor = System.Drawing.SystemColors.Control;
            this.Command3.Cursor = System.Windows.Forms.Cursors.Default;
            this.Command3.Font = new System.Drawing.Font("Arial", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Command3.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Command3.Location = new System.Drawing.Point(-157, 132);
            this.Command3.Name = "Command3";
            this.Command3.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Command3.Size = new System.Drawing.Size(89, 25);
            this.Command3.TabIndex = 128;
            this.Command3.Text = "Rotate 270";
            this.Command3.UseVisualStyleBackColor = false;
            // 
            // Command2
            // 
            this.Command2.BackColor = System.Drawing.SystemColors.Control;
            this.Command2.Cursor = System.Windows.Forms.Cursors.Default;
            this.Command2.Font = new System.Drawing.Font("Arial", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Command2.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Command2.Location = new System.Drawing.Point(-157, 100);
            this.Command2.Name = "Command2";
            this.Command2.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Command2.Size = new System.Drawing.Size(89, 25);
            this.Command2.TabIndex = 127;
            this.Command2.Text = "Rotate 180";
            this.Command2.UseVisualStyleBackColor = false;
            // 
            // Command1
            // 
            this.Command1.BackColor = System.Drawing.SystemColors.Control;
            this.Command1.Cursor = System.Windows.Forms.Cursors.Default;
            this.Command1.Font = new System.Drawing.Font("Arial", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Command1.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Command1.Location = new System.Drawing.Point(-157, 68);
            this.Command1.Name = "Command1";
            this.Command1.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Command1.Size = new System.Drawing.Size(89, 25);
            this.Command1.TabIndex = 126;
            this.Command1.Text = "Rotate 90";
            this.Command1.UseVisualStyleBackColor = false;
            // 
            // button8
            // 
            this.button8.Location = new System.Drawing.Point(12, 2);
            this.button8.Name = "button8";
            this.button8.Size = new System.Drawing.Size(129, 60);
            this.button8.TabIndex = 148;
            this.button8.Text = "Read QR Code Demo";
            this.button8.UseVisualStyleBackColor = true;
            this.button8.Click += new System.EventHandler(this.button8_Click);
            // 
            // axScanner1
            // 
            this.axScanner1.Enabled = true;
            this.axScanner1.Location = new System.Drawing.Point(4, 86);
            this.axScanner1.Name = "axScanner1";
            this.axScanner1.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("axScanner1.OcxState")));
            this.axScanner1.Size = new System.Drawing.Size(365, 442);
            this.axScanner1.TabIndex = 147;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.Button7);
            this.groupBox3.Controls.Add(this.Button1);
            this.groupBox3.Location = new System.Drawing.Point(383, 331);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(498, 93);
            this.groupBox3.TabIndex = 149;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "If BarCodeReadFullPage cannot detect the barcode, try use for Advanced Barcode re" +
                "ader ";
            // 
            // Button1
            // 
            this.Button1.Location = new System.Drawing.Point(6, 27);
            this.Button1.Name = "Button1";
            this.Button1.Size = new System.Drawing.Size(170, 60);
            this.Button1.TabIndex = 145;
            this.Button1.Text = "Read PDF Demo";
            this.Button1.UseVisualStyleBackColor = true;
            this.Button1.Click += new System.EventHandler(this.Button1_Click_1);
            // 
            // Button7
            // 
            this.Button7.Location = new System.Drawing.Point(188, 27);
            this.Button7.Name = "Button7";
            this.Button7.Size = new System.Drawing.Size(287, 59);
            this.Button7.TabIndex = 146;
            this.Button7.Text = "scan barcode for Advanced Barcode reader  ";
            this.Button7.UseVisualStyleBackColor = true;
            this.Button7.Click += new System.EventHandler(this.Button7_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1056, 588);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.button8);
            this.Controls.Add(this.axScanner1);
            this.Controls.Add(this.GroupBox2);
            this.Controls.Add(this.GroupBox1);
            this.Controls.Add(this.Button6);
            this.Controls.Add(this.Button5);
            this.Controls.Add(this.Button4);
            this.Controls.Add(this.Button3);
            this.Controls.Add(this.Button2);
            this.Controls.Add(this.Command12);
            this.Controls.Add(this.Command11);
            this.Controls.Add(this.Command7);
            this.Controls.Add(this.Command6);
            this.Controls.Add(this.Command8);
            this.Controls.Add(this.Command9);
            this.Controls.Add(this.Command10);
            this.Controls.Add(this.Command5);
            this.Controls.Add(this.Command20);
            this.Controls.Add(this.Command3);
            this.Controls.Add(this.Command2);
            this.Controls.Add(this.Command1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.GroupBox2.ResumeLayout(false);
            this.GroupBox2.PerformLayout();
            this.GroupBox1.ResumeLayout(false);
            this.GroupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.axScanner1)).EndInit();
            this.groupBox3.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        internal System.Windows.Forms.GroupBox GroupBox2;
        internal System.Windows.Forms.CheckBox chkscanmore;
        internal System.Windows.Forms.TextBox txtcustomrotation;
        internal System.Windows.Forms.Label Label3;
        internal System.Windows.Forms.CheckBox chkScan90Rotation;
        internal System.Windows.Forms.CheckBox chkScan45CounterRotation;
        internal System.Windows.Forms.CheckBox chkScan45Rotation;
        internal System.Windows.Forms.CheckBox chkscannorotation;
        internal System.Windows.Forms.CheckBox chkScanMultiple;
        internal System.Windows.Forms.ComboBox cbopdfimagecount;
        internal System.Windows.Forms.GroupBox GroupBox1;
        internal System.Windows.Forms.Label Label2;
        internal System.Windows.Forms.ComboBox cbopdfpage;
        internal System.Windows.Forms.Label Label1;
        internal System.Windows.Forms.Button Button6;
        internal System.Windows.Forms.Button Button5;
        internal System.Windows.Forms.Button Button4;
        internal System.Windows.Forms.Button Button3;
        internal System.Windows.Forms.Button Button2;
        public System.Windows.Forms.Button Command12;
        public System.Windows.Forms.Button Command11;
        public System.Windows.Forms.Button Command7;
        public System.Windows.Forms.Button Command6;
        public System.Windows.Forms.Button Command8;
        public System.Windows.Forms.Button Command9;
        public System.Windows.Forms.Button Command10;
        public System.Windows.Forms.ToolTip ToolTip1;
        public System.Windows.Forms.Button Command5;
        internal System.Windows.Forms.ColorDialog ColorDialog1;
        public System.Windows.Forms.Button Command20;
        public System.Windows.Forms.Button Command3;
        internal System.Windows.Forms.OpenFileDialog OpenFileDialog1;
        public System.Windows.Forms.Button Command2;
        public System.Windows.Forms.Button Command1;
        private AxSCANNERLib.AxScanner axScanner1;
        internal System.Windows.Forms.Button button8;
        private System.Windows.Forms.GroupBox groupBox3;
        internal System.Windows.Forms.Button Button7;
        internal System.Windows.Forms.Button Button1;
    }
}

