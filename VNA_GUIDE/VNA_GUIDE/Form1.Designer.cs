namespace VNA_GUIDE
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
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea11 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend11 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series15 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.Series series16 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea12 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend12 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series17 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea13 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend13 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series18 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea14 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend14 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series19 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea15 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend15 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series20 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.Series series21 = new System.Windows.Forms.DataVisualization.Charting.Series();
            this.serialPort1 = new System.IO.Ports.SerialPort(this.components);
            this.Kps_txtbox = new System.Windows.Forms.TextBox();
            this.Kis_txtbox = new System.Windows.Forms.TextBox();
            this.Kds_txtbox = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.chart1 = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.send = new System.Windows.Forms.Button();
            this.label12 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.Angle_txtbox = new System.Windows.Forms.TextBox();
            this.pwm_txtbox = new System.Windows.Forms.TextBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.Setangle_txt = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.gpb_rs232 = new System.Windows.Forms.GroupBox();
            this.btDisConnectVNA = new System.Windows.Forms.Button();
            this.btConnectVNA = new System.Windows.Forms.Button();
            this.ComboBox_baud2 = new System.Windows.Forms.ComboBox();
            this.label47 = new System.Windows.Forms.Label();
            this.ComboBox_COM = new System.Windows.Forms.ComboBox();
            this.label6 = new System.Windows.Forms.Label();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.btDisConnectMCU = new System.Windows.Forms.Button();
            this.btConnectMCU = new System.Windows.Forms.Button();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.label11 = new System.Windows.Forms.Label();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.label14 = new System.Windows.Forms.Label();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.button5 = new System.Windows.Forms.Button();
            this.txtG2 = new System.Windows.Forms.TextBox();
            this.label27 = new System.Windows.Forms.Label();
            this.txtR = new System.Windows.Forms.TextBox();
            this.label26 = new System.Windows.Forms.Label();
            this.textNP = new System.Windows.Forms.Label();
            this.label24 = new System.Windows.Forms.Label();
            this.scanVNA = new System.Windows.Forms.Button();
            this.labelProgress = new System.Windows.Forms.Label();
            this.btnPlotCharrt = new System.Windows.Forms.Button();
            this.label19 = new System.Windows.Forms.Label();
            this.label22 = new System.Windows.Forms.Label();
            this.txtSpanFreq = new System.Windows.Forms.TextBox();
            this.txtStepValue = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.txtSetAngle = new System.Windows.Forms.TextBox();
            this.rb_Counter_Clockwise = new System.Windows.Forms.RadioButton();
            this.rb_Clockwise = new System.Windows.Forms.RadioButton();
            this.label18 = new System.Windows.Forms.Label();
            this.cbNumberPoint = new System.Windows.Forms.ComboBox();
            this.btnStop = new System.Windows.Forms.Button();
            this.btnAuto = new System.Windows.Forms.Button();
            this.btnScan = new System.Windows.Forms.Button();
            this.label17 = new System.Windows.Forms.Label();
            this.label16 = new System.Windows.Forms.Label();
            this.txtStopFreq = new System.Windows.Forms.TextBox();
            this.label15 = new System.Windows.Forms.Label();
            this.txtStartFreq = new System.Windows.Forms.TextBox();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.Command = new System.Windows.Forms.GroupBox();
            this.serialPort2 = new System.IO.Ports.SerialPort(this.components);
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tabControl3 = new System.Windows.Forms.TabControl();
            this.tabPage5 = new System.Windows.Forms.TabPage();
            this.tabPage6 = new System.Windows.Forms.TabPage();
            this.chart5 = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.label25 = new System.Windows.Forms.Label();
            this.label23 = new System.Windows.Forms.Label();
            this.txtFolderS21 = new System.Windows.Forms.TextBox();
            this.txtFolderS11 = new System.Windows.Forms.TextBox();
            this.button4 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.tabControl2 = new System.Windows.Forms.TabControl();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.chart2 = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.tabPage4 = new System.Windows.Forms.TabPage();
            this.chart4 = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.UpdateCOM = new System.Windows.Forms.Button();
            this.groupBox10 = new System.Windows.Forms.GroupBox();
            this.btnPlotChartFromDatabase = new System.Windows.Forms.Button();
            this.label21 = new System.Windows.Forms.Label();
            this.txtAngleSelect = new System.Windows.Forms.TextBox();
            this.groupBox9 = new System.Windows.Forms.GroupBox();
            this.label8 = new System.Windows.Forms.Label();
            this.txtFreqIndex = new System.Windows.Forms.TextBox();
            this.button3 = new System.Windows.Forms.Button();
            this.btnExport = new System.Windows.Forms.Button();
            this.groupBox8 = new System.Windows.Forms.GroupBox();
            this.btnSetAngle = new System.Windows.Forms.Button();
            this.txtInitAngle = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.groupBox7 = new System.Windows.Forms.GroupBox();
            this.label20 = new System.Windows.Forms.Label();
            this.txtScanEnable = new System.Windows.Forms.TextBox();
            this.txt_ResAngle = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.textBoxPwm = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.chart3 = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.timerScan = new System.Windows.Forms.Timer(this.components);
            this.timerAutoMode = new System.Windows.Forms.Timer(this.components);
            this.timerValid = new System.Windows.Forms.Timer(this.components);
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.richTextBox2 = new System.Windows.Forms.RichTextBox();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            ((System.ComponentModel.ISupportInitialize)(this.chart1)).BeginInit();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.gpb_rs232.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.Command.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabControl3.SuspendLayout();
            this.tabPage5.SuspendLayout();
            this.tabPage6.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.chart5)).BeginInit();
            this.tabControl2.SuspendLayout();
            this.tabPage3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.chart2)).BeginInit();
            this.tabPage4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.chart4)).BeginInit();
            this.groupBox10.SuspendLayout();
            this.groupBox9.SuspendLayout();
            this.groupBox8.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.groupBox7.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.chart3)).BeginInit();
            this.groupBox6.SuspendLayout();
            this.SuspendLayout();
            // 
            // serialPort1
            // 
            this.serialPort1.ReadBufferSize = 256;
            this.serialPort1.ReceivedBytesThreshold = 32;
            this.serialPort1.DataReceived += new System.IO.Ports.SerialDataReceivedEventHandler(this.serialPort1_DataReceived);
            // 
            // Kps_txtbox
            // 
            this.Kps_txtbox.Location = new System.Drawing.Point(192, 31);
            this.Kps_txtbox.Margin = new System.Windows.Forms.Padding(2);
            this.Kps_txtbox.Name = "Kps_txtbox";
            this.Kps_txtbox.Size = new System.Drawing.Size(76, 20);
            this.Kps_txtbox.TabIndex = 2;
            this.Kps_txtbox.Text = "0.02";
            // 
            // Kis_txtbox
            // 
            this.Kis_txtbox.Location = new System.Drawing.Point(312, 31);
            this.Kis_txtbox.Margin = new System.Windows.Forms.Padding(2);
            this.Kis_txtbox.Name = "Kis_txtbox";
            this.Kis_txtbox.Size = new System.Drawing.Size(76, 20);
            this.Kis_txtbox.TabIndex = 3;
            this.Kis_txtbox.Text = "0";
            this.Kis_txtbox.TextChanged += new System.EventHandler(this.Ki_txtbox_TextChanged);
            // 
            // Kds_txtbox
            // 
            this.Kds_txtbox.Location = new System.Drawing.Point(430, 31);
            this.Kds_txtbox.Margin = new System.Windows.Forms.Padding(2);
            this.Kds_txtbox.Name = "Kds_txtbox";
            this.Kds_txtbox.Size = new System.Drawing.Size(76, 20);
            this.Kds_txtbox.TabIndex = 4;
            this.Kds_txtbox.Text = "0.2";
            this.Kds_txtbox.TextChanged += new System.EventHandler(this.Kd_txtbox_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(159, 31);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(20, 13);
            this.label1.TabIndex = 5;
            this.label1.Text = "Kp";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(279, 31);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(16, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "Ki";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(397, 31);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(20, 13);
            this.label3.TabIndex = 7;
            this.label3.Text = "Kd";
            // 
            // chart1
            // 
            this.chart1.BackImageAlignment = System.Windows.Forms.DataVisualization.Charting.ChartImageAlignmentStyle.Top;
            chartArea11.Name = "ChartArea1";
            this.chart1.ChartAreas.Add(chartArea11);
            legend11.Name = "Legend1";
            this.chart1.Legends.Add(legend11);
            this.chart1.Location = new System.Drawing.Point(-2, 5);
            this.chart1.Margin = new System.Windows.Forms.Padding(2);
            this.chart1.Name = "chart1";
            series15.BorderWidth = 3;
            series15.ChartArea = "ChartArea1";
            series15.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            series15.Legend = "Legend1";
            series15.Name = "S11";
            series15.YValuesPerPoint = 2;
            series16.BorderWidth = 3;
            series16.ChartArea = "ChartArea1";
            series16.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            series16.Legend = "Legend1";
            series16.Name = "S21";
            this.chart1.Series.Add(series15);
            this.chart1.Series.Add(series16);
            this.chart1.Size = new System.Drawing.Size(736, 549);
            this.chart1.TabIndex = 14;
            this.chart1.Text = "chart1";
            // 
            // send
            // 
            this.send.Location = new System.Drawing.Point(544, 28);
            this.send.Margin = new System.Windows.Forms.Padding(2);
            this.send.Name = "send";
            this.send.Size = new System.Drawing.Size(76, 23);
            this.send.TabIndex = 15;
            this.send.Text = "send";
            this.send.UseVisualStyleBackColor = true;
            this.send.Click += new System.EventHandler(this.send_Click);
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(18, 22);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(75, 13);
            this.label12.TabIndex = 25;
            this.label12.Text = "Angle(Degree)";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(22, 51);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(30, 13);
            this.label13.TabIndex = 26;
            this.label13.Text = "Pwm";
            // 
            // Angle_txtbox
            // 
            this.Angle_txtbox.Location = new System.Drawing.Point(96, 22);
            this.Angle_txtbox.Name = "Angle_txtbox";
            this.Angle_txtbox.ReadOnly = true;
            this.Angle_txtbox.Size = new System.Drawing.Size(71, 20);
            this.Angle_txtbox.TabIndex = 29;
            // 
            // pwm_txtbox
            // 
            this.pwm_txtbox.Location = new System.Drawing.Point(96, 48);
            this.pwm_txtbox.Name = "pwm_txtbox";
            this.pwm_txtbox.ReadOnly = true;
            this.pwm_txtbox.Size = new System.Drawing.Size(71, 20);
            this.pwm_txtbox.TabIndex = 30;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.Setangle_txt);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.Kds_txtbox);
            this.groupBox2.Controls.Add(this.Kps_txtbox);
            this.groupBox2.Controls.Add(this.Kis_txtbox);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.send);
            this.groupBox2.Location = new System.Drawing.Point(3, 10);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(658, 76);
            this.groupBox2.TabIndex = 40;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Pid Parameter";
            // 
            // Setangle_txt
            // 
            this.Setangle_txt.Location = new System.Drawing.Point(68, 31);
            this.Setangle_txt.Margin = new System.Windows.Forms.Padding(2);
            this.Setangle_txt.Name = "Setangle_txt";
            this.Setangle_txt.Size = new System.Drawing.Size(76, 20);
            this.Setangle_txt.TabIndex = 12;
            this.Setangle_txt.Text = "0";
            this.Setangle_txt.TextChanged += new System.EventHandler(this.Setpoint_txt_TextChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(14, 31);
            this.label5.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(46, 13);
            this.label5.TabIndex = 13;
            this.label5.Text = "Setpoint";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.Angle_txtbox);
            this.groupBox3.Controls.Add(this.label12);
            this.groupBox3.Controls.Add(this.pwm_txtbox);
            this.groupBox3.Controls.Add(this.label13);
            this.groupBox3.Location = new System.Drawing.Point(702, 5);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(183, 76);
            this.groupBox3.TabIndex = 41;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Respone";
            // 
            // gpb_rs232
            // 
            this.gpb_rs232.Controls.Add(this.btDisConnectVNA);
            this.gpb_rs232.Controls.Add(this.btConnectVNA);
            this.gpb_rs232.Controls.Add(this.ComboBox_baud2);
            this.gpb_rs232.Controls.Add(this.label47);
            this.gpb_rs232.Controls.Add(this.ComboBox_COM);
            this.gpb_rs232.Controls.Add(this.label6);
            this.gpb_rs232.Location = new System.Drawing.Point(3, 6);
            this.gpb_rs232.Name = "gpb_rs232";
            this.gpb_rs232.Size = new System.Drawing.Size(343, 76);
            this.gpb_rs232.TabIndex = 42;
            this.gpb_rs232.TabStop = false;
            this.gpb_rs232.Text = "Serial VNA";
            // 
            // btDisConnectVNA
            // 
            this.btDisConnectVNA.Enabled = false;
            this.btDisConnectVNA.Location = new System.Drawing.Point(234, 45);
            this.btDisConnectVNA.Name = "btDisConnectVNA";
            this.btDisConnectVNA.Size = new System.Drawing.Size(90, 25);
            this.btDisConnectVNA.TabIndex = 20;
            this.btDisConnectVNA.Text = "Disconnect";
            this.btDisConnectVNA.UseVisualStyleBackColor = true;
            this.btDisConnectVNA.Click += new System.EventHandler(this.btDisConnectVNA_Click);
            // 
            // btConnectVNA
            // 
            this.btConnectVNA.Location = new System.Drawing.Point(234, 14);
            this.btConnectVNA.Name = "btConnectVNA";
            this.btConnectVNA.Size = new System.Drawing.Size(90, 25);
            this.btConnectVNA.TabIndex = 19;
            this.btConnectVNA.Text = "Connect";
            this.btConnectVNA.UseVisualStyleBackColor = true;
            this.btConnectVNA.Click += new System.EventHandler(this.btConnectVNA_Click);
            // 
            // ComboBox_baud2
            // 
            this.ComboBox_baud2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.ComboBox_baud2.FormattingEnabled = true;
            this.ComboBox_baud2.Items.AddRange(new object[] {
            "9600",
            "19200",
            "38400",
            "57600",
            "115200"});
            this.ComboBox_baud2.Location = new System.Drawing.Point(96, 46);
            this.ComboBox_baud2.Name = "ComboBox_baud2";
            this.ComboBox_baud2.Size = new System.Drawing.Size(121, 21);
            this.ComboBox_baud2.TabIndex = 18;
            // 
            // label47
            // 
            this.label47.AutoSize = true;
            this.label47.Location = new System.Drawing.Point(21, 52);
            this.label47.Name = "label47";
            this.label47.Size = new System.Drawing.Size(56, 13);
            this.label47.TabIndex = 17;
            this.label47.Text = "Baud rate:";
            // 
            // ComboBox_COM
            // 
            this.ComboBox_COM.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.ComboBox_COM.FormattingEnabled = true;
            this.ComboBox_COM.Items.AddRange(new object[] {
            "COM1",
            "COM2",
            "COM3",
            "COM4",
            "COM5",
            "COM6",
            "COM7",
            "COM8",
            "COM9"});
            this.ComboBox_COM.Location = new System.Drawing.Point(96, 16);
            this.ComboBox_COM.Name = "ComboBox_COM";
            this.ComboBox_COM.Size = new System.Drawing.Size(121, 21);
            this.ComboBox_COM.TabIndex = 14;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(21, 26);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(58, 13);
            this.label6.TabIndex = 13;
            this.label6.Text = "Serial Port:";
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.btDisConnectMCU);
            this.groupBox4.Controls.Add(this.btConnectMCU);
            this.groupBox4.Controls.Add(this.comboBox1);
            this.groupBox4.Controls.Add(this.label11);
            this.groupBox4.Controls.Add(this.comboBox2);
            this.groupBox4.Controls.Add(this.label14);
            this.groupBox4.Location = new System.Drawing.Point(352, 6);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(343, 76);
            this.groupBox4.TabIndex = 43;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Serial MCU";
            // 
            // btDisConnectMCU
            // 
            this.btDisConnectMCU.Enabled = false;
            this.btDisConnectMCU.Location = new System.Drawing.Point(234, 45);
            this.btDisConnectMCU.Name = "btDisConnectMCU";
            this.btDisConnectMCU.Size = new System.Drawing.Size(90, 25);
            this.btDisConnectMCU.TabIndex = 20;
            this.btDisConnectMCU.Text = "Disconnect";
            this.btDisConnectMCU.UseVisualStyleBackColor = true;
            this.btDisConnectMCU.Click += new System.EventHandler(this.btDisConnectMCU_Click);
            // 
            // btConnectMCU
            // 
            this.btConnectMCU.Location = new System.Drawing.Point(234, 14);
            this.btConnectMCU.Name = "btConnectMCU";
            this.btConnectMCU.Size = new System.Drawing.Size(90, 25);
            this.btConnectMCU.TabIndex = 19;
            this.btConnectMCU.Text = "Connect";
            this.btConnectMCU.UseVisualStyleBackColor = true;
            this.btConnectMCU.Click += new System.EventHandler(this.btConnectMCU_Click);
            // 
            // comboBox1
            // 
            this.comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Items.AddRange(new object[] {
            "9600",
            "19200",
            "38400",
            "57600",
            "115200"});
            this.comboBox1.Location = new System.Drawing.Point(96, 46);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(121, 21);
            this.comboBox1.TabIndex = 18;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(21, 52);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(56, 13);
            this.label11.TabIndex = 17;
            this.label11.Text = "Baud rate:";
            // 
            // comboBox2
            // 
            this.comboBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox2.FormattingEnabled = true;
            this.comboBox2.Items.AddRange(new object[] {
            "COM1",
            "COM2",
            "COM3",
            "COM4",
            "COM5",
            "COM6",
            "COM7",
            "COM8",
            "COM9"});
            this.comboBox2.Location = new System.Drawing.Point(96, 16);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(121, 21);
            this.comboBox2.TabIndex = 14;
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(21, 26);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(58, 13);
            this.label14.TabIndex = 13;
            this.label14.Text = "Serial Port:";
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.button5);
            this.groupBox5.Controls.Add(this.txtG2);
            this.groupBox5.Controls.Add(this.label27);
            this.groupBox5.Controls.Add(this.txtR);
            this.groupBox5.Controls.Add(this.label26);
            this.groupBox5.Controls.Add(this.textNP);
            this.groupBox5.Controls.Add(this.label24);
            this.groupBox5.Controls.Add(this.scanVNA);
            this.groupBox5.Controls.Add(this.labelProgress);
            this.groupBox5.Controls.Add(this.btnPlotCharrt);
            this.groupBox5.Controls.Add(this.label19);
            this.groupBox5.Controls.Add(this.label22);
            this.groupBox5.Controls.Add(this.txtSpanFreq);
            this.groupBox5.Controls.Add(this.txtStepValue);
            this.groupBox5.Controls.Add(this.label9);
            this.groupBox5.Controls.Add(this.txtSetAngle);
            this.groupBox5.Controls.Add(this.rb_Counter_Clockwise);
            this.groupBox5.Controls.Add(this.rb_Clockwise);
            this.groupBox5.Controls.Add(this.label18);
            this.groupBox5.Controls.Add(this.cbNumberPoint);
            this.groupBox5.Controls.Add(this.btnStop);
            this.groupBox5.Controls.Add(this.btnAuto);
            this.groupBox5.Controls.Add(this.btnScan);
            this.groupBox5.Controls.Add(this.label17);
            this.groupBox5.Controls.Add(this.label16);
            this.groupBox5.Controls.Add(this.txtStopFreq);
            this.groupBox5.Controls.Add(this.label15);
            this.groupBox5.Controls.Add(this.txtStartFreq);
            this.groupBox5.Location = new System.Drawing.Point(3, 88);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(217, 449);
            this.groupBox5.TabIndex = 44;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "VNA";
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(3, 255);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(91, 25);
            this.button5.TabIndex = 59;
            this.button5.Text = "Refresh";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // txtG2
            // 
            this.txtG2.Location = new System.Drawing.Point(122, 220);
            this.txtG2.Name = "txtG2";
            this.txtG2.Size = new System.Drawing.Size(71, 20);
            this.txtG2.TabIndex = 57;
            this.txtG2.Text = "10";
            // 
            // label27
            // 
            this.label27.AutoSize = true;
            this.label27.Location = new System.Drawing.Point(68, 223);
            this.label27.Name = "label27";
            this.label27.Size = new System.Drawing.Size(38, 13);
            this.label27.TabIndex = 56;
            this.label27.Text = "Gain 2";
            // 
            // txtR
            // 
            this.txtR.Location = new System.Drawing.Point(122, 194);
            this.txtR.Name = "txtR";
            this.txtR.Size = new System.Drawing.Size(71, 20);
            this.txtR.TabIndex = 55;
            this.txtR.Text = "10";
            // 
            // label26
            // 
            this.label26.AutoSize = true;
            this.label26.Location = new System.Drawing.Point(53, 197);
            this.label26.Name = "label26";
            this.label26.Size = new System.Drawing.Size(54, 13);
            this.label26.TabIndex = 54;
            this.label26.Text = "Distant(m)";
            // 
            // textNP
            // 
            this.textNP.AutoSize = true;
            this.textNP.Location = new System.Drawing.Point(163, 306);
            this.textNP.Name = "textNP";
            this.textNP.Size = new System.Drawing.Size(24, 13);
            this.textNP.TabIndex = 53;
            this.textNP.Text = "0/0";
            // 
            // label24
            // 
            this.label24.AutoSize = true;
            this.label24.Location = new System.Drawing.Point(31, 305);
            this.label24.Name = "label24";
            this.label24.Size = new System.Drawing.Size(126, 13);
            this.label24.TabIndex = 52;
            this.label24.Text = "Progress Number of point";
            // 
            // scanVNA
            // 
            this.scanVNA.Location = new System.Drawing.Point(114, 410);
            this.scanVNA.Name = "scanVNA";
            this.scanVNA.Size = new System.Drawing.Size(73, 25);
            this.scanVNA.TabIndex = 51;
            this.scanVNA.Text = "ScanVNA";
            this.scanVNA.UseVisualStyleBackColor = true;
            this.scanVNA.Click += new System.EventHandler(this.scanVNA_Click);
            // 
            // labelProgress
            // 
            this.labelProgress.AutoSize = true;
            this.labelProgress.Location = new System.Drawing.Point(163, 284);
            this.labelProgress.Name = "labelProgress";
            this.labelProgress.Size = new System.Drawing.Size(24, 13);
            this.labelProgress.TabIndex = 47;
            this.labelProgress.Text = "0/0";
            // 
            // btnPlotCharrt
            // 
            this.btnPlotCharrt.Location = new System.Drawing.Point(24, 410);
            this.btnPlotCharrt.Name = "btnPlotCharrt";
            this.btnPlotCharrt.Size = new System.Drawing.Size(72, 25);
            this.btnPlotCharrt.TabIndex = 46;
            this.btnPlotCharrt.Text = "Plot";
            this.btnPlotCharrt.UseVisualStyleBackColor = true;
            this.btnPlotCharrt.Click += new System.EventHandler(this.button2_Click_1);
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Location = new System.Drawing.Point(53, 86);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(62, 13);
            this.label19.TabIndex = 49;
            this.label19.Text = "Span(MHZ)";
            // 
            // label22
            // 
            this.label22.AutoSize = true;
            this.label22.Location = new System.Drawing.Point(30, 283);
            this.label22.Name = "label22";
            this.label22.Size = new System.Drawing.Size(79, 13);
            this.label22.TabIndex = 25;
            this.label22.Text = "Progress Scan ";
            // 
            // txtSpanFreq
            // 
            this.txtSpanFreq.Location = new System.Drawing.Point(122, 83);
            this.txtSpanFreq.Name = "txtSpanFreq";
            this.txtSpanFreq.ReadOnly = true;
            this.txtSpanFreq.Size = new System.Drawing.Size(71, 20);
            this.txtSpanFreq.TabIndex = 50;
            this.txtSpanFreq.Text = "600";
            // 
            // txtStepValue
            // 
            this.txtStepValue.Location = new System.Drawing.Point(122, 168);
            this.txtStepValue.Name = "txtStepValue";
            this.txtStepValue.Size = new System.Drawing.Size(71, 20);
            this.txtStepValue.TabIndex = 48;
            this.txtStepValue.Text = "5";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(12, 171);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(100, 13);
            this.label9.TabIndex = 47;
            this.label9.Text = "Step Angle(Degree)";
            // 
            // txtSetAngle
            // 
            this.txtSetAngle.Location = new System.Drawing.Point(122, 140);
            this.txtSetAngle.Name = "txtSetAngle";
            this.txtSetAngle.Size = new System.Drawing.Size(71, 20);
            this.txtSetAngle.TabIndex = 44;
            this.txtSetAngle.Text = "360";
            // 
            // rb_Counter_Clockwise
            // 
            this.rb_Counter_Clockwise.AutoSize = true;
            this.rb_Counter_Clockwise.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.rb_Counter_Clockwise.Location = new System.Drawing.Point(15, 338);
            this.rb_Counter_Clockwise.Name = "rb_Counter_Clockwise";
            this.rb_Counter_Clockwise.Size = new System.Drawing.Size(113, 17);
            this.rb_Counter_Clockwise.TabIndex = 43;
            this.rb_Counter_Clockwise.Text = "Counter-Clockwise";
            this.rb_Counter_Clockwise.UseVisualStyleBackColor = true;
            // 
            // rb_Clockwise
            // 
            this.rb_Clockwise.AutoSize = true;
            this.rb_Clockwise.Checked = true;
            this.rb_Clockwise.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.rb_Clockwise.Location = new System.Drawing.Point(15, 318);
            this.rb_Clockwise.Name = "rb_Clockwise";
            this.rb_Clockwise.Size = new System.Drawing.Size(73, 17);
            this.rb_Clockwise.TabIndex = 42;
            this.rb_Clockwise.TabStop = true;
            this.rb_Clockwise.Text = "Clockwise";
            this.rb_Clockwise.UseVisualStyleBackColor = true;
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Location = new System.Drawing.Point(12, 142);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(100, 13);
            this.label18.TabIndex = 39;
            this.label18.Text = "Stop Angle(Degree)";
            // 
            // cbNumberPoint
            // 
            this.cbNumberPoint.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbNumberPoint.FormattingEnabled = true;
            this.cbNumberPoint.Items.AddRange(new object[] {
            "11",
            "51",
            "101",
            "201",
            "401",
            "801",
            "1024",
            "1601",
            "3201",
            "4501",
            "6401"});
            this.cbNumberPoint.Location = new System.Drawing.Point(122, 112);
            this.cbNumberPoint.Name = "cbNumberPoint";
            this.cbNumberPoint.Size = new System.Drawing.Size(71, 21);
            this.cbNumberPoint.TabIndex = 21;
            // 
            // btnStop
            // 
            this.btnStop.Location = new System.Drawing.Point(24, 374);
            this.btnStop.Name = "btnStop";
            this.btnStop.Size = new System.Drawing.Size(50, 25);
            this.btnStop.TabIndex = 37;
            this.btnStop.Text = "Stop";
            this.btnStop.UseVisualStyleBackColor = true;
            this.btnStop.Click += new System.EventHandler(this.btnStop_Click);
            // 
            // btnAuto
            // 
            this.btnAuto.Location = new System.Drawing.Point(80, 374);
            this.btnAuto.Name = "btnAuto";
            this.btnAuto.Size = new System.Drawing.Size(44, 25);
            this.btnAuto.TabIndex = 36;
            this.btnAuto.Text = "Auto";
            this.btnAuto.UseVisualStyleBackColor = true;
            this.btnAuto.Click += new System.EventHandler(this.btnAuto_Click);
            // 
            // btnScan
            // 
            this.btnScan.Location = new System.Drawing.Point(130, 374);
            this.btnScan.Name = "btnScan";
            this.btnScan.Size = new System.Drawing.Size(57, 25);
            this.btnScan.TabIndex = 21;
            this.btnScan.Text = "Manual";
            this.btnScan.UseVisualStyleBackColor = true;
            this.btnScan.Click += new System.EventHandler(this.btnScan_Click);
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(30, 114);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(82, 13);
            this.label17.TabIndex = 34;
            this.label17.Text = "Number of point";
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(53, 60);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(59, 13);
            this.label16.TabIndex = 32;
            this.label16.Text = "Stop(MHZ)";
            // 
            // txtStopFreq
            // 
            this.txtStopFreq.Location = new System.Drawing.Point(122, 57);
            this.txtStopFreq.Name = "txtStopFreq";
            this.txtStopFreq.Size = new System.Drawing.Size(71, 20);
            this.txtStopFreq.TabIndex = 33;
            this.txtStopFreq.Text = "1200";
            this.txtStopFreq.TextChanged += new System.EventHandler(this.txtStopFreq_TextChanged);
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(53, 34);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(59, 13);
            this.label15.TabIndex = 31;
            this.label15.Text = "Start(MHZ)";
            // 
            // txtStartFreq
            // 
            this.txtStartFreq.Location = new System.Drawing.Point(122, 31);
            this.txtStartFreq.Name = "txtStartFreq";
            this.txtStartFreq.Size = new System.Drawing.Size(71, 20);
            this.txtStartFreq.TabIndex = 31;
            this.txtStartFreq.Text = "700";
            this.txtStartFreq.TextChanged += new System.EventHandler(this.txtStartFreq_TextChanged);
            // 
            // richTextBox1
            // 
            this.richTextBox1.Location = new System.Drawing.Point(6, 13);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.Size = new System.Drawing.Size(800, 168);
            this.richTextBox1.TabIndex = 45;
            this.richTextBox1.Text = "";
            // 
            // Command
            // 
            this.Command.Controls.Add(this.richTextBox1);
            this.Command.Location = new System.Drawing.Point(11, 712);
            this.Command.Name = "Command";
            this.Command.Size = new System.Drawing.Size(812, 187);
            this.Command.TabIndex = 46;
            this.Command.TabStop = false;
            this.Command.Text = "Command";
            // 
            // serialPort2
            // 
            this.serialPort2.ReadBufferSize = 2048;
            this.serialPort2.ReceivedBytesThreshold = 11;
            this.serialPort2.DataReceived += new System.IO.Ports.SerialDataReceivedEventHandler(this.serialPort2_DataReceived);
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Location = new System.Drawing.Point(11, 12);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1625, 694);
            this.tabControl1.TabIndex = 47;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.tabControl3);
            this.tabPage1.Controls.Add(this.label25);
            this.tabPage1.Controls.Add(this.label23);
            this.tabPage1.Controls.Add(this.txtFolderS21);
            this.tabPage1.Controls.Add(this.txtFolderS11);
            this.tabPage1.Controls.Add(this.button4);
            this.tabPage1.Controls.Add(this.button2);
            this.tabPage1.Controls.Add(this.tabControl2);
            this.tabPage1.Controls.Add(this.UpdateCOM);
            this.tabPage1.Controls.Add(this.groupBox10);
            this.tabPage1.Controls.Add(this.groupBox9);
            this.tabPage1.Controls.Add(this.btnExport);
            this.tabPage1.Controls.Add(this.groupBox8);
            this.tabPage1.Controls.Add(this.button1);
            this.tabPage1.Controls.Add(this.gpb_rs232);
            this.tabPage1.Controls.Add(this.groupBox4);
            this.tabPage1.Controls.Add(this.groupBox3);
            this.tabPage1.Controls.Add(this.groupBox5);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(1617, 668);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Main Tab";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // tabControl3
            // 
            this.tabControl3.Controls.Add(this.tabPage5);
            this.tabControl3.Controls.Add(this.tabPage6);
            this.tabControl3.Location = new System.Drawing.Point(228, 89);
            this.tabControl3.Name = "tabControl3";
            this.tabControl3.SelectedIndex = 0;
            this.tabControl3.Size = new System.Drawing.Size(750, 571);
            this.tabControl3.TabIndex = 58;
            // 
            // tabPage5
            // 
            this.tabPage5.Controls.Add(this.chart1);
            this.tabPage5.Location = new System.Drawing.Point(4, 22);
            this.tabPage5.Name = "tabPage5";
            this.tabPage5.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage5.Size = new System.Drawing.Size(742, 545);
            this.tabPage5.TabIndex = 0;
            this.tabPage5.Text = "S11 & S21";
            this.tabPage5.UseVisualStyleBackColor = true;
            // 
            // tabPage6
            // 
            this.tabPage6.Controls.Add(this.chart5);
            this.tabPage6.Location = new System.Drawing.Point(4, 22);
            this.tabPage6.Name = "tabPage6";
            this.tabPage6.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage6.Size = new System.Drawing.Size(742, 545);
            this.tabPage6.TabIndex = 1;
            this.tabPage6.Text = "Gain";
            this.tabPage6.UseVisualStyleBackColor = true;
            // 
            // chart5
            // 
            this.chart5.BackImageAlignment = System.Windows.Forms.DataVisualization.Charting.ChartImageAlignmentStyle.Top;
            chartArea12.Name = "ChartArea1";
            this.chart5.ChartAreas.Add(chartArea12);
            legend12.Name = "Legend1";
            this.chart5.Legends.Add(legend12);
            this.chart5.Location = new System.Drawing.Point(4, 3);
            this.chart5.Margin = new System.Windows.Forms.Padding(2);
            this.chart5.Name = "chart5";
            series17.BorderWidth = 3;
            series17.ChartArea = "ChartArea1";
            series17.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            series17.Legend = "Legend1";
            series17.Name = "Gain";
            series17.YValuesPerPoint = 2;
            this.chart5.Series.Add(series17);
            this.chart5.Size = new System.Drawing.Size(736, 549);
            this.chart5.TabIndex = 15;
            this.chart5.Text = "chart5";
            // 
            // label25
            // 
            this.label25.AutoSize = true;
            this.label25.Location = new System.Drawing.Point(6, 577);
            this.label25.Name = "label25";
            this.label25.Size = new System.Drawing.Size(129, 13);
            this.label25.TabIndex = 57;
            this.label25.Text = "Floder Contain S21 DATA";
            // 
            // label23
            // 
            this.label23.AutoSize = true;
            this.label23.Location = new System.Drawing.Point(6, 540);
            this.label23.Name = "label23";
            this.label23.Size = new System.Drawing.Size(129, 13);
            this.label23.TabIndex = 54;
            this.label23.Text = "Floder Contain S11 DATA";
            // 
            // txtFolderS21
            // 
            this.txtFolderS21.Location = new System.Drawing.Point(6, 594);
            this.txtFolderS21.Name = "txtFolderS21";
            this.txtFolderS21.Size = new System.Drawing.Size(178, 20);
            this.txtFolderS21.TabIndex = 56;
            // 
            // txtFolderS11
            // 
            this.txtFolderS11.Location = new System.Drawing.Point(6, 556);
            this.txtFolderS11.Name = "txtFolderS11";
            this.txtFolderS11.Size = new System.Drawing.Size(178, 20);
            this.txtFolderS11.TabIndex = 54;
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(190, 593);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(24, 25);
            this.button4.TabIndex = 55;
            this.button4.Text = "...";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(190, 554);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(24, 25);
            this.button2.TabIndex = 54;
            this.button2.Text = "...";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click_3);
            // 
            // tabControl2
            // 
            this.tabControl2.Controls.Add(this.tabPage3);
            this.tabControl2.Controls.Add(this.tabPage4);
            this.tabControl2.Location = new System.Drawing.Point(980, 89);
            this.tabControl2.Name = "tabControl2";
            this.tabControl2.SelectedIndex = 0;
            this.tabControl2.Size = new System.Drawing.Size(630, 571);
            this.tabControl2.TabIndex = 53;
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.chart2);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage3.Size = new System.Drawing.Size(622, 545);
            this.tabPage3.TabIndex = 0;
            this.tabPage3.Text = "polar S11";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // chart2
            // 
            this.chart2.BorderlineWidth = 10;
            chartArea13.Name = "ChartArea1";
            this.chart2.ChartAreas.Add(chartArea13);
            legend13.Name = "Legend1";
            this.chart2.Legends.Add(legend13);
            this.chart2.Location = new System.Drawing.Point(7, 4);
            this.chart2.Name = "chart2";
            series18.BorderWidth = 3;
            series18.ChartArea = "ChartArea1";
            series18.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Polar;
            series18.LabelBorderWidth = 2;
            series18.Legend = "Legend1";
            series18.Name = "Series1";
            series18.ShadowColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            series18.YValuesPerPoint = 6;
            this.chart2.Series.Add(series18);
            this.chart2.Size = new System.Drawing.Size(609, 535);
            this.chart2.TabIndex = 35;
            this.chart2.Text = "chart2";
            this.chart2.UseWaitCursor = true;
            // 
            // tabPage4
            // 
            this.tabPage4.Controls.Add(this.chart4);
            this.tabPage4.Location = new System.Drawing.Point(4, 22);
            this.tabPage4.Name = "tabPage4";
            this.tabPage4.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage4.Size = new System.Drawing.Size(622, 545);
            this.tabPage4.TabIndex = 1;
            this.tabPage4.Text = "polar S21";
            this.tabPage4.UseVisualStyleBackColor = true;
            // 
            // chart4
            // 
            this.chart4.BorderlineWidth = 10;
            chartArea14.Name = "ChartArea1";
            this.chart4.ChartAreas.Add(chartArea14);
            legend14.Name = "Legend1";
            this.chart4.Legends.Add(legend14);
            this.chart4.Location = new System.Drawing.Point(7, 5);
            this.chart4.Name = "chart4";
            series19.BorderWidth = 3;
            series19.ChartArea = "ChartArea1";
            series19.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Polar;
            series19.LabelBorderWidth = 2;
            series19.Legend = "Legend1";
            series19.Name = "Series1";
            series19.ShadowColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            series19.YValuesPerPoint = 6;
            this.chart4.Series.Add(series19);
            this.chart4.Size = new System.Drawing.Size(609, 535);
            this.chart4.TabIndex = 36;
            this.chart4.Text = "chart4";
            this.chart4.UseWaitCursor = true;
            // 
            // UpdateCOM
            // 
            this.UpdateCOM.Location = new System.Drawing.Point(1520, 15);
            this.UpdateCOM.Name = "UpdateCOM";
            this.UpdateCOM.Size = new System.Drawing.Size(77, 25);
            this.UpdateCOM.TabIndex = 21;
            this.UpdateCOM.Text = "Update COM";
            this.UpdateCOM.UseVisualStyleBackColor = true;
            this.UpdateCOM.Click += new System.EventHandler(this.UpdateCOM_Click);
            // 
            // groupBox10
            // 
            this.groupBox10.Controls.Add(this.btnPlotChartFromDatabase);
            this.groupBox10.Controls.Add(this.label21);
            this.groupBox10.Controls.Add(this.txtAngleSelect);
            this.groupBox10.Location = new System.Drawing.Point(1080, 5);
            this.groupBox10.Name = "groupBox10";
            this.groupBox10.Size = new System.Drawing.Size(214, 76);
            this.groupBox10.TabIndex = 52;
            this.groupBox10.TabStop = false;
            this.groupBox10.Text = "Plot chart";
            // 
            // btnPlotChartFromDatabase
            // 
            this.btnPlotChartFromDatabase.Location = new System.Drawing.Point(21, 45);
            this.btnPlotChartFromDatabase.Name = "btnPlotChartFromDatabase";
            this.btnPlotChartFromDatabase.Size = new System.Drawing.Size(72, 25);
            this.btnPlotChartFromDatabase.TabIndex = 53;
            this.btnPlotChartFromDatabase.Text = "Plot chart";
            this.btnPlotChartFromDatabase.UseVisualStyleBackColor = true;
            this.btnPlotChartFromDatabase.Click += new System.EventHandler(this.btnPlotChartFromDatabase_Click);
            // 
            // label21
            // 
            this.label21.AutoSize = true;
            this.label21.Location = new System.Drawing.Point(18, 22);
            this.label21.Name = "label21";
            this.label21.Size = new System.Drawing.Size(67, 13);
            this.label21.TabIndex = 25;
            this.label21.Text = "Select Angle";
            // 
            // txtAngleSelect
            // 
            this.txtAngleSelect.Location = new System.Drawing.Point(119, 19);
            this.txtAngleSelect.Name = "txtAngleSelect";
            this.txtAngleSelect.Size = new System.Drawing.Size(71, 20);
            this.txtAngleSelect.TabIndex = 51;
            this.txtAngleSelect.Text = "10";
            // 
            // groupBox9
            // 
            this.groupBox9.Controls.Add(this.label8);
            this.groupBox9.Controls.Add(this.txtFreqIndex);
            this.groupBox9.Controls.Add(this.button3);
            this.groupBox9.Location = new System.Drawing.Point(1300, 5);
            this.groupBox9.Name = "groupBox9";
            this.groupBox9.Size = new System.Drawing.Size(214, 76);
            this.groupBox9.TabIndex = 47;
            this.groupBox9.TabStop = false;
            this.groupBox9.Text = "Plot polar";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(18, 22);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(86, 13);
            this.label8.TabIndex = 25;
            this.label8.Text = "Frequence index";
            // 
            // txtFreqIndex
            // 
            this.txtFreqIndex.Location = new System.Drawing.Point(119, 19);
            this.txtFreqIndex.Name = "txtFreqIndex";
            this.txtFreqIndex.Size = new System.Drawing.Size(71, 20);
            this.txtFreqIndex.TabIndex = 51;
            this.txtFreqIndex.Text = "50";
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(21, 45);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(71, 25);
            this.button3.TabIndex = 48;
            this.button3.Text = "plot polar";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click_1);
            // 
            // btnExport
            // 
            this.btnExport.Location = new System.Drawing.Point(114, 631);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(91, 25);
            this.btnExport.TabIndex = 47;
            this.btnExport.Text = "export Excel";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // groupBox8
            // 
            this.groupBox8.Controls.Add(this.btnSetAngle);
            this.groupBox8.Controls.Add(this.txtInitAngle);
            this.groupBox8.Controls.Add(this.label10);
            this.groupBox8.Location = new System.Drawing.Point(891, 5);
            this.groupBox8.Name = "groupBox8";
            this.groupBox8.Size = new System.Drawing.Size(183, 76);
            this.groupBox8.TabIndex = 42;
            this.groupBox8.TabStop = false;
            this.groupBox8.Text = "Set Angle";
            // 
            // btnSetAngle
            // 
            this.btnSetAngle.Location = new System.Drawing.Point(31, 47);
            this.btnSetAngle.Name = "btnSetAngle";
            this.btnSetAngle.Size = new System.Drawing.Size(50, 22);
            this.btnSetAngle.TabIndex = 46;
            this.btnSetAngle.Text = "Set";
            this.btnSetAngle.UseVisualStyleBackColor = true;
            this.btnSetAngle.Click += new System.EventHandler(this.btnSetAngle_Click);
            // 
            // txtInitAngle
            // 
            this.txtInitAngle.Location = new System.Drawing.Point(96, 22);
            this.txtInitAngle.Name = "txtInitAngle";
            this.txtInitAngle.Size = new System.Drawing.Size(71, 20);
            this.txtInitAngle.TabIndex = 29;
            this.txtInitAngle.Text = "0";
            this.txtInitAngle.TextChanged += new System.EventHandler(this.txtInitAngle_TextChanged);
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(18, 22);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(75, 13);
            this.label10.TabIndex = 25;
            this.label10.Text = "Angle(Degree)";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(8, 631);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(91, 25);
            this.button1.TabIndex = 45;
            this.button1.Text = "Refresh";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.checkBox1);
            this.tabPage2.Controls.Add(this.groupBox7);
            this.tabPage2.Controls.Add(this.chart3);
            this.tabPage2.Controls.Add(this.groupBox2);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(1617, 668);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "PID Tab";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(1225, 21);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(98, 17);
            this.checkBox1.TabIndex = 43;
            this.checkBox1.Text = "checkBox Plot ";
            this.checkBox1.UseVisualStyleBackColor = true;
            this.checkBox1.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
            // 
            // groupBox7
            // 
            this.groupBox7.Controls.Add(this.label20);
            this.groupBox7.Controls.Add(this.txtScanEnable);
            this.groupBox7.Controls.Add(this.txt_ResAngle);
            this.groupBox7.Controls.Add(this.label4);
            this.groupBox7.Controls.Add(this.textBoxPwm);
            this.groupBox7.Controls.Add(this.label7);
            this.groupBox7.Location = new System.Drawing.Point(667, 10);
            this.groupBox7.Name = "groupBox7";
            this.groupBox7.Size = new System.Drawing.Size(541, 76);
            this.groupBox7.TabIndex = 42;
            this.groupBox7.TabStop = false;
            this.groupBox7.Text = "Respone";
            // 
            // label20
            // 
            this.label20.AutoSize = true;
            this.label20.Location = new System.Drawing.Point(369, 34);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(68, 13);
            this.label20.TabIndex = 32;
            this.label20.Text = "Scan Enable";
            // 
            // txtScanEnable
            // 
            this.txtScanEnable.Location = new System.Drawing.Point(443, 31);
            this.txtScanEnable.Name = "txtScanEnable";
            this.txtScanEnable.ReadOnly = true;
            this.txtScanEnable.Size = new System.Drawing.Size(71, 20);
            this.txtScanEnable.TabIndex = 31;
            // 
            // txt_ResAngle
            // 
            this.txt_ResAngle.Location = new System.Drawing.Point(97, 31);
            this.txt_ResAngle.Name = "txt_ResAngle";
            this.txt_ResAngle.ReadOnly = true;
            this.txt_ResAngle.Size = new System.Drawing.Size(71, 20);
            this.txt_ResAngle.TabIndex = 29;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(19, 31);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(75, 13);
            this.label4.TabIndex = 25;
            this.label4.Text = "Angle(Degree)";
            // 
            // textBoxPwm
            // 
            this.textBoxPwm.Location = new System.Drawing.Point(268, 31);
            this.textBoxPwm.Name = "textBoxPwm";
            this.textBoxPwm.ReadOnly = true;
            this.textBoxPwm.Size = new System.Drawing.Size(71, 20);
            this.textBoxPwm.TabIndex = 30;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(194, 34);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(30, 13);
            this.label7.TabIndex = 26;
            this.label7.Text = "Pwm";
            // 
            // chart3
            // 
            this.chart3.BackImageAlignment = System.Windows.Forms.DataVisualization.Charting.ChartImageAlignmentStyle.Top;
            chartArea15.Name = "ChartArea1";
            this.chart3.ChartAreas.Add(chartArea15);
            legend15.Name = "Legend1";
            this.chart3.Legends.Add(legend15);
            this.chart3.Location = new System.Drawing.Point(3, 91);
            this.chart3.Margin = new System.Windows.Forms.Padding(2);
            this.chart3.Name = "chart3";
            series20.ChartArea = "ChartArea1";
            series20.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            series20.Legend = "Legend1";
            series20.Name = "SetAngle";
            series21.ChartArea = "ChartArea1";
            series21.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            series21.Legend = "Legend1";
            series21.Name = "Angle";
            this.chart3.Series.Add(series20);
            this.chart3.Series.Add(series21);
            this.chart3.Size = new System.Drawing.Size(1221, 549);
            this.chart3.TabIndex = 41;
            this.chart3.Text = "chart3";
            // 
            // timerScan
            // 
            this.timerScan.Interval = 1000;
            this.timerScan.Tick += new System.EventHandler(this.timer2_Tick_1);
            // 
            // timerAutoMode
            // 
            this.timerAutoMode.Interval = 600;
            this.timerAutoMode.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // timerValid
            // 
            this.timerValid.Tick += new System.EventHandler(this.timerValid_Tick);
            // 
            // groupBox6
            // 
            this.groupBox6.Controls.Add(this.richTextBox2);
            this.groupBox6.Location = new System.Drawing.Point(829, 712);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Size = new System.Drawing.Size(803, 187);
            this.groupBox6.TabIndex = 47;
            this.groupBox6.TabStop = false;
            this.groupBox6.Text = "Command VNA";
            // 
            // richTextBox2
            // 
            this.richTextBox2.Location = new System.Drawing.Point(6, 13);
            this.richTextBox2.Name = "richTextBox2";
            this.richTextBox2.Size = new System.Drawing.Size(791, 168);
            this.richTextBox2.TabIndex = 45;
            this.richTextBox2.Text = "";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1637, 918);
            this.Controls.Add(this.groupBox6);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.Command);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.chart1)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.gpb_rs232.ResumeLayout(false);
            this.gpb_rs232.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            this.Command.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabControl3.ResumeLayout(false);
            this.tabPage5.ResumeLayout(false);
            this.tabPage6.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.chart5)).EndInit();
            this.tabControl2.ResumeLayout(false);
            this.tabPage3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.chart2)).EndInit();
            this.tabPage4.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.chart4)).EndInit();
            this.groupBox10.ResumeLayout(false);
            this.groupBox10.PerformLayout();
            this.groupBox9.ResumeLayout(false);
            this.groupBox9.PerformLayout();
            this.groupBox8.ResumeLayout(false);
            this.groupBox8.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.groupBox7.ResumeLayout(false);
            this.groupBox7.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.chart3)).EndInit();
            this.groupBox6.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.IO.Ports.SerialPort serialPort1;
        private System.Windows.Forms.TextBox Kps_txtbox;
        private System.Windows.Forms.TextBox Kis_txtbox;
        private System.Windows.Forms.TextBox Kds_txtbox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DataVisualization.Charting.Chart chart1;
        private System.Windows.Forms.Button send;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.TextBox Angle_txtbox;
        private System.Windows.Forms.TextBox pwm_txtbox;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.GroupBox gpb_rs232;
        private System.Windows.Forms.Button btDisConnectVNA;
        private System.Windows.Forms.Button btConnectVNA;
        private System.Windows.Forms.ComboBox ComboBox_baud2;
        private System.Windows.Forms.Label label47;
        internal System.Windows.Forms.ComboBox ComboBox_COM;
        internal System.Windows.Forms.Label label6;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.Button btDisConnectMCU;
        private System.Windows.Forms.Button btConnectMCU;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Label label11;
        internal System.Windows.Forms.ComboBox comboBox2;
        internal System.Windows.Forms.Label label14;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.TextBox txtStopFreq;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.TextBox txtStartFreq;
        internal System.Windows.Forms.ComboBox cbNumberPoint;
        private System.Windows.Forms.Button btnStop;
        private System.Windows.Forms.Button btnAuto;
        private System.Windows.Forms.Button btnScan;
        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.RadioButton rb_Counter_Clockwise;
        private System.Windows.Forms.RadioButton rb_Clockwise;
        private System.Windows.Forms.TextBox txtSetAngle;
        private System.Windows.Forms.GroupBox Command;
        private System.IO.Ports.SerialPort serialPort2;
        private System.Windows.Forms.TextBox Setangle_txt;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.DataVisualization.Charting.Chart chart3;
        private System.Windows.Forms.GroupBox groupBox7;
        private System.Windows.Forms.TextBox txt_ResAngle;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox textBoxPwm;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox txtStepValue;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.GroupBox groupBox8;
        private System.Windows.Forms.Button btnSetAngle;
        private System.Windows.Forms.TextBox txtInitAngle;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label19;
        private System.Windows.Forms.TextBox txtSpanFreq;
        private System.Windows.Forms.Button btnPlotCharrt;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.Timer timerScan;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Timer timerAutoMode;
        private System.Windows.Forms.GroupBox groupBox9;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox txtFreqIndex;
        private System.Windows.Forms.Label label20;
        private System.Windows.Forms.TextBox txtScanEnable;
        private System.Windows.Forms.GroupBox groupBox10;
        private System.Windows.Forms.Label label21;
        private System.Windows.Forms.TextBox txtAngleSelect;
        private System.Windows.Forms.Label labelProgress;
        private System.Windows.Forms.Label label22;
        private System.Windows.Forms.Button btnPlotChartFromDatabase;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.Button scanVNA;
        private System.Windows.Forms.Timer timerValid;
        private System.Windows.Forms.Label textNP;
        private System.Windows.Forms.Label label24;
        private System.Windows.Forms.Button UpdateCOM;
        private System.Windows.Forms.TabControl tabControl2;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.DataVisualization.Charting.Chart chart2;
        private System.Windows.Forms.TabPage tabPage4;
        private System.Windows.Forms.DataVisualization.Charting.Chart chart4;
        private System.Windows.Forms.GroupBox groupBox6;
        private System.Windows.Forms.RichTextBox richTextBox2;
        private System.Windows.Forms.TextBox txtFolderS21;
        private System.Windows.Forms.TextBox txtFolderS11;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Label label25;
        private System.Windows.Forms.Label label23;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.TabControl tabControl3;
        private System.Windows.Forms.TabPage tabPage5;
        private System.Windows.Forms.TabPage tabPage6;
        private System.Windows.Forms.DataVisualization.Charting.Chart chart5;
        private System.Windows.Forms.TextBox txtG2;
        private System.Windows.Forms.Label label27;
        private System.Windows.Forms.TextBox txtR;
        private System.Windows.Forms.Label label26;
        private System.Windows.Forms.Button button5;
    }
}

