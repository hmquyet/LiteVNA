using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using static System.Windows.Forms.LinkLabel;
using OfficeOpenXml;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Runtime.InteropServices;
using System.Threading;
using System.IO;
using System.Data.SqlTypes;


namespace VNA_GUIDE
{
    public partial class Form1 : Form
    {
        private double time, set_angle, angle, pwm, Iref, Idc, Ts, time_angle, n, step_enable, tick;
        List<Tuple<byte, string>> pairs = new List<Tuple<byte, string>>();
        string rev_data;
        string rev1_data;
        int Enable_Scan;
        float Old_Trigger_Scan;
        int ctnAutoScan = 0 ;
        float InitAngleSendToMCU;
        float SetAngleSendToMCU;
        string oldData;
        int modeScan;
        string filePath = @"D:\NextwaveIndustries\liteVNA\DATA\DATA.xlsx";
        List<string> Listdata = new List<string>();

        List<double> parameterListS21 = new List<double>();
   

        List<byte> m_tx_cmd = new List<byte>();
        List<byte> m_rx_cmd = new List<byte>();
        public const byte CMD_V2_NOP = 0x00;          // no operation
        public const byte CMD_V2_INDICATE = 0x0D;     // device always replies with ascii '2' (0x32)
        public const byte CMD_V2_READ_1 = 0x10;       // read 1 byte from register
        public const byte CMD_V2_READ_2 = 0x11;       // read 2 bytes from register
        public const byte CMD_V2_READ_4 = 0x12;       // read 4 bytes from register
        public const byte CMD_V2_READ_8 = 0x13;       // read 8 bytes from register
        public const byte CMD_V2_READ_FIFO = 0x18;    // read n bytes from an FIFO
        public const byte CMD_V2_WRITE_1 = 0x20;      // write 1 byte to a register
        public const byte CMD_V2_WRITE_2 = 0x21;      // write 2 bytes to a register
        public const byte CMD_V2_WRITE_4 = 0x22;      // write 4 bytes to a register
        public const byte CMD_V2_WRITE_8 = 0x23;      // write 8 bytes to a register
        public const byte CMD_V2_WRITE_FIFO = 0x28;   // write n bytes to an FIFO
        private static Thread excelThread;
        int ctnAutoExcel = 2;

        public Form1()
        {
            InitializeComponent();
            //PlotData();

        }


        private void button3_Click(object sender, EventArgs e)
        {
            serialPort1.Write("V");
            string dataToSend = new string('*', 32);
            serialPort1.Write(dataToSend);

            step_enable = 0;
        }

        private void Kd_txtbox_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            serialPort1.Write("R");
            string dataToSend = new string('*', 32);
            serialPort1.Write(dataToSend);

            step_enable = 1;

        }

        private void Ki_txtbox_TextChanged(object sender, EventArgs e)
        {

        }


        private void CLOSE_Click(object sender, EventArgs e)
        {

        }

        private void btDisConnectVNA_Click(object sender, EventArgs e)
        {
            serialPort1.Close(); 
            btConnectVNA.Enabled = true;
            btDisConnectVNA.Enabled = false;
            ComboBox_COM.Enabled = true;
            ComboBox_baud2.Enabled = true;
            btConnectVNA.ForeColor = Color.Indigo;
            btDisConnectVNA.ForeColor = Color.Black;
            DisabledForm();

        }

        private void btDisConnectMCU_Click(object sender, EventArgs e)
        {
            serialPort2.Close();
            btConnectMCU.Enabled = true;
            btDisConnectMCU.Enabled = false;
            comboBox2.Enabled = true;
            comboBox1.Enabled = true;
            btConnectMCU.ForeColor = Color.Indigo;
            btDisConnectMCU.ForeColor = Color.Black;
            DisabledForm();
        }
        int filterTriggerScan = 0;
        private void ProcessData()
        {

            // Draw PID
            if (rev1_data != null && rev1_data.StartsWith("A"))
            {
                rev1_data = rev1_data.Substring(1);

                // Validate if rev_data can be converted to a double
                if (double.TryParse(rev1_data, out double angle))
                {
                    Angle_txtbox.Text = rev1_data;
                    txt_ResAngle.Text = rev1_data;
                    if (checkBox1.Checked)
                    {
                        chart3.Series["SetAngle"].Points.AddXY(time_angle * Ts, set_angle);
                        chart3.Series["Angle"].Points.AddXY(time_angle * Ts, angle);

                        if (time_angle * Ts > chart3.ChartAreas[0].AxisX.Maximum)
                        {
                            chart3.ChartAreas[0].AxisX.Maximum += time_angle * Ts;
                        }
                        if (angle > chart3.ChartAreas[0].AxisY.Maximum)
                        {
                            chart3.ChartAreas[0].AxisY.Maximum = angle * 1.2;
                        }
                        if (angle < chart3.ChartAreas[0].AxisY.Minimum)
                        {
                            chart3.ChartAreas[0].AxisY.Minimum = angle * 1.2;
                        }

                        time_angle = time_angle + 1;

                    }
                    
                }
               
            }
            else if (rev1_data != null && rev1_data.StartsWith("u"))
            {
                rev1_data = rev1_data.Substring(1);
                if (double.TryParse(rev1_data, out double angle))
                {
                    
                    pwm_txtbox.Text = rev1_data;
                    textBoxPwm.Text = rev1_data;
                }
            }
            else if (rev1_data != null && rev1_data.StartsWith("X"))
            {
                string substring = rev1_data.Substring(1);
                if (float.TryParse(substring, out float Trigger_Scan))
                {
                    txtScanEnable.Text = Trigger_Scan.ToString();

                    if (Trigger_Scan == 1 && Trigger_Scan != Old_Trigger_Scan)
                    {
                        if(modeScan == 0)
                        {
                            
                                // Write VNA
                                timerScan.Enabled = true;
                                DoneScan = false;
                                Listdata.Clear();
                                SendParameter();
                                richTextBox1.AppendText("requesting " + int.Parse(cbNumberPoint.SelectedItem.ToString()) + " points .." + "\n");
                                SendScanData();
                           
                            
                        }
                        else if(modeScan == 1)
                        {
                           
                                ctnAutoExcel++;
                                timerAutoMode.Enabled = true;
                                timerScan.Enabled = true;
                                DoneScan = false;
                                AutoScan = true;
                               
                            

                        }

                        filterTriggerScan++;
                    }
                    Old_Trigger_Scan = Trigger_Scan;
                }  
            }

        }



        private void serialPort2_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            if (serialPort2.IsOpen)
            {
                try
                {
                    // Read data from the serial port
                    rev1_data = serialPort2.ReadLine();
                }
                catch (IOException)
                {
                    // Handle the specific case where the I/O operation is aborted
                    // You can set a flag or take appropriate action here
                    return;
                }

                // Invoke the processing on the UI thread
                if (!string.IsNullOrEmpty(rev1_data))
                {
                    this.Invoke(new Action(() => ProcessData()));
                }
            }

        }


        int ctnPoint = 0;
        int ctnScan = 0;
        bool DoneScan = false;
        bool AutoScan = false;


        private void serialPort1_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {

            byte[] buffer = new byte[32]; // Mảng byte để lưu trữ dữ liệu
            int bytesRead = serialPort1.Read(buffer, 0, buffer.Length);

            if (bytesRead >= 32)
            {
                string hexString = BitConverter.ToString(buffer).Replace("-", "");

                // Biến đổi dữ liệu thành chuỗi hex và hiển thị trên RichTextBox

                if (ctnPoint < int.Parse(cbNumberPoint.SelectedItem.ToString())) // int.Parse(cbNumberPoint.SelectedItem.ToString())
                {
                    if (DoneScan == false)
                    {

                        Dictionary<string, double> decodedValues = DecodeFrame(hexString);

                        int freqIndex = int.Parse(decodedValues["freqIndex"].ToString());
                        //if (freqIndex == ctnPoint)
                        //{
                            Listdata.Add(hexString);
                            richTextBox1.AppendText("rx " + ctnPoint + ": " + hexString + "\n");
                            labelProgress.Text =  (ctnPoint + 1).ToString() + "/" + int.Parse(cbNumberPoint.Text).ToString();
                            ctnPoint++;
                        //}
                    }
                }
                else if (ctnPoint >= int.Parse(cbNumberPoint.SelectedItem.ToString()))
                {

                    DoneScan = true;
                    timerScan.Enabled = false;


                    ctnPoint = 0;
                    ctnScan = 0;
                    // Clear buffer
                    richTextBox1.AppendText("tx: " + "leave USB data mode\n");
                    serialPort1.Write(new byte[] { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 }, 0, 8);
                    serialPort1.Write(new byte[] { 0x20, 0x26, 0x02 }, 0, 3);



                }



            }


        }
        private void StartListeningSerialPort1()
        {
            // Đăng ký lại sự kiện DataReceived để tiếp tục nhận dữ liệu
            serialPort1.DataReceived += serialPort1_DataReceived;
        }
        private void StopListeningSerialPort1()
        {
            // Gỡ bỏ sự kiện DataReceived để ngừng nhận dữ liệu
            serialPort1.DataReceived -= serialPort1_DataReceived;
        }

        private void btConnectVNA_Click(object sender, EventArgs e)
        {
            //timer1.Interval = 3000;
            //timer1.Enabled = true;
            if (ComboBox_baud2.SelectedItem == null)
            {
                MessageBox.Show("Bạn chưa chọn baudrate.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                // Xử lý các thao tác khác tại đây nếu có baudrate được chọn

                if (!serialPort1.IsOpen)
                {
                    try
                    {
                        Control.CheckForIllegalCrossThreadCalls = false;
                        serialPort1.PortName = ComboBox_COM.SelectedItem.ToString();
                        serialPort1.BaudRate = int.Parse(ComboBox_baud2.SelectedItem.ToString());
                        serialPort1.Open();
                        btConnectVNA.Enabled = false;
                        btDisConnectVNA.Enabled = true;

                        ComboBox_COM.Enabled = false;
                        ComboBox_baud2.Enabled = false;
                        btConnectVNA.ForeColor = Color.Black;
                        btDisConnectVNA.ForeColor = Color.Indigo;

                        EnabledForm();
                        //requesting();
                        //SendParameter();
                    }
                    catch (UnauthorizedAccessException)
                    {
                        MessageBox.Show("Access to the port is denied. Please check if the port is already in use or you do not have the necessary permissions.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }


        }
        private void btConnectMCU_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem == null)
            {
                MessageBox.Show("Bạn chưa chọn baudrate.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                if (!serialPort2.IsOpen)
                {
                    try
                    {
                        Control.CheckForIllegalCrossThreadCalls = false;
                        serialPort2.PortName = comboBox2.SelectedItem.ToString();
                        serialPort2.BaudRate = int.Parse(comboBox1.SelectedItem.ToString());
                        serialPort2.Open();
                        btConnectMCU.Enabled = false;
                        btDisConnectMCU.Enabled = true;
                        comboBox2.Enabled = false;
                        comboBox1.Enabled = false;
                        btConnectMCU.ForeColor = Color.Indigo;
                        btDisConnectMCU.ForeColor = Color.Black;
                        EnabledForm();
                    }
                    catch (UnauthorizedAccessException)
                    {
                        MessageBox.Show("Access to the port is denied. Please check if the port is already in use or you do not have the necessary permissions.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            foreach (var series in chart1.Series)
            {
                series.Points.Clear();
            }
            foreach (var series in chart3.Series)
            {
                series.Points.Clear();
            }
            richTextBox1.Clear();
            Listdata.Clear();
        }

        private void btnSetAngle_Click(object sender, EventArgs e)
        {

            InitAngleSendToMCU = float.Parse(txtInitAngle.Text.ToString());
            string dataToSend = new string('*', 24);
            serialPort2.Write("I");
            serialPort2.Write(txtInitAngle.Text.PadLeft(8, '0'));
            serialPort2.Write(dataToSend);
        }

        private void btnAuto_Click(object sender, EventArgs e)
        {
            modeScan = 1;

            float CurentAngle = float.Parse(Angle_txtbox.Text.ToString());
            SetAngleSendToMCU = CurentAngle;
            float StopAngle = float.Parse(txtSetAngle.Text.ToString());
            float StepAngle = float.Parse(txtStepValue.Text.ToString());

            ctnAutoScan = Convert.ToInt32((StopAngle - CurentAngle) / StepAngle);
            SetAngleSendToMCU += StepAngle;

            serialPort2.Write("R");
            serialPort2.Write(SetAngleSendToMCU.ToString().PadLeft(8, '0'));
            
            if (rb_Clockwise.Checked)
            {
                serialPort2.Write("00000001");
            }
            else if (rb_Counter_Clockwise.Checked)
            {
                serialPort2.Write("00000000");
            }
            string dataToSend = new string('*', 16);
            serialPort2.Write(dataToSend);

        }



        private void btnScan_Click(object sender, EventArgs e)
        {
            modeScan = 0;
            // Write MCU
           serialPort2.Write("R");

           serialPort2.Write(txtSetAngle.Text.PadLeft(8, '0'));
            if (rb_Clockwise.Checked)
            {
                serialPort2.Write("00000001");
            }
            else if (rb_Counter_Clockwise.Checked)
            {
                serialPort2.Write("00000000");
            }
            string dataToSend = new string('*', 16);
            serialPort2.Write(dataToSend);

            //Write VNA



        }

        private void SendScanData()
        {
            serialPort1.Write(new byte[] { 0x18, 0x30, 0x00 }, 0, 3);

        }
        private void btnStop_Click(object sender, EventArgs e)
        {
            //serialPort2.Write("S");
            //serialPort2.Write("00000000");
            //serialPort2.Write(txtSetAngle.Text.PadLeft(8, '0'));
            //if (rb_Clockwise.Checked)
            //{
            //    serialPort2.Write("00000001");
            //}
            //else if (rb_Counter_Clockwise.Checked)
            //{
            //    serialPort2.Write("00000000");
            //}

            timerScan.Enabled = false;
            ctnPoint = 0;
            ctnScan = 0;
            serialPort1.Write(new byte[] { 0x20, 0x26, 0x02 }, 0, 3);
        }
        private void SendParameter()
        {
            byte[] cmdBuffer;
            byte[] cmdBytes;
            //  Number Of Point
            int NumberOfPoint = int.Parse(cbNumberPoint.SelectedItem.ToString());
            // Clear buffer
            serialPort1.Write(new byte[] { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 }, 0, 8);
            // USB Mode
            serialPort1.Write(new byte[] { 0x20, 0x26, 0x03 }, 0, 3);
            //  Start Frequency
            long StartFreq = int.Parse(txtStartFreq.Text.ToString()) * 1000000;
            cmdBuffer = BitConverter.GetBytes(StartFreq);
            cmdBytes = new byte[] { 0x23, 0x00, cmdBuffer[0], cmdBuffer[1], cmdBuffer[2], cmdBuffer[3], cmdBuffer[4], cmdBuffer[5], cmdBuffer[6], cmdBuffer[7] };
            serialPort1.Write(cmdBytes, 0, 10);
            //richTextBox1.AppendText(cmdBytes.ToString());
            //  Step Frequency
            long StopFreq = int.Parse(txtStopFreq.Text.ToString()) * 1000000;
            long stepFreq = (StopFreq - StartFreq) / (NumberOfPoint - 1);
            cmdBuffer = BitConverter.GetBytes(stepFreq);
            cmdBytes = new byte[] { 0x23, 0x10, cmdBuffer[0], cmdBuffer[1], cmdBuffer[2], cmdBuffer[3], cmdBuffer[4], cmdBuffer[5], cmdBuffer[6], cmdBuffer[7] };
            serialPort1.Write(cmdBytes, 0, 10);
            //richTextBox1.AppendText(cmdBytes.ToString());
            //  sweep of Point
            cmdBuffer = BitConverter.GetBytes(NumberOfPoint);
            cmdBytes = new byte[] { 0x21, 0x20, cmdBuffer[0], cmdBuffer[1] };
            serialPort1.Write(cmdBytes, 0, 4);
            //richTextBox1.AppendText(cmdBytes.ToString());
            //  Values per Frequency
            short valuesPerFreq = 1;
            cmdBuffer = BitConverter.GetBytes(valuesPerFreq);
            cmdBytes = new byte[] { 0x21, 0x22, cmdBuffer[0], 0x00 };
            serialPort1.Write(cmdBytes, 0, 4);
            //richTextBox1.AppendText(cmdBytes.ToString());

        }

        private void button7_Click(object sender, EventArgs e)
        {
            serialPort1.Write("S");
            string dataToSend = new string('*', 32);
            serialPort1.Write(dataToSend);

        }
        private void Form1_Load(object sender, EventArgs e)
        {
            InitializeSerialPort1();
            InitializeSerialPort2();
            ComboBox_baud2.SelectedIndex = 0;
            comboBox1.SelectedIndex = 4;
            cbNumberPoint.SelectedIndex = 2;
            DisabledForm();

            // Init default mode
            //Setangle_txt.Text = "0";
            //Kps_txtbox.Text = "0.0004";
            //Kis_txtbox.Text = "0.002";
            //Kds_txtbox.Text = "0.00001";

            time_angle = 0;
            n = 5;
            Ts = 0.01 * n;
        }


        private StringBuilder receivedDataSeral1 = new StringBuilder();



        private void button2_Click_1(object sender, EventArgs e)
        {
            foreach (var series in chart1.Series)
            {
                series.Points.Clear();
            }
            PlotData();
        }

        private void txtStopFreq_TextChanged(object sender, EventArgs e)
        {
            txtSpanFreq.Text = (int.Parse(txtStopFreq.Text.ToString()) - int.Parse(txtStartFreq.Text.ToString())).ToString();
        }

        private void txtStartFreq_TextChanged(object sender, EventArgs e)
        {
            txtSpanFreq.Text = (int.Parse(txtStopFreq.Text.ToString()) - int.Parse(txtStartFreq.Text.ToString())).ToString();
        }

        private void DisabledForm()
        {
           
            groupBox5.Enabled = false;
            groupBox10.Enabled = false;
            groupBox9.Enabled = false;
            groupBox8.Enabled = false;
            btnExport.Enabled   = false;
            groupBox2.Enabled = false;
            groupBox7.Enabled = false;

        }
        private void EnabledForm()
        {
           
            groupBox5.Enabled = true;
            groupBox10.Enabled = true;
            groupBox9.Enabled = true;
            groupBox8.Enabled = true;
            btnExport.Enabled = true;
            groupBox2.Enabled = true;
            groupBox7.Enabled = true;

        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            List<double> freqIndexList = new List<double>();
            List<double> parameterListS11 = new List<double>();
            foreach (string line in Listdata)
            {
                Dictionary<string, double> decodedValues = DecodeFrame(line);
                double parameter = Calculate(decodedValues["fwd0Re"], decodedValues["fwd0Im"], decodedValues["rev0Re"], decodedValues["rev0Im"], decodedValues["freqIndex"]);
                freqIndexList.Add(decodedValues["freqIndex"]);
                parameterListS11.Add(parameter);
            }
            ExportToExcel(freqIndexList, parameterListS11, filePath, 10);
        }

        private void InitializeSerialPort1()
        {

            ComboBox_COM.Items.Clear();

            string[] ports = SerialPort.GetPortNames();

            foreach (string port in ports)
            {
                ComboBox_COM.Items.Add(port);
            }

            if (ComboBox_COM.Items.Count > 0)
            {
                ComboBox_COM.SelectedIndex = 0;
            }
            //else
            //{
            //    MessageBox.Show("Can not find COM port.");
            //}
        }



        private void timer2_Tick_1(object sender, EventArgs e)
        {
            if (ctnPoint < int.Parse(cbNumberPoint.SelectedItem.ToString()))
            {

                //SendParameter();
                SendScanData();
                richTextBox1.AppendText("Scan: " + ctnScan + "\t");
                ctnScan++;

            }
            else
            {

                DoneScan = true;
                richTextBox1.AppendText(ctnPoint + " points received\n");
                ctnPoint = 0;
                ctnScan = 0;

                // Clear buffer
                serialPort1.Write(new byte[] { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 }, 0, 8);
                serialPort1.Write(new byte[] { 0x20, 0x26, 0x02 }, 0, 3);


                timerScan.Enabled = false;
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            
            if (DoneScan == true)
            {
                List<double> freqIndexList = new List<double>();
                List<double> parameterListS11 = new List<double>();
                richTextBox1.AppendText("Auto " + ctnAutoExcel + " Done\n");
                foreach (string line in Listdata)
                {
                    Dictionary<string, double> decodedValues = DecodeFrame(line);
                    double parameter = Calculate(decodedValues["fwd0Re"], decodedValues["fwd0Im"], decodedValues["rev0Re"], decodedValues["rev0Im"], decodedValues["freqIndex"]);
                    freqIndexList.Add(decodedValues["freqIndex"]);
                    parameterListS11.Add(parameter);
                }
                ExportToExcel(freqIndexList, parameterListS11, filePath, ctnAutoExcel);

                if (ctnAutoExcel <= ctnAutoScan+2)
                {
                    freqIndexList.Clear();
                    parameterListS11.Clear();
                    Listdata.Clear();

                    //ctnAutoExcel++;
                    //timerScan.Enabled = true;
                    //DoneScan = false;
                    //AutoScan = true;
                    //SendScanData();

                    float StepAngle = float.Parse(txtStepValue.Text.ToString());
                    SetAngleSendToMCU += StepAngle;
                    serialPort2.Write("R");
                    serialPort2.Write(SetAngleSendToMCU.ToString().PadLeft(8, '0'));
                    if (rb_Clockwise.Checked)
                    {
                        serialPort2.Write("00000001");
                    }
                    else if (rb_Counter_Clockwise.Checked)
                    {
                        serialPort2.Write("00000000");
                    }
                    string dataToSend = new string('*', 16);
                    serialPort2.Write(dataToSend);

                }
                else
                {
                    ctnAutoScan = 0;
                    ctnAutoExcel = 2;
                    AutoScan = false;
                    richTextBox1.AppendText("Auto Mode Done \n");
                    System.Media.SystemSounds.Exclamation.Play();
                    Thread.Sleep(1000);
                    timerAutoMode.Enabled = false;
                }
            }
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            foreach (var series in chart2.Series)
            {
                series.Points.Clear();
            }
            PlotPolar(int.Parse(txtFreqIndex.Text.ToString()));
        }

        private void txtInitAngle_TextChanged(object sender, EventArgs e)
        {

        }

        private void send_Click(object sender, EventArgs e)
        {
            serialPort2.Write("P");
            serialPort2.Write(Setangle_txt.Text.PadLeft(8, '0'));
            serialPort2.Write(Kps_txtbox.Text.PadLeft(8, '0'));
            serialPort2.Write(Kis_txtbox.Text.PadLeft(8, '0'));
            serialPort2.Write(Kds_txtbox.Text.PadLeft(8, '0'));
            set_angle = int.Parse(Setangle_txt.Text);
        }

        private void btnPlotChartFromDatabase_Click(object sender, EventArgs e)
        {
            foreach (var series in chart1.Series)
            {
                series.Points.Clear();
            }
            PlotChartFromExcel(int.Parse(txtAngleSelect.Text.ToString()));
        }

        private void InitializeSerialPort2()
        {

            comboBox2.Items.Clear();

            string[] ports = SerialPort.GetPortNames();

            foreach (string port in ports)
            {
                comboBox2.Items.Add(port);
            }

            if (comboBox2.Items.Count > 0)
            {
                comboBox2.SelectedIndex = 0;
            }
            //else
            //{
            //    MessageBox.Show("Can not find COM port.");
            //}
        }

        private void scanVNA_Click(object sender, EventArgs e)
        {
            // Write VNA
            timerScan.Enabled = true;
            DoneScan = false;
            Listdata.Clear();
            SendParameter();
            richTextBox1.AppendText("requesting " + int.Parse(cbNumberPoint.SelectedItem.ToString()) + " points .." + "\n");
            SendScanData();
        }

        private void button2_Click_2(object sender, EventArgs e)
        {
            ctnAutoExcel++;
            timerAutoMode.Enabled = true;
            timerScan.Enabled = true;
            DoneScan = false;
            AutoScan = true;
        }

        private void requesting()
        {
            serialPort1.Write(new byte[] { 0x10, 0xf2 }, 0, 2);
            serialPort1.Write(new byte[] { 0x10, 0xf3 }, 0, 2);
            serialPort1.Write(new byte[] { 0x10, 0xf4 }, 0, 2);
            serialPort1.Write(new byte[] { 0x11, 0x5c }, 0, 2);
            richTextBox1.AppendText("tx: polling - requesting hw-revision and fw-version\n");
            richTextBox1.AppendText("tx: 10 f2 10 f3 10 f4 11 5c\n");
        }
        private void Setpoint_txt_TextChanged(object sender, EventArgs e)
        {

        }

        public static void ExportToExcel(List<double> freqIndexList, List<double> parameterList, string filePath, int parameterColumn)
        {
            if (freqIndexList == null) throw new ArgumentNullException(nameof(freqIndexList));
            if (parameterList == null) throw new ArgumentNullException(nameof(parameterList));
            if (freqIndexList.Count != parameterList.Count) throw new ArgumentException("Lists must have the same length");
            if (parameterColumn < 2) throw new ArgumentOutOfRangeException(nameof(parameterColumn), "parameterColumn must be 2 or greater");

            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                // Initialize Excel application
                excelApp = new Excel.Application();
                excelApp.DisplayAlerts = false;
                excelApp.ScreenUpdating = false;
                excelApp.EnableEvents = false;

                // Open the existing workbook if it exists, otherwise create a new one
                if (System.IO.File.Exists(filePath))
                {
                    workbook = excelApp.Workbooks.Open(filePath);
                }
                else
                {
                    workbook = excelApp.Workbooks.Add();
                }

                worksheet = workbook.Worksheets[1];



                // Clear existing data in the column (excluding the header)
                Excel.Range columnRange = worksheet.Columns[parameterColumn];
                //Excel.Range cellsToClear = columnRange.Cells;
                //cellsToClear.ClearContents();
                columnRange.Delete();
                worksheet.Cells[1, 1] = "Frequency Index";

                worksheet.Cells[1, parameterColumn] = parameterColumn.ToString();
                // Add data
                for (int i = 0; i < freqIndexList.Count; i++)
                {
                    worksheet.Cells[i + 2, 1] = freqIndexList[i];
                    worksheet.Cells[i + 2, parameterColumn] = parameterList[i];
                }

                // Save workbook
                workbook.SaveAs(filePath);

            }
            catch (Exception ex)
            {
                // Handle exceptions
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
            finally
            {
                // Release COM objects to avoid memory leaks
                if (workbook != null)
                {
                    workbook.Close(false);
                    Marshal.ReleaseComObject(workbook);
                }
                if (worksheet != null)
                {
                    Marshal.ReleaseComObject(worksheet);
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }
            }

        }


        public class Complex
        {
            public double Real { get; set; }
            public double Imaginary { get; set; }

            public Complex(double real, double imaginary)
            {
                Real = real;
                Imaginary = imaginary;
            }

            public static Complex operator /(Complex c1, Complex c2)
            {
                double real = (c1.Real * c2.Real + c1.Imaginary * c2.Imaginary) / (c2.Real * c2.Real + c2.Imaginary * c2.Imaginary);
                double imaginary = (c1.Imaginary * c2.Real - c1.Real * c2.Imaginary) / (c2.Real * c2.Real + c2.Imaginary * c2.Imaginary);
                return new Complex(real, imaginary);
            }

            public double Magnitude
            {
                get { return Math.Sqrt(Real * Real + Imaginary * Imaginary); }
            }
        }
        private Dictionary<string, double> DecodeFrame(string hexString)
        {
            byte[] byteData = FromHexString(hexString);

            double fwd0Re = BitConverter.ToInt32(byteData, 0);
            double fwd0Im = BitConverter.ToInt32(byteData, 4);
            double rev0Re = BitConverter.ToInt32(byteData, 8);
            double rev0Im = BitConverter.ToInt32(byteData, 12);
            double rev1Re = BitConverter.ToInt32(byteData, 16);
            double rev1Im = BitConverter.ToInt32(byteData, 20);
            int freqIndex = BitConverter.ToUInt16(byteData, 24);
            int freq = freqIndex * (int.Parse(txtSpanFreq.Text.ToString()) / (int.Parse(cbNumberPoint.SelectedItem.ToString()) - 1)) + int.Parse(txtStartFreq.Text.ToString());
            return new Dictionary<string, double>
            {
                { "fwd0Re", fwd0Re },
                { "fwd0Im", fwd0Im },
                { "rev0Re", rev0Re },
                { "rev0Im", rev0Im },
                { "rev1Re", rev1Re },
                { "rev1Im", rev1Im },
                { "freqIndex", freq }
            };
        }

        private byte[] FromHexString(string hexString)
        {
            int length = hexString.Length;
            byte[] bytes = new byte[length / 2];
            for (int i = 0; i < length; i += 2)
            {
                bytes[i / 2] = Convert.ToByte(hexString.Substring(i, 2), 16);
            }
            return bytes;
        }

        private double Calculate(double fwd0Re, double fwd0Im, double rev0Re, double rev0Im, double freqIndex)
        {
            Complex A1 = new Complex(fwd0Re, fwd0Im);
            Complex A2 = new Complex(rev0Re, rev0Im);
            Complex ratio = A2 / A1;
            return 20 * Math.Log10(ratio.Magnitude);
        }

        private void PlotChartFromExcel(int AngleIndex)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            List<string> FreqList = new List<string>();
            List<string> parameterListS11 = new List<string>();

            try
            {
                using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheets = excelPackage.Workbook.Worksheets;
                    string sheetName = "Sheet1";
                    ExcelWorksheet worksheet = worksheets[sheetName];
                    if (worksheet == null)
                    {
                        MessageBox.Show($"Worksheet '{sheetName}' does not exist.");
                        return;
                    }

                    if (worksheet.Dimension == null)
                    {
                        MessageBox.Show("Worksheet is empty.");
                        return;
                    }

                    int rowCount = worksheet.Dimension.End.Row;

                    for (int row = 2; row <= rowCount; row++)
                    {
                        string cellValue = worksheet.Cells[AngleIndex, row].Value?.ToString();
                        string cellFreqValue = worksheet.Cells[1, row].Value?.ToString();

                        if (!string.IsNullOrEmpty(cellValue) && !string.IsNullOrEmpty(cellFreqValue))
                        {
                            if (double.TryParse(cellValue, out double parameter) && int.TryParse(cellFreqValue, out int freq))
                            {
                                parameterListS11.Add(cellValue);
                                FreqList.Add(cellFreqValue);
                            }
                            else
                            {
                                MessageBox.Show($"Invalid data at column {row}");
                                return;
                            }
                        }
                    }
                }

                for (int i = 0; i < FreqList.Count; i++)
                {
                    chart1.Series["S11"].Points.AddXY(int.Parse(FreqList[i]), double.Parse(parameterListS11[i]));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}");
            }

        }
            private void PlotData()
            {
            List<double> freqIndexList = new List<double>();
            List<double> parameterListS11 = new List<double>();

            //using (StreamReader reader = new StreamReader(fileName))
            //{
            //    string line;
            //    while ((line = reader.ReadLine()) != null)
            //    {
            //        Listdata.Add(line);
            //    }
            //}

            foreach (string line in Listdata)
            {
                Console.WriteLine(line);
            }


            foreach (string line in Listdata)
            {
                Dictionary<string, double> decodedValues = DecodeFrame(line);
                double parameter = Calculate(decodedValues["fwd0Re"], decodedValues["fwd0Im"], decodedValues["rev0Re"], decodedValues["rev0Im"], decodedValues["freqIndex"]);
                freqIndexList.Add(decodedValues["freqIndex"]);
                parameterListS11.Add(parameter);

                Console.WriteLine("fwd0Re: " + decodedValues["fwd0Re"]);
                Console.WriteLine("fwd0Im: " + decodedValues["fwd0Im"]);
                Console.WriteLine("rev0Re: " + decodedValues["rev0Re"]);
                Console.WriteLine("rev0Im: " + decodedValues["rev0Im"]);
                Console.WriteLine("freqIndex: " + decodedValues["freqIndex"]);


            }

            for (int i = 0; i < freqIndexList.Count; i++)
            {
                chart1.Series["S11"].Points.AddXY(freqIndexList[i], parameterListS11[i]);
                Console.WriteLine(freqIndexList[i]);
                Console.WriteLine(parameterListS11[i]);
            }


        }
        private void PlotPolar(int FreqIndex)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            List<string> AngleList = new List<string>();
            List<string> parameterListS11 = new List<string>();

            try
            {
                using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheets = excelPackage.Workbook.Worksheets;
                    string sheetName = "Sheet1";
                    ExcelWorksheet worksheet = worksheets[sheetName];
                    if (worksheet == null)
                    {
                        MessageBox.Show($"Worksheet '{sheetName}' does not exist.");
                        return;
                    }

                    if (worksheet.Dimension == null)
                    {
                        MessageBox.Show("Worksheet is empty.");
                        return;
                    }

                    int columnCount = worksheet.Dimension.End.Column;

                    for (int col = 2; col <= columnCount; col++) 
                    {
                        string cellValue = worksheet.Cells[FreqIndex, col].Value?.ToString();
                        string cellAngleValue = worksheet.Cells[1, col].Value?.ToString();

                        if (!string.IsNullOrEmpty(cellValue) && !string.IsNullOrEmpty(cellAngleValue))
                        {
                            if (double.TryParse(cellValue, out double parameter) && int.TryParse(cellAngleValue, out int angle))
                            {
                                parameterListS11.Add(cellValue);
                                AngleList.Add(cellAngleValue);
                            }
                            else
                            {
                                MessageBox.Show($"Invalid data at column {col}");
                                return;
                            }
                        }
                    }
                }
                
                for (int i = 0; i < AngleList.Count; i++)
                {
                    chart2.Series["Series1"].Points.AddXY(int.Parse(AngleList[i]), double.Parse(parameterListS11[i]));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}");
            }
        }

    }
}
