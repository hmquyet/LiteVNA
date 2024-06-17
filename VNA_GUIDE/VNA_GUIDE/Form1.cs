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
using System.Timers;
using Microsoft.Office.Interop.Excel;


namespace VNA_GUIDE
{
    public partial class Form1 : Form
    {
        private System.Timers.Timer myTimer;
        // Khai báo một biến để đánh dấu khi cần dừng Timer
        private bool stopTimerScan = false;

        private double time, set_angle, angle, pwm, Iref, Idc, Ts, time_angle, n, step_enable, tick;
        List<Tuple<byte, string>> pairs = new List<Tuple<byte, string>>();

        string rev_data;
        string rev1_data;
        int Enable_Scan;
        float Old_Trigger_Scan;
        float Trigger_Scan;
        int ctnAutoScan = 0 ;
        float InitAngleSendToMCU;
        float SetAngleSendToMCU;
        string oldData;
        int modeScan;
        string filePathS11 = @"D:\NextwaveIndustries\liteVNA\DATA\DATA_S11.xlsx";
        string filePathS21 = @"D:\NextwaveIndustries\liteVNA\DATA\DATA_S21.xlsx";
        List<string> Listdata = new List<string>();

   

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
            // Khởi tạo Timer từ System.Windows.Forms.Timer


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
            if (!serialPort2.IsOpen)
            {
                DisabledForm();
            }
                
            
            

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
            if (!serialPort1.IsOpen)
            {
                DisabledForm();
            }
        }
        int filterTriggerScan = 0;
        private void ScrollToBottom(RichTextBox richTextBox)
        {
            richTextBox.SelectionStart = richTextBox.Text.Length;
            richTextBox.ScrollToCaret();
        }
        float oldpwm =0;
        // Hàm xử lý khi Timer kích hoạt
       

        // Hàm bắt đầu Timer
        private void StartMyTimer()
        {

            // Bắt đầu Timer
            timerValid.Start();
        }

        // Hàm dừng Timer
        private void StopMyTimer()
        {
            // Dừng Timer
            timerValid?.Dispose();
        }

        // Hàm để bắt đầu xử lý
        private void serialPort2_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            if (serialPort2.IsOpen)
            {
                    string data = serialPort2.ReadLine();
                //richTextBox1.AppendText(data);
                //ScrollToBottom(richTextBox1);

                // Draw PID
                if ( data.StartsWith("A"))
                {
                    data = data.Substring(1);

                    // Validate if rev_data can be converted to a double
                    if (double.TryParse(data, out double angle))
                    {
                        if (Angle_txtbox.Text != data)
                            Angle_txtbox.Text = data;

                        if (txt_ResAngle.Text != data)
                            txt_ResAngle.Text = data;

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

                            time_angle += 1;
                        }
                    }
                }
                else if (data.StartsWith("u"))
                {
                    data = data.Substring(1);
                    if (double.TryParse(data, out double pwm))
                    {
                        if (pwm_txtbox.Text != data)
                            pwm_txtbox.Text = data;
                        

                        if (textBoxPwm.Text != data)
                            textBoxPwm.Text = data;
                    }
                    float.TryParse(data, out oldpwm);  
                }
                else if (data.StartsWith("X"))
                {
                    data = data.Substring(1);
                    if (float.TryParse(data, out float Trigger_temp_Scan))
                    {
                        if (txtScanEnable.Text != Trigger_temp_Scan.ToString())
                            txtScanEnable.Text = Trigger_temp_Scan.ToString();
                            Old_Trigger_Scan = Trigger_temp_Scan;
                        if (Trigger_temp_Scan == 1)
                        {
                            Trigger_Scan = 1;
                        }
                    }
                }

            }
        }

        int ctnPoint = 0;
        int ctnScan = 0;
        bool DoneScan = false;
        bool AutoScan = false;


        //private void serialPort1_DataReceived(object sender, SerialDataReceivedEventArgs e)
        //{
        //    byte[] buffer = new byte[32]; // Mảng byte để lưu trữ dữ liệu
        //    int bytesRead = serialPort1.Read(buffer, 0, buffer.Length);

        //    if (bytesRead >= 32)
        //    {
        //        string hexString = BitConverter.ToString(buffer).Replace("-", "");

        //        if (ctnPoint < int.Parse(cbNumberPoint.SelectedItem.ToString())) // int.Parse(cbNumberPoint.SelectedItem.ToString())
        //        {
        //            if (DoneScan == false)
        //            {
        //                string lastTwoChars = hexString.Substring(hexString.Length - 2);
        //                string lastfourChars = hexString.Substring(hexString.Length - 4);

        //                //Dictionary<string, double> decodedValues = DecodeFrame(hexString);
        //                //int freqIndex = int.Parse(decodedValues["freqIndex"].ToString());

        //                Console.Write(hexString + "\n");

        //                if ("00000040" != hexString.Substring(0, 8))
        //                {
        //                    if (lastTwoChars == "00")
        //                    {

        //                        hexString = "00" + hexString.Substring(0, hexString.Length - 2);

        //                        Listdata.Add(hexString);
        //                        ctnPoint++;


        //                    }
        //                    else if (lastfourChars == "0000")
        //                    {
        //                        hexString = "0000" + hexString.Substring(0, hexString.Length - 4);
        //                        Listdata.Add(hexString);
        //                        ctnPoint++;
        //                    }
        //                }
        //                else
        //                {
        //                    Listdata.Add(hexString);
        //                    ctnPoint++;
        //                }
        //                Console.Write("resutl:" + hexString + "\n");
        //                //richTextBox2.AppendText("rx " + ctnPoint + ": " + hexString + "\n");
        //                //ScrollToBottom(richTextBox2);


        //            }
        //        }
        //        else if (ctnPoint >= int.Parse(cbNumberPoint.SelectedItem.ToString()))
        //        {
        //            DoneScan = true;
        //            timerScan.Enabled = false;
        //            ctnPoint = 0;
        //            ctnScan = 0;
        //            // Clear buffer

        //            richTextBox1.AppendText("tx: " + "leave USB data mode\n");
        //            Console.Write("leave USB data mode\n");
        //            ScrollToBottom(richTextBox1);
        //            serialPort1.Write(new byte[] { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 }, 0, 8);
        //            serialPort1.Write(new byte[] { 0x20, 0x26, 0x02 }, 0, 3);

        //        }
        //        textNP.Text = (ctnPoint).ToString() + "/" + int.Parse(cbNumberPoint.SelectedItem.ToString()).ToString();
        //    }


        //}

        private byte CalculateChecksum(byte[] data)
        {
            byte checksum = 0x46; // Giá trị khởi tạo checksum, có thể thay đổi tùy thuộc vào thiết kế của bạn
            for (int i = 0; i < data.Length - 1; i++) // Trừ 1 vì byte cuối cùng là giá trị checksum cần kiểm tra
            {
                checksum = (byte)((checksum ^ ((checksum << 1) | 1u)) ^ data[i]);
            }
            return checksum;
        }

        //private void serialPort1_DataReceived(object sender, SerialDataReceivedEventArgs e)
        //{

        //    byte[] buffer = new byte[32];
        //    int bytesRead = serialPort1.Read(buffer, 0, buffer.Length);

        //    string hexString = BitConverter.ToString(buffer, 0, 31).Replace("-", "");
        //    Console.Write(hexString + "\n");
        //    if (bytesRead >= 32)
        //    {
        //        byte receivedChecksum = buffer[31]; // Giả sử checksum là byte cuối cùng
        //        byte calculatedChecksum = CalculateChecksum(buffer);

        //        if (calculatedChecksum == receivedChecksum)
        //        {
        //            // Xử lý dữ liệu nếu checksum hợp lệ
        //            // Chỉ lấy 31 byte đầu

        //            if (ctnPoint < int.Parse(cbNumberPoint.SelectedItem.ToString())) // int.Parse(cbNumberPoint.SelectedItem.ToString())
        //            {
        //                if (DoneScan == false)
        //                {
        //                    //Dictionary<string, double> decodedValues = DecodeFrame(hexString);
        //                    //int freqIndex = int.Parse(decodedValues["freqIndex"].ToString());
        //                    Listdata.Add(hexString);
        //                    ctnPoint++;

        //                }
        //            }
        //            else if (ctnPoint == int.Parse(cbNumberPoint.SelectedItem.ToString()))
        //            {
        //                double max_index_freq ;
        //                List<double> freqIndexListBuffer = new List<double>();
        //                foreach (string line in Listdata)
        //                {
        //                    Dictionary<string, double> decodedValues = DecodeFrame(line);
        //                    freqIndexListBuffer.Add(decodedValues["freqIndex"]);
        //                }
        //                max_index_freq = (int)freqIndexListBuffer.Max();

        //                Console.Write("\n"+"max freq" + max_index_freq );
        //                Console.Write("\n" + "max freq" + (int.Parse(cbNumberPoint.SelectedItem.ToString()) - 1)+ "\n");

        //                //if (max_index_freq == (int.Parse(cbNumberPoint.SelectedItem.ToString())-1))
        //                //{
        //                    DoneScan = true;
        //                    timerScan.Enabled = false;
        //                    ctnPoint = 0;
        //                    ctnScan = 0;
        //                    // Clear buffer

        //                    richTextBox1.AppendText("tx: " + "leave USB data mode\n");
        //                    Console.Write("leave USB data mode\n");
        //                    ScrollToBottom(richTextBox1);
        //                    serialPort1.Write(new byte[] { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 }, 0, 8);
        //                    serialPort1.Write(new byte[] { 0x20, 0x26, 0x02 }, 0, 3);


        //                //}



        //            }
        //            textNP.Text = (ctnPoint).ToString() + "/" + int.Parse(cbNumberPoint.SelectedItem.ToString()).ToString();
        //        }
        //        else
        //        {
        //            // Báo lỗi nếu checksum không hợp lệ


        //            Dictionary<string, double> decodedValues = DecodeFrame(hexString);
        //            int freqIndex = int.Parse(decodedValues["freqIndex"].ToString());


        //            if ( freqIndex< int.Parse(cbNumberPoint.SelectedItem.ToString()) && 0 <= freqIndex)
        //                {

        //                    Listdata.Add(hexString);
        //                    ctnPoint++;

        //                }
        //            else{
        //                Console.WriteLine("Checksum error");

        //            }



        //        }
        //    }
        //    //Console.Write(hexString + "\n");
        //}

        private void serialPort1_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {

            try
            {
                if (DoneScan == false)
                {


                    int totalBytesToRead = 32 * int.Parse(cbNumberPoint.SelectedItem.ToString()); // Độ dài tổng của dữ liệu mong đợi
                    byte[] buffer = new byte[totalBytesToRead];
                    int totalBytesRead = 0;

                    while (totalBytesRead < totalBytesToRead && serialPort1.IsOpen)
                    {
                        int bytesRead = serialPort1.Read(buffer, totalBytesRead, totalBytesToRead - totalBytesRead);
                        if (bytesRead == 0)
                        {
                            throw new Exception("Không nhận được thêm dữ liệu từ cổng serial.");
                        }
                        totalBytesRead += bytesRead;
                    }

                    if (serialPort1.IsOpen)
                    {
                        Console.WriteLine(BitConverter.ToString(buffer).Replace("-", ""));

                        int bytesPerSegment = 32;
                        for (int i = 0; i < totalBytesRead; i += bytesPerSegment)
                        {
                            int length = Math.Min(bytesPerSegment, totalBytesRead - i);
                            string hexString = BitConverter.ToString(buffer, i, length).Replace("-", "");

                            Dictionary<string, double> decodedValues = DecodeFrame(hexString);
                            int freqIndex = int.Parse(decodedValues["freqIndex"].ToString());

                            double parameterS11 = CalculateS11(decodedValues["rev0Re"], decodedValues["rev0Im"], decodedValues["rev1Re"], decodedValues["rev1Im"], decodedValues["freqIndex"]);
                            double parameterS21 = CalculateS21(decodedValues["rev0Re"], decodedValues["rev0Im"], decodedValues["rev1Re"], decodedValues["rev1Im"], decodedValues["freqIndex"]);
                            Console.Write(freqIndex + ": " + parameterS11+"\t" + parameterS21 + "\t");

                            Listdata.Add(hexString);

                            if (freqIndex == (int.Parse(cbNumberPoint.SelectedItem.ToString()) - 1))
                            {
                                DoneScan = true;
                                ctnScan = 0;
                                richTextBox1.AppendText("leave USB data mode\n");
                                serialPort1.Write(new byte[] { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 }, 0, 8);
                                serialPort1.Write(new byte[] { 0x20, 0x26, 0x02 }, 0, 3);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
               
            }

        }


        private void StartListeningSerialPort1()
        {
            serialPort1.DataReceived += serialPort1_DataReceived;
        }
        private void StopListeningSerialPort1()
        {
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
                        if (!serialPort2.IsOpen)
                        {
                            EnabledForm();
                        }
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
            foreach (var series in chart2.Series)
            {
                series.Points.Clear();
            }
            foreach (var series in chart3.Series)
            {
                series.Points.Clear();
            }
            foreach (var series in chart4.Series)
            {
                series.Points.Clear();
            }
            ctnAutoScan = 0;
            ctnAutoExcel = 2;
            timerScan.Enabled = false;
            timerAutoMode.Enabled = false;
            timerValid.Enabled = false;
            richTextBox1.Clear();
            richTextBox2.Clear();
            Listdata.Clear();
            StopMyTimer();
        }

        private void btnSetAngle_Click(object sender, EventArgs e)
        {
            if (serialPort2.IsOpen)
            {
                InitAngleSendToMCU = float.Parse(txtInitAngle.Text.ToString());
                CurentAngle = (int)InitAngleSendToMCU;
                serialPort2.Write("P");
                serialPort2.Write("00000000");
                serialPort2.Write("00000000");
                serialPort2.Write("00000000");
                serialPort2.Write("00000000");
                Thread.Sleep(200);
                string dataToSend = new string('*', 24);
                serialPort2.Write("I");
                serialPort2.Write(txtInitAngle.Text.PadLeft(8, '0'));
                serialPort2.Write(dataToSend);

                Thread.Sleep(200);
                serialPort2.Write("P");
                serialPort2.Write(txtInitAngle.Text.PadLeft(8, '0'));
                serialPort2.Write(Kps_txtbox.Text.PadLeft(8, '0'));
                serialPort2.Write(Kis_txtbox.Text.PadLeft(8, '0'));
                serialPort2.Write(Kds_txtbox.Text.PadLeft(8, '0'));
            }
            else
            {
                MessageBox.Show($"serialPort not available");
            }
        }
        int CurentAngle = 0;
        private void btnAuto_Click(object sender, EventArgs e)
        {
            if (serialPort2.IsOpen && serialPort1.IsOpen)
            {
                modeScan = 1;
                StartMyTimer();

                SetAngleSendToMCU = CurentAngle;
                int StopAngle = int.Parse(txtSetAngle.Text.ToString());
                int StepAngle = int.Parse(txtStepValue.Text.ToString());

                ctnAutoScan = (StopAngle - CurentAngle) / StepAngle;
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
                MessageBox.Show($"serialPort not available");
            }
        }



        private void btnScan_Click(object sender, EventArgs e)
        {
            if (serialPort2.IsOpen && serialPort1.IsOpen)
            {

            modeScan = 0;
            StartMyTimer();
           
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
            else
            {
           
                 MessageBox.Show($"serialPort not available");
               
            }

        }

        private void SendScanData()
        {

            serialPort1.Write(new byte[] { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 }, 0, 8);
            serialPort1.Write(new byte[] { 0x20, 0x26, 0x02 }, 0, 3);
            //SendParameter();
            serialPort1.Write(new byte[] { 0x18, 0x30, 0x00 }, 0, 3);

        }
        private void btnStop_Click(object sender, EventArgs e)
        {
            if (serialPort2.IsOpen && serialPort1.IsOpen)
            {
                serialPort2.Write("S");
                serialPort2.Write("00000000");
                serialPort2.Write(txtSetAngle.Text.PadLeft(8, '0'));
                if (rb_Clockwise.Checked)
                {
                    serialPort2.Write("00000001");
                }
                else if (rb_Counter_Clockwise.Checked)
                {
                    serialPort2.Write("00000000");
                }
                ctnAutoScan = 0;
                ctnAutoExcel = 2;
                timerScan.Enabled = false;
                timerAutoMode.Enabled = false;
                timerValid.Enabled = false;
                ctnPoint = 0;
                ctnScan = 0;
                serialPort1.Write(new byte[] { 0x20, 0x26, 0x02 }, 0, 3);
            }
            else
            {
                MessageBox.Show($"serialPort not available");
            }
        }
        private void SendParameter()
        {
            //byte[] cmdBuffer;
            //byte[] cmdBytes;
            ////  Number Of Point
            //int NumberOfPoint = int.Parse(cbNumberPoint.SelectedItem.ToString());
            //// Clear bufferF
            //serialPort1.Write(new byte[] { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 }, 0, 8);
            //// USB Mode
            //serialPort1.Write(new byte[] { 0x20, 0x26, 0x03 }, 0, 3);
            ////  Start Frequency
            //long StartFreq = int.Parse(txtStartFreq.Text.ToString()) * 1000000;
            //cmdBuffer = BitConverter.GetBytes(StartFreq);
            //cmdBytes = new byte[] { 0x23, 0x00, cmdBuffer[0], cmdBuffer[1], cmdBuffer[2], cmdBuffer[3], cmdBuffer[4], cmdBuffer[5], cmdBuffer[6], cmdBuffer[7] };
            //serialPort1.Write(cmdBytes, 0, 10);
            ////richTextBox1.AppendText(cmdBytes.ToString());
            ////  Step Frequency
            //long StopFreq = int.Parse(txtStopFreq.Text.ToString()) * 1000000;
            //long stepFreq = (StopFreq - StartFreq) / (NumberOfPoint - 1);
            //cmdBuffer = BitConverter.GetBytes(stepFreq);
            //cmdBytes = new byte[] { 0x23, 0x10, cmdBuffer[0], cmdBuffer[1], cmdBuffer[2], cmdBuffer[3], cmdBuffer[4], cmdBuffer[5], cmdBuffer[6], cmdBuffer[7] };
            //serialPort1.Write(cmdBytes, 0, 10);
            ////richTextBox1.AppendText(cmdBytes.ToString());
            ////  sweep of Point
            //cmdBuffer = BitConverter.GetBytes(NumberOfPoint);
            //cmdBytes = new byte[] { 0x21, 0x20, cmdBuffer[0], cmdBuffer[1] };
            //serialPort1.Write(cmdBytes, 0, 4);
            ////richTextBox1.AppendText(cmdBytes.ToString());
            ////  Values per Frequency
            //short valuesPerFreq = 1;
            //cmdBuffer = BitConverter.GetBytes(valuesPerFreq);
            //cmdBytes = new byte[] { 0x21, 0x22, cmdBuffer[0], 0x00 };
            //serialPort1.Write(cmdBytes, 0, 4);
            ////richTextBox1.AppendText(cmdBytes.ToString());
            ///

            List<byte> cmdList = new List<byte>();
            // Lấy số điểm đo từ giao diện người dùng
            int NumberOfPoint = int.Parse(cbNumberPoint.SelectedItem.ToString());

            // Tần số bắt đầu
            long StartFreq = int.Parse(txtStartFreq.Text.ToString()) * 1000000;
            byte[] startFreqBytes = BitConverter.GetBytes(StartFreq);

            // Tần số kết thúc và bước tần số
            long StopFreq = int.Parse(txtStopFreq.Text.ToString()) * 1000000;
            long stepFreq = (StopFreq - StartFreq) / (NumberOfPoint - 1);
            byte[] stepFreqBytes = BitConverter.GetBytes(stepFreq);

            // Số điểm quét
            byte[] numberOfPointBytes = BitConverter.GetBytes(NumberOfPoint);

            // Số giá trị trên mỗi tần số
            short valuesPerFreq = 1;
            byte[] valuesPerFreqBytes = BitConverter.GetBytes(valuesPerFreq);
            // Clear buffer
            cmdList.AddRange(new byte[] { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 });

            // USB Mode
            cmdList.AddRange(new byte[] { 0x20, 0x26, 0x03 });

            // Start Frequency
            cmdList.Add(0x23);
            cmdList.Add(0x00);
            cmdList.AddRange(startFreqBytes);

            // Step Frequency
            cmdList.Add(0x23);
            cmdList.Add(0x10);
            cmdList.AddRange(stepFreqBytes);

            // Sweep of Point
            cmdList.Add(0x21);
            cmdList.Add(0x20);
            cmdList.AddRange(numberOfPointBytes.Take(2)); // Chỉ lấy 2 byte đầu tiên vì số điểm quét là kiểu int (nhưng chỉ cần 2 byte)

            // Values per Frequency
            cmdList.Add(0x21);
            cmdList.Add(0x22);
            cmdList.Add(valuesPerFreqBytes[0]);
            cmdList.Add(0x00);

            // Chuyển đổi List<byte> thành mảng byte[]
            byte[] cmdBytes = cmdList.ToArray();

            serialPort1.Write(cmdBytes, 0, cmdBytes.Length);

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
            try
            {
                int startFreq = int.Parse(txtStartFreq.Text);
                int stopFreq = int.Parse(txtStopFreq.Text);
                txtSpanFreq.Text = (stopFreq - startFreq).ToString();
            }
            catch (FormatException ex)
            {

            }
        }

        private void txtStartFreq_TextChanged(object sender, EventArgs e)
        {
            //txtSpanFreq.Text = (int.Parse(txtStopFreq.Text.ToString()) - int.Parse(txtStartFreq.Text.ToString())).ToString();

            try
            {
                int startFreq = int.Parse(txtStartFreq.Text);
                int stopFreq = int.Parse(txtStopFreq.Text);
                txtSpanFreq.Text = (stopFreq - startFreq).ToString();
            }
            catch (FormatException ex)
            {
               
            }
        }

        private void DisabledForm()
        {
           
            groupBox5.Enabled = false;
            groupBox8.Enabled = false;
            btnExport.Enabled   = false;
            groupBox2.Enabled = false;
            groupBox7.Enabled = false;

        }
        private void EnabledForm()
        {
           
            groupBox5.Enabled = true;

            groupBox8.Enabled = true;
            btnExport.Enabled = true;
            groupBox2.Enabled = true;
            groupBox7.Enabled = true;

        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            List<double> freqIndexList = new List<double>();
            List<double> parameterListS11 = new List<double>();
            List<double> parameterListS21 = new List<double>();
            int lineNumber = 0;

            foreach (string line in Listdata)
            {
                Dictionary<string, double> decodedValues = DecodeFrame(line);
                double parameter = CalculateS11(decodedValues["fwd0Re"], decodedValues["fwd0Im"], decodedValues["rev0Re"], decodedValues["rev0Im"], decodedValues["freqIndex"]);
                double parameterS21 = CalculateS21(decodedValues["rev0Re"], decodedValues["rev0Im"], decodedValues["rev1Re"], decodedValues["rev1Im"], decodedValues["freqIndex"]);
                freqIndexList.Add(lineNumber);
                parameterListS11.Add(parameter);
                parameterListS21.Add(parameterS21);
                lineNumber++;
            }
            ExportToExcel(freqIndexList, parameterListS11, filePathS11, 10,60);
            ExportToExcel(freqIndexList, parameterListS21, filePathS21, 10, 60);
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
            if (DoneScan ==false) //ctnPoint < int.Parse(cbNumberPoint.SelectedItem.ToString())
            {
                
                //SendParameter();
                SendScanData();
                richTextBox1.AppendText("Scan: " + ctnScan + "\t");
                ScrollToBottom(richTextBox1);
                ctnScan++;

            }
            //else
            //{

            //    DoneScan = true;
            //    richTextBox1.AppendText(ctnPoint + " points received\n");
            //    ScrollToBottom(richTextBox1);
            //    ctnPoint = 0;
            //    ctnScan = 0;
            //    // Clear buffer
            //    serialPort1.Write(new byte[] { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 }, 0, 8);
            //    serialPort1.Write(new byte[] { 0x20, 0x26, 0x02 }, 0, 3);

            //    timerScan.Enabled = false;


            //}

        }
       
        private void timer1_Tick(object sender, EventArgs e)
        {
            
            if (DoneScan == true)
            {
                List<double> freqIndexListS11 = new List<double>();
                List<double> freqIndexListS21 = new List<double>();
                List<double> parameterListS11 = new List<double>();
                List<double> parameterListS21 = new List<double>();
                richTextBox1.AppendText("Auto " + (ctnAutoExcel-1) + " Done\n");

                int lineNumber = 0;
                foreach (string line in Listdata)
                {
                    Dictionary<string, double> decodedValues = DecodeFrame(line);
                    double parameterS11 = CalculateS11(decodedValues["fwd0Re"], decodedValues["fwd0Im"], decodedValues["rev0Re"], decodedValues["rev0Im"], decodedValues["freqIndex"]);
                    freqIndexListS11.Add(decodedValues["freqIndex"]);
                    parameterListS11.Add(parameterS11);

                }

                //for (int i = 0; i <= 100; i++)
                //{
                //    Console.Write(freqIndexListS11[i] + ":\t"); //+ parameterListS11[i] + "\t"

                //}
                //Console.Write("\n");
                foreach (string line in Listdata)
                {

                    Dictionary<string, double> decodedValues = DecodeFrame(line);
                    double parameterS21 = CalculateS21(decodedValues["rev0Re"], decodedValues["rev0Im"], decodedValues["rev1Re"], decodedValues["rev1Im"], decodedValues["freqIndex"]);
                    freqIndexListS21.Add(decodedValues["freqIndex"]);
                    parameterListS21.Add(parameterS21);
                }
                // error this is

                //(freqIndexListS11, parameterListS11) = SortDataBasedOnZero(freqIndexListS11, parameterListS11);


                foreach (string line in Listdata)
                {
                    Dictionary<string, double> decodedValues = DecodeFrame(line);
                    double parameterS11 = CalculateS11(decodedValues["fwd0Re"], decodedValues["fwd0Im"], decodedValues["rev0Re"], decodedValues["rev0Im"], decodedValues["freqIndex"]);
                    //reqIndexList.Add(lineNumber);
                    //Console.Write(ctnS11 + "\t");

                    freqIndexListS11[(int)decodedValues["freqIndex"]] = decodedValues["freqIndex"];
                    parameterListS11[(int)decodedValues["freqIndex"]] = parameterS11;

                }

                //for (int i = 0; i <= 100; i++)
                //{
                //    Console.Write(freqIndexListS11[i] + ":\t");

                //}
                //(freqIndexListS21, parameterListS21) = SortDataBasedOnZero(freqIndexListS21, parameterListS21);



                foreach (string line in Listdata)
                {
                    Dictionary<string, double> decodedValues = DecodeFrame(line);
                    double parameterS21 = CalculateS21(decodedValues["rev0Re"], decodedValues["rev0Im"], decodedValues["rev1Re"], decodedValues["rev1Im"], decodedValues["freqIndex"]);
                    //Console.WriteLine(ctnS21);
                    freqIndexListS21[(int)decodedValues["freqIndex"]] = decodedValues["freqIndex"];
                    parameterListS21[(int)decodedValues["freqIndex"]] = parameterS21;

                }

                ExportToExcel(freqIndexListS11, parameterListS11, filePathS11, ctnAutoExcel, (ctnAutoExcel -1)* int.Parse(txtStepValue.Text));
                ExportToExcel(freqIndexListS21, parameterListS21, filePathS21, ctnAutoExcel, (ctnAutoExcel - 1) * int.Parse(txtStepValue.Text));


                if ((ctnAutoExcel-2) < (ctnAutoScan-1))
                {
                    modeScan = 1;
                    StartMyTimer();
                    Listdata.Clear();
                    ctnAutoExcel++;
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
                    timerAutoMode.Enabled = false;
                    timerScan.Enabled = false;

                }
                else
                {
                    
                    ctnAutoScan = 0;
                    ctnAutoExcel = 2;
                    AutoScan = false;
                    richTextBox1.AppendText("Auto Mode Done \n");
                    System.Media.SystemSounds.Exclamation.Play();
                    Thread.Sleep(500);
                    timerAutoMode.Enabled = false;
                    ctnScan = 0;
                    serialPort1.Write(new byte[] { 0x20, 0x26, 0x02 }, 0, 3);
                }
                labelProgress.Text = (ctnAutoExcel - 1).ToString() + "/" + ctnAutoScan.ToString();
            }
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
           
            foreach (var series in chart2.Series)
            {
                series.Points.Clear();
            }
            foreach (var series in chart4.Series)
            {
                series.Points.Clear();
            }

            PlotPolarS11(int.Parse(txtFreqIndex.Text.ToString()));
            PlotPolarS21(int.Parse(txtFreqIndex.Text.ToString()));
        }

        private void txtInitAngle_TextChanged(object sender, EventArgs e)
        {

        }

        private void send_Click(object sender, EventArgs e)
        {
            if (serialPort2.IsOpen)
            {
                serialPort2.Write("P");
                serialPort2.Write(Setangle_txt.Text.PadLeft(8, '0'));
                serialPort2.Write(Kps_txtbox.Text.PadLeft(8, '0'));
                serialPort2.Write(Kis_txtbox.Text.PadLeft(8, '0'));
                serialPort2.Write(Kds_txtbox.Text.PadLeft(8, '0'));
                set_angle = int.Parse(Setangle_txt.Text);
            }
            else
            {
                MessageBox.Show($"serialPort not available");
            }
        }

        private void btnPlotChartFromDatabase_Click(object sender, EventArgs e)
        {
            foreach (var series in chart1.Series)
            {
                series.Points.Clear();
            }
            PlotChartS11FromExcel(int.Parse(txtAngleSelect.Text.ToString()));
            PlotChartS21FromExcel(int.Parse(txtAngleSelect.Text.ToString()));
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
            else
            {
                MessageBox.Show("Can not find COM port.");
            }
        }

        private void scanVNA_Click(object sender, EventArgs e)
        {
            if ( serialPort1.IsOpen)
            {
                // Write VNA
                timerScan.Enabled = true;
                DoneScan = false;
                Listdata.Clear();
                SendParameter();
                richTextBox1.AppendText("requesting " + int.Parse(cbNumberPoint.SelectedItem.ToString()) + " points .." + "\n");
                ScrollToBottom(richTextBox1);
                SendScanData();
            }
            else
            {
                MessageBox.Show($"serialPort not available");
            }
        }

        private void button2_Click_2(object sender, EventArgs e)
        {
            ctnAutoExcel++;
            timerAutoMode.Enabled = true;
            timerScan.Enabled = true;
            DoneScan = false;
            AutoScan = true;
        }

        private void timerValid_Tick(object sender, EventArgs e)
        {
            // Thực hiện công việc khi Trigger_Scan được đặt thành 1
            if (Trigger_Scan == 1)
            {
                richTextBox1.AppendText("trigger: " + Trigger_Scan.ToString() + "\n");

                if (modeScan == 0)
                {
                    // Write VNA
                    timerScan.Enabled = true;
                    DoneScan = false;
                    Listdata.Clear();
                    SendParameter();
                    richTextBox1.AppendText("requesting " + int.Parse(cbNumberPoint.SelectedItem.ToString()) + " points .." + "\n");
                    ScrollToBottom(richTextBox1);
                    SendScanData();
                    Trigger_Scan = 0;

                }
                else if (modeScan == 1)
                {
                    if ((ctnAutoExcel - 2) < ctnAutoScan)
                    {
                        SendParameter();
                        SendScanData();
                        timerAutoMode.Enabled = true;
                        timerScan.Enabled = true;
                        DoneScan = false;
                        AutoScan = true;
                        Trigger_Scan = 0;
                    }
                }
                StopMyTimer();
            }
        }

        private void UpdateCOM_Click(object sender, EventArgs e)
        {
            InitializeSerialPort1();
            InitializeSerialPort2();
            System.Media.SystemSounds.Exclamation.Play();

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void requesting()
        {
            serialPort1.Write(new byte[] { 0x10, 0xf2 }, 0, 2);
            serialPort1.Write(new byte[] { 0x10, 0xf3 }, 0, 2);
            serialPort1.Write(new byte[] { 0x10, 0xf4 }, 0, 2);
            serialPort1.Write(new byte[] { 0x11, 0x5c }, 0, 2);
            richTextBox1.AppendText("tx: polling - requesting hw-revision and fw-version\n");
            richTextBox1.AppendText("tx: 10 f2 10 f3 10 f4 11 5c\n");
            ScrollToBottom(richTextBox1);
        }
        private void Setpoint_txt_TextChanged(object sender, EventArgs e)
        {

        }

        public static void ExportToExcel(List<double> freqIndexList, List<double> parameterList, string filePath, int parameterColumn,int DataColumn)
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

                worksheet.Cells[1, parameterColumn] = (DataColumn.ToString());
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

        private void button2_Click_3(object sender, EventArgs e)
        {
            folderBrowserDialog1.ShowDialog();
            txtFolderS11.Text = folderBrowserDialog1.SelectedPath;
            filePathS11 = @txtFolderS11.Text + @"\DATA_S11.xlsx";
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.ShowDialog();
            txtFolderS21.Text = folderBrowserDialog1.SelectedPath;
            filePathS21 = @txtFolderS21.Text+@"\DATA_S21.xlsx";
        }

        private void button5_Click(object sender, EventArgs e)
        {
            richTextBox2.Clear();
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
                { "freqIndex", freqIndex },
                { "freq", freq }
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

        private double CalculateS11(double fwd0Re, double fwd0Im, double rev0Re, double rev0Im, double freqIndex)
        {
            Complex A1 = new Complex(fwd0Re, fwd0Im);
            Complex A2 = new Complex(rev0Re, rev0Im);
            Complex ratio = A2 / A1;
            return 20 * Math.Log10(ratio.Magnitude);
        }
        private double CalculateS21(double rev0Re, double rev0Im, double rev1Re, double rev1Im, double freqIndex)
        {
            Complex A1 = new Complex(rev1Re, rev1Im);
            Complex A2 = new Complex(rev0Re, rev0Im);
            Complex ratio = A1 / A2;
            return 20 * Math.Log10(ratio.Magnitude);
        }

        private void PlotChartS11FromExcel(int AngleIndex)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            List<string> FreqList = new List<string>();
            List<string> parameterListS11 = new List<string>();


            try
            {
                using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(filePathS11)))
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
                        string cellValue = worksheet.Cells[ row, (AngleIndex / int.Parse(txtStepValue.Text.ToString())) + 1].Value?.ToString();
                        string cellFreqValue = worksheet.Cells[row, 1].Value?.ToString();

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
        private void PlotChartS21FromExcel(int AngleIndex)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            List<string> FreqList = new List<string>();
            List<string> parameterListS21 = new List<string>();


            try
            {
                using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(filePathS21)))
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
                        string cellValue = worksheet.Cells[row, (AngleIndex/ int.Parse(txtStepValue.Text.ToString())) + 1].Value?.ToString();
                        string cellFreqValue = worksheet.Cells[row, 1].Value?.ToString();

                        if (!string.IsNullOrEmpty(cellValue) && !string.IsNullOrEmpty(cellFreqValue))
                        {
                            if (double.TryParse(cellValue, out double parameter) && int.TryParse(cellFreqValue, out int freq))
                            {
                                parameterListS21.Add(cellValue);
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
                    chart1.Series["S21"].Points.AddXY(int.Parse(FreqList[i]), double.Parse(parameterListS21[i]));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}");
            }

        }
        private void PlotData()
            {
            int index0OfS11=0;
            List<double> freqIndexListS11 = new List<double>();
            List<double> freqIndexListS21 = new List<double>();
            List<double> parameterListS11 = new List<double>();
            List<double> parameterListS21 = new List<double>();

            foreach (string line in Listdata)
            {
                Dictionary<string, double> decodedValues = DecodeFrame(line);
                double parameterS11 = CalculateS11(decodedValues["fwd0Re"], decodedValues["fwd0Im"], decodedValues["rev0Re"], decodedValues["rev0Im"], decodedValues["freqIndex"]);
                freqIndexListS11.Add(decodedValues["freqIndex"]);
                parameterListS11.Add(parameterS11);
               
            }


            
            foreach (string line in Listdata)
            {

                Dictionary<string, double> decodedValues = DecodeFrame(line);
                double parameterS21 = CalculateS21(decodedValues["rev0Re"], decodedValues["rev0Im"], decodedValues["rev1Re"], decodedValues["rev1Im"], decodedValues["freqIndex"]);
                freqIndexListS21.Add(decodedValues["freqIndex"]);
                parameterListS21.Add(parameterS21);
            }

            //for (int i = 0; i < freqIndexListS11.Count; i++)
            //{
            //    Console.Write(freqIndexListS11[i] + ":\t");

            //}
            //Console.Write("\n");
            //// error this is

            //(freqIndexListS11, parameterListS11) = SortDataBasedOnZero(freqIndexListS11, parameterListS11);

            Console.Write("\n");
            foreach (string line in Listdata)
            {
                Dictionary<string, double> decodedValues = DecodeFrame(line);
                double parameterS11 = CalculateS11(decodedValues["fwd0Re"], decodedValues["fwd0Im"], decodedValues["rev0Re"], decodedValues["rev0Im"], decodedValues["freqIndex"]);
                //reqIndexList.Add(lineNumber);
                //Console.Write(ctnS11 + "\t");

                freqIndexListS11[(int)decodedValues["freqIndex"]] = decodedValues["freqIndex"];
                parameterListS11[(int)decodedValues["freqIndex"]] = parameterS11;
                Console.Write(parameterS11 + "\t");
            }


            //(freqIndexListS21, parameterListS21) = SortDataBasedOnZero(freqIndexListS21, parameterListS21);



            Console.Write("\n");
            foreach (string line in Listdata)
            {
                Dictionary<string, double> decodedValues = DecodeFrame(line);
                double parameterS21 = CalculateS21(decodedValues["rev0Re"], decodedValues["rev0Im"], decodedValues["rev1Re"], decodedValues["rev1Im"], decodedValues["freqIndex"]);
                
                freqIndexListS21[(int)decodedValues["freqIndex"]] = decodedValues["freqIndex"];
                parameterListS21[(int)decodedValues["freqIndex"]] = parameterS21;
                Console.Write(parameterS21+"\t");

            }



            for (int i = 0; i < freqIndexListS11.Count; i++)
            {

                chart1.Series["S11"].Points.AddXY(freqIndexListS11[i], parameterListS11[i]);
            }
            Console.Write("\n");
            for (int i = 0; i < freqIndexListS21.Count; i++)
            {
              
                chart1.Series["S21"].Points.AddXY(freqIndexListS21[i], parameterListS21[i]);

            }


        }
        static (List<double>, List<double>) SortDataBasedOnZero(List<double> data1, List<double> data2)
        {
            int zeroIndex = data1.IndexOf(0);

            if (zeroIndex == -1)
            {
                return (data1,data2);
            }

            List<double> result1 = new List<double>();
            List<double> result2 = new List<double>();
            for (int i = zeroIndex; i < data1.Count; i++)
            {
                result1.Add(data1[i]);
                result2.Add(data2[i]);
            }
            for (int i = 0; i < zeroIndex; i++)
            {
                result1.Add(data1[i]);
                result2.Add(data2[i]);
            }

            return (result1, result2);
        }
        private void PlotPolarS11(int FreqIndex)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            List<string> AngleList = new List<string>();
            List<string> parameterListS11 = new List<string>();

            try
            {
                using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(filePathS11)))
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

        private void PlotPolarS21(int FreqIndex)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            List<string> AngleList = new List<string>();
            List<string> parameterListS11 = new List<string>();

            try
            {
                using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(filePathS21)))
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
                    chart4.Series["Series1"].Points.AddXY(int.Parse(AngleList[i]), double.Parse(parameterListS11[i]));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}");
            }
        }

    }
}
