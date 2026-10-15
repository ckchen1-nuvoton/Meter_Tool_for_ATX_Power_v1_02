using Microsoft.Win32;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Resources;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using static System.Net.Mime.MediaTypeNames;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Excel = Microsoft.Office.Interop.Excel;
using Ivi.Visa;
using NationalInstruments.Visa;
using static System.Collections.Specialized.BitVector32;
using static OfficeOpenXml.ExcelErrorValue;
using System.Text.RegularExpressions;



namespace NADC_COM_Tool_v1_00
{
    public partial class Form1 : Form
    {
        private bool isOpen = false;
        private List<byte> sendBuffer = new List<byte>();
        private int sendCount = 0;
        //private bool isReport = false;
        static bool isReport = false;
        private bool isCalcVTHD = false;
        private bool isCalcITHD = false;

        private Thread thread1, thread2, thread3;
        private bool isThread1Running = false;
        private bool isThread2Running = false;
        private bool isThread3Running = false;

        private System.Windows.Forms.Timer timer;
        //private Random random;
        private int xValue1, xValue2, xValue3;
        private Series series1, series2, series3;

        public static int freq_int = 0;
        public static float vrms_float = 0;
        public static float irms_float = 0;
        public static float power_float = 0;
        public static float pf_value_float = 0;
        public static float v_thd_float = 0;
        public static float i_thd_float = 0;
        public static float vbulk_float = 0;

        List<int>   freq_curve_data = new List<int>();   // 创建一个空的列表用于存储数据
        List<float> vrms_curve_data = new List<float>(); // 创建一个空的列表用于存储数据
        List<float> irms_curve_data = new List<float>(); // 创建一个空的列表用于存储数据
        List<float> powe_curve_data = new List<float>(); // 创建一个空的列表用于存储数据
        List<float> pfva_curve_data = new List<float>(); // 创建一个空的列表用于存储数据
        List<float> vthd_curve_data = new List<float>(); // 创建一个空的列表用于存储数据
        List<float> ithd_curve_data = new List<float>(); // 创建一个空的列表用于存储数据

        // Nuvoton
        private string PowerMeter66202_Address = "USB0::0x0A69::0x0835::A66200110413::INSTR"; // CHROMA 66202 的 USB 地址
        private string DC_Load63202A_Address   = "USB0::0x0A69::0x0879::63202AL01017::INSTR"; // CHROMA 63202A 的 USB 地址
        
        // AcBel
        //private string DC_Load63202A_Address = "USB0::0x0A69::0x087C::63206EL01029::INSTR";
        //private string PowerMeter66202_Address = "USB0::0x0A69::0x087D::662050008757::INSTR"; // CHROMA 66202 的 USB 地址
        private List<byte> receiveBuffer = new List<byte>();  // 用于累积接收到的数据
        private string targetTextBox = "";  // 用于存储被点击的按钮的信息
        private int expectedLength = 19;  // 默认期望接收19字节，用于auto-report
        private int indexer = 0;

        public Form1()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void DataReceivedHandler(object sender, SerialDataReceivedEventArgs e)
        {
            SerialPort sp = (SerialPort)sender;
            int bytesToRead = sp.BytesToRead;
            byte[] buffer = new byte[bytesToRead];
            sp.Read(buffer, 0, bytesToRead);

            // 将接收到的数据加入全局缓冲区
            receiveBuffer.AddRange(buffer);

            // 检查是否有足够的数据可以处理
            ProcessBuffer();
        }

        private void ProcessBuffer()
        {
            // 处理19字节的auto-report数据
            while (receiveBuffer.Count >= expectedLength)
            {
                if (expectedLength == 19)
                {
                    // 处理19字节的数据
                    byte[] autoReportData = receiveBuffer.GetRange(0, 19).ToArray();
                    receiveBuffer.RemoveRange(0, 19);   // 移除已处理的数据
                    
                    ProcessAutoReport(autoReportData);
                }
                else if (expectedLength == 15)
                {
                    // 处理19字节的数据
                    byte[] autoReportData = receiveBuffer.GetRange(0, 15).ToArray();
                    //receiveBuffer.RemoveRange(0, 15);   // 移除已处理的数据
                    receiveBuffer.Clear();  // 清除整个缓冲区，等待下一个有效数据帧

                    // 根据 targetTextBox 更新对应的控件
                    ProcessCaliData(autoReportData);

                    // 恢复19字节接收模式
                    //expectedLength = 19;

                    // 發送開始命令
                    //byte[] myBytes14 = { 0x49, 0x02, 0x81, 0x7F, 0xB5, 0x0D };
                    //serialPort1.Write(myBytes14, 0, myBytes14.Length); // 將 6 個 Byte 數據寫出去
                    //Thread.Sleep(100);  // Delay 100ms
                }
                else if (expectedLength == 4)
                {
                    // 处理4字节的数据
                    byte[] commandResponse = receiveBuffer.GetRange(0, 4).ToArray();
                    receiveBuffer.RemoveRange(0, 4);   // 移除已处理的数据

                    // 根据 targetTextBox 更新对应的控件
                    ProcessDataForTextBox(commandResponse, targetTextBox);

                    // 恢复19字节接收模式
                    expectedLength = 19;

                    // 發送開始命令
                    byte[] myBytes14 = { 0x49, 0x02, 0x81, 0x7F, 0xB5, 0x0D };
                    serialPort1.Write(myBytes14, 0, myBytes14.Length); // 將 6 個 Byte 數據寫出去
                    Thread.Sleep(100);  // Delay 100ms
                }
                else if (expectedLength == 7)
                {
                    // 处理4字节的数据
                    byte[] commandResponse = receiveBuffer.GetRange(0, 7).ToArray();
                    receiveBuffer.RemoveRange(0, 7);   // 移除已处理的数据

                    // 根据 targetTextBox 更新对应的控件
                    ProcessDataForTextBox(commandResponse, targetTextBox);

                    // 恢复19字节接收模式
                    expectedLength = 19;

                    // 發送開始命令
                    byte[] myBytes14 = { 0x49, 0x02, 0x81, 0x7F, 0xB5, 0x0D };
                    serialPort1.Write(myBytes14, 0, myBytes14.Length); // 將 6 個 Byte 數據寫出去
                    Thread.Sleep(100);  // Delay 100ms
                }
            }
        }

        private void ProcessCaliData(byte[] buffer)
        {
            Console.WriteLine("-------------------------------------------------------------------------");
            // Debug: 将读取到的数据转换为字符串
            string result = BitConverter.ToString(buffer, 0, buffer.Length).Replace("-", " ");
            Console.WriteLine(result);
            Console.WriteLine("-------------------------------------------------------------------------");

            // 在這裡處理讀取到的資料
            if ((buffer[0] == 0x49) && (buffer[1] == 0x0C))
            {
                uint Vrms_int32 = (uint)(buffer[2] << 24 | buffer[3] << 16 | buffer[4] << 8 | buffer[5]);
                float Vrms_float = (float)(Vrms_int32 / 1000.0);
                Console.WriteLine("Vrms : " + Vrms_float);
                //p3_vrms_box.Text = Vrms_float.ToString();

                uint Irms_int32 = (uint)(buffer[6] << 24 | buffer[7] << 16 | buffer[8] << 8 | buffer[9]);
                float Irms_float = (float)(Irms_int32 / 128000.0);
                Console.WriteLine("Irms : " + Irms_float);
                //p3_irms_box.Text = Irms_float.ToString();

                uint Pavg_int32 = (uint)(buffer[10] << 24 | buffer[11] << 16 | buffer[12] << 8 | buffer[13]);
                float Pavg_float = (float)(Pavg_int32 / 1000.0);
                Console.WriteLine("Pavg : " + Pavg_float);
                //p3_pavg_box.Text = Pavg_float.ToString();


                // 根据 targetTextBox 判断显示在哪个 TextBox
                if (targetTextBox == "p3_vrms_box")
                {
                    // 更新 UI
                    Action action = () =>
                    {
                        /*
                        p3_vrms_box.Clear();
                        p3_irms_box.Clear();
                        p3_pavg_box.Clear();
                        Thread.Sleep(100);  // Delay 100ms
                        p3_vrms_box.Text = Vrms_float.ToString();
                        p3_irms_box.Text = Irms_float.ToString();
                        p3_pavg_box.Text = Pavg_float.ToString();
                        */
                    };
                    this.Invoke(action);
                }
                else if (targetTextBox == "p1_vrms_box")
                {
                    // 更新 UI
                    Action action = () =>
                    {
                        /*
                        p1_vrms_box.Clear();
                        p1_irms_box.Clear();
                        p1_pavg_box.Clear();
                        Thread.Sleep(100);  // Delay 100ms
                        p1_vrms_box.Text = Vrms_float.ToString();
                        p1_irms_box.Text = Irms_float.ToString();
                        p1_pavg_box.Text = Pavg_float.ToString();
                        */
                    };
                    this.Invoke(action);
                }
                else if (targetTextBox == "p2_vrms_box")
                {
                    // 更新 UI
                    Action action = () =>
                    {
                        /*
                        p2_vrms_box.Clear();
                        p2_irms_box.Clear();
                        p2_pavg_box.Clear();
                        Thread.Sleep(100);  // Delay 100ms
                        p2_vrms_box.Text = Vrms_float.ToString();
                        p2_irms_box.Text = Irms_float.ToString();
                        p2_pavg_box.Text = Pavg_float.ToString();
                        */
                    };
                    this.Invoke(action);
                }
                else if (targetTextBox == "p4_vrms_box")
                {
                    // 更新 UI
                    Action action = () =>
                    {
                        /*
                        p4_vrms_box.Clear();
                        p4_irms_box.Clear();
                        p4_pavg_box.Clear();
                        Thread.Sleep(100);  // Delay 100ms
                        p4_vrms_box.Text = Vrms_float.ToString();
                        p4_irms_box.Text = Irms_float.ToString();
                        p4_pavg_box.Text = Pavg_float.ToString();
                        */
                    };
                    this.Invoke(action);
                }

                // 清除串口接收缓冲区中的数据
                serialPort1.DiscardInBuffer();
                receiveBuffer.Clear();  // 清除整个缓冲区，等待下一个有效数据帧
            }
        }

        // 处理 auto-report 数据
        private void ProcessAutoReport(byte[] buffer)
        {
            // Debug: 将读取到的数据转换为字符串
            string result = BitConverter.ToString(buffer, 0, buffer.Length).Replace("-", " ");
            Console.WriteLine(result);

            // 检查第一个字节为 0x49，第二个字节为 0x10
            if (buffer[0] == 0x49 && buffer[1] == 0x10)
            {
                // 处理 auto-report 数据
                int auto_report_vrms = buffer[4] << 16 | buffer[3] << 8 | buffer[2];
                vrms_float = (float)(auto_report_vrms / 1000.0);
                Console.WriteLine("auto_report_vrms : " + vrms_float);

                int auto_report_irms = buffer[7] << 16 | buffer[6] << 8 | buffer[5];
                irms_float = (float)(auto_report_irms / 128000.0);
                Console.WriteLine("auto_report_irms : " + irms_float);

                int auto_report_power = buffer[10] << 16 | buffer[9] << 8 | buffer[8];
                power_float = (float)(auto_report_power / 1000.0);
                Console.WriteLine("auto_report_power : " + power_float);

                int auto_report_frequency = buffer[13] << 16 | buffer[12] << 8 | buffer[11];
                freq_int = auto_report_frequency / 1000;
                Console.WriteLine("auto_report_frequency : " + freq_int);

                int auto_report_vbulk = buffer[16] << 16 | buffer[15] << 8 | buffer[14];
                vbulk_float = (float)(auto_report_vbulk / 1000.0);
                Console.WriteLine("auto_report_vbulk : " + vbulk_float);

                // 更新 UI
                Action action = () =>
                {
                    freq_box.Text = freq_int.ToString();

                    vrms_box.ForeColor = Color.FromArgb(255, 40, 53, 118);
                    vrms_box.Text = vrms_float.ToString("N2");

                    irms_box.ForeColor = Color.FromArgb(255, 2, 97, 97);
                    irms_box.Text = irms_float.ToString("N5");

                    watt_box.ForeColor = Color.FromArgb(255, 115, 73, 48);
                    watt_box.Text = power_float.ToString("N3");

                    vbulk_box.ForeColor = Color.FromArgb(255, 115, 73, 48);
                    vbulk_box.Text = vbulk_float.ToString("N2");
                };
                this.Invoke(action);
            }
        }

        private void ProcessDataForTextBox(byte[] data, string targetTextBox)
        {
            // Debug: 将读取到的数据转换为字符串并输出到控制台
            string result = BitConverter.ToString(data, 0, data.Length).Replace("-", " ");
            Console.WriteLine("Received Data for " + targetTextBox + ": " + result);

            // 根据 targetTextBox 判断显示在哪个 TextBox
            if (targetTextBox == "chip_ssid_box")
            {
                if (chip_ssid_box.InvokeRequired)
                {
                    chip_ssid_box.Invoke(new Action(() =>
                    {
                        // 检查第一个字节为 0x49，第二个字节为 0x01
                        if (data[0] == 0x49 && data[1] == 0x01)
                        {
                            byte thirdByte = data[2];
                            Console.WriteLine("Value : " + thirdByte);
                            string dataString = thirdByte.ToString("D2");
                            chip_ssid_box.AppendText(dataString);
                        }
                    }));
                }
            }
            else if (targetTextBox == "fw_ver_box")
            {
                if (fw_ver_box.InvokeRequired)
                {
                    fw_ver_box.Invoke(new Action(() =>
                    {
                        // 检查第一个字节为 0x49，第二个字节为 0x01
                        if (data[0] == 0x49 && data[1] == 0x01)
                        {
                            byte thirdByte = data[2];
                            Console.WriteLine("Value : " + thirdByte);
                            string dataString = thirdByte.ToString("D2");
                            fw_ver_box.AppendText(dataString);
                        }
                    }));
                }
            }
            else if (targetTextBox == "status_box")
            {
                if (fw_ver_box.InvokeRequired)
                {
                    fw_ver_box.Invoke(new Action(() =>
                    {
                        // 检查第一个字节为 0x49，第二个字节为 0x01
                        if (data[0] == 0x49 && data[1] == 0x01)
                        {
                            byte thirdByte = data[2];
                            Console.WriteLine("Value : " + thirdByte);
                            string dataString = thirdByte.ToString("D2");
                            //string dataString = "0x" + thirdByte.ToString("X2"); // "X2" 格式化為兩位十六進位，並加上 "0x" 前綴
                            status_box.AppendText(dataString);
                        }
                    }));
                }
            }
            else if (targetTextBox == "p1_vrms_box")
            {
                if (p1_vrms_box.InvokeRequired)
                {
                    p1_vrms_box.Invoke(new Action(() =>
                    {
                        // 检查第一个字节为 0x49，第二个字节为 0x04
                        if (data[0] == 0x49 && data[1] == 0x04)
                        {
                            uint Vrms_int32 = (uint)(data[2] << 24 | data[3] << 16 | data[4] << 8 | data[5]);
                            float Vrms_float = (float)(Vrms_int32 / 1000.0);
                            Console.WriteLine("Value : " + Vrms_float);
                            p1_vrms_box.Text = Vrms_float.ToString();
                        }
                    }));
                }
            }
            else if (targetTextBox == "p1_irms_box")
            {
                if (p1_irms_box.InvokeRequired)
                {
                    p1_irms_box.Invoke(new Action(() =>
                    {
                        // 检查第一个字节为 0x49，第二个字节为 0x04
                        if (data[0] == 0x49 && data[1] == 0x04)
                        {
                            uint Irms_int32 = (uint)(data[2] << 24 | data[3] << 16 | data[4] << 8 | data[5]);
                            float Irms_float = (float)(Irms_int32 / 128000.0);
                            Console.WriteLine("Value : " + Irms_float);
                            p1_irms_box.Text = Irms_float.ToString();
                        }
                    }));
                }
            }
            else if (targetTextBox == "p1_pavg_box")
            {
                if (p1_pavg_box.InvokeRequired)
                {
                    p1_pavg_box.Invoke(new Action(() =>
                    {
                        // 检查第一个字节为 0x49，第二个字节为 0x04
                        if (data[0] == 0x49 && data[1] == 0x04)
                        {
                            uint Pavg_int32 = (uint)(data[2] << 24 | data[3] << 16 | data[4] << 8 | data[5]);
                            float Pavg_float = (float)(Pavg_int32 / 1000.0);
                            Console.WriteLine("Value : " + Pavg_float);
                            p1_pavg_box.Text = Pavg_float.ToString();
                        }
                    }));
                }
            }
            else if (targetTextBox == "p2_vrms_box")
            {
                if (p2_vrms_box.InvokeRequired)
                {
                    p2_vrms_box.Invoke(new Action(() =>
                    {
                        // 检查第一个字节为 0x49，第二个字节为 0x04
                        if (data[0] == 0x49 && data[1] == 0x04)
                        {
                            uint Vrms_int32 = (uint)(data[2] << 24 | data[3] << 16 | data[4] << 8 | data[5]);
                            float Vrms_float = (float)(Vrms_int32 / 1000.0);
                            Console.WriteLine("Value : " + Vrms_float);
                            p2_vrms_box.Text = Vrms_float.ToString();
                        }
                    }));
                }
            }
            else if (targetTextBox == "p2_irms_box")
            {
                if (p2_irms_box.InvokeRequired)
                {
                    p2_irms_box.Invoke(new Action(() =>
                    {
                        // 检查第一个字节为 0x49，第二个字节为 0x04
                        if (data[0] == 0x49 && data[1] == 0x04)
                        {
                            uint Irms_int32 = (uint)(data[2] << 24 | data[3] << 16 | data[4] << 8 | data[5]);
                            float Irms_float = (float)(Irms_int32 / 128000.0);
                            Console.WriteLine("Value : " + Irms_float);
                            p2_irms_box.Text = Irms_float.ToString();
                        }
                    }));
                }
            }
            else if (targetTextBox == "p2_pavg_box")
            {
                if (p2_pavg_box.InvokeRequired)
                {
                    p2_pavg_box.Invoke(new Action(() =>
                    {
                        // 检查第一个字节为 0x49，第二个字节为 0x04
                        if (data[0] == 0x49 && data[1] == 0x04)
                        {
                            uint Pavg_int32 = (uint)(data[2] << 24 | data[3] << 16 | data[4] << 8 | data[5]);
                            float Pavg_float = (float)(Pavg_int32 / 1000.0);
                            Console.WriteLine("Value : " + Pavg_float);
                            p2_pavg_box.Text = Pavg_float.ToString();
                        }
                    }));
                }
            }
            else if (targetTextBox == "p3_vrms_box")
            {
                if (p3_vrms_box.InvokeRequired)
                {
                    p3_vrms_box.Invoke(new Action(() =>
                    {
                        // 检查第一个字节为 0x49，第二个字节为 0x04
                        if (data[0] == 0x49 && data[1] == 0x04)
                        {
                            uint Vrms_int32 = (uint)(data[2] << 24 | data[3] << 16 | data[4] << 8 | data[5]);
                            float Vrms_float = (float)(Vrms_int32 / 1000.0);
                            Console.WriteLine("Value : " + Vrms_float);
                            p3_vrms_box.Text = Vrms_float.ToString();
                        }
                    }));
                }
            }
            else if (targetTextBox == "p3_irms_box")
            {
                if (p3_irms_box.InvokeRequired)
                {
                    p3_irms_box.Invoke(new Action(() =>
                    {
                        // 检查第一个字节为 0x49，第二个字节为 0x04
                        if (data[0] == 0x49 && data[1] == 0x04)
                        {
                            uint Irms_int32 = (uint)(data[2] << 24 | data[3] << 16 | data[4] << 8 | data[5]);
                            float Irms_float = (float)(Irms_int32 / 128000.0);
                            Console.WriteLine("Value : " + Irms_float);
                            p3_irms_box.Text = Irms_float.ToString();
                        }
                    }));
                }
            }
            else if (targetTextBox == "p3_pavg_box")
            {
                if (p3_pavg_box.InvokeRequired)
                {
                    p3_pavg_box.Invoke(new Action(() =>
                    {
                        // 检查第一个字节为 0x49，第二个字节为 0x04
                        if (data[0] == 0x49 && data[1] == 0x04)
                        {
                            uint Pavg_int32 = (uint)(data[2] << 24 | data[3] << 16 | data[4] << 8 | data[5]);
                            float Pavg_float = (float)(Pavg_int32 / 1000.0);
                            Console.WriteLine("Value : " + Pavg_float);
                            p3_pavg_box.Text = Pavg_float.ToString();
                        }
                    }));
                }
            }
            else if (targetTextBox == "p4_vrms_box")
            {
                if (p4_vrms_box.InvokeRequired)
                {
                    p4_vrms_box.Invoke(new Action(() =>
                    {
                        // 检查第一个字节为 0x49，第二个字节为 0x04
                        if (data[0] == 0x49 && data[1] == 0x04)
                        {
                            uint Vrms_int32 = (uint)(data[2] << 24 | data[3] << 16 | data[4] << 8 | data[5]);
                            float Vrms_float = (float)(Vrms_int32 / 1000.0);
                            Console.WriteLine("Value : " + Vrms_float);
                            p4_vrms_box.Text = Vrms_float.ToString();
                        }
                    }));
                }
            }
            else if (targetTextBox == "p4_irms_box")
            {
                if (p4_irms_box.InvokeRequired)
                {
                    p4_irms_box.Invoke(new Action(() =>
                    {
                        // 检查第一个字节为 0x49，第二个字节为 0x04
                        if (data[0] == 0x49 && data[1] == 0x04)
                        {
                            uint Irms_int32 = (uint)(data[2] << 24 | data[3] << 16 | data[4] << 8 | data[5]);
                            float Irms_float = (float)(Irms_int32 / 128000.0);
                            Console.WriteLine("Value : " + Irms_float);
                            p4_irms_box.Text = Irms_float.ToString();
                        }
                    }));
                }
            }
            else if (targetTextBox == "p4_pavg_box")
            {
                if (p4_pavg_box.InvokeRequired)
                {
                    p4_pavg_box.Invoke(new Action(() =>
                    {
                        // 检查第一个字节为 0x49，第二个字节为 0x04
                        if (data[0] == 0x49 && data[1] == 0x04)
                        {
                            uint Pavg_int32 = (uint)(data[2] << 24 | data[3] << 16 | data[4] << 8 | data[5]);
                            float Pavg_float = (float)(Pavg_int32 / 1000.0);
                            Console.WriteLine("Value : " + Pavg_float);
                            p4_pavg_box.Text = Pavg_float.ToString();
                        }
                    }));
                }
            }
        }

        private void LoadComPorts()
        {
            port_cbb.Items.Clear();

            string[] allPorts = SerialPort.GetPortNames();  // e.g., COM1, COM3, COM101

            foreach (string port in allPorts)
            {
                // 使用正則表達式過濾 COM0 ~ COM99
                Match match = Regex.Match(port, @"COM(\d+)");
                if (match.Success)
                {
                    int portNum = int.Parse(match.Groups[1].Value);
                    if (portNum >= 0 && portNum <= 99)
                    {
                        port_cbb.Items.Add(port);
                    }
                }
            }

            if (port_cbb.Items.Count > 0)
                port_cbb.SelectedIndex = 0;  // 預設選第一個
        }

        private void button1_Click(object sender, EventArgs e)
        { 
            try
            {
                if (serialPort1.IsOpen == false)
                {
                    serialPort1.PortName = port_cbb.Text;
                    //serialPort1.BaudRate = Convert.ToInt32(speed_cbb.Text);
                    //serialPort1.DataBits = Convert.ToInt32(data_cbb.Text);
                    serialPort1.BaudRate = 115200;
                    serialPort1.DataBits = 8;
                    serialPort1.Parity = Parity.None;
                    serialPort1.StopBits = StopBits.One;
                    serialPort1.RtsEnable = false;
                    serialPort1.DtrEnable = false;
                    /*
                    switch (parity_cbb.SelectedIndex)
                    {
                        // None, Odd, Even
                        case 0:
                            serialPort1.Parity = Parity.None;
                            break;
                        case 1:
                            serialPort1.Parity = Parity.Odd;
                            break;
                        case 2:
                            serialPort1.Parity = Parity.Even;
                            break;
                        default:
                            serialPort1.Parity = Parity.None;
                            break;
                    }
                    
                    switch (stop_cbb.SelectedIndex)
                    {
                        // 1, 1.5, 2 bit
                        case 0:
                            serialPort1.StopBits = StopBits.One;
                            break;
                        case 1:
                            serialPort1.StopBits = StopBits.OnePointFive;
                            break;
                        case 2:
                            serialPort1.StopBits = StopBits.Two;
                            break;
                        default:
                            serialPort1.StopBits = StopBits.One;
                            break;
                    }
                    */

                    serialPort1.Open();
                    isOpen = true;
                    connection_btn.Text = "Close";
                    StatusLabel2.Text = $"Connected with {serialPort1.PortName} ";
                    Thread.Sleep(100);  // Delay 100ms

                    // 清除串口接收缓冲区中的数据
                    serialPort1.DiscardInBuffer();
                    receiveBuffer.Clear();  // 清除整个缓冲区，等待下一个有效数据帧

                    // 發送停止命令
                    //byte[] myBytes1 = { 0x49, 0x01, 0x71, 0x00, 0xD4, 0x19, 0x0D };
                    //serialPort1.Write(myBytes1, 0, myBytes1.Length); // 將 7 個 Byte 數據寫出去
                    //Thread.Sleep(100);  // Delay 100ms
                    //

                    // 開啟計算線程
                    //thread1 = new Thread(new ThreadStart(DoWork1));
                    //thread1.Start();
                    //isThread1Running = true;

                    // Chart            
                    // 启动 Timer 控件
                    timer.Start();

                    // 绑定 DataReceived 事件处理程序
                    serialPort1.DataReceived += new SerialDataReceivedEventHandler(DataReceivedHandler);
                }
                else
                {
                    serialPort1.Close();
                    isOpen = false;
                    connection_btn.Text = "Open";
                    StatusLabel2.Text = $"Disconnected with {serialPort1.PortName} ";
                    Thread.Sleep(100);  // Delay 100ms

                    // 清除串口接收缓冲区中的数据
                    serialPort1.DiscardInBuffer();
                    receiveBuffer.Clear();  // 清除整个缓冲区，等待下一个有效数据帧

                    // 停止計算線程
                    //thread1.Abort();
                    //isThread1Running = false;

                    // Chart            
                    // 停止 Timer 控件
                    timer.Stop();

                    // 取消绑定 DataReceived 事件处理程序
                    serialPort1.DataReceived -= new SerialDataReceivedEventHandler(DataReceivedHandler);
                }
            }
            catch (Exception ex)
            {
                // 处理异常
                Console.WriteLine("Error : " + ex.Message);
                StatusLabel2.Text = $"Cannot open {serialPort1.PortName} ";
                MessageBox.Show($"Cannot open {serialPort1.PortName} ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }
        
        private void Timer_Tick(object sender, EventArgs e)
        {
            // 儲存最新400筆資料
            //freq_curve_data.Add(freq_int);
            //if (freq_curve_data.Count > 400)
                //freq_curve_data.RemoveAt(0); // 删除最旧的数据（位于列表的首位）
            //powe_curve_data.Add(power_float);
            //if (powe_curve_data.Count > 400)
                //powe_curve_data.RemoveAt(0); // 删除最旧的数据（位于列表的首位）
            //pfva_curve_data.Add(pf_value_float);
            //if (pfva_curve_data.Count > 400)
                //pfva_curve_data.RemoveAt(0); // 删除最旧的数据（位于列表的首位）
            //vthd_curve_data.Add(v_thd_float);
            //if (vthd_curve_data.Count > 400)
                //vthd_curve_data.RemoveAt(0); // 删除最旧的数据（位于列表的首位）
            //ithd_curve_data.Add(i_thd_float);
            //if (ithd_curve_data.Count > 400)
                //ithd_curve_data.RemoveAt(0); // 删除最旧的数据（位于列表的首位）


            // 讀取400筆vrms，並且算出max跟min
            vrms_curve_data.Add(vrms_float);
            if (vrms_curve_data.Count > 400)
                vrms_curve_data.RemoveAt(0); // 删除最旧的数据（位于列表的首位）

            // 找到最大值和最小值            
            float vrms_maxValue = vrms_curve_data.Max();
            float vrms_minValue = vrms_curve_data.Min();
            //Console.WriteLine("vrms_maxValue : " + vrms_maxValue);
            //Console.WriteLine("vrms_minValue : " + vrms_minValue);

            // 假设 e.Data 是一个 float 类型的数据数组
            irms_curve_data.Add(irms_float);
            if (irms_curve_data.Count > 400)        
                irms_curve_data.RemoveAt(0); // 删除最旧的数据（位于列表的首位）

            // 找到最大值和最小值
            float irms_maxValue = irms_curve_data.Max();
            float irms_minValue = irms_curve_data.Min();
            //Console.WriteLine("irms_maxValue : " + irms_maxValue);
            //Console.WriteLine("irms_minValue : " + irms_minValue);

            // 假设 e.Data 是一个 float 类型的数据数组
            powe_curve_data.Add(power_float);
            if (powe_curve_data.Count > 400)
                powe_curve_data.RemoveAt(0); // 删除最旧的数据（位于列表的首位）

            // 找到最大值和最小值
            float powe_maxValue = powe_curve_data.Max();
            float powe_minValue = powe_curve_data.Min();
            //Console.WriteLine("powe_maxValue : " + powe_maxValue);
            //Console.WriteLine("powe_minValue : " + powe_minValue);

            // 将数据添加到曲线中
            series1.Points.AddXY(xValue1, vrms_float);
            //chart1.Series["Series1"].Color = Color.Blue;

            // 设置 X 轴的范围，仅显示最新的100个数据点
            chart1.ChartAreas[0].AxisX.Minimum = xValue1 - 400;
            chart1.ChartAreas[0].AxisX.Maximum = xValue1;
            chart1.ChartAreas[0].AxisY.LabelStyle.Format = "0.00";
            chart1.ChartAreas[0].AxisY.Interval = 0;
            //chart1.ChartAreas[0].AxisY.Title = "Vrms";

            if (vrms_maxValue > 60) 
            {
                chart1.ChartAreas[0].AxisY.Maximum = vrms_maxValue*1.01;
                chart1.ChartAreas[0].AxisY.Minimum = vrms_minValue*0.99;
            }
            //else
            //{
                // 设置Y轴自动调整最小值和最大值
                //chart1.ChartAreas[0].AxisY.Maximum = double.NaN;
                //chart1.ChartAreas[0].AxisY.Minimum = double.NaN;
            //}

            // 将数据添加到曲线中
            series2.Points.AddXY(xValue2, irms_float);
            //chart2.Series["Series2"].Color = Color.Green;

            // 设置 X 轴的范围，仅显示最新的100个数据点
            chart2.ChartAreas[0].AxisX.Minimum = xValue2 - 400;
            chart2.ChartAreas[0].AxisX.Maximum = xValue2;
            chart2.ChartAreas[0].AxisY.LabelStyle.Format = "0.000";
            chart2.ChartAreas[0].AxisY.Interval = 0;
            //chart2.ChartAreas[0].AxisY.Title = "Irms";

            if (irms_maxValue > 0 && (irms_maxValue != irms_minValue))
            {
                chart2.ChartAreas[0].AxisY.Maximum = irms_maxValue;
                chart2.ChartAreas[0].AxisY.Minimum = irms_minValue;
            }
            else
            {
                // 设置Y轴自动调整最小值和最大值
                chart2.ChartAreas[0].AxisY.Maximum = double.NaN;
                chart2.ChartAreas[0].AxisY.Minimum = double.NaN;
            }

            // 将数据添加到曲线中
            series3.Points.AddXY(xValue3, power_float);
            //chart3.Series["Series3"].Color = Color.Red;

            // 设置 X 轴的范围，仅显示最新的100个数据点
            chart3.ChartAreas[0].AxisX.Minimum = xValue3 - 400;
            chart3.ChartAreas[0].AxisX.Maximum = xValue3;
            chart3.ChartAreas[0].AxisY.LabelStyle.Format = "0.00";
            chart3.ChartAreas[0].AxisY.Interval = 0;
            //chart3.ChartAreas[0].AxisY.Title = "Power";

            if (powe_maxValue > 0 && (powe_maxValue != powe_minValue))
            {
                chart3.ChartAreas[0].AxisY.Maximum = powe_maxValue;
                chart3.ChartAreas[0].AxisY.Minimum = powe_minValue;
            }
            else
            {
                // 设置Y轴自动调整最小值和最大值
                chart3.ChartAreas[0].AxisY.Maximum = double.NaN;
                chart3.ChartAreas[0].AxisY.Minimum = double.NaN;
            }

            // 刷新图表
            chart1.Refresh();
            chart2.Refresh();
            chart3.Refresh();

            // 增加 X 值
            xValue1++;
            xValue2++;
            xValue3++;
        }

        /*
        public class CRC8
        {
            private const byte Polynomial = 0x07; 

            public byte CalculateCRC(byte[] data)
            {
                byte crc = 0;
                foreach (byte b in data)
                {
                    crc ^= b;
                    for (int i = 0; i < 8; i++)
                    {
                        if ((crc & 0x80) != 0)
                        {
                            crc = (byte)((crc << 1) ^ Polynomial);
                        }
                        else
                        {
                            crc <<= 1;
                        }
                    }
                }
                return crc;
            }
        }
        */

        public class SumZeroCalculator
        {
            public byte CalculateSumZero(byte[] data)
            {
                byte sum = 0;

                // Summing the data
                foreach (byte b in data)
                {
                    sum += b;
                }

                // Two's complement
                byte sumZero = (byte)(0xFF - sum + 1);
                return sumZero;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            serialLoad();

            LoadComPorts();

            //this.vth_order_box.Text = "1";
            //this.ith_order_box.Text = "1";

            /*
            this.cali_v1_box.Text = "232.8309";
            this.cali_i1_box.Text = "0.350387";
            this.cali_p1_box.Text = "69.50793";

            this.cali_v2_box.Text = "232.6768";
            this.cali_i2_box.Text = "1.797165";
            this.cali_p2_box.Text = "412.9108";

            this.cali_v3_box.Text = "233.0053";
            this.cali_i3_box.Text = "3.638794";
            this.cali_p3_box.Text = "841.5959";

            this.cali_v4_box.Text = "233.4083";
            this.cali_i4_box.Text = "5.044907";
            this.cali_p4_box.Text = "1171.313";
            */

            // 設置 RichTextBox 的預設內容
            curr_100p_box.Text = "230"; //"106.6";  //"82";
            delay_sec_box.Text = "20";  //"20";
            repeat_box.Text    = "1";   //"1";

            // 初始化随机数生成器y
            //random = new Random();

            // 创建曲线的 Series
            series1 = new Series();
            series1.ChartType = SeriesChartType.Line;
            series1.BorderWidth = 3;
            series1.Color = Color.FromArgb(255, 40, 53, 118);
            chart1.Series.Add(series1);

            series2 = new Series();
            series2.ChartType = SeriesChartType.Line;
            series2.BorderWidth = 3;
            series2.Color = Color.FromArgb(255, 2, 97, 97);
            chart2.Series.Add(series2);

            series3 = new Series();
            series3.ChartType = SeriesChartType.Line;
            series3.BorderWidth = 3;
            series3.Color = Color.FromArgb(255, 115, 73, 48);
            chart3.Series.Add(series3);

            // 创建 Timer 控件，每100毫秒触发一次 Tick 事件
            timer = new System.Windows.Forms.Timer();
            timer.Interval = 100;
            timer.Tick += Timer_Tick;

            // 初始化Chart
            // 将数据添加到曲线中
            series1.Points.AddXY(xValue1, vrms_float);
            //chart1.Series["Series1"].Color = Color.Blue;

            // 设置 X 轴的范围，仅显示最新的100个数据点
            chart1.ChartAreas[0].AxisX.Minimum = xValue1 - 400;
            chart1.ChartAreas[0].AxisX.Maximum = xValue1;

            chart1.ChartAreas[0].AxisY.Minimum = Double.NaN;
            chart1.ChartAreas[0].AxisY.Maximum = Double.NaN;
            chart1.ChartAreas[0].AxisY.IsStartedFromZero = false;

            // 将数据添加到曲线中
            series2.Points.AddXY(xValue2, irms_float);
            //chart2.Series["Series2"].Color = Color.Green;

            // 设置 X 轴的范围，仅显示最新的100个数据点
            chart2.ChartAreas[0].AxisX.Minimum = xValue2 - 400;
            chart2.ChartAreas[0].AxisX.Maximum = xValue2;

            chart2.ChartAreas[0].AxisY.Minimum = Double.NaN;
            chart2.ChartAreas[0].AxisY.Maximum = Double.NaN;
            chart2.ChartAreas[0].AxisY.IsStartedFromZero = false;

            // 将数据添加到曲线中
            series3.Points.AddXY(xValue3, power_float);
            //chart3.Series["Series3"].Color = Color.Red;

            // 设置 X 轴的范围，仅显示最新的100个数据点
            chart3.ChartAreas[0].AxisX.Minimum = xValue3 - 400;
            chart3.ChartAreas[0].AxisX.Maximum = xValue3;

            chart3.ChartAreas[0].AxisY.Minimum = Double.NaN;
            chart3.ChartAreas[0].AxisY.Maximum = Double.NaN;
            chart3.ChartAreas[0].AxisY.IsStartedFromZero = false;

            // 刷新图表
            chart1.Refresh();
            chart2.Refresh();
            chart3.Refresh();
        }

        private void serialLoad()
        {
            /*
            EncodingInfo[] encodingInfos = Encoding.GetEncodings();
            RegistryKey keyCom = Registry.LocalMachine.OpenSubKey(@"Hardware\DeviceMap\SerialComm");

            string[] sSubKeys = keyCom.GetValueNames();
            port_cbb.Items.Clear();
            foreach (var sValue in sSubKeys)
            {
                string portName = (string)keyCom.GetValue(sValue);
                port_cbb.Items.Add(portName);
            }
            */

            EncodingInfo[] encodingInfos = Encoding.GetEncodings();
            RegistryKey keyCom = Registry.LocalMachine.OpenSubKey(@"HARDWARE\DeviceMap\SerialComm");

            if (keyCom != null)
            {
                try
                {
                    string[] sSubKeys = keyCom.GetValueNames();
                    port_cbb.Items.Clear();

                    foreach (var sValue in sSubKeys)
                    {
                        string portName = (string)keyCom.GetValue(sValue);
                        port_cbb.Items.Add(portName);
                    }

                    // 設定 ComboBox 控制項的預設選項
                    this.port_cbb.SelectedIndex   = 0;
                    //this.speed_cbb.SelectedIndex  = 6;
                    //this.parity_cbb.SelectedIndex = 0;
                    //this.data_cbb.SelectedIndex   = 3;
                    //this.stop_cbb.SelectedIndex   = 0;

                    keyCom.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"發生例外: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("找不到指定的註冊表金鑰。");
            }
        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {

        }

        private void groupBox4_Enter(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {                
                if ((!string.IsNullOrEmpty(cali_v1_box.Text)) && (!string.IsNullOrEmpty(cali_i1_box.Text)) && (!string.IsNullOrEmpty(cali_p1_box.Text))) // 如果 RichTextBox 的內容不為空
                {
                    // 更改按鈕的背景顏色
                    point1_button.BackColor = SystemColors.Control;

                    // 清除串口接收缓冲区中的数据
                    serialPort1.DiscardInBuffer();
                    receiveBuffer.Clear();  // 清除整个缓冲区，等待下一个有效数据帧

                    string text1 = cali_v1_box.Text;
                    float cali_v1_rms;
                    float.TryParse(text1, out cali_v1_rms);
                    // Console.WriteLine(cali_v1_rms);

                    cali_v1_rms = cali_v1_rms * (float)1000.0;
                    int intValue1 = Convert.ToInt32(cali_v1_rms);     // 將浮點數轉換為整數

                    byte[] byteArray1 = new byte[4];             
                    byteArray1[0] = (byte)((intValue1 >> 24) & 0xFF); // 取得高位元組
                    byteArray1[1] = (byte)((intValue1 >> 16) & 0xFF); // 取得次高位元組
                    byteArray1[2] = (byte)((intValue1 >>  8) & 0xFF); // 取得次低位元組
                    byteArray1[3] = (byte)((intValue1 >>  0) & 0xFF); // 取得低位元組

                    // SumZero
                    SumZeroCalculator SZ_1 = new SumZeroCalculator();
                    byte[] data_1 = new byte[] { 0x49, 0x05, 0x41, byteArray1[0], byteArray1[1], byteArray1[2], byteArray1[3] };
                    byte SumZero_1 = SZ_1.CalculateSumZero(data_1);
                    //Console.WriteLine("SumZero Value: 0x" + SumZero_1.ToString("X2"));                                        

                    // 设置目标为 chip_ssid_box
                    targetTextBox = "p1_vrms_box";
                    expectedLength = 15;

                    // 向串口发送数据 
                    byte[] myBytes1 = { 0x49, 0x05, 0x41, byteArray1[0], byteArray1[1], byteArray1[2], byteArray1[3], SumZero_1 }; 
                    serialPort1.Write(myBytes1, 0, myBytes1.Length);
                    Thread.Sleep(200);  // Delay 100ms
                    //await Task.Delay(500);  // 非同步延遲

                    // Read status
                    //byte[] myBytes2 = { 0x49, 0x01, 0x85, 0x31, 0x0D }; 
                    //serialPort1.Write(myBytes2, 0, myBytes2.Length); 
                    //Thread.Sleep(100);  // Delay 100ms

                    /*
                    // 从串口读取数据
                    if (serialPort1.BytesToRead >= 4)
                    {
                        byte[] buffer = new byte[4];
                        serialPort1.Read(buffer, 0, 4);

                        // Debug
                        //string result = BitConverter.ToString(buffer, 0, buffer.Length).Replace("-", " "); // 將讀取到的數據轉換為字串
                        //Console.WriteLine(result);

                        // 检查第一个字节为 0x49，第二个字节为 0x01
                        if (buffer[0] == 0x49 && buffer[1] == 0x01)
                        {
                            Console.WriteLine("Status Value: 0x" + buffer[2].ToString("X2"));

                            if(buffer[2]==4)
                                MessageBox.Show("Sum-Zero Error ! ", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    */

                    /*******************************************************************/

                    // 清除串口接收缓冲区中的数据
                    serialPort1.DiscardInBuffer();
                    receiveBuffer.Clear();  // 清除整个缓冲区，等待下一个有效数据帧

                    string text2 = cali_i1_box.Text;
                    float cali_i1_rms;
                    float.TryParse(text2, out cali_i1_rms);
                    //Console.WriteLine(cali_i1_vrms);

                    cali_i1_rms = cali_i1_rms * (float)128000.0;
                    int intValue2 = Convert.ToInt32(cali_i1_rms);     // 將浮點數轉換為整數

                    byte[] byteArray2 = new byte[4];
                    byteArray2[0] = (byte)((intValue2 >> 24) & 0xFF); // 取得高位元組
                    byteArray2[1] = (byte)((intValue2 >> 16) & 0xFF); // 取得次高位元組
                    byteArray2[2] = (byte)((intValue2 >>  8) & 0xFF); // 取得次低位元組
                    byteArray2[3] = (byte)((intValue2 >>  0) & 0xFF); // 取得低位元組

                    // SumZero
                    SumZeroCalculator SZ_2 = new SumZeroCalculator();
                    byte[] data_2 = new byte[] { 0x49, 0x05, 0x43, byteArray2[0], byteArray2[1], byteArray2[2], byteArray2[3] };
                    byte SumZero_2 = SZ_2.CalculateSumZero(data_2);
                    //Console.WriteLine("SumZero Value: 0x" + SumZero_2.ToString("X2"));

                    // 向串口发送数据 
                    byte[] myBytes3 = { 0x49, 0x05, 0x43, byteArray2[0], byteArray2[1], byteArray2[2], byteArray2[3], SumZero_2 }; 
                    serialPort1.Write(myBytes3, 0, myBytes3.Length);
                    Thread.Sleep(200);  // Delay 100ms
                    //await Task.Delay(500);  // 非同步延遲

                    // Read status
                    //byte[] myBytes4 = { 0x49, 0x01, 0x85, 0x31, 0x0D }; 
                    //serialPort1.Write(myBytes4, 0, myBytes4.Length); 
                    //Thread.Sleep(100);  // Delay 100ms

                    /*
                    // 从串口读取数据
                    if (serialPort1.BytesToRead >= 4)
                    {
                        byte[] buffer = new byte[4];
                        serialPort1.Read(buffer, 0, 4);

                        // Debug
                        //string result = BitConverter.ToString(buffer, 0, buffer.Length).Replace("-", " "); // 將讀取到的數據轉換為字串
                        //Console.WriteLine(result);

                        // 检查第一个字节为 0x49，第二个字节为 0x01
                        if (buffer[0] == 0x49 && buffer[1] == 0x01)
                        {
                            Console.WriteLine("Status Value: 0x" + buffer[2].ToString("X2"));

                            if (buffer[2] == 4)
                                MessageBox.Show("Sum-Zero Error ! ", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    */

                    /*******************************************************************/

                    // 清除串口接收缓冲区中的数据
                    serialPort1.DiscardInBuffer();
                    receiveBuffer.Clear();  // 清除整个缓冲区，等待下一个有效数据帧

                    string text3 = cali_p1_box.Text;
                    float cali_p1_rms;
                    float.TryParse(text3, out cali_p1_rms);
                    //Console.WriteLine(cali_p1_rms);

                    cali_p1_rms = cali_p1_rms * (float)1000.0;
                    int intValue3 = Convert.ToInt32(cali_p1_rms);     // 將浮點數轉換為整數

                    byte[] byteArray3 = new byte[4];
                    byteArray3[0] = (byte)((intValue3 >> 24) & 0xFF); // 取得高位元組
                    byteArray3[1] = (byte)((intValue3 >> 16) & 0xFF); // 取得次高位元組
                    byteArray3[2] = (byte)((intValue3 >>  8) & 0xFF); // 取得次低位元組
                    byteArray3[3] = (byte)((intValue3 >>  0) & 0xFF); // 取得低位元組

                    // SumZero
                    SumZeroCalculator SZ_3 = new SumZeroCalculator();
                    byte[] data_3 = new byte[] { 0x49, 0x05, 0x45, byteArray3[0], byteArray3[1], byteArray3[2], byteArray3[3] };
                    byte SumZero_3 = SZ_3.CalculateSumZero(data_3);
                    //Console.WriteLine("SumZero Value: 0x" + SumZero_3.ToString("X2"));

                    // 向串口发送数据 
                    byte[] myBytes5 = { 0x49, 0x05, 0x45, byteArray3[0], byteArray3[1], byteArray3[2], byteArray3[3], SumZero_3 }; 
                    serialPort1.Write(myBytes5, 0, myBytes5.Length);
                    Thread.Sleep(200);  // Delay 100ms
                    //await Task.Delay(500);  // 非同步延遲

                    // Read status
                    //byte[] myBytes6 = { 0x49, 0x01, 0x85, 0x31, 0x0D }; 
                    //serialPort1.Write(myBytes6, 0, myBytes6.Length); 
                    //Thread.Sleep(100);  // Delay 100ms

                    /*
                    // 从串口读取数据
                    if (serialPort1.BytesToRead >= 4)
                    {
                        byte[] buffer = new byte[4];
                        serialPort1.Read(buffer, 0, 4);

                        // Debug
                        //string result = BitConverter.ToString(buffer, 0, buffer.Length).Replace("-", " "); // 將讀取到的數據轉換為字串
                        //Console.WriteLine(result);

                        // 检查第一个字节为 0x49，第二个字节为 0x01
                        if (buffer[0] == 0x49 && buffer[1] == 0x01)
                        {
                            Console.WriteLine("Status Value: 0x" + buffer[2].ToString("X2"));

                            if (buffer[2] == 4)
                                MessageBox.Show("Sum-Zero Error ! ", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    */

                    /*******************************************************************/

                    // 清除串口接收缓冲区中的数据
                    serialPort1.DiscardInBuffer();
                    receiveBuffer.Clear();  // 清除整个缓冲区，等待下一个有效数据帧

                    // 向串口发送数据 
                    byte[] myBytes7 = { 0x49, 0x01, 0x47, 0x6F }; 
                    serialPort1.Write(myBytes7, 0, myBytes7.Length);
                    Thread.Sleep(4000);  // Delay 4000ms
                    //await Task.Delay(3000);  // 非同步延遲

                    /*
                    // 从串口读取数据
                    if (serialPort1.BytesToRead >= 15)
                    {
                        byte[] buffer = new byte[15];
                        serialPort1.Read(buffer, 0, 15);

                        // Debug
                        string result3 = BitConverter.ToString(buffer, 0, buffer.Length).Replace("-", " "); // 將讀取到的數據轉換為字串
                        Console.WriteLine(result3);

                        // 在這裡處理讀取到的資料
                        if ((buffer[0] == 0x49) && (buffer[1] == 0x0C))
                        {
                            uint Vrms_int32 = (uint)(buffer[2] << 24 | buffer[3] << 16 | buffer[4] << 8 | buffer[5]);
                            float Vrms_float = (float)(Vrms_int32 / 1000.0);
                            Console.WriteLine("Vrms : " + Vrms_float);
                            p1_vrms_box.Text = Vrms_float.ToString();

                            uint Irms_int32 = (uint)(buffer[6] << 24 | buffer[7] << 16 | buffer[8] << 8 | buffer[9]);
                            float Irms_float = (float)(Irms_int32 / 128000.0);                           
                            Console.WriteLine("Irms : " + Irms_float);
                            p1_irms_box.Text = Irms_float.ToString();

                            uint Pavg_int32 = (uint)(buffer[10] << 24 | buffer[11] << 16 | buffer[12] << 8 | buffer[13]);
                            float Pavg_float = (float)(Pavg_int32 / 1000.0);
                            Console.WriteLine("Pavg : " + Pavg_float);
                            p1_pavg_box.Text = Pavg_float.ToString();
                        }
                    }
                    */

                    /*******************************************************************/

                    // 清除串口接收缓冲区中的数据
                    serialPort1.DiscardInBuffer();
                    receiveBuffer.Clear();  // 清除整个缓冲区，等待下一个有效数据帧

                    // 向串口发送数据 
                    byte[] myBytes8 = { 0x49, 0x01, 0x49, 0x6D }; 
                    serialPort1.Write(myBytes8, 0, myBytes8.Length);
                    Thread.Sleep(200);  // Delay 100ms
                    //await Task.Delay(500);  // 非同步延遲

                    // 恢复19字节接收模式
                    expectedLength = 19;

                    // 更改按鈕的背景顏色
                    point1_button.BackColor = Color.LightBlue;

                    // Read status
                    //byte[] myBytes9 = { 0x49, 0x01, 0x85, 0x31, 0x0D };
                    //serialPort1.Write(myBytes9, 0, myBytes9.Length); 
                    //Thread.Sleep(100);  // Delay 100ms

                    /*
                    // 从串口读取数据
                    if (serialPort1.BytesToRead >= 4)
                    {
                        byte[] buffer = new byte[4];
                        serialPort1.Read(buffer, 0, 4);

                        // Debug
                        //string result = BitConverter.ToString(buffer, 0, buffer.Length).Replace("-", " "); // 將讀取到的數據轉換為字串
                        //Console.WriteLine(result);

                        // 检查第一个字节为 0x49，第二个字节为 0x01
                        if (buffer[0] == 0x49 && buffer[1] == 0x01)
                        {
                            Console.WriteLine("Status Value: 0x" + buffer[2].ToString("X2"));

                            if (buffer[2] == 4)
                                MessageBox.Show("Sum-Zero Error ! ", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    */
                }
            }
        }

        private void groupBox5_Enter(object sender, EventArgs e)
        {

        }

        private void data_cbb_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void speed_cbb_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void StatusLabel2_Click(object sender, EventArgs e)
        {

        }

        private void rts_chb_CheckedChanged(object sender, EventArgs e)
        {   
            /*
            if (rts_chb.CheckState == CheckState.Checked)
            {
                serialPort1.RtsEnable = true;
            }
            else
            {
                serialPort1.RtsEnable = false;
            }
            */
        }

        private void dtr_chb_CheckedChanged(object sender, EventArgs e)
        {
            /*
            if (dtr_chb.CheckState == CheckState.Checked)
            {
                serialPort1.DtrEnable = true;
            }
            else
            {
                serialPort1.DtrEnable = false;
            }
            */
        }

        private void sendData()
        {
            serialPort1.Write(sendBuffer.ToArray(), 0, sendBuffer.Count);
            //sendCount += sendBuffer.Count;
            //sencount_tssl.Text = sendCount.ToString();
        }


        private void button5_Click(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                // 清空 RichTextBox 控制元件內容
                chip_ssid_box.Clear();

                // 清除串口接收缓冲区中的数据
                serialPort1.DiscardInBuffer();
                receiveBuffer.Clear();  // 清除整个缓冲区，等待下一个有效数据帧

                // 發送停止命令 
                byte[] myBytes1 = { 0x49, 0x01, 0x83, 0x33 };
                serialPort1.Write(myBytes1, 0, myBytes1.Length);
                Thread.Sleep(200);  // Delay 100ms
                //await Task.Delay(200);  // 非同步延遲

                // 设置目标为 chip_ssid_box
                targetTextBox = "chip_ssid_box";  
                expectedLength = 4;

                // 向串口发送数据
                byte[] myBytes = { 0x49, 0x01, 0x10, 0xA6 };
                serialPort1.Write(myBytes, 0, myBytes.Length);
                Thread.Sleep(200);  // Delay 100ms
                //await Task.Delay(200);  // 非同步延遲

                /*
                if (serialPort1.BytesToRead >= 4)
                {
                    byte[] buffer = new byte[4];
                    serialPort1.Read(buffer, 0, 4);
                    
                    // Debug
                    string result = BitConverter.ToString(buffer, 0, buffer.Length).Replace("-", " "); // 將讀取到的數據轉換為字串
                    Console.WriteLine(result);

                    // 检查第一个字节为 0x49，第二个字节为 0x01
                    if (buffer[0] == 0x49 && buffer[1] == 0x01)
                    {
                        byte thirdByte = buffer[2];
                        Console.WriteLine("Value : " + thirdByte);
                        string dataString = thirdByte.ToString("D2");
                        chip_ssid_box.AppendText(dataString);
                    }                    
                }
                */
            }
        }

        private void chip_ssid_box_TextChanged(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                // 清空 RichTextBox 控制元件內容
                fw_ver_box.Clear();

                // 清除串口接收缓冲区中的数据
                serialPort1.DiscardInBuffer();
                receiveBuffer.Clear();  // 清除整个缓冲区，等待下一个有效数据帧

                // 發送停止命令 
                byte[] myBytes1 = { 0x49, 0x01, 0x83, 0x33 };
                serialPort1.Write(myBytes1, 0, myBytes1.Length);
                Thread.Sleep(200);  // Delay 100ms
                //await Task.Delay(200);  // 非同步延遲

                // 设置目标为 fw_ver_box
                targetTextBox = "fw_ver_box";  
                expectedLength = 4; 

                // 向串口发送数据
                byte[] myBytes = { 0x49, 0x01, 0x12, 0xA4 };
                serialPort1.Write(myBytes, 0, myBytes.Length);
                Thread.Sleep(200);  // Delay 100ms
                //await Task.Delay(200);  // 非同步延遲

                /*
                if (serialPort1.BytesToRead >= 4)
                {
                    byte[] buffer = new byte[4];
                    serialPort1.Read(buffer, 0, 4);                    
                    
                    // Debug
                    //string result = BitConverter.ToString(buffer, 0, buffer.Length).Replace("-", " "); // 將讀取到的數據轉換為字串
                    //Console.WriteLine(result);

                    // 检查第一个字节为 0x49，第二个字节为 0x01
                    if (buffer[0] == 0x49 && buffer[1] == 0x01)
                    {
                        byte thirdByte = buffer[2];
                        Console.WriteLine("Value : " + thirdByte);
                        string dataString = thirdByte.ToString("D2");
                        fw_ver_box.AppendText(dataString);
                    }
                }
                */
            }
        }

        private void textBox13_TextChanged(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                // 清空 RichTextBox 控制元件內容
                p1_vrms_box.Clear();

                // 清除串口接收缓冲区中的数据
                serialPort1.DiscardInBuffer();
                receiveBuffer.Clear();  // 清除整个缓冲区，等待下一个有效数据帧

                // 發送停止命令 
                byte[] myBytes1 = { 0x49, 0x01, 0x83, 0x33 };
                serialPort1.Write(myBytes1, 0, myBytes1.Length);
                Thread.Sleep(200);  // Delay 100ms
                //await Task.Delay(200);  // 非同步延遲

                // 设置目标为 fw_ver_box
                targetTextBox = "p1_vrms_box";
                expectedLength = 7;

                // 向串口发送数据 
                byte[] myBytes = { 0x49, 0x01, 0x40, 0x76 }; 
                serialPort1.Write(myBytes, 0, myBytes.Length);
                Thread.Sleep(200);  // Delay 100ms
                //await Task.Delay(200);  // 非同步延遲

                /*
                if (serialPort1.BytesToRead >= 7)
                {
                    byte[] buffer = new byte[7];
                    serialPort1.Read(buffer, 0, 7);                    
                    
                    // Debug
                    string result = BitConverter.ToString(buffer, 0, buffer.Length).Replace("-", " "); // 將讀取到的數據轉換為字串
                    Console.WriteLine(result);

                    // 检查第一个字节为 0x49，第二个字节为 0x04
                    if (buffer[0] == 0x49 && buffer[1] == 0x04)
                    {
                        uint Vrms_int32 = (uint)(buffer[2] << 24 | buffer[3] << 16 | buffer[4] << 8 | buffer[5]);
                        float Vrms_float = (float)(Vrms_int32 / 1000.0);
                        Console.WriteLine("Value : " + Vrms_float);
                        p1_vrms_box.Text = Vrms_float.ToString();
                    }
                }
                */
            }
        }

        private void port_cbb_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label6_Click_1(object sender, EventArgs e)
        {

        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {

        }

        
        private void button2_Click(object sender, EventArgs e)
        {
            /*
            if (serialPort1.IsOpen)
            {                
                if (isReport)
                {
                    // 切換狀態
                    isReport = false;
                    auto_report_btn.Text = "Start";

                    // 停止計算線程
                    thread1.Abort();
                    isThread1Running = false;

                    // 發送停止命令 
                    byte[] myBytes1 = { 0x49, 0x01, 0x83, 0x29, 0x0D };
                    serialPort1.Write(myBytes1, 0, myBytes1.Length);
                    Thread.Sleep(100);  // Delay 100ms    

                    // Chart            
                    // 停止 Timer 控件
                    timer.Stop();
                }
                else
                {
                    // 切換狀態
                    isReport = true;
                    auto_report_btn.Text = "Stop";

                    // 開啟計算線程
                    //thread1 = new Thread(new ThreadStart(DoWork1));
                    thread1.Start();
                    isThread1Running = true;
                    
                    // 向串口发送数据 
                    byte[] myBytes = { 0x49, 0x02, 0x81, 0x7F, 0x32, 0x0D };
                    serialPort1.Write(myBytes, 0, myBytes.Length);
                    Thread.Sleep(100);  // Delay 100ms

                    // Chart            
                    // 启动 Timer 控件
                    timer.Start();
                }              
            }
            */
        }

        /*
        private void DoWork1()
        {
            try
            {
                while (true)
                {
                    try
                    {
                        if (serialPort1.IsOpen)
                        {
                            // 清除串口接收缓冲区中的数据
                            serialPort1.DiscardInBuffer();
                            receiveBuffer.Clear();  // 清除整个缓冲区，等待下一个有效数据帧

                            Thread.Sleep(20);  // Delay 20ms

                            // 49 0C 00 03 86 80 00 05 38 71 00 09 59 76 49
                            // 49 10 00 01 87 E6 00 01 9B FA 00 01 38 E4 00
                            if (serialPort1.BytesToRead >= 19)
                            {
                                byte[] buffer = new byte[19];
                                serialPort1.Read(buffer, 0, 19);

                                // Debug
                                string result = BitConverter.ToString(buffer, 0, buffer.Length).Replace("-", " "); // 將讀取到的數據轉換為字串
                                Console.WriteLine(result);

                                // 检查第一个字节为 0xAE，第二个字节为 0x0F
                                if (buffer[0] == 0x49 && buffer[1] == 0x10)
                                {
                                    //int auto_report_frequency = buffer[2];
                                    //freq_int = auto_report_frequency;
                                    //Console.WriteLine("auto_report_frequency : " + freq_int);

                                    int auto_report_vrms = buffer[4] << 16 | buffer[3] << 8 | buffer[2];
                                    vrms_float = (float)(auto_report_vrms / 1000.0);
                                    Console.WriteLine("auto_report_vrms : " + vrms_float);

                                    int auto_report_irms = buffer[7] << 16 | buffer[6] << 8 | buffer[5];
                                    irms_float = (float)(auto_report_irms / 128000.0);
                                    Console.WriteLine("auto_report_irms : " + irms_float);

                                    int auto_report_power = buffer[10] << 16 | buffer[9] << 8 | buffer[8];
                                    power_float = (float)(auto_report_power / 1000.0);
                                    Console.WriteLine("auto_report_power : " + power_float);

                                    int auto_report_frequency = buffer[13] << 16 | buffer[12] << 8 | buffer[11];
                                    freq_int = auto_report_frequency / 1000;
                                    Console.WriteLine("auto_report_frequency : " + freq_int);

                                    Action action = () =>
                                    {
                                        freq_box.Text = freq_int.ToString();

                                        vrms_box.ForeColor = Color.FromArgb(255, 40, 53, 118);
                                        vrms_box.Text = vrms_float.ToString("N2");

                                        irms_box.ForeColor = Color.FromArgb(255, 2, 97, 97);
                                        irms_box.Text = irms_float.ToString("N5");

                                        watt_box.ForeColor = Color.FromArgb(255, 115, 73, 48);
                                        watt_box.Text = power_float.ToString("N3");
                                    };
                                    this.Invoke(action);
                                }
                            }
                        }
                        else
                        {
                            // Handle the case where the port is closed
                        }
                    }
                    catch (InvalidOperationException ex)
                    {
                        // Handle the exception (e.g., log the error, display a message to the user)
                    }
                }
            }
            catch (OperationCanceledException)
            {
                // 被取消時不處理
            }
        }
        */

        /*
        private void DoWork1()
        {
            try
            {
                while (true)
                {
                    // 清除串口接收缓冲区中的数据
                    serialPort1.DiscardInBuffer();
                    receiveBuffer.Clear();  // 清除整个缓冲区，等待下一个有效数据帧

                    Thread.Sleep(100);  // Delay 100ms

                    if (serialPort1.BytesToRead >= 28)
                    {
                        byte[] buffer = new byte[28];
                        serialPort1.Read(buffer, 0, 28);

                        // Debug
                        string result = BitConverter.ToString(buffer, 0, buffer.Length).Replace("-", " "); // 將讀取到的數據轉換為字串
                        Console.WriteLine(result);

                        // 检查第一个字节为 0x49，第二个字节为 0x01
                        if (buffer[0] == 0x49 && buffer[1] == 0x19)
                        {
                            int auto_report_frequency = buffer[2];
                            freq_int = auto_report_frequency;
                            Console.WriteLine("auto_report_frequency : " + freq_int);

                            int auto_report_vrms = buffer[3] << 24 | buffer[4] << 16 | buffer[5] << 8 | buffer[6];
                            vrms_float = (float)(auto_report_vrms / 65536.0);
                            Console.WriteLine("auto_report_vrms : " + vrms_float);

                            int auto_report_irms = buffer[7] << 24 | buffer[8] << 16 | buffer[9] << 8 | buffer[10];
                            irms_float = (float)(auto_report_irms / 65536.0);
                            Console.WriteLine("auto_report_irms : " + irms_float);

                            int auto_report_power = buffer[11] << 24 | buffer[12] << 16 | buffer[13] << 8 | buffer[14];
                            power_float = (float)(auto_report_power / 65536.0);
                            Console.WriteLine("auto_report_power : " + power_float);

                            int auto_report_pf_value = buffer[15] << 24 | buffer[16] << 16 | buffer[17] << 8 | buffer[18];
                            pf_value_float = (float)(auto_report_pf_value / 65536.0);
                            Console.WriteLine("auto_report_power : " + pf_value_float);

                            int auto_report_v_thd = buffer[19] << 24 | buffer[20] << 16 | buffer[21] << 8 | buffer[22];
                            v_thd_float = (float)(auto_report_v_thd / 65536.0);
                            Console.WriteLine("auto_report_v_thd : " + v_thd_float);

                            int auto_report_i_thd = buffer[23] << 24 | buffer[24] << 16 | buffer[25] << 8 | buffer[26];
                            i_thd_float = (float)(auto_report_i_thd / 65536.0);
                            Console.WriteLine("auto_report_i_thd : " + i_thd_float);
                                               
                            Action action = () =>
                            {
                                freq_box.Text  = freq_int.ToString();
                                vrms_box.Text  = vrms_float.ToString();
                                irms_box.Text  = irms_float.ToString();
                                watt_box.Text  = power_float.ToString();
                                //pf_box.Text    = pf_value_float.ToString();
                                //v_thd_box.Text = v_thd_float.ToString();
                                //i_thd_box.Text = i_thd_float.ToString();
                            };
                            this.Invoke(action);                   
                        }
                    }
                }
            }
            catch (OperationCanceledException)
            {
                // 被取消時不處理
            }
        }
        */


        private void button9_Click(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                // 清空 RichTextBox 控制元件內容
                p1_irms_box.Clear();

                // 清除串口接收缓冲区中的数据
                serialPort1.DiscardInBuffer();
                receiveBuffer.Clear();  // 清除整个缓冲区，等待下一个有效数据帧

                // 發送停止命令 
                byte[] myBytes1 = { 0x49, 0x01, 0x83, 0x33 };
                serialPort1.Write(myBytes1, 0, myBytes1.Length);
                Thread.Sleep(200);  // Delay 100ms
                //await Task.Delay(200);  // 非同步延遲  

                // 设置目标为 p1_irms_box
                targetTextBox = "p1_irms_box";
                expectedLength = 7;

                // 向串口发送数据
                byte[] myBytes = { 0x49, 0x01, 0x42, 0x74 }; 
                serialPort1.Write(myBytes, 0, myBytes.Length);
                Thread.Sleep(200);  // Delay 100ms
                //await Task.Delay(200);  // 非同步延遲

                /*
                if (serialPort1.BytesToRead >= 7)
                {
                    byte[] buffer = new byte[7];
                    serialPort1.Read(buffer, 0, 7);                    
                    
                    // Debug
                    string result = BitConverter.ToString(buffer, 0, buffer.Length).Replace("-", " "); // 將讀取到的數據轉換為字串
                    Console.WriteLine(result);

                    // 检查第一个字节为 0x49，第二个字节为 0x04
                    if (buffer[0] == 0x49 && buffer[1] == 0x04)
                    {
                        uint Irms_int32 = (uint)(buffer[2] << 24 | buffer[3] << 16 | buffer[4] << 8 | buffer[5]);
                        float Irms_float = (float)(Irms_int32 / 128000.0);
                        Console.WriteLine("Value : " + Irms_float);
                        p1_irms_box.Text = Irms_float.ToString();
                    }
                }
                */
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                // 清空 RichTextBox 控制元件內容
                p4_vrms_box.Clear();

                // 清除串口接收缓冲区中的数据
                serialPort1.DiscardInBuffer();
                receiveBuffer.Clear();  // 清除整个缓冲区，等待下一个有效数据帧

                // 發送停止命令 
                byte[] myBytes1 = { 0x49, 0x01, 0x83, 0x33 };
                serialPort1.Write(myBytes1, 0, myBytes1.Length);
                Thread.Sleep(200);  // Delay 100ms
                //await Task.Delay(200);  // 非同步延遲

                // 设置目标为 p4_vrms_box
                targetTextBox = "p4_vrms_box";
                expectedLength = 7;

                // 向串口发送数据
                byte[] myBytes = { 0x49, 0x01, 0x50, 0x66 }; 
                serialPort1.Write(myBytes, 0, myBytes.Length);
                Thread.Sleep(200);  // Delay 100ms
                //await Task.Delay(200);  // 非同步延遲

                /*
                if (serialPort1.BytesToRead >= 7)
                {
                    byte[] buffer = new byte[7];
                    serialPort1.Read(buffer, 0, 7);                    
                    
                    // Debug
                    string result = BitConverter.ToString(buffer, 0, buffer.Length).Replace("-", " "); // 將讀取到的數據轉換為字串
                    Console.WriteLine(result);

                    // 检查第一个字节为 0x49，第二个字节为 0x04
                    if (buffer[0] == 0x49 && buffer[1] == 0x04)
                    {
                        uint Vrms_int32 = (uint)(buffer[2] << 24 | buffer[3] << 16 | buffer[4] << 8 | buffer[5]);
                        float Vrms_float = (float)(Vrms_int32 / 1000.0);
                        Console.WriteLine("Value : " + Vrms_float);
                        p4_vrms_box.Text = Vrms_float.ToString();
                    }
                }
                */
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                // 清空 RichTextBox 控制元件內容
                p4_irms_box.Clear();

                // 清除串口接收缓冲区中的数据
                serialPort1.DiscardInBuffer();
                receiveBuffer.Clear();  // 清除整个缓冲区，等待下一个有效数据帧

                // 發送停止命令 
                byte[] myBytes1 = { 0x49, 0x01, 0x83, 0x33 };
                serialPort1.Write(myBytes1, 0, myBytes1.Length);
                Thread.Sleep(200);  // Delay 100ms
                //await Task.Delay(200);  // 非同步延遲 

                // 设置目标为 p4_irms_box
                targetTextBox = "p4_irms_box";
                expectedLength = 7;

                // 向串口发送数据
                byte[] myBytes = { 0x49, 0x01, 0x52, 0x64 };
                serialPort1.Write(myBytes, 0, myBytes.Length);
                Thread.Sleep(200);  // Delay 100ms
                //await Task.Delay(200);  // 非同步延遲

                /*
                if (serialPort1.BytesToRead >= 7)
                {
                    byte[] buffer = new byte[7];
                    serialPort1.Read(buffer, 0, 7);                    
                    
                    // Debug
                    string result = BitConverter.ToString(buffer, 0, buffer.Length).Replace("-", " "); // 將讀取到的數據轉換為字串
                    Console.WriteLine(result);

                    // 检查第一个字节为 0x49，第二个字节为 0x04
                    if (buffer[0] == 0x49 && buffer[1] == 0x04)
                    {
                        uint Irms_int32 = (uint)(buffer[2] << 24 | buffer[3] << 16 | buffer[4] << 8 | buffer[5]);
                        float Irms_float = (float)(Irms_int32 / 128000.0);
                        Console.WriteLine("Value : " + Irms_float);
                        p4_irms_box.Text = Irms_float.ToString();
                    }
                }
                */
            }
        }
               


        private void button1_Click_1(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                try
                {
                    Excel.Application excelApp = new Excel.Application();
                    Excel.Workbook workbook = excelApp.Workbooks.Add();
                    Excel.Worksheet worksheet = workbook.ActiveSheet;

                    // 將資料寫入工作表
                    worksheet.Cells[1, 1].Value = "Freq";
                    worksheet.Cells[1, 2].Value = "Vrms";
                    worksheet.Cells[1, 3].Value = "Irms";
                    worksheet.Cells[1, 4].Value = "Pow";
                    worksheet.Cells[1, 5].Value = "PF";
                    worksheet.Cells[1, 6].Value = "VTHD";
                    worksheet.Cells[1, 7].Value = "ITHD";

                    for (int i = 0; i < 400; i++)
                    {
                        worksheet.Cells[i + 2, 1].Value = freq_curve_data[i];
                        worksheet.Cells[i + 2, 2].Value = vrms_curve_data[i];
                        worksheet.Cells[i + 2, 3].Value = irms_curve_data[i];
                        worksheet.Cells[i + 2, 4].Value = powe_curve_data[i];
                        worksheet.Cells[i + 2, 5].Value = pfva_curve_data[i];
                        worksheet.Cells[i + 2, 6].Value = vthd_curve_data[i];
                        worksheet.Cells[i + 2, 7].Value = ithd_curve_data[i];                        
                    }

                    // 取得現在時間並格式化為檔名
                    string fileName = DateTime.Now.ToString("yyyyMMdd_HHmmss");

                    // 取得目前應用程式執行的目錄路徑
                    string currentDirectory = System.IO.Directory.GetCurrentDirectory();

                    // 組合完整的檔案路徑
                    string filePath = System.IO.Path.Combine(currentDirectory, $"{fileName}.xlsx");

                    // 儲存檔案
                    workbook.SaveAs(filePath);

                    // 關閉工作簿和 Excel 應用程式
                    workbook.Close();
                    excelApp.Quit();

                    // 釋放 COM 物件
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

                     MessageBox.Show("File saved successfully ! ", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                  
                }
                catch (Exception ex)
                {
                    // 处理异常
                    Console.WriteLine("Error : " + ex.Message);
                    MessageBox.Show("Failed to save file ! ", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void chart2_Click(object sender, EventArgs e)
        {

        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        private void label32_Click(object sender, EventArgs e)
        {

        }

        private void label31_Click(object sender, EventArgs e)
        {

        }

        private void label39_Click(object sender, EventArgs e)
        {

        }

        private void button13_Click(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                // 清空 RichTextBox 控制元件內容
                p3_vrms_box.Clear();

                // 清除串口接收缓冲区中的数据
                serialPort1.DiscardInBuffer();
                receiveBuffer.Clear();  // 清除整个缓冲区，等待下一个有效数据帧

                // 發送停止命令 
                byte[] myBytes1 = { 0x49, 0x01, 0x83, 0x33 };
                serialPort1.Write(myBytes1, 0, myBytes1.Length);
                Thread.Sleep(200);  // Delay 100ms
                //await Task.Delay(200);  // 非同步延遲

                // 设置目标为 p4_irms_box
                targetTextBox = "p3_vrms_box";
                expectedLength = 7;

                // 向串口发送数据
                byte[] myBytes = { 0x49, 0x01, 0x20, 0x96 }; 
                serialPort1.Write(myBytes, 0, myBytes.Length);
                Thread.Sleep(200);  // Delay 100ms
                //await Task.Delay(200);  // 非同步延遲

                /*
                if (serialPort1.BytesToRead >= 7)
                {
                    byte[] buffer = new byte[7];
                    serialPort1.Read(buffer, 0, 7);

                    // Debug
                    string result = BitConverter.ToString(buffer, 0, buffer.Length).Replace("-", " "); // 將讀取到的數據轉換為字串
                    Console.WriteLine(result);

                    // 检查第一个字节为 0x49，第二个字节为 0x04
                    if (buffer[0] == 0x49 && buffer[1] == 0x04)
                    {
                        uint Vrms_int32 = (uint)(buffer[2] << 24 | buffer[3] << 16 | buffer[4] << 8 | buffer[5]);
                        float Vrms_float = (float)(Vrms_int32 / 1000.0);
                        Console.WriteLine("Value : " + Vrms_float);
                        p3_vrms_box.Text = Vrms_float.ToString();
                    }
                }
                */
            }
        }

        private void button11_Click_1(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                // 清空 RichTextBox 控制元件內容
                p3_irms_box.Clear();

                // 清除串口接收缓冲区中的数据
                serialPort1.DiscardInBuffer();
                receiveBuffer.Clear();  // 清除整个缓冲区，等待下一个有效数据帧

                // 發送停止命令 
                byte[] myBytes1 = { 0x49, 0x01, 0x83, 0x33 };
                serialPort1.Write(myBytes1, 0, myBytes1.Length);
                Thread.Sleep(200);  // Delay 100ms
                //await Task.Delay(200);  // 非同步延遲   

                // 设置目标为 p3_irms_box
                targetTextBox = "p3_irms_box";
                expectedLength = 7;

                // 向串口发送数据
                byte[] myBytes = { 0x49, 0x01, 0x22, 0x94 }; 
                serialPort1.Write(myBytes, 0, myBytes.Length);
                Thread.Sleep(200);  // Delay 100ms
                //await Task.Delay(200);  // 非同步延遲

                /*
                if (serialPort1.BytesToRead >= 7)
                {
                    byte[] buffer = new byte[7];
                    serialPort1.Read(buffer, 0, 7);

                    // Debug
                    string result = BitConverter.ToString(buffer, 0, buffer.Length).Replace("-", " "); // 將讀取到的數據轉換為字串
                    Console.WriteLine(result);

                    // 检查第一个字节为 0x49，第二个字节为 0x04
                    if (buffer[0] == 0x49 && buffer[1] == 0x04)
                    {
                        uint Irms_int32 = (uint)(buffer[2] << 24 | buffer[3] << 16 | buffer[4] << 8 | buffer[5]);
                        float Irms_float = (float)(Irms_int32 / 128000.0);
                        Console.WriteLine("Value : " + Irms_float);
                        p3_irms_box.Text = Irms_float.ToString();
                    }
                }
                */
            }
        }

        

        private void cali_i1_box_TextChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                if ((!string.IsNullOrEmpty(cali_v3_box.Text)) && (!string.IsNullOrEmpty(cali_i3_box.Text)) && (!string.IsNullOrEmpty(cali_p3_box.Text))) // 如果 RichTextBox 的內容不為空
                {
                    // 更改按鈕的背景顏色
                    point3_button.BackColor = SystemColors.Control;

                    // 清除串口接收缓冲区中的数据
                    serialPort1.DiscardInBuffer();
                    receiveBuffer.Clear();  // 清除整个缓冲区，等待下一个有效数据帧

                    string text1 = cali_v3_box.Text;
                    float cali_v3_rms;
                    float.TryParse(text1, out cali_v3_rms);
                    //Console.WriteLine(cali_v3_rms);

                    cali_v3_rms = cali_v3_rms * (float)1000.0;
                    int intValue1 = Convert.ToInt32(cali_v3_rms);     // 將浮點數轉換為整數

                    byte[] byteArray1 = new byte[4];
                    byteArray1[0] = (byte)((intValue1 >> 24) & 0xFF); // 取得高位元組
                    byteArray1[1] = (byte)((intValue1 >> 16) & 0xFF); // 取得次高位元組
                    byteArray1[2] = (byte)((intValue1 >>  8) & 0xFF); // 取得次低位元組
                    byteArray1[3] = (byte)((intValue1 >>  0) & 0xFF); // 取得低位元組

                    // SumZero
                    SumZeroCalculator SZ_1 = new SumZeroCalculator();
                    byte[] data_1 = new byte[] { 0x49, 0x05, 0x21, byteArray1[0], byteArray1[1], byteArray1[2], byteArray1[3] };
                    byte SumZero_1 = SZ_1.CalculateSumZero(data_1);
                    //Console.WriteLine("SumZero Value: 0x" + SumZero_1.ToString("X2"));

                    // 设置目标为 chip_ssid_box
                    targetTextBox = "p3_vrms_box";
                    expectedLength = 15; 

                    // 向串口发送数据
                    byte[] myBytes1 = { 0x49, 0x05, 0x21, byteArray1[0], byteArray1[1], byteArray1[2], byteArray1[3], SumZero_1 }; 
                    serialPort1.Write(myBytes1, 0, myBytes1.Length); 
                    Thread.Sleep(200);  // Delay 100ms
                    //await Task.Delay(500);  // 非同步延遲

                    // Read status
                    //byte[] myBytes2 = { 0x49, 0x01, 0x85, 0x31, 0x0D }; 
                    //serialPort1.Write(myBytes2, 0, myBytes2.Length);
                    //Thread.Sleep(100);  // Delay 100ms

                    /*
                    // 从串口读取数据
                    if (serialPort1.BytesToRead >= 4)
                    {
                        byte[] buffer = new byte[4];
                        serialPort1.Read(buffer, 0, 4);

                        // Debug
                        //string result = BitConverter.ToString(buffer, 0, buffer.Length).Replace("-", " "); // 將讀取到的數據轉換為字串
                        //Console.WriteLine(result);

                        // 检查第一个字节为 0x49，第二个字节为 0x01
                        if (buffer[0] == 0x49 && buffer[1] == 0x01)
                        {
                            Console.WriteLine("Status Value: 0x" + buffer[2].ToString("X2"));

                            if (buffer[2] == 4)
                                MessageBox.Show("Sum-Zero Error ! ", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    */


                    /*******************************************************************/

                    // 清除串口接收缓冲区中的数据
                    serialPort1.DiscardInBuffer();
                    receiveBuffer.Clear();  // 清除整个缓冲区，等待下一个有效数据帧

                    string text2 = cali_i3_box.Text;
                    float cali_i3_rms;
                    float.TryParse(text2, out cali_i3_rms);
                    //Console.WriteLine(cali_i1_vrms);

                    cali_i3_rms = cali_i3_rms * (float)128000.0;
                    int intValue2 = Convert.ToInt32(cali_i3_rms);     // 將浮點數轉換為整數

                    byte[] byteArray2 = new byte[4];
                    byteArray2[0] = (byte)((intValue2 >> 24) & 0xFF); // 取得高位元組
                    byteArray2[1] = (byte)((intValue2 >> 16) & 0xFF); // 取得次高位元組
                    byteArray2[2] = (byte)((intValue2 >>  8) & 0xFF); // 取得次低位元組
                    byteArray2[3] = (byte)((intValue2 >>  0) & 0xFF); // 取得低位元組

                    // SumZero
                    SumZeroCalculator SZ_2 = new SumZeroCalculator();
                    byte[] data_2 = new byte[] { 0x49, 0x05, 0x23, byteArray2[0], byteArray2[1], byteArray2[2], byteArray2[3] };
                    byte SumZero_2 = SZ_2.CalculateSumZero(data_2);
                    //Console.WriteLine("SumZero Value: 0x" + SumZero_2.ToString("X2"));

                    // 向串口发送数据
                    byte[] myBytes3 = { 0x49, 0x05, 0x23, byteArray2[0], byteArray2[1], byteArray2[2], byteArray2[3], SumZero_2 }; 
                    serialPort1.Write(myBytes3, 0, myBytes3.Length);
                    Thread.Sleep(200);  // Delay 100ms
                    //await Task.Delay(500);  // 非同步延遲

                    // Read status
                    //byte[] myBytes4 = { 0x49, 0x01, 0x85, 0x31, 0x0D }; 
                    //serialPort1.Write(myBytes4, 0, myBytes4.Length); 
                    //Thread.Sleep(100);  // Delay 100ms

                    /*
                    // 从串口读取数据
                    if (serialPort1.BytesToRead >= 4)
                    {
                        byte[] buffer = new byte[4];
                        serialPort1.Read(buffer, 0, 4);

                        // Debug
                        //string result = BitConverter.ToString(buffer, 0, buffer.Length).Replace("-", " "); // 將讀取到的數據轉換為字串
                        //Console.WriteLine(result);

                        // 检查第一个字节为 0x49，第二个字节为 0x01
                        if (buffer[0] == 0x49 && buffer[1] == 0x01)
                        {
                            Console.WriteLine("Status Value: 0x" + buffer[2].ToString("X2"));

                            if (buffer[2] == 4)
                                MessageBox.Show("Sum-Zero Error ! ", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    */

                    /*******************************************************************/

                    // 清除串口接收缓冲区中的数据
                    serialPort1.DiscardInBuffer();
                    receiveBuffer.Clear();  // 清除整个缓冲区，等待下一个有效数据帧

                    string text3 = cali_p3_box.Text;
                    float cali_p3_rms;
                    float.TryParse(text3, out cali_p3_rms);
                    //Console.WriteLine(cali_p3_rms);

                    cali_p3_rms = cali_p3_rms * (float)1000.0;
                    int intValue3 = Convert.ToInt32(cali_p3_rms);     // 將浮點數轉換為整數

                    byte[] byteArray3 = new byte[4];
                    byteArray3[0] = (byte)((intValue3 >> 24) & 0xFF); // 取得高位元組
                    byteArray3[1] = (byte)((intValue3 >> 16) & 0xFF); // 取得次高位元組
                    byteArray3[2] = (byte)((intValue3 >>  8) & 0xFF); // 取得次低位元組
                    byteArray3[3] = (byte)((intValue3 >>  0) & 0xFF); // 取得低位元組

                    // SumZero
                    SumZeroCalculator SZ_3 = new SumZeroCalculator();
                    byte[] data_3 = new byte[] { 0x49, 0x05, 0x25, byteArray3[0], byteArray3[1], byteArray3[2], byteArray3[3] };
                    byte SumZero_3 = SZ_3.CalculateSumZero(data_3);
                    //Console.WriteLine("SumZero Value: 0x" + SumZero_3.ToString("X2"));

                    // 向串口发送数据
                    byte[] myBytes5 = { 0x49, 0x05, 0x25, byteArray3[0], byteArray3[1], byteArray3[2], byteArray3[3], SumZero_3 }; 
                    serialPort1.Write(myBytes5, 0, myBytes5.Length);
                    Thread.Sleep(200);  // Delay 100ms
                    //await Task.Delay(500);  // 非同步延遲

                    // Read status
                    //byte[] myBytes6 = { 0x49, 0x01, 0x85, 0x31, 0x0D }; 
                    //serialPort1.Write(myBytes6, 0, myBytes6.Length); 
                    //Thread.Sleep(100);  // Delay 100ms

                    /*
                    // 从串口读取数据
                    if (serialPort1.BytesToRead >= 4)
                    {
                        byte[] buffer = new byte[4];
                        serialPort1.Read(buffer, 0, 4);

                        // Debug
                        //string result = BitConverter.ToString(buffer, 0, buffer.Length).Replace("-", " "); // 將讀取到的數據轉換為字串
                        //Console.WriteLine(result);

                        // 检查第一个字节为 0x49，第二个字节为 0x01
                        if (buffer[0] == 0x49 && buffer[1] == 0x01)
                        {
                            Console.WriteLine("Status Value: 0x" + buffer[2].ToString("X2"));

                            if (buffer[2] == 4)
                                MessageBox.Show("Sum-Zero Error ! ", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    */

                    /*******************************************************************/

                    // 清除串口接收缓冲区中的数据
                    serialPort1.DiscardInBuffer();
                    receiveBuffer.Clear();  // 清除整个缓冲区，等待下一个有效数据帧

                    // 向串口发送数据
                    byte[] myBytes7 = { 0x49, 0x01, 0x27, 0x8F }; 
                    serialPort1.Write(myBytes7, 0, myBytes7.Length);
                    Thread.Sleep(4000);  // Delay 4000ms
                    //await Task.Delay(3000);  // 非同步延遲

                    /*
                    // 从串口读取数据
                    if (serialPort1.BytesToRead >= 15)
                    {
                        byte[] buffer = new byte[15];
                        serialPort1.Read(buffer, 0, 15);

                        // Debug
                        string result3 = BitConverter.ToString(buffer, 0, buffer.Length).Replace("-", " "); // 將讀取到的數據轉換為字串
                        Console.WriteLine(result3);

                        // 在這裡處理讀取到的資料
                        if ((buffer[0] == 0x49) && (buffer[1] == 0x0C))
                        {
                            uint Vrms_int32 = (uint)(buffer[2] << 24 | buffer[3] << 16 | buffer[4] << 8 | buffer[5]);
                            float Vrms_float = (float)(Vrms_int32 / 1000.0);
                            Console.WriteLine("Vrms : " + Vrms_float);
                            p3_vrms_box.Text = Vrms_float.ToString();

                            uint Irms_int32 = (uint)(buffer[6] << 24 | buffer[7] << 16 | buffer[8] << 8 | buffer[9]);
                            float Irms_float = (float)(Irms_int32 / 128000.0);
                            Console.WriteLine("Irms : " + Irms_float);
                            p3_irms_box.Text = Irms_float.ToString();

                            uint Pavg_int32 = (uint)(buffer[10] << 24 | buffer[11] << 16 | buffer[12] << 8 | buffer[13]);
                            float Pavg_float = (float)(Pavg_int32 / 1000.0);
                            Console.WriteLine("Pavg : " + Pavg_float);
                            p3_pavg_box.Text = Pavg_float.ToString();
                        }
                    }
                    */

                    /*******************************************************************/

                    // 清除串口接收缓冲区中的数据
                    serialPort1.DiscardInBuffer();
                    receiveBuffer.Clear();  // 清除整个缓冲区，等待下一个有效数据帧

                    // 向串口发送数据
                    byte[] myBytes8 = { 0x49, 0x01, 0x29, 0x8D }; 
                    serialPort1.Write(myBytes8, 0, myBytes8.Length);
                    Thread.Sleep(200);  // Delay 100ms
                    //await Task.Delay(500);  // 非同步延遲

                    // 恢复19字节接收模式
                    expectedLength = 19;

                    // 更改按鈕的背景顏色
                    point3_button.BackColor = Color.LightBlue;

                    // 發送開始命令
                    //byte[] myBytes14 = { 0x49, 0x02, 0x81, 0x7F, 0xB5, 0x0D };
                    //serialPort1.Write(myBytes14, 0, myBytes14.Length); // 將 6 個 Byte 數據寫出去
                    //Thread.Sleep(100);  // Delay 100ms

                    // Read status
                    //byte[] myBytes9 = { 0x49, 0x01, 0x85, 0x31, 0x0D }; 
                    //serialPort1.Write(myBytes9, 0, myBytes9.Length); 
                    //Thread.Sleep(100);  // Delay 100ms

                    /*
                    // 从串口读取数据
                    if (serialPort1.BytesToRead >= 4)
                    {
                        byte[] buffer = new byte[4];
                        serialPort1.Read(buffer, 0, 4);

                        // Debug
                        //string result = BitConverter.ToString(buffer, 0, buffer.Length).Replace("-", " "); // 將讀取到的數據轉換為字串
                        //Console.WriteLine(result);

                        // 检查第一个字节为 0x49，第二个字节为 0x01
                        if (buffer[0] == 0x49 && buffer[1] == 0x01)
                        {
                            Console.WriteLine("Status Value: 0x" + buffer[2].ToString("X2"));

                            if (buffer[2] == 4)
                                MessageBox.Show("Sum-Zero Error ! ", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    */
                }
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void vth_order_box_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                // 清空 RichTextBox 控制元件內容
                p1_pavg_box.Clear();

                // 清除串口接收缓冲区中的数据
                serialPort1.DiscardInBuffer();
                receiveBuffer.Clear();  // 清除整个缓冲区，等待下一个有效数据帧

                // 發送停止命令 
                byte[] myBytes1 = { 0x49, 0x01, 0x83, 0x33 };
                serialPort1.Write(myBytes1, 0, myBytes1.Length);
                Thread.Sleep(200);  // Delay 100ms
                //await Task.Delay(200);  // 非同步延遲

                // 设置目标为 p1_pavg_box
                targetTextBox = "p1_pavg_box";
                expectedLength = 7;

                // 向串口发送数据
                byte[] myBytes = { 0x49, 0x01, 0x44, 0x72 }; 
                serialPort1.Write(myBytes, 0, myBytes.Length);
                Thread.Sleep(200);  // Delay 100ms
                //await Task.Delay(200);  // 非同步延遲

                /*
                if (serialPort1.BytesToRead >= 7)
                {
                    byte[] buffer = new byte[7];
                    serialPort1.Read(buffer, 0, 7);

                    // Debug
                    string result = BitConverter.ToString(buffer, 0, buffer.Length).Replace("-", " "); // 將讀取到的數據轉換為字串
                    Console.WriteLine(result);

                    // 检查第一个字节为 0x49，第二个字节为 0x04
                    if (buffer[0] == 0x49 && buffer[1] == 0x04)
                    {
                        uint Pavg_int32 = (uint)(buffer[2] << 24 | buffer[3] << 16 | buffer[4] << 8 | buffer[5]);
                        float Pavg_float = (float)(Pavg_int32 / 1000.0);
                        Console.WriteLine("Value : " + Pavg_float);
                        p1_pavg_box.Text = Pavg_float.ToString();
                    }
                }
                */
            }
        }

        private void button3_Click_2(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                // 清空 RichTextBox 控制元件內容
                p4_pavg_box.Clear();

                // 清除串口接收缓冲区中的数据
                serialPort1.DiscardInBuffer();
                receiveBuffer.Clear();  // 清除整个缓冲区，等待下一个有效数据帧

                // 發送停止命令 
                byte[] myBytes1 = { 0x49, 0x01, 0x83, 0x33 };
                serialPort1.Write(myBytes1, 0, myBytes1.Length);
                Thread.Sleep(200);  // Delay 100ms
                //await Task.Delay(200);  // 非同步延遲

                // 设置目标为 p4_pavg_box
                targetTextBox = "p4_pavg_box";
                expectedLength = 7;

                // 向串口发送数据
                byte[] myBytes = { 0x49, 0x01, 0x54, 0x62 }; 
                serialPort1.Write(myBytes, 0, myBytes.Length);
                Thread.Sleep(200);  // Delay 100ms
                //await Task.Delay(200);  // 非同步延遲

                /*
                if (serialPort1.BytesToRead >= 7)
                {
                    byte[] buffer = new byte[7];
                    serialPort1.Read(buffer, 0, 7);

                    // Debug
                    string result = BitConverter.ToString(buffer, 0, buffer.Length).Replace("-", " "); // 將讀取到的數據轉換為字串
                    Console.WriteLine(result);

                    // 检查第一个字节为 0x49，第二个字节为 0x04
                    if (buffer[0] == 0x49 && buffer[1] == 0x04)
                    {
                        uint Pavg_int32 = (uint)(buffer[2] << 24 | buffer[3] << 16 | buffer[4] << 8 | buffer[5]);
                        float Pavg_float = (float)(Pavg_int32 / 1000.0);
                        Console.WriteLine("Value : " + Pavg_float);
                        p4_pavg_box.Text = Pavg_float.ToString();
                    }
                }
                */
            }
        }

        private void point4_button_Click(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                if ((!string.IsNullOrEmpty(cali_v4_box.Text)) && (!string.IsNullOrEmpty(cali_i4_box.Text)) && (!string.IsNullOrEmpty(cali_p4_box.Text))) // 如果 RichTextBox 的內容不為空
                {
                    // 更改按鈕的背景顏色
                    point4_button.BackColor = SystemColors.Control;

                    // 清除串口接收缓冲区中的数据
                    serialPort1.DiscardInBuffer();
                    receiveBuffer.Clear();  // 清除整个缓冲区，等待下一个有效数据帧

                    // 讀取text內容
                    string text1 = cali_v4_box.Text;
                    float cali_v4_rms;
                    float.TryParse(text1, out cali_v4_rms);
                    // Console.WriteLine(cali_v4_rms);

                    cali_v4_rms = cali_v4_rms * (float)1000.0;
                    int intValue1 = Convert.ToInt32(cali_v4_rms);     // 將浮點數轉換為整數

                    byte[] byteArray1 = new byte[4];
                    byteArray1[0] = (byte)((intValue1 >> 24) & 0xFF); // 取得高位元組
                    byteArray1[1] = (byte)((intValue1 >> 16) & 0xFF); // 取得次高位元組
                    byteArray1[2] = (byte)((intValue1 >> 8) & 0xFF);  // 取得次低位元組
                    byteArray1[3] = (byte)((intValue1 >> 0) & 0xFF);  // 取得低位元組

                    // SumZero
                    SumZeroCalculator SZ_1 = new SumZeroCalculator();
                    byte[] data_1 = new byte[] { 0x49, 0x05, 0x51, byteArray1[0], byteArray1[1], byteArray1[2], byteArray1[3] };
                    byte SumZero_1 = SZ_1.CalculateSumZero(data_1);
                    //Console.WriteLine("SumZero Value: 0x" + SumZero_1.ToString("X2"));                                        

                    // 设置目标为 chip_ssid_box
                    targetTextBox = "p4_vrms_box";
                    expectedLength = 15;  

                    // 向串口发送数据 
                    byte[] myBytes1 = { 0x49, 0x05, 0x51, byteArray1[0], byteArray1[1], byteArray1[2], byteArray1[3], SumZero_1 };
                    serialPort1.Write(myBytes1, 0, myBytes1.Length);
                    Thread.Sleep(200);  // Delay 100ms
                    //await Task.Delay(500);  // 非同步延遲

                    // Read status
                    //byte[] myBytes2 = { 0x49, 0x01, 0x85, 0x31, 0x0D };
                    //serialPort1.Write(myBytes2, 0, myBytes2.Length);
                    //Thread.Sleep(100);  // Delay 100ms

                    /*
                    // 从串口读取数据
                    if (serialPort1.BytesToRead >= 4)
                    {
                        byte[] buffer = new byte[4];
                        serialPort1.Read(buffer, 0, 4);

                        // Debug
                        //string result = BitConverter.ToString(buffer, 0, buffer.Length).Replace("-", " "); // 將讀取到的數據轉換為字串
                        //Console.WriteLine(result);

                        // 检查第一个字节为 0x49，第二个字节为 0x01
                        if (buffer[0] == 0x49 && buffer[1] == 0x01)
                        {
                            Console.WriteLine("Status Value: 0x" + buffer[2].ToString("X2"));

                            if (buffer[2] == 4)
                                MessageBox.Show("Sum-Zero Error ! ", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    */

                    /*******************************************************************/

                    // 清除串口接收缓冲区中的数据
                    serialPort1.DiscardInBuffer();
                    receiveBuffer.Clear();  // 清除整个缓冲区，等待下一个有效数据帧

                    string text2 = cali_i4_box.Text;
                    float cali_i4_rms;
                    float.TryParse(text2, out cali_i4_rms);
                    //Console.WriteLine(cali_i4_vrms);

                    cali_i4_rms = cali_i4_rms * (float)128000.0;
                    int intValue2 = Convert.ToInt32(cali_i4_rms);     // 將浮點數轉換為整數

                    byte[] byteArray2 = new byte[4];
                    byteArray2[0] = (byte)((intValue2 >> 24) & 0xFF); // 取得高位元組
                    byteArray2[1] = (byte)((intValue2 >> 16) & 0xFF); // 取得次高位元組
                    byteArray2[2] = (byte)((intValue2 >> 8) & 0xFF);  // 取得次低位元組
                    byteArray2[3] = (byte)((intValue2 >> 0) & 0xFF);  // 取得低位元組

                    // SumZero
                    SumZeroCalculator SZ_2 = new SumZeroCalculator();
                    byte[] data_2 = new byte[] { 0x49, 0x05, 0x53, byteArray2[0], byteArray2[1], byteArray2[2], byteArray2[3] };
                    byte SumZero_2 = SZ_2.CalculateSumZero(data_2);
                    //Console.WriteLine("SumZero Value: 0x" + SumZero_2.ToString("X2"));

                    // 向串口发送数据 
                    byte[] myBytes3 = { 0x49, 0x05, 0x53, byteArray2[0], byteArray2[1], byteArray2[2], byteArray2[3], SumZero_2 };
                    serialPort1.Write(myBytes3, 0, myBytes3.Length);
                    Thread.Sleep(200);  // Delay 100ms
                    //await Task.Delay(500);  // 非同步延遲

                    // Read status
                    //byte[] myBytes4 = { 0x49, 0x01, 0x85, 0x31, 0x0D };
                    //serialPort1.Write(myBytes4, 0, myBytes4.Length);
                    //Thread.Sleep(100);  // Delay 100ms

                    /*
                    // 从串口读取数据
                    if (serialPort1.BytesToRead >= 4)
                    {
                        byte[] buffer = new byte[4];
                        serialPort1.Read(buffer, 0, 4);

                        // Debug
                        //string result = BitConverter.ToString(buffer, 0, buffer.Length).Replace("-", " "); // 將讀取到的數據轉換為字串
                        //Console.WriteLine(result);

                        // 检查第一个字节为 0x49，第二个字节为 0x01
                        if (buffer[0] == 0x49 && buffer[1] == 0x01)
                        {
                            Console.WriteLine("Status Value: 0x" + buffer[2].ToString("X2"));

                            if (buffer[2] == 4)
                                MessageBox.Show("Sum-Zero Error ! ", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    */

                    /*******************************************************************/

                    // 清除串口接收缓冲区中的数据
                    serialPort1.DiscardInBuffer();
                    receiveBuffer.Clear();  // 清除整个缓冲区，等待下一个有效数据帧

                    string text3 = cali_p4_box.Text;
                    float cali_p4_rms;
                    float.TryParse(text3, out cali_p4_rms);
                    //Console.WriteLine(cali_p4_rms);

                    cali_p4_rms = cali_p4_rms * (float)1000.0;
                    int intValue3 = Convert.ToInt32(cali_p4_rms);     // 將浮點數轉換為整數

                    byte[] byteArray3 = new byte[4];
                    byteArray3[0] = (byte)((intValue3 >> 24) & 0xFF); // 取得高位元組
                    byteArray3[1] = (byte)((intValue3 >> 16) & 0xFF); // 取得次高位元組
                    byteArray3[2] = (byte)((intValue3 >> 8) & 0xFF);  // 取得次低位元組
                    byteArray3[3] = (byte)((intValue3 >> 0) & 0xFF);  // 取得低位元組

                    // SumZero
                    SumZeroCalculator SZ_3 = new SumZeroCalculator();
                    byte[] data_3 = new byte[] { 0x49, 0x05, 0x55, byteArray3[0], byteArray3[1], byteArray3[2], byteArray3[3] };
                    byte SumZero_3 = SZ_3.CalculateSumZero(data_3);
                    //Console.WriteLine("SumZero Value: 0x" + SumZero_3.ToString("X2"));

                    // 向串口发送数据 
                    byte[] myBytes5 = { 0x49, 0x05, 0x55, byteArray3[0], byteArray3[1], byteArray3[2], byteArray3[3], SumZero_3 };
                    serialPort1.Write(myBytes5, 0, myBytes5.Length);
                    Thread.Sleep(200);  // Delay 100ms
                    //await Task.Delay(500);  // 非同步延遲

                    // Read status
                    //byte[] myBytes6 = { 0x49, 0x01, 0x85, 0x31, 0x0D };
                    //serialPort1.Write(myBytes6, 0, myBytes6.Length);
                    //Thread.Sleep(100);  // Delay 100ms

                    /*
                    // 从串口读取数据
                    if (serialPort1.BytesToRead >= 4)
                    {
                        byte[] buffer = new byte[4];
                        serialPort1.Read(buffer, 0, 4);

                        // Debug
                        //string result = BitConverter.ToString(buffer, 0, buffer.Length).Replace("-", " "); // 將讀取到的數據轉換為字串
                        //Console.WriteLine(result);

                        // 检查第一个字节为 0x49，第二个字节为 0x01
                        if (buffer[0] == 0x49 && buffer[1] == 0x01)
                        {
                            Console.WriteLine("Status Value: 0x" + buffer[2].ToString("X2"));

                            if (buffer[2] == 4)
                                MessageBox.Show("Sum-Zero Error ! ", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    */

                    /*******************************************************************/

                    // 清除串口接收缓冲区中的数据
                    serialPort1.DiscardInBuffer();
                    receiveBuffer.Clear();  // 清除整个缓冲区，等待下一个有效数据帧

                    // 向串口发送数据 
                    byte[] myBytes7 = { 0x49, 0x01, 0x57, 0x5F };
                    serialPort1.Write(myBytes7, 0, myBytes7.Length);
                    Thread.Sleep(4000);  // Delay 4000ms
                    //await Task.Delay(3000);  // 非同步延遲

                    /*
                    // 从串口读取数据
                    if (serialPort1.BytesToRead >= 15)
                    {
                        byte[] buffer = new byte[15];
                        serialPort1.Read(buffer, 0, 15);

                        // Debug
                        string result3 = BitConverter.ToString(buffer, 0, buffer.Length).Replace("-", " "); // 將讀取到的數據轉換為字串
                        Console.WriteLine(result3);

                        // 在這裡處理讀取到的資料
                        if ((buffer[0] == 0x49) && (buffer[1] == 0x0C))
                        {
                            uint Vrms_int32 = (uint)(buffer[2] << 24 | buffer[3] << 16 | buffer[4] << 8 | buffer[5]);
                            float Vrms_float = (float)(Vrms_int32 / 1000.0);
                            Console.WriteLine("Vrms : " + Vrms_float);
                            p4_vrms_box.Text = Vrms_float.ToString();

                            uint Irms_int32 = (uint)(buffer[6] << 24 | buffer[7] << 16 | buffer[8] << 8 | buffer[9]);
                            float Irms_float = (float)(Irms_int32 / 128000.0);
                            Console.WriteLine("Irms : " + Irms_float);
                            p4_irms_box.Text = Irms_float.ToString();

                            uint Pavg_int32 = (uint)(buffer[10] << 24 | buffer[11] << 16 | buffer[12] << 8 | buffer[13]);
                            float Pavg_float = (float)(Pavg_int32 / 1000.0);
                            Console.WriteLine("Pavg : " + Pavg_float);
                            p4_pavg_box.Text = Pavg_float.ToString();
                        }
                    }
                    */

                    /*******************************************************************/

                    // 清除串口接收缓冲区中的数据
                    serialPort1.DiscardInBuffer();
                    receiveBuffer.Clear();  // 清除整个缓冲区，等待下一个有效数据帧

                    // 向串口发送数据 
                    byte[] myBytes8 = { 0x49, 0x01, 0x59, 0x5D };
                    serialPort1.Write(myBytes8, 0, myBytes8.Length);
                    Thread.Sleep(200);  // Delay 100ms
                    //await Task.Delay(500);  // 非同步延遲

                    // 恢复19字节接收模式
                    expectedLength = 19;

                    // 更改按鈕的背景顏色
                    point4_button.BackColor = Color.LightBlue;

                    // Read status
                    //byte[] myBytes9 = { 0x49, 0x01, 0x85, 0x31, 0x0D };
                    //serialPort1.Write(myBytes9, 0, myBytes9.Length);
                    //Thread.Sleep(100);  // Delay 100ms

                    /*
                    // 从串口读取数据
                    if (serialPort1.BytesToRead >= 4)
                    {
                        byte[] buffer = new byte[4];
                        serialPort1.Read(buffer, 0, 4);

                        // Debug
                        //string result = BitConverter.ToString(buffer, 0, buffer.Length).Replace("-", " "); // 將讀取到的數據轉換為字串
                        //Console.WriteLine(result);

                        // 检查第一个字节为 0x49，第二个字节为 0x01
                        if (buffer[0] == 0x49 && buffer[1] == 0x01)
                        {
                            Console.WriteLine("Status Value: 0x" + buffer[2].ToString("X2"));

                            if (buffer[2] == 4)
                                MessageBox.Show("Sum-Zero Error ! ", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    */
                }
            }
        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click_2(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                if ((!string.IsNullOrEmpty(cali_v2_box.Text)) && (!string.IsNullOrEmpty(cali_i2_box.Text)) && (!string.IsNullOrEmpty(cali_p2_box.Text))) // 如果 RichTextBox 的內容不為空
                {
                    // 更改按鈕的背景顏色
                    point2_button.BackColor = SystemColors.Control;

                    // 清除串口接收缓冲区中的数据
                    serialPort1.DiscardInBuffer();
                    receiveBuffer.Clear();  // 清除整个缓冲区，等待下一个有效数据帧

                    string text1 = cali_v2_box.Text;
                    float cali_v2_rms;
                    float.TryParse(text1, out cali_v2_rms);
                    //Console.WriteLine(cali_v2_rms);

                    cali_v2_rms = cali_v2_rms * (float)1000.0;
                    int intValue1 = Convert.ToInt32(cali_v2_rms);     // 將浮點數轉換為整數

                    byte[] byteArray1 = new byte[4];
                    byteArray1[0] = (byte)((intValue1 >> 24) & 0xFF); // 取得高位元組
                    byteArray1[1] = (byte)((intValue1 >> 16) & 0xFF); // 取得次高位元組
                    byteArray1[2] = (byte)((intValue1 >> 8) & 0xFF); // 取得次低位元組
                    byteArray1[3] = (byte)((intValue1 >> 0) & 0xFF); // 取得低位元組

                    // SumZero
                    SumZeroCalculator SZ_1 = new SumZeroCalculator();
                    byte[] data_1 = new byte[] { 0x49, 0x05, 0x31, byteArray1[0], byteArray1[1], byteArray1[2], byteArray1[3] };
                    byte SumZero_1 = SZ_1.CalculateSumZero(data_1);
                    //Console.WriteLine("SumZero Value: 0x" + SumZero_1.ToString("X2"));

                    // 设置目标为 chip_ssid_box
                    //targetTextBox = "p2_vrms_box";
                    //expectedLength = 15;

                    // 向串口发送数据
                    byte[] myBytes1 = { 0x49, 0x05, 0x31, byteArray1[0], byteArray1[1], byteArray1[2], byteArray1[3], SumZero_1 };
                    serialPort1.Write(myBytes1, 0, myBytes1.Length);
                    Thread.Sleep(200);  // Delay 100ms
                    //await Task.Delay(500);  // 非同步延遲

                    // Read status
                    //byte[] myBytes2 = { 0x49, 0x01, 0x85, 0x31, 0x0D };
                    //serialPort1.Write(myBytes2, 0, myBytes2.Length);
                    //Thread.Sleep(100);  // Delay 100ms

                    /*
                    // 从串口读取数据
                    if (serialPort1.BytesToRead >= 4)
                    {
                        byte[] buffer = new byte[4];
                        serialPort1.Read(buffer, 0, 4);

                        // Debug
                        //string result = BitConverter.ToString(buffer, 0, buffer.Length).Replace("-", " "); // 將讀取到的數據轉換為字串
                        //Console.WriteLine(result);

                        // 检查第一个字节为 0x49，第二个字节为 0x01
                        if (buffer[0] == 0x49 && buffer[1] == 0x01)
                        {
                            Console.WriteLine("Status Value: 0x" + buffer[2].ToString("X2"));

                            if (buffer[2] == 4)
                                MessageBox.Show("Sum-Zero Error ! ", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    */

                    /*******************************************************************/

                    // 清除串口接收缓冲区中的数据
                    serialPort1.DiscardInBuffer();
                    receiveBuffer.Clear();  // 清除整个缓冲区，等待下一个有效数据帧

                    string text2 = cali_i2_box.Text;
                    float cali_i2_rms;
                    float.TryParse(text2, out cali_i2_rms);
                    //Console.WriteLine(cali_i2_vrms);

                    cali_i2_rms = cali_i2_rms * (float)128000.0;
                    int intValue2 = Convert.ToInt32(cali_i2_rms);     // 將浮點數轉換為整數

                    byte[] byteArray2 = new byte[4];
                    byteArray2[0] = (byte)((intValue2 >> 24) & 0xFF); // 取得高位元組
                    byteArray2[1] = (byte)((intValue2 >> 16) & 0xFF); // 取得次高位元組
                    byteArray2[2] = (byte)((intValue2 >> 8) & 0xFF); // 取得次低位元組
                    byteArray2[3] = (byte)((intValue2 >> 0) & 0xFF); // 取得低位元組

                    // SumZero
                    SumZeroCalculator SZ_2 = new SumZeroCalculator();
                    byte[] data_2 = new byte[] { 0x49, 0x05, 0x33, byteArray2[0], byteArray2[1], byteArray2[2], byteArray2[3] };
                    byte SumZero_2 = SZ_2.CalculateSumZero(data_2);
                    //Console.WriteLine("SumZero Value: 0x" + SumZero_2.ToString("X2"));

                    // 向串口发送数据
                    byte[] myBytes3 = { 0x49, 0x05, 0x33, byteArray2[0], byteArray2[1], byteArray2[2], byteArray2[3], SumZero_2 };
                    serialPort1.Write(myBytes3, 0, myBytes3.Length);
                    Thread.Sleep(200);  // Delay 100ms
                    //await Task.Delay(500);  // 非同步延遲

                    // Read status
                    //byte[] myBytes4 = { 0x49, 0x01, 0x85, 0x31, 0x0D };
                    //serialPort1.Write(myBytes4, 0, myBytes4.Length);
                    //Thread.Sleep(100);  // Delay 100ms

                    /*
                    // 从串口读取数据
                    if (serialPort1.BytesToRead >= 4)
                    {
                        byte[] buffer = new byte[4];
                        serialPort1.Read(buffer, 0, 4);

                        // Debug
                        //string result = BitConverter.ToString(buffer, 0, buffer.Length).Replace("-", " "); // 將讀取到的數據轉換為字串
                        //Console.WriteLine(result);

                        // 检查第一个字节为 0x49，第二个字节为 0x01
                        if (buffer[0] == 0x49 && buffer[1] == 0x01)
                        {
                            Console.WriteLine("Status Value: 0x" + buffer[2].ToString("X2"));

                            if (buffer[2] == 4)
                                MessageBox.Show("Sum-Zero Error ! ", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    */

                    /*******************************************************************/

                    // 清除串口接收缓冲区中的数据
                    serialPort1.DiscardInBuffer();
                    receiveBuffer.Clear();  // 清除整个缓冲区，等待下一个有效数据帧

                    string text3 = cali_p2_box.Text;
                    float cali_p2_rms;
                    float.TryParse(text3, out cali_p2_rms);
                    //Console.WriteLine(cali_p2_rms);

                    cali_p2_rms = cali_p2_rms * (float)1000.0;
                    int intValue3 = Convert.ToInt32(cali_p2_rms);     // 將浮點數轉換為整數

                    byte[] byteArray3 = new byte[4];
                    byteArray3[0] = (byte)((intValue3 >> 24) & 0xFF); // 取得高位元組
                    byteArray3[1] = (byte)((intValue3 >> 16) & 0xFF); // 取得次高位元組
                    byteArray3[2] = (byte)((intValue3 >> 8) & 0xFF); // 取得次低位元組
                    byteArray3[3] = (byte)((intValue3 >> 0) & 0xFF); // 取得低位元組

                    // SumZero
                    SumZeroCalculator SZ_3 = new SumZeroCalculator();
                    byte[] data_3 = new byte[] { 0x49, 0x05, 0x35, byteArray3[0], byteArray3[1], byteArray3[2], byteArray3[3] };
                    byte SumZero_3 = SZ_3.CalculateSumZero(data_3);
                    //Console.WriteLine("SumZero Value: 0x" + SumZero_3.ToString("X2"));

                    // 向串口发送数据
                    byte[] myBytes5 = { 0x49, 0x05, 0x35, byteArray3[0], byteArray3[1], byteArray3[2], byteArray3[3], SumZero_3 };
                    serialPort1.Write(myBytes5, 0, myBytes5.Length);
                    Thread.Sleep(200);  // Delay 100ms
                    //await Task.Delay(500);  // 非同步延遲

                    // Read status
                    //byte[] myBytes6 = { 0x49, 0x01, 0x85, 0x31, 0x0D };
                    //serialPort1.Write(myBytes6, 0, myBytes6.Length);
                    //Thread.Sleep(100);  // Delay 100ms

                    /*
                    // 从串口读取数据
                    if (serialPort1.BytesToRead >= 4)
                    {
                        byte[] buffer = new byte[4];
                        serialPort1.Read(buffer, 0, 4);

                        // Debug
                        //string result = BitConverter.ToString(buffer, 0, buffer.Length).Replace("-", " "); // 將讀取到的數據轉換為字串
                        //Console.WriteLine(result);

                        // 检查第一个字节为 0x49，第二个字节为 0x01
                        if (buffer[0] == 0x49 && buffer[1] == 0x01)
                        {
                            Console.WriteLine("Status Value: 0x" + buffer[2].ToString("X2"));

                            if (buffer[2] == 4)
                                MessageBox.Show("Sum-Zero Error ! ", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    */

                    /*******************************************************************/

                    // 清除串口接收缓冲区中的数据
                    serialPort1.DiscardInBuffer();
                    receiveBuffer.Clear();  // 清除整个缓冲区，等待下一个有效数据帧

                    // 向串口发送数据
                    byte[] myBytes7 = { 0x49, 0x01, 0x37, 0x7F };
                    serialPort1.Write(myBytes7, 0, myBytes7.Length);
                    Thread.Sleep(4000);  // Delay 4000ms
                    //await Task.Delay(3000);  // 非同步延遲

                    /*
                    // 从串口读取数据
                    if (serialPort1.BytesToRead >= 15)
                    {
                        byte[] buffer = new byte[15];
                        serialPort1.Read(buffer, 0, 15);

                        // Debug
                        string result3 = BitConverter.ToString(buffer, 0, buffer.Length).Replace("-", " "); // 將讀取到的數據轉換為字串
                        Console.WriteLine(result3);

                        // 在這裡處理讀取到的資料
                        if ((buffer[0] == 0x49) && (buffer[1] == 0x0C))
                        {
                            uint Vrms_int32 = (uint)(buffer[2] << 24 | buffer[3] << 16 | buffer[4] << 8 | buffer[5]);
                            float Vrms_float = (float)(Vrms_int32 / 1000.0);
                            Console.WriteLine("Vrms : " + Vrms_float);
                            p2_vrms_box.Text = Vrms_float.ToString();

                            uint Irms_int32 = (uint)(buffer[6] << 24 | buffer[7] << 16 | buffer[8] << 8 | buffer[9]);
                            float Irms_float = (float)(Irms_int32 / 128000.0);
                            Console.WriteLine("Irms : " + Irms_float);
                            p2_irms_box.Text = Irms_float.ToString();

                            uint Pavg_int32 = (uint)(buffer[10] << 24 | buffer[11] << 16 | buffer[12] << 8 | buffer[13]);
                            float Pavg_float = (float)(Pavg_int32 / 1000.0);
                            Console.WriteLine("Pavg : " + Pavg_float);
                            p2_pavg_box.Text = Pavg_float.ToString();
                        }
                    }
                    */

                    /*******************************************************************/

                    // 清除串口接收缓冲区中的数据
                    serialPort1.DiscardInBuffer();
                    receiveBuffer.Clear();  // 清除整个缓冲区，等待下一个有效数据帧

                    // 向串口发送数据
                    byte[] myBytes8 = { 0x49, 0x01, 0x39, 0x7D };
                    serialPort1.Write(myBytes8, 0, myBytes8.Length);
                    Thread.Sleep(200);  // Delay 100ms
                    //await Task.Delay(500);  // 非同步延遲

                    // 恢复19字节接收模式
                    expectedLength = 19;

                    // 更改按鈕的背景顏色
                    point2_button.BackColor = Color.LightBlue;
                  
                    // Read status
                    //byte[] myBytes9 = { 0x49, 0x01, 0x85, 0x31, 0x0D };
                    //serialPort1.Write(myBytes9, 0, myBytes9.Length);
                    //Thread.Sleep(100);  // Delay 100ms

                    /*
                    // 从串口读取数据
                    if (serialPort1.BytesToRead >= 4)
                    {
                        byte[] buffer = new byte[4];
                        serialPort1.Read(buffer, 0, 4);

                        // Debug
                        //string result = BitConverter.ToString(buffer, 0, buffer.Length).Replace("-", " "); // 將讀取到的數據轉換為字串
                        //Console.WriteLine(result);

                        // 检查第一个字节为 0x49，第二个字节为 0x01
                        if (buffer[0] == 0x49 && buffer[1] == 0x01)
                        {
                            Console.WriteLine("Status Value: 0x" + buffer[2].ToString("X2"));

                            if (buffer[2] == 4)
                                MessageBox.Show("Sum-Zero Error ! ", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    */
                }
            }
        }

        private void button6_Click_1(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                // 清空 RichTextBox 控制元件內容
                p2_vrms_box.Clear();

                // 清除串口接收缓冲区中的数据
                serialPort1.DiscardInBuffer();
                receiveBuffer.Clear();  // 清除整个缓冲区，等待下一个有效数据帧

                // 發送停止命令 
                byte[] myBytes1 = { 0x49, 0x01, 0x83, 0x33 };
                serialPort1.Write(myBytes1, 0, myBytes1.Length);
                Thread.Sleep(200);  // Delay 100ms
                //await Task.Delay(200);  // 非同步延遲

                // 设置目标为 p2_vrms_box
                targetTextBox = "p2_vrms_box";
                expectedLength = 7;

                // 向串口发送数据
                byte[] myBytes = { 0x49, 0x01, 0x30, 0x86 };
                serialPort1.Write(myBytes, 0, myBytes.Length);
                Thread.Sleep(200);  // Delay 100ms
                //await Task.Delay(200);  // 非同步延遲

                /*
                if (serialPort1.BytesToRead >= 7)
                {
                    byte[] buffer = new byte[7];
                    serialPort1.Read(buffer, 0, 7);

                    // Debug
                    string result = BitConverter.ToString(buffer, 0, buffer.Length).Replace("-", " "); // 將讀取到的數據轉換為字串
                    Console.WriteLine(result);

                    // 检查第一个字节为 0x49，第二个字节为 0x04
                    if (buffer[0] == 0x49 && buffer[1] == 0x04)
                    {
                        uint Vrms_int32 = (uint)(buffer[2] << 24 | buffer[3] << 16 | buffer[4] << 8 | buffer[5]);
                        float Vrms_float = (float)(Vrms_int32 / 1000.0);
                        Console.WriteLine("Value : " + Vrms_float);
                        p2_vrms_box.Text = Vrms_float.ToString();
                    }
                }
                */
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                // 清空 RichTextBox 控制元件內容
                p2_irms_box.Clear();

                // 清除串口接收缓冲区中的数据
                serialPort1.DiscardInBuffer();
                receiveBuffer.Clear();  // 清除整个缓冲区，等待下一个有效数据帧

                // 發送停止命令 
                byte[] myBytes1 = { 0x49, 0x01, 0x83, 0x33 };
                serialPort1.Write(myBytes1, 0, myBytes1.Length);
                Thread.Sleep(200);  // Delay 100ms
                //await Task.Delay(200);  // 非同步延遲

                // 设置目标为 p2_irms_box
                targetTextBox = "p2_irms_box";
                expectedLength = 7;

                // 向串口发送数据
                byte[] myBytes = { 0x49, 0x01, 0x32, 0x84 };
                serialPort1.Write(myBytes, 0, myBytes.Length);
                Thread.Sleep(200);  // Delay 100ms
                //await Task.Delay(200);  // 非同步延遲

                /*
                if (serialPort1.BytesToRead >= 7)
                {
                    byte[] buffer = new byte[7];
                    serialPort1.Read(buffer, 0, 7);

                    // Debug
                    string result = BitConverter.ToString(buffer, 0, buffer.Length).Replace("-", " "); // 將讀取到的數據轉換為字串
                    Console.WriteLine(result);

                    // 检查第一个字节为 0x49，第二个字节为 0x04
                    if (buffer[0] == 0x49 && buffer[1] == 0x04)
                    {
                        uint Irms_int32 = (uint)(buffer[2] << 24 | buffer[3] << 16 | buffer[4] << 8 | buffer[5]);
                        float Irms_float = (float)(Irms_int32 / 128000.0);
                        Console.WriteLine("Value : " + Irms_float);
                        p2_irms_box.Text = Irms_float.ToString();
                    }
                }
                */
            }
        }

        private void button1_Click_3(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                // 清空 RichTextBox 控制元件內容
                p2_pavg_box.Clear();

                // 清除串口接收缓冲区中的数据
                serialPort1.DiscardInBuffer();
                receiveBuffer.Clear();  // 清除整个缓冲区，等待下一个有效数据帧

                // 發送停止命令 
                byte[] myBytes1 = { 0x49, 0x01, 0x83, 0x33 };
                serialPort1.Write(myBytes1, 0, myBytes1.Length);
                Thread.Sleep(200);  // Delay 100ms
                //await Task.Delay(200);  // 非同步延遲

                // 设置目标为 p2_pavg_box
                targetTextBox = "p2_pavg_box";
                expectedLength = 7;

                // 向串口发送数据
                byte[] myBytes = { 0x49, 0x01, 0x34, 0x82 };
                serialPort1.Write(myBytes, 0, myBytes.Length);
                Thread.Sleep(200);  // Delay 100ms
                //await Task.Delay(200);  // 非同步延遲

                /*
                if (serialPort1.BytesToRead >= 7)
                {
                    byte[] buffer = new byte[7];
                    serialPort1.Read(buffer, 0, 7);

                    // Debug
                    string result = BitConverter.ToString(buffer, 0, buffer.Length).Replace("-", " "); // 將讀取到的數據轉換為字串
                    Console.WriteLine(result);

                    // 检查第一个字节为 0x49，第二个字节为 0x04
                    if (buffer[0] == 0x49 && buffer[1] == 0x04)
                    {
                        uint Pavg_int32 = (uint)(buffer[2] << 24 | buffer[3] << 16 | buffer[4] << 8 | buffer[5]);
                        float Pavg_float = (float)(Pavg_int32 / 1000.0);
                        Console.WriteLine("Value : " + Pavg_float);
                        p2_pavg_box.Text = Pavg_float.ToString();
                    }
                }
                */
            }
        }

        private void groupBox8_Enter(object sender, EventArgs e)
        {

        }

        async private void button12_Click(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                try
                {
                    // 使用系統時間動態生成檔名，格式化為 "yyyyMMdd_HHmmss"
                    string fileName = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    string filePath = $"output_{fileName}.txt"; // 組合成完整檔案路徑

                    // 如果文件已經存在，刪除它以創建一個新的
                    if (File.Exists(filePath))
                    {
                        File.Delete(filePath);
                    }

                    // 初始化一個空文件
                    using (StreamWriter sw = new StreamWriter(filePath))
                    {
                        // 寫入表頭
                        //sw.WriteLine("Voltage RMS");
                    }

                    // 讀取100%電流規格
                    string text1 = curr_100p_box.Text;
                    float maxCurrent;
                    float.TryParse(text1, out maxCurrent);
                    Console.WriteLine(maxCurrent);

                    // 讀取Delay sencond
                    string text2 = delay_sec_box.Text;
                    int delaySec;
                    int.TryParse(text2, out delaySec);
                    Console.WriteLine(delaySec);

                    // 讀取repeat
                    string text3 = repeat_box.Text;
                    int repeatTimes;
                    int.TryParse(text3, out repeatTimes);
                    Console.WriteLine(repeatTimes);

                    // 使用 NationalInstruments.VisaNS.ResourceManager 來避免命名衝突
                    NationalInstruments.Visa.ResourceManager resourceManager = new NationalInstruments.Visa.ResourceManager();

                    for (int repeat = 0; repeat < repeatTimes; repeat++)
                    {
                        // 從5%到100%，每次增加5%
                        for (int percentage = 5; percentage <= 100; percentage += 5)
                        {
                            //if (percentage == 25)
                            // percentage = 23;

                            // 打開與設備的會話
                            using (MessageBasedSession session = (MessageBasedSession)resourceManager.Open(DC_Load63202A_Address))
                            {
                                double currentValue = Math.Round((maxCurrent * percentage) / 100.0, 2); //保留2位小數
                                Console.WriteLine($"Setting current to {percentage}% of {maxCurrent}A: {currentValue}A");

                                // 發送SCPI命令來設定電流值
                                // 假設你已經有一個名為 session 的儀器控制對象
                                string command = $"CURR:STAT:L1 {currentValue}\n";
                                session.RawIO.Write(command);
                            }

                            // 等待
                            await Task.Delay(delaySec * 1000);  // 非同步延遲

                            // 打開與設備的會話
                            using (MessageBasedSession session = (MessageBasedSession)resourceManager.Open(PowerMeter66202_Address))
                            {
                                // 發送 SCPI 命令來讀取電壓 RMS 值
                                session.RawIO.Write("MEAS:VOLT:RMS?\n");
                                string voltageRmsResponse = session.RawIO.ReadString();

                                // 發送 SCPI 命令來讀取電流 RMS 值
                                session.RawIO.Write("MEAS:CURR:RMS?\n");
                                string currentRmsResponse = session.RawIO.ReadString();

                                // 發送 SCPI 命令來讀取實際功率
                                session.RawIO.Write("MEAS:POW:REAL?\n");
                                string powerResponse = session.RawIO.ReadString();

                                // 清空 RichTextBox 控制元件內容
                                chroma_box.Clear();

                                // 將讀取到的數據顯示在 richTextBox1 中
                                chroma_box.Text = voltageRmsResponse + currentRmsResponse + powerResponse;
                                //chroma_box.AppendText(percentage.ToString());
                                chroma_box.AppendText($"{percentage:F2}%\n");
                                chroma_box.AppendText($"{repeat+1}\n");

                                // 增加格式化後的數據顯示
                                //chroma_box.AppendText("\n\n");
                                //chroma_box.AppendText(vrms_float.ToString("N2") + "\n");
                                //chroma_box.AppendText(irms_float.ToString("N5") + "\n");
                                //chroma_box.AppendText(power_float.ToString("N3") + "\n");


                                // 模擬得到的6筆數值數據
                                double[] values = new double[6];

                                double voltageRmsValue;
                                double.TryParse(voltageRmsResponse, out voltageRmsValue);
                                values[0] = voltageRmsValue;

                                double currentRmsValue;
                                double.TryParse(currentRmsResponse, out currentRmsValue);
                                values[1] = currentRmsValue;

                                double powerValue;
                                double.TryParse(powerResponse, out powerValue);
                                values[2] = powerValue;

                                values[3] = vrms_float;
                                values[4] = irms_float;
                                values[5] = power_float;

                                values[3] = Math.Round(vrms_float, 2); // 保留2位小數
                                values[4] = Math.Round(irms_float, 5); // 保留5位小數
                                values[5] = Math.Round(power_float, 3); // 保留3位小數

                                // 清空 RichTextBox 控制元件內容
                                error_rate.Clear();

                                // 將讀取到的數據顯示在 richTextBox2 中
                                if (voltageRmsValue != 0)
                                {
                                    double voltage_error_rate;
                                    voltage_error_rate = ((voltageRmsValue - vrms_float) / voltageRmsValue) * 100.0;
                                    error_rate.AppendText($"{voltage_error_rate:F2}%\n");
                                }

                                if (currentRmsValue != 0)
                                {
                                    double current_error_rate;
                                    current_error_rate = ((currentRmsValue - irms_float) / currentRmsValue) * 100.0;
                                    error_rate.AppendText($"{current_error_rate:F2}%\n");
                                }

                                if (powerValue != 0)
                                {
                                    double power_error_rate;
                                    power_error_rate = ((powerValue - power_float) / powerValue) * 100.0;
                                    error_rate.AppendText($"{power_error_rate:F2}%\n");
                                }

                                // 讀取當前的文件內容並拆分成多行
                                string[] fileLines = File.Exists(filePath) ? File.ReadAllLines(filePath) : new string[6];

                                // 確保文件中有6行（對應6筆數據）
                                if (fileLines.Length < 6)
                                {
                                    Array.Resize(ref fileLines, 6); // 如果不足6行，則擴展到6行
                                }

                                // 打開文件以追加方式寫入
                                using (StreamWriter sw = new StreamWriter(filePath))
                                {
                                    for (int i = 0; i < 6; i++)
                                    {
                                        string formattedValue;

                                        // 根據索引設置不同的小數位格式
                                        if (i == 0)
                                        {
                                            formattedValue = values[i].ToString("F2");  // 第一筆到小數點第2位
                                        }
                                        else if (i == 1)
                                        {
                                            formattedValue = values[i].ToString("F4");  // 第二筆到小數點第4位
                                        }
                                        else if (i == 2)
                                        {
                                            formattedValue = values[i].ToString("F2");  // 第三筆到小數點第2位
                                        }
                                        else if (i == 3)
                                        {
                                            formattedValue = values[i].ToString("F2");  // 第四筆到小數點第2位
                                        }
                                        else if (i == 4)
                                        {
                                            formattedValue = values[i].ToString("F5");  // 第五筆到小數點第5位
                                        }
                                        else // i == 5
                                        {
                                            formattedValue = values[i].ToString("F3");  // 第六筆到小數點第3位
                                        }

                                        // 將當前數據附加到每一行的末尾，如果行為null則設置為空字符串
                                        if (string.IsNullOrEmpty(fileLines[i]))
                                        {
                                            fileLines[i] = formattedValue; // 新行
                                        }
                                        else
                                        {
                                            fileLines[i] += $"\t{formattedValue}"; // 將數值附加到現有行，使用制表符分隔
                                        }
                                    }

                                    // 將修改後的內容寫回文件
                                    foreach (var line in fileLines)
                                    {
                                        sw.WriteLine(line);
                                    }
                                }

                            }

                            //if (percentage == 23)
                            //percentage = 25;
                        }
                    }

                    // 打開與設備的會話
                    using (MessageBasedSession session = (MessageBasedSession)resourceManager.Open(DC_Load63202A_Address))
                    {
                        // 發送 SCPI 命令來讀取電壓 RMS 值
                        session.RawIO.Write("CURR:STAT:L1 0\n");
                    }

                    // 顯示完成操作的訊息框
                    MessageBox.Show("Successfully written to the file !", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                catch (VisaException ex)
                {
                    MessageBox.Show("Confirm connection.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        async private void button14_Click(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                try
                {
                    // 使用系統時間動態生成檔名，格式化為 "yyyyMMdd_HHmmss"
                    string fileName = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    string filePath = $"output_{fileName}.txt"; // 組合成完整檔案路徑

                    // 如果文件已經存在，刪除它以創建一個新的
                    if (File.Exists(filePath))
                    {
                        File.Delete(filePath);
                    }

                    // 初始化一個空文件
                    using (StreamWriter sw = new StreamWriter(filePath))
                    {
                        // 寫入表頭
                        //sw.WriteLine("Voltage RMS");
                    }

                    // 讀取100%電流規格
                    string text1 = curr_100p_box.Text;
                    float maxCurrent;
                    float.TryParse(text1, out maxCurrent);
                    Console.WriteLine(maxCurrent);

                    // 讀取Delay sencond
                    string text2 = delay_sec_box.Text;
                    int delaySec;
                    int.TryParse(text2, out delaySec);
                    Console.WriteLine(delaySec);

                    // 使用 NationalInstruments.VisaNS.ResourceManager 來避免命名衝突
                    NationalInstruments.Visa.ResourceManager resourceManager = new NationalInstruments.Visa.ResourceManager();

                    // 從5%到100%，每次增加5%
                    for (int percentage = 5; percentage <= 100; percentage += 5)
                    {
                        if (percentage == 25)
                            percentage = 23;

                        // 打開與設備的會話
                        using (MessageBasedSession session = (MessageBasedSession)resourceManager.Open(DC_Load63202A_Address))
                        {
                            double currentValue = Math.Round((maxCurrent * percentage) / 100.0, 2); //保留2位小數
                            Console.WriteLine($"Setting current to {percentage}% of {maxCurrent}A: {currentValue}A");

                            // 發送SCPI命令來設定電流值
                            // 假設你已經有一個名為 session 的儀器控制對象
                            string command = $"CURR:STAT:L1 {currentValue}\n";
                            session.RawIO.Write(command);
                        }

                        // 等待
                        await Task.Delay(delaySec * 1000);  // 非同步延遲

                        // 打開與設備的會話
                        using (MessageBasedSession session = (MessageBasedSession)resourceManager.Open(PowerMeter66202_Address))
                        {
                            // 發送 SCPI 命令來讀取電壓 RMS 值
                            session.RawIO.Write("MEAS:VOLT:DC?\n");
                            string voltageRmsResponse = session.RawIO.ReadString();

                            // 發送 SCPI 命令來讀取電流 RMS 值
                            session.RawIO.Write("MEAS:CURR:DC?\n");
                            string currentRmsResponse = session.RawIO.ReadString();

                            // 發送 SCPI 命令來讀取實際功率
                            session.RawIO.Write("MEAS:POW:DC?\n");
                            string powerResponse = session.RawIO.ReadString();

                            // 清空 RichTextBox 控制元件內容
                            chroma_box.Clear();

                            // 將讀取到的數據顯示在 richTextBox1 中
                            chroma_box.Text = voltageRmsResponse + currentRmsResponse + powerResponse;
                            //chroma_box.AppendText(percentage.ToString());
                            chroma_box.AppendText($"{percentage:F2}%\n");

                            // 增加格式化後的數據顯示
                            //chroma_box.AppendText("\n\n");
                            //chroma_box.AppendText(vrms_float.ToString("N2") + "\n");
                            //chroma_box.AppendText(irms_float.ToString("N5") + "\n");
                            //chroma_box.AppendText(power_float.ToString("N3") + "\n");


                            // 模擬得到的6筆數值數據
                            double[] values = new double[6];

                            double voltageRmsValue;
                            double.TryParse(voltageRmsResponse, out voltageRmsValue);
                            values[0] = voltageRmsValue;

                            double currentRmsValue;
                            double.TryParse(currentRmsResponse, out currentRmsValue);
                            values[1] = currentRmsValue;

                            double powerValue;
                            double.TryParse(powerResponse, out powerValue);
                            values[2] = powerValue;

                            values[3] = vrms_float;
                            values[4] = irms_float;
                            values[5] = power_float;

                            values[3] = Math.Round(vrms_float, 2); // 保留2位小數
                            values[4] = Math.Round(irms_float, 5); // 保留5位小數
                            values[5] = Math.Round(power_float, 3); // 保留3位小數

                            // 清空 RichTextBox 控制元件內容
                            error_rate.Clear();

                            // 將讀取到的數據顯示在 richTextBox2 中
                            double voltage_error_rate;
                            voltage_error_rate = ((voltageRmsValue - vrms_float) / voltageRmsValue) * 100.0;
                            error_rate.AppendText(voltage_error_rate.ToString());

                            double current_error_rate;
                            current_error_rate = ((currentRmsValue - irms_float) / currentRmsValue) * 100.0;
                            error_rate.AppendText(current_error_rate.ToString());

                            double power_error_rate;
                            power_error_rate = ((powerValue - power_float) / powerValue) * 100.0;
                            error_rate.AppendText(power_error_rate.ToString());

                            // 讀取當前的文件內容並拆分成多行
                            string[] fileLines = File.Exists(filePath) ? File.ReadAllLines(filePath) : new string[6];

                            // 確保文件中有6行（對應6筆數據）
                            if (fileLines.Length < 6)
                            {
                                Array.Resize(ref fileLines, 6); // 如果不足6行，則擴展到6行
                            }

                            // 打開文件以追加方式寫入
                            using (StreamWriter sw = new StreamWriter(filePath))
                            {
                                for (int i = 0; i < 6; i++)
                                {
                                    string formattedValue;

                                    // 根據索引設置不同的小數位格式
                                    if (i == 0)
                                    {
                                        formattedValue = values[i].ToString("F2");  // 第一筆到小數點第2位
                                    }
                                    else if (i == 1)
                                    {
                                        formattedValue = values[i].ToString("F4");  // 第二筆到小數點第4位
                                    }
                                    else if (i == 2)
                                    {
                                        formattedValue = values[i].ToString("F2");  // 第三筆到小數點第2位
                                    }
                                    else if (i == 3)
                                    {
                                        formattedValue = values[i].ToString("F2");  // 第四筆到小數點第2位
                                    }
                                    else if (i == 4)
                                    {
                                        formattedValue = values[i].ToString("F5");  // 第五筆到小數點第5位
                                    }
                                    else // i == 5
                                    {
                                        formattedValue = values[i].ToString("F3");  // 第六筆到小數點第3位
                                    }

                                    // 將當前數據附加到每一行的末尾，如果行為null則設置為空字符串
                                    if (string.IsNullOrEmpty(fileLines[i]))
                                    {
                                        fileLines[i] = formattedValue; // 新行
                                    }
                                    else
                                    {
                                        fileLines[i] += $"\t{formattedValue}"; // 將數值附加到現有行，使用制表符分隔
                                    }
                                }

                                // 將修改後的內容寫回文件
                                foreach (var line in fileLines)
                                {
                                    sw.WriteLine(line);
                                }
                            }

                        }

                        if (percentage == 23)
                            percentage = 25;
                    }

                    // 打開與設備的會話
                    using (MessageBasedSession session = (MessageBasedSession)resourceManager.Open(DC_Load63202A_Address))
                    {
                        // 發送 SCPI 命令來讀取電壓 RMS 值
                        session.RawIO.Write("CURR:STAT:L1 0\n");
                    }

                    // 顯示完成操作的訊息框
                    MessageBox.Show("Successfully written to the file !", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                catch (VisaException ex)
                {
                    MessageBox.Show("Confirm connection.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                try
                {
                    // 清空 RichTextBox 控制元件內容
                    cali_v3_box.Clear();
                    cali_i3_box.Clear();
                    cali_p3_box.Clear();

                    // 使用 NationalInstruments.VisaNS.ResourceManager 來避免命名衝突
                    NationalInstruments.Visa.ResourceManager resourceManager = new NationalInstruments.Visa.ResourceManager();

                    // 打開與設備的會話
                    using (MessageBasedSession session = (MessageBasedSession)resourceManager.Open(PowerMeter66202_Address))
                    {
                        // 發送 SCPI 命令來讀取電壓 RMS 值
                        session.RawIO.Write("MEAS:VOLT:RMS?\n");
                        string voltageRmsResponse = session.RawIO.ReadString();

                        // 發送 SCPI 命令來讀取電流 RMS 值
                        session.RawIO.Write("MEAS:CURR:RMS?\n");
                        string currentRmsResponse = session.RawIO.ReadString();

                        // 發送 SCPI 命令來讀取實際功率
                        session.RawIO.Write("MEAS:POW:REAL?\n");
                        string powerResponse = session.RawIO.ReadString();

                        // 將讀取到的數據顯示在 richTextBox1 中
                        cali_v3_box.Text = voltageRmsResponse;
                        cali_i3_box.Text = currentRmsResponse;
                        cali_p3_box.Text = powerResponse;
                    }
                }
                catch (VisaException ex)
                {
                    MessageBox.Show("Confirm connection.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void button16_Click(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                try
                {
                    // 清空 RichTextBox 控制元件內容
                    cali_v2_box.Clear();
                    cali_i2_box.Clear();
                    cali_p2_box.Clear();

                    // 使用 NationalInstruments.VisaNS.ResourceManager 來避免命名衝突
                    NationalInstruments.Visa.ResourceManager resourceManager = new NationalInstruments.Visa.ResourceManager();

                    // 打開與設備的會話
                    using (MessageBasedSession session = (MessageBasedSession)resourceManager.Open(PowerMeter66202_Address))
                    {
                        // 發送 SCPI 命令來讀取電壓 RMS 值
                        session.RawIO.Write("MEAS:VOLT:RMS?\n");
                        string voltageRmsResponse = session.RawIO.ReadString();

                        // 發送 SCPI 命令來讀取電流 RMS 值
                        session.RawIO.Write("MEAS:CURR:RMS?\n");
                        string currentRmsResponse = session.RawIO.ReadString();

                        // 發送 SCPI 命令來讀取實際功率
                        session.RawIO.Write("MEAS:POW:REAL?\n");
                        string powerResponse = session.RawIO.ReadString();

                        // 將讀取到的數據顯示在 richTextBox1 中
                        cali_v2_box.Text = voltageRmsResponse;
                        cali_i2_box.Text = currentRmsResponse;
                        cali_p2_box.Text = powerResponse;
                    }
                }
                catch (VisaException ex)
                {
                    MessageBox.Show("Confirm connection.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void button17_Click(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                try
                {
                    // 清空 RichTextBox 控制元件內容
                    cali_v1_box.Clear();
                    cali_i1_box.Clear();
                    cali_p1_box.Clear();

                    // 使用 NationalInstruments.VisaNS.ResourceManager 來避免命名衝突
                    NationalInstruments.Visa.ResourceManager resourceManager = new NationalInstruments.Visa.ResourceManager();

                    // 打開與設備的會話
                    using (MessageBasedSession session = (MessageBasedSession)resourceManager.Open(PowerMeter66202_Address))
                    {
                        // 發送 SCPI 命令來讀取電壓 RMS 值
                        session.RawIO.Write("MEAS:VOLT:RMS?\n");
                        string voltageRmsResponse = session.RawIO.ReadString();

                        // 發送 SCPI 命令來讀取電流 RMS 值
                        session.RawIO.Write("MEAS:CURR:RMS?\n");
                        string currentRmsResponse = session.RawIO.ReadString();

                        // 發送 SCPI 命令來讀取實際功率
                        session.RawIO.Write("MEAS:POW:REAL?\n");
                        string powerResponse = session.RawIO.ReadString();

                        // 將讀取到的數據顯示在 richTextBox1 中
                        cali_v1_box.Text = voltageRmsResponse;
                        cali_i1_box.Text = currentRmsResponse;
                        cali_p1_box.Text = powerResponse;
                    }
                }
                catch (VisaException ex)
                {
                    MessageBox.Show("Confirm connection.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void button18_Click(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                try
                {
                    // 清空 RichTextBox 控制元件內容
                    cali_v4_box.Clear();
                    cali_i4_box.Clear();
                    cali_p4_box.Clear();

                    // 使用 NationalInstruments.VisaNS.ResourceManager 來避免命名衝突
                    NationalInstruments.Visa.ResourceManager resourceManager = new NationalInstruments.Visa.ResourceManager();

                    // 打開與設備的會話
                    using (MessageBasedSession session = (MessageBasedSession)resourceManager.Open(PowerMeter66202_Address))
                    {
                        // 發送 SCPI 命令來讀取電壓 RMS 值
                        session.RawIO.Write("MEAS:VOLT:RMS?\n");
                        string voltageRmsResponse = session.RawIO.ReadString();

                        // 發送 SCPI 命令來讀取電流 RMS 值
                        session.RawIO.Write("MEAS:CURR:RMS?\n");
                        string currentRmsResponse = session.RawIO.ReadString();

                        // 發送 SCPI 命令來讀取實際功率
                        session.RawIO.Write("MEAS:POW:REAL?\n");
                        string powerResponse = session.RawIO.ReadString();

                        // 將讀取到的數據顯示在 richTextBox1 中
                        cali_v4_box.Text = voltageRmsResponse;
                        cali_i4_box.Text = currentRmsResponse;
                        cali_p4_box.Text = powerResponse;
                    }
                }
                catch (VisaException ex)
                {
                    MessageBox.Show("Confirm connection.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void button19_Click(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                // 清空 RichTextBox 控制元件內容
                status_box.Clear();

                // 清除串口接收缓冲区中的数据
                serialPort1.DiscardInBuffer();
                receiveBuffer.Clear();  // 清除整个缓冲区，等待下一个有效数据帧

                // 發送停止命令 
                byte[] myBytes1 = { 0x49, 0x01, 0x83, 0x33 };
                serialPort1.Write(myBytes1, 0, myBytes1.Length);
                Thread.Sleep(200);  // Delay 100ms
                //await Task.Delay(200);  // 非同步延遲

                // 设置目标为 chip_ssid_box
                targetTextBox = "status_box";
                expectedLength = 4;

                // 向串口发送数据
                byte[] myBytes = { 0x49, 0x01, 0x85, 0x31 };
                serialPort1.Write(myBytes, 0, myBytes.Length);
                Thread.Sleep(200);  // Delay 100ms
                //await Task.Delay(200);  // 非同步延遲

                /*
                if (serialPort1.BytesToRead >= 4)
                {
                    byte[] buffer = new byte[4];
                    serialPort1.Read(buffer, 0, 4);
                    
                    // Debug
                    string result = BitConverter.ToString(buffer, 0, buffer.Length).Replace("-", " "); // 將讀取到的數據轉換為字串
                    Console.WriteLine(result);

                    // 检查第一个字节为 0x49，第二个字节为 0x01
                    if (buffer[0] == 0x49 && buffer[1] == 0x01)
                    {
                        byte thirdByte = buffer[2];
                        Console.WriteLine("Value : " + thirdByte);
                        string dataString = thirdByte.ToString("D2");
                        chip_ssid_box.AppendText(dataString);
                    }                    
                }
                */
            }
        }

        private void delay_sec_box_TextChanged(object sender, EventArgs e)
        {

        }

        /*
        private async void button20_Click(object sender, EventArgs e)
        {
            string filePath = "power_data.txt";

            try
            {
                // 使用異步方式寫入檔案
                using (StreamWriter writer = new StreamWriter(filePath, true))
                {
                    for (int i = 0; i < 6000; i++) // 假設寫入10次
                    {
                        writer.WriteLine(power_float.ToString()); // 寫入power_float
                        writer.Flush(); // 立即寫入檔案
                        await Task.Delay(100); // 延遲100ms，避免阻塞UI
                    }
                }

                MessageBox.Show("資料寫入完成！");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"寫入檔案時發生錯誤：{ex.Message}");
            }
        }
        */

        private void button5_Click_1(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                // 清空 RichTextBox 控制元件內容
                p3_pavg_box.Clear();

                // 清除串口接收缓冲区中的数据
                serialPort1.DiscardInBuffer();
                receiveBuffer.Clear();  // 清除整个缓冲区，等待下一个有效数据帧

                // 發送停止命令 
                byte[] myBytes1 = { 0x49, 0x01, 0x83, 0x33 };
                serialPort1.Write(myBytes1, 0, myBytes1.Length);
                Thread.Sleep(200);  // Delay 100ms
                //await Task.Delay(200);  // 非同步延遲 

                // 设置目标为 p3_pavg_box
                targetTextBox = "p3_pavg_box";
                expectedLength = 7;

                // 向串口发送数据
                byte[] myBytes = { 0x49, 0x01, 0x24, 0x92 }; 
                serialPort1.Write(myBytes, 0, myBytes.Length);
                Thread.Sleep(200);  // Delay 100ms
                //await Task.Delay(200);  // 非同步延遲

                /*
                if (serialPort1.BytesToRead >= 7)
                {
                    byte[] buffer = new byte[7];
                    serialPort1.Read(buffer, 0, 7);

                    // Debug
                    string result = BitConverter.ToString(buffer, 0, buffer.Length).Replace("-", " "); // 將讀取到的數據轉換為字串
                    Console.WriteLine(result);

                    // 检查第一个字节为 0x49，第二个字节为 0x04
                    if (buffer[0] == 0x49 && buffer[1] == 0x04)
                    {
                        uint Pavg_int32 = (uint)(buffer[2] << 24 | buffer[3] << 16 | buffer[4] << 8 | buffer[5]);
                        float Pavg_float = (float)(Pavg_int32 / 1000.0);
                        Console.WriteLine("Value : " + Pavg_float);
                        p3_pavg_box.Text = Pavg_float.ToString();
                    }
                }
                */
            }
        }

        private void label37_Click(object sender, EventArgs e)
        {

        }

        private void vrms_box_TextChanged(object sender, EventArgs e)
        {

        }

        private void chart1_Click(object sender, EventArgs e)
        {

        }

        private void chart1_Click_1(object sender, EventArgs e)
        {

        }

       

        private void textBox19_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
