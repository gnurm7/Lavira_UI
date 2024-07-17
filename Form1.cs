using CefSharp;
using CefSharp.DevTools.Network;
using CefSharp.SchemeHandler;
using CefSharp.WinForms;
using DocumentFormat.OpenXml.Drawing.Charts;
using GMap.NET;
using GMap.NET.MapProviders;
using System;
using System.Diagnostics;
using System.IO;
using System.IO.Ports;
using System.Net.Sockets;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Timers;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Spreadsheet;
using static System.Windows.Forms.AxHost;

namespace arayüz_örnek
{


    public partial class Form1 : Form
    {

        private SerialPort serialPort;
        //başlangıç unity
        private Process unityProcess; // Unity uygulamasını temsil edecek Process nesnesi
        private IntPtr unityWindowHandle; // Unity penceresinin handle'ı

        [DllImport("user32.dll")]
        static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

        [DllImport("user32.dll", SetLastError = true, EntryPoint = "SetWindowLong")]
        private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true, ExactSpelling = true)]
        public static extern int GetWindowLongA(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true, ExactSpelling = true)]
        public static extern int SetWindowLongA(IntPtr hWnd, int nIndex, int dwNewLong);
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true, ExactSpelling = true)]
        public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, int uFlags);

        private const int GWL_STYLE = -16;
        private const int WS_VISIBLE = 0x10000000;
        private const int WS_POPUP = unchecked((int)0x80000000);
        const int WS_BORDER = 8388608;
        const int WS_DLGFRAME = 4194304;
        const int WS_CAPTION = WS_BORDER | WS_DLGFRAME;
        const int WS_SYSMENU = 524288;
        const int WS_THICKFRAME = 262144;
        const int WS_MINIMIZE = 536870912;
        const int WS_MAXIMIZEBOX = 65536;
        const int GWL_EXSTYLE = (int)-20L;
        const int WS_EX_DLGMODALFRAME = (int)0x1L;
        const int SWP_NOMOVE = 0x2;
        const int SWP_NOSIZE = 0x1;
        const int SWP_FRAMECHANGED = 0x20;
        public void MakeExternalWindowBorderless(IntPtr MainWindowHandle)
        {
            int Style = 0;
            Style = GetWindowLongA(MainWindowHandle, GWL_STYLE);
            Style = Style & ~WS_CAPTION;
            Style = Style & ~WS_SYSMENU;
            Style = Style & ~WS_THICKFRAME;
            Style = Style & ~WS_MINIMIZE;
            Style = Style & ~WS_MAXIMIZEBOX;
            SetWindowLongA(MainWindowHandle, GWL_STYLE, Style);
            Style = GetWindowLongA(MainWindowHandle, GWL_EXSTYLE);
            SetWindowLongA(MainWindowHandle, GWL_EXSTYLE, Style | WS_EX_DLGMODALFRAME);
            SetWindowPos(MainWindowHandle, new IntPtr(0), 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_FRAMECHANGED);
        }
        //end unity
        public Form1()
        {
            InitializeComponent();
            DisableAllControls(this);
            chromiumWebBrowser1.LoadHtml(File.ReadAllText(@Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"/index.html"));
        }


        SerialPort stream_sendDataToHYI = new SerialPort();
        private IntPtr FindWindowByProcessId(int processId)
        {
            Process[] processes = Process.GetProcesses();
            foreach (Process process in processes)
            {
                if (process.Id == processId)
                {
                    return process.MainWindowHandle;
                }
            }
            return IntPtr.Zero;
        }
        private void SetWindowStyle(IntPtr hWnd, int style)
        {
            SetWindowLong(hWnd, GWL_STYLE, style);
        }
        Process simApplication;
        void Kill(string app)
        {
            foreach (var process in Process.GetProcessesByName(app))
            {
                process.Kill();
            }
        }
        //Loglama Düzeltilcek
        //Grafiklere İsim Eklenilcek
        //HYI Test Edilcek
        //HYI Verilerinin Yerleri Düzeltilcek
        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                timer2.Start();
                Kill("3DSim");
                Kill("3DSim.exe");
            }
            catch (Exception) { }
            try
            {
                string _3DSimExePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\roketunity\3DSim_New\3DSim";
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.FileName = _3DSimExePath;
                startInfo.WindowStyle = ProcessWindowStyle.Maximized;
                simApplication = Process.Start(startInfo);
                simApplication.WaitForInputIdle();
                MoveWindow(simApplication.MainWindowHandle, 0, 0, panel_unity.Width, panel_unity.Height, true);
                SetParent(simApplication.MainWindowHandle, panel_unity.Handle);
                MakeExternalWindowBorderless(simApplication.MainWindowHandle);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Hata: " + ex.Message);
            }
            dataGridView1.ColumnCount = 21;
            dataGridView1.Columns[0].Name = "Sayac";
            dataGridView1.Columns[1].Name = "İrtifa";
            dataGridView1.Columns[2].Name = "Roket GPS irtifa";
            dataGridView1.Columns[3].Name = "Roket Boylam";
            dataGridView1.Columns[4].Name = "Roket Enlem";
            dataGridView1.Columns[5].Name = "Görev Yükü GPS irtifa";
            dataGridView1.Columns[6].Name = "Görev Yükü Enlem";
            dataGridView1.Columns[7].Name = "Görev Yükü Boylam ";
            dataGridView1.Columns[8].Name = "JireskopX";
            dataGridView1.Columns[9].Name = "JireskopY";
            dataGridView1.Columns[10].Name = "JireskopZ";
            dataGridView1.Columns[11].Name = "ivmeX";
            dataGridView1.Columns[12].Name = "ivmeY";
            dataGridView1.Columns[13].Name = "ivmeZ";
            dataGridView1.Columns[14].Name = "Durum";
            dataGridView1.Columns[15].Name = "CRC";
            dataGridView1.Columns[16].Name = "Pil Gerilim";
            dataGridView1.Columns[17].Name = "Sıcaklık";
            dataGridView1.Columns[18].Name = "Hız";
            dataGridView1.Columns[19].Name = "Tarih";
            dataGridView1.Columns[20].Name = "Saat";
            string[] ports = SerialPort.GetPortNames();
            comboBox1.Items.AddRange(ports);
            comboBox2.Items.AddRange(ports);
            MAP.CacheLocation = "C:\\Users\\Asus\\Desktop\\Cache";
            MAP.DragButton = MouseButtons.Left;
            GMaps.Instance.Mode = AccessMode.ServerAndCache;
            MAP.Manager.CancelTileCaching();
            MAP.HoldInvalidation = false;
            MAP.Zoom = 15;
            MAP.Position = new GMap.NET.PointLatLng(40.74208412258973, 29.942214732057536);
            MAP.MapProvider = GoogleSatelliteMapProvider.Instance;
            timer3.Start();
        }
        //openPortToSendData
        private void openButton_Click(object sender, EventArgs e)
        {
            button2.PerformClick();
            string[] ports = SerialPort.GetPortNames();
            comboBox1.Items.AddRange(ports);
            serialPort = new SerialPort(comboBox1.SelectedItem.ToString(), 9600, Parity.None, 8, StopBits.One);
            serialPort.DataReceived += SerialPort_DataReceived;
            if (!serialPort.IsOpen)
            {
                serialPort.PortName = comboBox1.SelectedItem.ToString(); // Bağlantı noktasını belirtin
                serialPort.BaudRate = 9600; // Bağlantı hızını belirtin
                serialPort.Parity = Parity.None; // Parity
                serialPort.DataBits = 8; // Data bits
                serialPort.StopBits = StopBits.One; // Stop bits
                serialPort.Handshake = Handshake.None; // Handshake

                try
                {
                    serialPort.Open();
                }
                catch (UnauthorizedAccessException ex)
                {
                    MessageBox.Show("Erişim reddedildi: " + ex.Message);
                }
            }
            else
                MessageBox.Show("Yavaş lan gaç tane basıyon");
        } 
        private void DisplayReceivedData(string data)
        {
            try
            {
                if (InvokeRequired) BeginInvoke(new Action<string>(DisplayReceivedData), data);
                else
                {
                    data.Replace("*", "");
                    data.Replace("+", "");
                    string[] veriler = data.Split(',');
                    veriler[0] = veriler[0].Replace("*", "");
                    veriler[0] = veriler[20].Replace("+", "");
                    Sayac.Text = veriler[0];
                    İrtifa.Text = veriler[1];
                    RoketGPSirtifa.Text = veriler[2];
                    roketBoylam.Text = veriler[3];
                    roketEnlem.Text = veriler[4];
                    GörevYüküGPSirtifa1.Text = veriler[5];
                    GörevYüküEnlem.Text = veriler[6];
                    GörevYüküBoylam.Text = veriler[7];
                    JireskopX.Text = veriler[8];
                    JireskopY.Text = veriler[9];
                    JireskopZ.Text = veriler[10];
                    ivmeX.Text = veriler[11];
                    ivmeY.Text = veriler[12];
                    ivmeZ.Text = veriler[13];
                    Durum.Text = veriler[14];
                    CRC.Text = veriler[15];
                    pilgerilim.Text = veriler[16];
                    sicaklik.Text = veriler[17];
                    hiz.Text = veriler[18];
                    tarih.Text = veriler[19];
                    saat.Text = veriler[20];
                    yunuslamaacisi.Text = "23";
                    if (richTextBox2.Text.Length >= 2000000)
                        ID.Text = "284994";
                    string _ID = "284994";
                    string _Sayac = veriler[0].Replace("*", "");
                    string _İrtifa = veriler[1];
                    string _RoketGPSirtifa = veriler[2];
                    string _RoketBoylam = veriler[3];
                    string _RoketEnlem = veriler[4];
                    string _GörevYüküGPSirtifa1 = veriler[5];
                    string _GörevYüküEnlem = veriler[6];
                    string _GörevYüküBoylam = veriler[7];
                    string _JireskopX = veriler[8];
                    string _JireskopY = veriler[9];
                    string _JireskopZ = veriler[10];
                    string _IvmeX = veriler[11];
                    string _IvmeY = veriler[12];
                    string _IvmeZ = veriler[13];
                    string _Durum = veriler[14];
                    string _CRC = veriler[15];
                    string _PilGerilim = veriler[16];
                    string _Sicaklik = veriler[17];
                    string _Hiz = veriler[18];
                    string _Tarih = veriler[19];
                    string _Saat = veriler[20];
                    string _yunuslamaacisi = yunuslamaacisi.Text;
                    byte[] package = new byte[78];

                    package[0] = 0xFF;
                    package[1] = 0xFF;
                    package[2] = 0x54;
                    package[3] = 0x52;
                    byte teamID;
                    byte counter;
                    package[4] = byte.TryParse(_ID, out teamID) ? teamID : (byte)0; //teamID 
                    package[5] = byte.TryParse(_Sayac, out counter) ? counter : (byte)0; //counter

                    package[6] = getBytes(float.Parse(_İrtifa))[0];
                    package[7] = getBytes(float.Parse(_İrtifa))[1];
                    package[8] = getBytes(float.Parse(_İrtifa))[2];
                    package[9] = getBytes(float.Parse(_İrtifa))[3];
                    //gps altitude
                    package[10] = getBytes(float.Parse(_RoketGPSirtifa))[0];
                    package[11] = getBytes(float.Parse(_RoketGPSirtifa))[1];
                    package[12] = getBytes(float.Parse(_RoketGPSirtifa))[2];
                    package[13] = getBytes(float.Parse(_RoketGPSirtifa))[3];
                    //gps latitude
                    package[14] = getBytes(float.Parse(_RoketEnlem))[0];
                    package[15] = getBytes(float.Parse(_RoketEnlem))[1];
                    package[16] = getBytes(float.Parse(_RoketEnlem))[2];
                    package[17] = getBytes(float.Parse(_RoketEnlem))[3];
                    //gps longitude
                    package[18] = getBytes(float.Parse(_RoketBoylam))[0];
                    package[19] = getBytes(float.Parse(_RoketBoylam))[1];
                    package[20] = getBytes(float.Parse(_RoketBoylam))[2];
                    package[21] = getBytes(float.Parse(_RoketBoylam))[3];
                    //payload gps altitude             
                    package[22] = getBytes(float.Parse(_GörevYüküGPSirtifa1))[0];
                    package[23] = getBytes(float.Parse(_GörevYüküGPSirtifa1))[1];
                    package[24] = getBytes(float.Parse(_GörevYüküGPSirtifa1))[2];
                    package[25] = getBytes(float.Parse(_GörevYüküGPSirtifa1))[3];
                    //payload gps latitude             
                    package[26] = getBytes(float.Parse(_GörevYüküEnlem))[0];
                    package[27] = getBytes(float.Parse(_GörevYüküEnlem))[1];
                    package[28] = getBytes(float.Parse(_GörevYüküEnlem))[2];
                    package[29] = getBytes(float.Parse(_GörevYüküEnlem))[3];
                    //payload gps longitude           
                    package[30] = getBytes(float.Parse(_GörevYüküBoylam))[0];
                    package[31] = getBytes(float.Parse(_GörevYüküBoylam))[1];
                    package[32] = getBytes(float.Parse(_GörevYüküBoylam))[2];
                    package[33] = getBytes(float.Parse(_GörevYüküBoylam))[3];
                    //kademe gps irtifa (bizde yok)
                    package[34] = 0x00;
                    package[35] = 0x00;
                    package[36] = 0x00;
                    package[37] = 0x00;
                    //kademe gps enlem (bizde yok)
                    package[38] = 0x00;
                    package[39] = 0x00;
                    package[40] = 0x00;
                    package[41] = 0x00;
                    //kademe gps boylam (bizde yok)
                    package[42] = 0x00;
                    package[43] = 0x00;
                    package[44] = 0x00;
                    package[45] = 0x00;
                    //gyroscope x
                    package[46] = getBytes(float.Parse(_JireskopX))[0];
                    package[47] = getBytes(float.Parse(_JireskopX))[1];
                    package[48] = getBytes(float.Parse(_JireskopX))[2];
                    package[49] = getBytes(float.Parse(_JireskopX))[3];
                    //gyroscope y                      
                    package[50] = getBytes(float.Parse(_JireskopY))[0];
                    package[51] = getBytes(float.Parse(_JireskopY))[1];
                    package[52] = getBytes(float.Parse(_JireskopY))[2];
                    package[53] = getBytes(float.Parse(_JireskopY))[3];
                    //gyroscope z                    
                    package[54] = getBytes(float.Parse(_JireskopZ))[0];
                    package[55] = getBytes(float.Parse(_JireskopZ))[1];
                    package[56] = getBytes(float.Parse(_JireskopZ))[2];
                    package[57] = getBytes(float.Parse(_JireskopZ))[3];
                    //acceleration x                  
                    package[58] = getBytes(float.Parse(_IvmeX))[0];
                    package[59] = getBytes(float.Parse(_IvmeX))[1];
                    package[60] = getBytes(float.Parse(_IvmeX))[2];
                    package[61] = getBytes(float.Parse(_IvmeX))[3];
                    //acceleration y                   
                    package[62] = getBytes(float.Parse(_IvmeY))[0];
                    package[63] = getBytes(float.Parse(_IvmeY))[1];
                    package[64] = getBytes(float.Parse(_IvmeY))[2];
                    package[65] = getBytes(float.Parse(_IvmeY))[3];
                    //acceleration z                  
                    package[66] = getBytes(float.Parse(_IvmeZ))[0];
                    package[67] = getBytes(float.Parse(_IvmeZ))[1];
                    package[68] = getBytes(float.Parse(_IvmeZ))[2];
                    package[69] = getBytes(float.Parse(_IvmeZ))[3];
                    //angle                            
                    package[70] = getBytes(float.Parse(_yunuslamaacisi))[0];
                    package[71] = getBytes(float.Parse(_yunuslamaacisi))[1];
                    package[72] = getBytes(float.Parse(_yunuslamaacisi))[2];
                    package[73] = getBytes(float.Parse(_yunuslamaacisi))[3];
                    //state                          
                    package[74] = getBytes(float.Parse(_Durum))[0];
                    //crc
                    package[75] = calculateCRC(package);
                    package[76] = 0x0D;
                    package[77] = 0x0A;
                    if (sendData) stream_sendDataToHYI.Write(package, 0, package.Length);
                    //udp portundan veri gönderilcek 
                    try
                    {
                        string aciX = JireskopX.Text;
                        string aciY = JireskopY.Text;
                        string aciZ = JireskopZ.Text;

                        string aci = aciX + "," + aciY + "," + aciZ;
                        try { new UdpClient().Send(Encoding.ASCII.GetBytes(aci), Encoding.ASCII.GetBytes(aci).Length, "127.0.0.1", 11000); } catch { }

                    }
                    catch (Exception e)
                    {
                        MessageBox.Show("Hata: " + e.ToString());
                    }

                    dataGridView1.Rows.Add(veriler);
                    pressure_chart.Series["Pressure"].Points.AddXY(Sayac.Text, İrtifa.Text);
                    pressure_chart.Series["P_Pressure"].Points.AddXY(Sayac.Text, GörevYüküGPSirtifa1.Text);
                    velocity_Chart.Series["Velocity"].Points.AddXY(Sayac.Text, hiz.Text);
                    altitude.Series["Altitude"].Points.AddXY(Sayac.Text, GörevYüküGPSirtifa1.Text);
                    altitude.Series["P_Altitude"].Points.AddXY(Sayac.Text, İrtifa.Text);
                    double lat = Convert.ToDouble(roketBoylam.Text.Replace('.', ','));
                    double longt = Convert.ToDouble(roketEnlem.Text.Replace('.', ','));
                    MAP.Position = new GMap.NET.PointLatLng(lat, longt);
                    MAP.MinZoom = 15;
                    MAP.MaxZoom = 15;
                    MAP.Zoom = 15; chromiumWebBrowser1.EvaluateScriptAsync("delLastMark();");
                    chromiumWebBrowser1.EvaluateScriptAsync("setmark(" + roketBoylam.Text + "," + roketEnlem.Text + "," + roketBoylam.Text + "," + roketEnlem.Text + ");");

                }
            }
            catch (Exception e)
            {
                Log(e.Message);
            }

        }
        void Log(string log)
        {
            richTextBox1.Text = "Log:" + log + Environment.NewLine;
        }
        public void openPortToSendData(string port)
        {
            try
            {
                stream_sendDataToHYI.PortName = port;
                stream_sendDataToHYI.BaudRate = 19200;
                stream_sendDataToHYI.Parity = Parity.None;
                stream_sendDataToHYI.DataBits = 8;
                stream_sendDataToHYI.StopBits = StopBits.One;
                stream_sendDataToHYI.Open();
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        byte calculateCRC(byte[] package)
        {
            int check_sum = 0;
            for (int i = 4; i < 75; i++)
            {
                check_sum += package[i];
            }
            return Convert.ToByte(check_sum % 256);
        }

        private byte[] getBytes(float value)
        {
            var buffer = BitConverter.GetBytes(value);
            //if (!BitConverter.IsLittleEndian)
            //{
            //    return buffer;
            //}
            return new[] { buffer[0], buffer[1], buffer[2], buffer[3] };
        }
        string a = "";
        private void SerialPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            int index = a.IndexOf("*");
            if (index >= 0)
            {
                a = a.Substring(index); // 0. indekse kadar olan verileri siliyoruz.
            }
            a += serialPort.ReadExisting();
            if (a.Contains("*"))
            {
                if (a.Contains("+"))
                {
                    DisplayReceivedData(a);
                    a = "";
                }
                else if (a.Length > 120)
                {
                    a = "";
                }
            }
            else if (a.Contains("+") && !a.Contains("*"))
            {
                a = "";
            }
        }
        private void closeButton_Click(object sender, EventArgs e)
        {
            try
            {
                button1.PerformClick();
            }
            catch (Exception) { }
            serialPort.Close();
            CSVOut();
        }
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView1.ColumnCount = 21;
            dataGridView1.Columns[0].Name = "Sayac";
            dataGridView1.Columns[1].Name = "İrtifa";
            dataGridView1.Columns[2].Name = "Roket GPS irtifa";
            dataGridView1.Columns[3].Name = "Roket Boylam";
            dataGridView1.Columns[4].Name = "Roket Enlem";
            dataGridView1.Columns[5].Name = "Görev Yükü GPS irtifa";
            dataGridView1.Columns[6].Name = "Görev Yükü Enlem";
            dataGridView1.Columns[7].Name = "Görev Yükü Boylam ";
            dataGridView1.Columns[8].Name = "JireskopX";
            dataGridView1.Columns[9].Name = "JireskopY";
            dataGridView1.Columns[10].Name = "JireskopZ";
            dataGridView1.Columns[11].Name = "ivmeX";
            dataGridView1.Columns[12].Name = "ivmeY";
            dataGridView1.Columns[13].Name = "ivmeZ";
            dataGridView1.Columns[14].Name = "Durum";
            dataGridView1.Columns[15].Name = "CRC";
            dataGridView1.Columns[16].Name = "Pil Gerilim";
            dataGridView1.Columns[17].Name = "Sıcaklık";
            dataGridView1.Columns[18].Name = "Hız";
            dataGridView1.Columns[19].Name = "Tarih";
            dataGridView1.Columns[20].Name = "Saat";
        }
        bool scaled = false;
        private readonly object Sim;

        private void connectButton_Click(object sender, EventArgs e)
        {
            if (connectButton.Text == "Connect")
            {
                chromiumWebBrowser1.Visible = true;
                MAP.Visible = false;
                connectButton.Text = "Disconnect";
            }
            else
            {
                chromiumWebBrowser1.Visible = false;
                connectButton.Text = "Connect";
                MAP.Visible = true;
            }
        }
        int unitywindowtime = 0;
        private void timer1_Tick(object sender, EventArgs e)
        {
            unitywindowtime++;
            if (unitywindowtime <= 5)
            {
                MoveWindow(simApplication.MainWindowHandle, 0, 0, panel_unity.Width, panel_unity.Height, true);
                SetParent(simApplication.MainWindowHandle, panel_unity.Handle);
                MakeExternalWindowBorderless(simApplication.MainWindowHandle);
            }
            else
            {
                timer1.Enabled = false;
                EnableAllControls(this);
            }
        }
        private void DisableAllControls(System.Windows.Forms.Control parent)
        {
            foreach (System.Windows.Forms.Control control in parent.Controls)
            {
                if (control.Name != connectButton.Name) control.Enabled = false;

                // Eğer kontrolün alt kontrolleri varsa, onları da devre dışı bırak
                if (control.HasChildren)
                {
                    DisableAllControls(control);
                }
            }
        }
        private void EnableAllControls(System.Windows.Forms.Control parent)
        {
            foreach (System.Windows.Forms.Control control in parent.Controls)
            {
                control.Enabled = true;

                // Eğer kontrolün alt kontrolleri varsa, onları da devre dışı bırak
                if (control.HasChildren)
                {
                    EnableAllControls(control);
                }
            }
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                CSVOut();
                Kill("My project (1)");
            }
            catch (Exception)
            { }
        }

        private void CSVOut()
        {
            try
            {
                using (StreamWriter sw = new StreamWriter(@Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"/data.csv"))
                {
                    for (int i = 0; i < dataGridView1.Columns.Count; i++)
                    {
                        sw.Write(dataGridView1.Columns[i].Name);
                        if (i < dataGridView1.Columns.Count - 1)
                        {
                            sw.Write(",");
                        }
                    }
                    sw.WriteLine();
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (!row.IsNewRow)
                        {
                            for (int i = 0; i < dataGridView1.Columns.Count; i++)
                            {
                                sw.Write(row.Cells[i].Value?.ToString());
                                if (i < dataGridView1.Columns.Count - 1)
                                {
                                    sw.Write(",");
                                }
                            }
                            sw.WriteLine();
                        }
                    }
                }
                Log("CSV dosyası başarıyla oluşturuldu.");
            }
            catch (Exception ex)
            {
                Log("CSV dosyası oluşturulurken bir hata oluştu: " + ex.Message);
            }
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            try
            {
                richTextBox1.SelectionStart = richTextBox1.Text.Length;
                richTextBox2.SelectionStart = richTextBox2.Text.Length;
                richTextBox1.ScrollToCaret();
                richTextBox2.ScrollToCaret();
                if (serialPort.IsOpen) richTextBox2.Text += "Data:" + a + Environment.NewLine;
            }
            catch (Exception)
            {
            }
        }
        bool sendData = false;
        public void closePort_sendData()
        {
            if (stream_sendDataToHYI.IsOpen)
            {
                stream_sendDataToHYI.Close();
                sendData = false;
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            sendData = true;
            try
            {
                openPortToSendData(comboBox2.SelectedItem.ToString());
            }
            catch (Exception) { }

        }

        private void timer3_Tick(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized;
            connectButton.PerformClick();
            Thread.Sleep(50);
            this.WindowState = FormWindowState.Minimized;
            Thread.Sleep(50);
            this.WindowState = FormWindowState.Maximized;
            Thread.Sleep(50);
            connectButton.PerformClick();
            Thread.Sleep(50);
            connectButton.PerformClick();
            this.Scale(1.295f);
            timer3.Stop();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            closePort_sendData();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (!richTextBox1.Visible)
            {
                dataGridView1.Visible = false;
                richTextBox1.Visible = true;
                richTextBox2.Visible = true;
            }
            else
            {
                richTextBox1.Visible = false;
                richTextBox2.Visible = false;
                dataGridView1.Visible = true;
            }
        }
    }

}