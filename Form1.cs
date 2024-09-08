using CefSharp;
using CefSharp.DevTools.CacheStorage;
using DocumentFormat.OpenXml.Bibliography;
using GMap.NET;
using GMap.NET.MapProviders;
using GMap.NET.WindowsForms.Markers;
using GMap.NET.WindowsForms;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Management;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data.Entity.Core.Common.CommandTrees.ExpressionBuilder;
namespace arayüz_örnek
{
    public partial class Form1 : Form
    {
        //Global Variables For General Purpose 
        private float groundStationLat = 38.398358f;//Atış Alanı için
        private float groundStationLng = 33.71108f;//Atış Alanı için
        private float vinsanLat = 40.7431761f;//Vinsan Testi için
        private float vinsanLng = 29.9412449f;//Vinsan Testi için
        internal readonly GMapOverlay objects = new GMapOverlay("objects");
        GMapOverlay markers = new GMapOverlay("markers");
        GMapMarker sat;
        GMapMarker station = new GMarkerGoogle(
             new PointLatLng(38.398358, 33.711087),//Aksaray Hisar Atış Alanı Koordinatları
             GMarkerGoogleType.red_dot);
        GMapOverlay polyOverlay = new GMapOverlay("polygons");
        private bool vinsan = false;
        private bool sendData = false;
        private SerialPort serialPortMain;
        private SerialPort serialPortPayload;
        private SerialPort serialPortHYI;
        int mainBaudRate = 9600;
        int payloadBaudRate = 9600;
        int HYIBaudRate = 19200;
        private int? unitywindowtime = 0;
        private byte durum, teamID = 23;
        private SerialPort stream_sendDataToHYI = new SerialPort();
        private string a = "";
        private short paket = 0;
        byte[] HYIPackage = new byte[78];
        //Global Variables For General Purpose
        private static List<byte> tempBuffer = new List<byte>();
        private static List<byte> tempBufferPayload = new List<byte>();
        //Global Variables For Unity
        private Process simApplication;
        private Process unityProcess;
        private IntPtr unityWindowHandle;
        private object packetcount;
        //Global Variables For Unity 
        //Global Functions and Imports For Unity
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
        //Global Functions and Imports For Unity
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
        //Global Functions and Imports For Unity
        public Form1()
        {
            if (vinsan)
            {
                groundStationLat = vinsanLat;
                groundStationLng = vinsanLng;
            }
            InitializeComponent();
            DisableAllControls(this);
            chromiumWebBrowser1.LoadHtml(File.ReadAllText(@Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"/index.html"));
        }
        private void Kill(string app)//Process Kill(Killing Unity Simulation App and Ground Station App)
        {
            foreach (var process in Process.GetProcessesByName(app))
                process.Kill();
        }
        public (float, float, float) QuaternionToEulerAngles(float w, float x, float y, float z)
        {
            // Quaternion'dan Euler açılarını hesaplayan fonksiyon  

            float sinr_cosp = 2 * (w * x + y * z);
            float cosr_cosp = 1 - 2 * (x * x + y * y);
            float roll = (float)Math.Atan2(sinr_cosp, cosr_cosp);

            float sinp = 2 * (w * y - z * x);
            float pitch;
            if (Math.Abs(sinp) >= 1)
                pitch = (float)(Math.Sign(sinp) * (Math.PI / 2)); // pitch = 90 veya -90 derece  
            else
                pitch = (float)Math.Asin(sinp);

            float siny_cosp = 2 * (w * z + x * y);
            float cosy_cosp = 1 - 2 * (y * y + z * z);
            float yaw = (float)Math.Atan2(siny_cosp, cosy_cosp);

            // Dönüşleri dereceye çevir  
            return (roll * (180f / (float)Math.PI), pitch * (180f / (float)Math.PI), yaw * (180f / (float)Math.PI));
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            try//Trying to kill old Unity Simulations and Timer2 tries to put new Unity App into UI
            {
                timer2.Start();
                timer4.Start();
                Kill("3D Sim");
            }
            catch (Exception ex) { Log(ex.ToString()); }
            //DataGridView Columns Name Setting Up
            try
            {
                dataGridView1.ColumnCount = 28;
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
                dataGridView1.Columns[21].Name = "Payload Sicaklik";
                dataGridView1.Columns[22].Name = "Payload İrtifa";
                dataGridView1.Columns[23].Name = "Basinc";
                dataGridView1.Columns[24].Name = "Nem";
                dataGridView1.Columns[25].Name = "Açı X";
                dataGridView1.Columns[26].Name = "Açı Y";
                dataGridView1.Columns[27].Name = "Açı Z";
            }
            catch (Exception ex) { Log(ex.ToString()); }
            //DataGridView Columns Name Setting Up
            ListComPorts();
            //Running up GMap with Cache for Offline use
            if (File.Exists(@Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"/Cache"))
                MAP.CacheLocation = @Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"/Cache";
            GMaps.Instance.Mode = AccessMode.ServerAndCache;
            MAP.MapProvider = GoogleSatelliteMapProvider.Instance;
            MAP.Position = new PointLatLng(station.Position.Lat, station.Position.Lng);//Aksaray Hisar Atış Alanı Koordinatları
            MAP.MinZoom = 3;
            MAP.MaxZoom = 20;
            MAP.Zoom = 17;
            MAP.Manager.CancelTileCaching();
            MAP.HoldInvalidation = false;
            markers.Markers.Add(station);
            MAP.Overlays.Add(markers);
            //Running up GMap with Cache for Offline use
            timer3.Start();//UI positioning fixer 
            //ss alınanacak grafikler
            pressure_chart.MouseClick += new MouseEventHandler(Chart_MouseClick);
            velocity_Chart.MouseClick += new MouseEventHandler(Chart_MouseClick);
            altitude.MouseClick += new MouseEventHandler(Chart_MouseClick);
        }
        void AddSatPointToGMAP(double lat, double lng)
        {
            sat = new GMarkerGoogle(
               new PointLatLng(lat, lng),
               GMarkerGoogleType.red_dot);
            sat.ToolTipMode = MarkerTooltipMode.Always;
            sat.ToolTipText = getDistance(new PointLatLng(sat.Position.Lat, sat.Position.Lng), new PointLatLng(station.Position.Lat, station.Position.Lng)).ToString("N2") + " m";
            markers.Markers.Add(sat);
            MAP.Overlays.Add(markers);
        }
        double _double(string x)
        {
            return Convert.ToDouble(x);
        }
        void UpdateGMap(string lat1, string lng1, string lat2, string lng2)
        {
            polyOverlay.Polygons.Clear();
            markers.Markers.Clear();
            markers.Markers.Add(station);
            AddSatPointToGMAP(_double(lat1), _double(lng1));
            AddPayloadPointToGMAP(_double(lat2), _double(lng2));
            List<PointLatLng> points = new List<PointLatLng>();
            points.Add(new PointLatLng(_double(lat1), _double(lng1)));
            points.Add(new PointLatLng(_double(lat2), _double(lng2)));
            points.Add(new PointLatLng(station.Position.Lat, station.Position.Lng));
            GMapPolygon polygon = new GMapPolygon(points, "mypolygon");
            polygon.Stroke = new Pen(System.Drawing.Color.Red, 3);
            polygon.Fill = new SolidBrush(System.Drawing.Color.Transparent);
            polyOverlay.Polygons.Add(polygon);
            MAP.Overlays.Add(polyOverlay);
            MAP.Refresh();
        }
        void AddPayloadPointToGMAP(double lat, double lng)
        {
            GMapMarker payload = new GMarkerGoogle(
                new PointLatLng(lat, lng),
                GMarkerGoogleType.red_dot);
            payload.ToolTipMode = MarkerTooltipMode.Always;
            payload.ToolTipText = getDistance(new PointLatLng(payload.Position.Lat, payload.Position.Lng), new PointLatLng(sat.Position.Lat, sat.Position.Lng)).ToString("N2") + " m";
            markers.Markers.Add(payload);
            MAP.Overlays.Add(markers);
        }
        //iki mesafe arası uzaklık
        double getDistance(PointLatLng p1, PointLatLng p2)
        {
            var R = 6378137;
            var dLat = rad(p2.Lat - p1.Lat);
            var dLong = rad(p2.Lng - p1.Lng);
            var a = Math.Sin(dLat / 2) * Math.Sin(dLat / 2) +
              Math.Cos(rad(p1.Lat)) * Math.Cos(rad(p2.Lat)) *
              Math.Sin(dLong / 2) * Math.Sin(dLong / 2);
            var c = 2 * Math.Atan2(Math.Sqrt(a), Math.Sqrt(1 - a));
            var d = R * c;
            return d;
        }
        double rad(double x)
        {
            return x * Math.PI / 180;
        }
        //grafiklerin ekran görüntüsünü alma
        private void Chart_MouseClick(object sender, MouseEventArgs e)
        {
            System.Windows.Forms.DataVisualization.Charting.Chart chart = sender as System.Windows.Forms.DataVisualization.Charting.Chart;
            if (chart != null)
            {
                int customWidth = 1920;  // Genişlik
                int customHeight = 1080; // Yükseklik 
                using (Bitmap bmp = new Bitmap(customWidth, customHeight))
                {
                    System.Drawing.Size originalSize = chart.Size;
                    chart.Size = new System.Drawing.Size(customWidth, customHeight);
                    chart.DrawToBitmap(bmp, new Rectangle(0, 0, customWidth, customHeight));
                    chart.Size = originalSize;
                    string directoryPath = @Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"/grafikekrangörüntüleri";
                    string filePath = Path.Combine(directoryPath, $"chart_screenshot_{DateTime.Now.ToString("yyyyMMdd_HHmmss")}.png");
                    if (!System.IO.Directory.Exists(directoryPath))
                        System.IO.Directory.CreateDirectory(directoryPath);
                    bmp.Save(filePath, System.Drawing.Imaging.ImageFormat.Png);
                    MessageBox.Show($"Ekran görüntüsü kaydedildi: {filePath}");
                }
            }
        }
        private void ListComPorts()//List All Com Ports To The ComboBoxes
        {
            MainComboBox.Items.Clear();
            HYIComboBox.Items.Clear();
            PayloadComboBox.Items.Clear();
            using (var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity WHERE Caption like '%(COM%'"))
            {
                var portnames = SerialPort.GetPortNames();
                var ports = searcher.Get().Cast<ManagementBaseObject>().ToList().Select(p => p["Caption"].ToString());

                var portList = portnames.Select(n => n + " - " + ports.FirstOrDefault(s => s.Contains(n))).ToList();

                foreach (string s in portList)
                {
                    MainComboBox.Items.Add(s);
                    HYIComboBox.Items.Add(s);
                    PayloadComboBox.Items.Add(s);
                }
            }
        }
        private void openButton_Click(object sender, EventArgs e)//Start Reading from Comport And Start Sending To The HYI
        {
            try
            {
                HYITimer.Enabled = true;
                HYITimer.Start();
                new Thread(new ThreadStart(Payload)).Start();
                new Thread(new ThreadStart(Main)).Start();
                new Thread(new ThreadStart(HYI)).Start();
                unitywindowtime = -1;
                timer1.Enabled = true;
                timer1.Start();
            }
            catch (Exception ex) { Log(ex.ToString()); }

        }
        void Main()
        {
            try
            {
                serialPortMain = new SerialPort(MainComboBox.SelectedItem.ToString().Split('-')[0], mainBaudRate, Parity.None, 8, StopBits.One);//Serial Port Settings
                serialPortMain.DataReceived += SerialPort_DataReceived;
                paket = short.Parse(Sayac.Text);
                timer2.Start();
                if (!serialPortMain.IsOpen)
                {
                    serialPortMain.PortName = MainComboBox.SelectedItem.ToString().Split('-')[0];
                    serialPortMain.BaudRate = mainBaudRate;
                    serialPortMain.Parity = Parity.None;
                    serialPortMain.DataBits = 8;
                    serialPortMain.StopBits = StopBits.One;
                    serialPortMain.Handshake = Handshake.None;
                    try
                    {
                        serialPortMain.Open();
                    }
                    catch (Exception ex) { Log(ex.ToString()); }

                }
                else
                    Log("Yavaş lan gaç tane basıyon");
            }
            catch (Exception ex) { Log(ex.ToString()); }
        }
        void Payload()
        {
            try
            {
                serialPortPayload = new SerialPort(PayloadComboBox.SelectedItem.ToString().Split('-')[0], payloadBaudRate, Parity.None, 8, StopBits.One);//Serial Port Settings
                serialPortPayload.DataReceived += SerialPortPayload_DataReceived1;
                if (!serialPortPayload.IsOpen)
                {
                    serialPortPayload.PortName = PayloadComboBox.SelectedItem.ToString().Split('-')[0];
                    serialPortPayload.BaudRate = payloadBaudRate;
                    serialPortPayload.Parity = Parity.None;
                    serialPortPayload.DataReceived += SerialPortPayload_DataReceived1;
                    serialPortPayload.DataBits = 8;
                    serialPortPayload.StopBits = StopBits.One;
                    serialPortPayload.Handshake = Handshake.None;
                    try
                    {
                        serialPortPayload.Open();
                    }
                    catch (Exception ex) { Log(ex.ToString()); }

                }
                else
                    Log("Yavaş lan gaç tane basıyon");

            }
            catch (Exception ex) { Log(ex.ToString()); }
        }

        private void SerialPortPayload_DataReceived1(object sender, SerialDataReceivedEventArgs e)
        {
            SerialPort sp = (SerialPort)sender;
            byte[] buffer = new byte[sp.BytesToRead];
            sp.Read(buffer, 0, buffer.Length);
            tempBufferPayload.AddRange(buffer);
            int ffCount = 0;
            int lastFFIndex = -1;
            for (int i = 0; i < tempBufferPayload.Count; i++)
            {
                if (tempBufferPayload[i] == 0xFF)
                {
                    ffCount++;
                    lastFFIndex = i;
                    if (ffCount == 3)
                    {
                        byte[] beforeFFs = tempBufferPayload.Take(lastFFIndex - 2).ToArray();
                        //foreach (byte b in beforeFFs) Log(b.ToString());
                        DisplayReceivedData1(beforeFFs);
                        tempBufferPayload.RemoveRange(0, lastFFIndex + 1);
                        ffCount = 0; // Reset count for next frame
                    }
                }
                else
                    ffCount = 0;
            }
        }
        int xyz = 1;
        private void DisplayReceivedData(byte[] t)
        {
            try
            {
                if (InvokeRequired)
                    BeginInvoke(new Action<byte[]>(DisplayReceivedData), t);
                else
                {
                    if (LogTextBox.Text.Length >= 2000000)
                        LogTextBox.Text = "";
                    paket++;
                    //teamID = t[0];// Karttan gelirse burayı aç, burdan gelmesini istiyorsan burayı kapat
                    string _Sayac = t[1].ToString();
                    string _Irtifa = BitConverter.ToSingle(t, 2).ToString();
                    string _RoketGPSirtifa = BitConverter.ToSingle(t, 6).ToString();
                    string _RoketEnlem = BitConverter.ToSingle(t, 10).ToString();
                    string _RoketBoylam = BitConverter.ToSingle(t, 14).ToString();
                    string _JireskopX = BitConverter.ToSingle(t, 18).ToString();
                    string _JireskopY = BitConverter.ToSingle(t, 22).ToString();
                    string _JireskopZ = BitConverter.ToSingle(t, 26).ToString();
                    string _IvmeX = BitConverter.ToSingle(t, 30).ToString();
                    string _IvmeY = BitConverter.ToSingle(t, 34).ToString();
                    string _IvmeZ = BitConverter.ToSingle(t, 38).ToString();
                    string _AciY = BitConverter.ToSingle(t, 42).ToString();
                    string _Durum = t[46].ToString();
                    string _CRC = t[47].ToString();
                    string _PilGerilim = BitConverter.ToSingle(t, 48).ToString();
                    string _Sicaklik = BitConverter.ToSingle(t, 52).ToString();
                    string _Hiz = BitConverter.ToSingle(t, 56).ToString();
                    string _Basinc = BitConverter.ToSingle(t, 60).ToString();
                    string _AciX = BitConverter.ToSingle(t, 64).ToString();
                    string _AciZ = BitConverter.ToSingle(t, 68).ToString();
                    string _AciW = BitConverter.ToSingle(t, 72).ToString();
                    DateTime now = DateTime.Now;
                    string _Tarih = now.ToString("yyyy-MM-dd");
                    string _Saat = now.ToString("HH:mm:ss");
                    Sayac.Text = _Sayac;
                    Basinc.Text = _Basinc;
                    İrtifa.Text = _Irtifa;
                    RoketGPSirtifa.Text = _RoketGPSirtifa;
                    roketBoylam.Text = _RoketBoylam;
                    roketEnlem.Text = _RoketEnlem;
                    JireskopX.Text = _JireskopX;
                    JireskopY.Text = _JireskopY;
                    JireskopZ.Text = _JireskopZ;
                    AciZ.Text = _AciZ;
                    AciX.Text = _AciX;
                    ivmeX.Text = _IvmeX;
                    ivmeY.Text = _IvmeY;
                    ivmeZ.Text = _IvmeZ;
                    Durum.Text = _Durum;
                    CRC.Text = _CRC;
                    pilgerilim.Text = _PilGerilim;
                    sicaklik.Text = _Sicaklik;
                    hiz.Text = _Hiz;
                    tarih.Text = _Tarih;
                    saat.Text = _Saat;
                    AciY.Text = _AciY;
                    AciW.Text = _AciW;
                    ID.Text = teamID.ToString();
                    for (int i = 0; i < HYIPackage.Length; i++)//Package'da boş değerleri yakalamak için unassigned value olarak FF atıyoruz
                        HYIPackage[i] = 0xFF;
                    //gride dataları ekle 
                    HYIPackage[0] = 0xFF;
                    HYIPackage[1] = 0xFF;
                    HYIPackage[2] = 0x54;
                    HYIPackage[3] = 0x52;
                    byte counter;
                    HYIPackage[4] = teamID; //teamID 
                    HYIPackage[5] = byte.TryParse(_Sayac, out counter) ? counter : (byte)0; //counter 
                    HYIPackage[6] = getBytes(float.Parse(_Irtifa))[0];
                    HYIPackage[7] = getBytes(float.Parse(_Irtifa))[1];
                    HYIPackage[8] = getBytes(float.Parse(_Irtifa))[2];
                    HYIPackage[9] = getBytes(float.Parse(_Irtifa))[3];
                    //gps altitude
                    HYIPackage[10] = getBytes(float.Parse(_RoketGPSirtifa))[0];
                    HYIPackage[11] = getBytes(float.Parse(_RoketGPSirtifa))[1];
                    HYIPackage[12] = getBytes(float.Parse(_RoketGPSirtifa))[2];
                    HYIPackage[13] = getBytes(float.Parse(_RoketGPSirtifa))[3];
                    //gps latitude
                    HYIPackage[14] = getBytes(float.Parse(_RoketEnlem))[0];
                    HYIPackage[15] = getBytes(float.Parse(_RoketEnlem))[1];
                    HYIPackage[16] = getBytes(float.Parse(_RoketEnlem))[2];
                    HYIPackage[17] = getBytes(float.Parse(_RoketEnlem))[3];
                    //gps longitude
                    HYIPackage[18] = getBytes(float.Parse(_RoketBoylam))[0];
                    HYIPackage[19] = getBytes(float.Parse(_RoketBoylam))[1];
                    HYIPackage[20] = getBytes(float.Parse(_RoketBoylam))[2];
                    HYIPackage[21] = getBytes(float.Parse(_RoketBoylam))[3];
                    //kademe gps irtifa (bizde yok)
                    HYIPackage[34] = 0x00;
                    HYIPackage[35] = 0x00;
                    HYIPackage[36] = 0x00;
                    HYIPackage[37] = 0x00;
                    //kademe gps enlem (bizde yok)
                    HYIPackage[38] = 0x00;
                    HYIPackage[39] = 0x00;
                    HYIPackage[40] = 0x00;
                    HYIPackage[41] = 0x00;
                    //kademe gps boylam (bizde yok)
                    HYIPackage[42] = 0x00;
                    HYIPackage[43] = 0x00;
                    HYIPackage[44] = 0x00;
                    HYIPackage[45] = 0x00;
                    //gyroscope x
                    HYIPackage[46] = getBytes(float.Parse(_JireskopX))[0];
                    HYIPackage[47] = getBytes(float.Parse(_JireskopX))[1];
                    HYIPackage[48] = getBytes(float.Parse(_JireskopX))[2];
                    HYIPackage[49] = getBytes(float.Parse(_JireskopX))[3];
                    //gyroscope y                      
                    HYIPackage[50] = getBytes(float.Parse(_JireskopY))[0];
                    HYIPackage[51] = getBytes(float.Parse(_JireskopY))[1];
                    HYIPackage[52] = getBytes(float.Parse(_JireskopY))[2];
                    HYIPackage[53] = getBytes(float.Parse(_JireskopY))[3];
                    //gyroscope z                    
                    HYIPackage[54] = getBytes(float.Parse(_JireskopZ))[0];
                    HYIPackage[55] = getBytes(float.Parse(_JireskopZ))[1];
                    HYIPackage[56] = getBytes(float.Parse(_JireskopZ))[2];
                    HYIPackage[57] = getBytes(float.Parse(_JireskopZ))[3];
                    //acceleration x                  
                    HYIPackage[58] = getBytes(float.Parse(_IvmeX))[0];
                    HYIPackage[59] = getBytes(float.Parse(_IvmeX))[1];
                    HYIPackage[60] = getBytes(float.Parse(_IvmeX))[2];
                    HYIPackage[61] = getBytes(float.Parse(_IvmeX))[3];
                    //acceleration y                   
                    HYIPackage[62] = getBytes(float.Parse(_IvmeY))[0];
                    HYIPackage[63] = getBytes(float.Parse(_IvmeY))[1];
                    HYIPackage[64] = getBytes(float.Parse(_IvmeY))[2];
                    HYIPackage[65] = getBytes(float.Parse(_IvmeY))[3];
                    //acceleration z                  
                    HYIPackage[66] = getBytes(float.Parse(_IvmeZ))[0];
                    HYIPackage[67] = getBytes(float.Parse(_IvmeZ))[1];
                    HYIPackage[68] = getBytes(float.Parse(_IvmeZ))[2];
                    HYIPackage[69] = getBytes(float.Parse(_IvmeZ))[3];
                    //angle                            
                    HYIPackage[70] = getBytes(float.Parse(_AciY))[0];
                    HYIPackage[71] = getBytes(float.Parse(_AciY))[1];
                    HYIPackage[72] = getBytes(float.Parse(_AciY))[2];
                    HYIPackage[73] = getBytes(float.Parse(_AciY))[3];
                    //state                          
                    HYIPackage[74] = byte.TryParse(_Durum, out durum) ? durum : (byte)0;
                    HYIPackage[76] = 0x0D;
                    HYIPackage[77] = 0x0A;
                    try
                    {
                        string aciX = AciX.Text.Replace(".", ",");
                        string aciY = AciY.Text.Replace(".", ",");
                        string aciZ = AciZ.Text.Replace(".", ",");
                        string aciW = AciW.Text.Replace(".", ",");
                        //var eulerAngles = QuaternionToEulerAngles(float.Parse(aciW), float.Parse(aciX), float.Parse(aciY), float.Parse(aciZ));
                        //string aci = (eulerAngles.Item3 * -1).ToString().Replace(",",".") + "," + (eulerAngles.Item2 * -1).ToString().Replace(",", ".") + "," + (eulerAngles.Item1 * -1).ToString().Replace(",", ".");
                        string aci = "";
                        //unitydeki roketın x y z sini döndürüyor
                        switch (xyz)
                        {
                            case 1:
                                aci = aciX + "*" + aciY + "*" + aciZ + "*" + aciW;
                                try { new UdpClient().Send(Encoding.ASCII.GetBytes(aci), Encoding.ASCII.GetBytes(aci).Length, "127.0.0.1", 11000); }//Sending 3D Angle Datas to Unity Simulation
                                catch (Exception ex) { Log(ex.ToString()); }
                                break;
                            case 2:
                                aci = aciX + "*" + aciZ + "*" + aciY + "*" + aciW;
                                try { new UdpClient().Send(Encoding.ASCII.GetBytes(aci), Encoding.ASCII.GetBytes(aci).Length, "127.0.0.1", 11000); }//Sending 3D Angle Datas to Unity Simulation
                                catch (Exception ex) { Log(ex.ToString()); }
                                break;
                            case 3:
                                aci = aciY + "*" + aciX + "*" + aciZ + "*" + aciW;
                                try { new UdpClient().Send(Encoding.ASCII.GetBytes(aci), Encoding.ASCII.GetBytes(aci).Length, "127.0.0.1", 11000); }//Sending 3D Angle Datas to Unity Simulation
                                catch (Exception ex) { Log(ex.ToString()); }
                                break;
                            case 4:
                                aci = aciY + "*" + aciZ + "*" + aciX + "*" + aciW;
                                try { new UdpClient().Send(Encoding.ASCII.GetBytes(aci), Encoding.ASCII.GetBytes(aci).Length, "127.0.0.1", 11000); }//Sending 3D Angle Datas to Unity Simulation
                                catch (Exception ex) { Log(ex.ToString()); }
                                break;
                            case 5:
                                aci = aciZ + "*" + aciX + "*" + aciY + "*" + aciW;
                                try { new UdpClient().Send(Encoding.ASCII.GetBytes(aci), Encoding.ASCII.GetBytes(aci).Length, "127.0.0.1", 11000); }//Sending 3D Angle Datas to Unity Simulation
                                catch (Exception ex) { Log(ex.ToString()); }
                                break;
                            case 6:
                                aci = aciZ + "*" + aciY + "*" + aciX + "*" + aciW;
                                try { new UdpClient().Send(Encoding.ASCII.GetBytes(aci), Encoding.ASCII.GetBytes(aci).Length, "127.0.0.1", 11000); }//Sending 3D Angle Datas to Unity Simulation
                                catch (Exception ex) { Log(ex.ToString()); }
                                break;
                            default:
                                xyz = 1;
                                break;
                        }
                        Log(aci);
                        label23.Visible = true;
                        //label23.Text = aci;
                        
                    }
                    catch (Exception ex) { Log(ex.ToString()); }
                    pressure_chart.Series["Pressure"].Points.AddXY(tarih.Text + saat.Text, Basinc.Text);
                    pressure_chart.Series["P_Pressure"].Points.AddXY(tarih.Text + saat.Text, P_Basinc.Text);
                    velocity_Chart.Series["Velocity"].Points.AddXY(tarih.Text + saat.Text, hiz.Text);
                    altitude.Series["Altitude"].Points.AddXY(tarih.Text + saat.Text, İrtifa.Text);
                    altitude.Series["P_Altitude"].Points.AddXY(tarih.Text + saat.Text, P_Altitude.Text);
                    GC.Collect();
                    if (PayloadComboBox.SelectedIndex == -1)
                    {
                        string[] row = new string[]
                        {
                            Sayac.Text,
                            İrtifa.Text,
                            RoketGPSirtifa.Text,
                            roketBoylam.Text,
                            roketEnlem.Text,
                            GörevYüküGPSirtifa1.Text,
                            GörevYüküEnlem.Text,
                            GörevYüküBoylam.Text,
                            JireskopX.Text,
                            JireskopY.Text,
                            JireskopZ.Text,
                            ivmeX.Text,
                            ivmeY.Text,
                            ivmeZ.Text,
                            Durum.Text,
                            CRC.Text,
                            pilgerilim.Text,
                            sicaklik.Text,
                            hiz.Text,
                            tarih.Text,
                            saat.Text,
                            P_Sicaklik.Text,
                            P_Altitude.Text,
                            P_Basinc.Text,
                            Nem.Text,
                            AciX.Text,
                            AciY.Text,
                            AciZ.Text
                        };
                        dataGridView1.Rows.Add(row);
                        chromiumWebBrowser1.EvaluateScriptAsync("delLastMark();");//GoogleMap(not GMap) Delete Marks
                        chromiumWebBrowser1.EvaluateScriptAsync("setmark(" + roketEnlem.Text.Replace(",", ".") + "," + roketBoylam.Text.Replace(",", ".") + "," + GörevYüküEnlem.Text.Replace(",", ".") + "," + GörevYüküBoylam.Text.Replace(",", ".") + ");");//GoogleMap Add Marks
                        double lat = _double(roketBoylam.Text.Replace('.', ','));
                        double longt = _double(roketEnlem.Text.Replace('.', ','));
                        double _lat = _double(GörevYüküEnlem.Text.Replace('.', ','));
                        double _longt = _double(GörevYüküBoylam.Text.Replace('.', ','));
                        UpdateGMap(roketEnlem.Text.Replace(".", ","), roketBoylam.Text.Replace(".", ","), GörevYüküEnlem.Text.Replace(".", ","), GörevYüküBoylam.Text.Replace(".", ","));
                        label23.Text = "setmark(" + roketEnlem.Text.Replace(",", ".") + "," + roketBoylam.Text.Replace(",", ".") + "," + GörevYüküEnlem.Text.Replace(",", ".") + "," + GörevYüküBoylam.Text.Replace(",", ".") + ");";
                    }
                }

            }
            catch (Exception ex) { Log(ex.ToString()); }
        }
        bool PackageCompleted(byte[] array)
        {
            if (HYIPackage[22] == 0xFF && HYIPackage[23] == 0xFF && HYIPackage[24] == 0xFF && HYIPackage[25] == 0xFF)
                return false;
            else
                return true;
        }
        private void DisplayReceivedData1(byte[] t)
        {
            try
            {
                if (InvokeRequired)
                    BeginInvoke(new Action<byte[]>(DisplayReceivedData1), t);
                else
                {
                    string _GörevYüküGPSirtifa = BitConverter.ToSingle(t, 0).ToString();
                    string _GörevYüküEnlem = BitConverter.ToSingle(t, 4).ToString();
                    string _GörevYüküBoylam = BitConverter.ToSingle(t, 8).ToString();
                    string _Sicaklik = BitConverter.ToSingle(t, 12).ToString();
                    string _Basinc = BitConverter.ToSingle(t, 16).ToString();
                    string _Nem = BitConverter.ToSingle(t, 20).ToString();
                    string _İrtifa = BitConverter.ToSingle(t, 24).ToString();
                    GörevYüküGPSirtifa1.Text = _GörevYüküGPSirtifa;
                    GörevYüküEnlem.Text = _GörevYüküEnlem;
                    GörevYüküBoylam.Text = _GörevYüküBoylam;
                    P_Sicaklik.Text = _Sicaklik;
                    P_Basinc.Text = _Basinc;
                    Nem.Text = _Nem;
                    P_Altitude.Text = _İrtifa;
                    //payload gps altitude
                    HYIPackage[22] = getBytes(float.Parse(_GörevYüküGPSirtifa))[0];
                    HYIPackage[23] = getBytes(float.Parse(_GörevYüküGPSirtifa))[1];
                    HYIPackage[24] = getBytes(float.Parse(_GörevYüküGPSirtifa))[2];
                    HYIPackage[25] = getBytes(float.Parse(_GörevYüküGPSirtifa))[3];
                    //payload gps latitude             
                    HYIPackage[26] = getBytes(float.Parse(_GörevYüküEnlem))[0];
                    HYIPackage[27] = getBytes(float.Parse(_GörevYüküEnlem))[1];
                    HYIPackage[28] = getBytes(float.Parse(_GörevYüküEnlem))[2];
                    HYIPackage[29] = getBytes(float.Parse(_GörevYüküEnlem))[3];
                    //payload gps longitude           
                    HYIPackage[30] = getBytes(float.Parse(_GörevYüküBoylam))[0];
                    HYIPackage[31] = getBytes(float.Parse(_GörevYüküBoylam))[1];
                    HYIPackage[32] = getBytes(float.Parse(_GörevYüküBoylam))[2];
                    HYIPackage[33] = getBytes(float.Parse(_GörevYüküBoylam))[3];
                    string[] row = new string[]
                    {
                        Sayac.Text,
                        İrtifa.Text,
                        RoketGPSirtifa.Text,
                        roketBoylam.Text,
                        roketEnlem.Text,
                        GörevYüküGPSirtifa1.Text,
                        GörevYüküEnlem.Text,
                        GörevYüküBoylam.Text,
                        JireskopX.Text,
                        JireskopY.Text,
                        JireskopZ.Text,
                        ivmeX.Text,
                        ivmeY.Text,
                        ivmeZ.Text,
                        Durum.Text,
                        CRC.Text,
                        pilgerilim.Text,
                        sicaklik.Text,
                        hiz.Text,
                        tarih.Text,
                        saat.Text,
                        P_Sicaklik.Text,
                        P_Altitude.Text,
                        P_Basinc.Text,
                        Nem.Text,
                        AciX.Text,
                        AciY.Text,
                        AciZ.Text
                    };
                    dataGridView1.Rows.Add(row);
                    chromiumWebBrowser1.EvaluateScriptAsync("delLastMark();");//GoogleMap(not GMap) Delete Marks
                    chromiumWebBrowser1.EvaluateScriptAsync("setmark(" + roketEnlem.Text.Replace(",", ".") + "," + roketBoylam.Text.Replace(",", ".") + "," + GörevYüküEnlem.Text.Replace(",", ".") + "," + GörevYüküBoylam.Text.Replace(",", ".") + ");");//GoogleMap Add Marks
                    double lat = _double(roketBoylam.Text.Replace('.', ','));
                    double longt = _double(roketEnlem.Text.Replace('.', ','));
                    double _lat = _double(GörevYüküEnlem.Text.Replace('.', ','));
                    double _longt = _double(GörevYüküBoylam.Text.Replace('.', ','));
                    UpdateGMap(roketEnlem.Text.Replace(".", ","), roketBoylam.Text.Replace(".", ","), GörevYüküEnlem.Text.Replace(".", ","), GörevYüküBoylam.Text.Replace(".", ","));
                }
            }
            catch (Exception ex)
            { Log(ex.ToString()); }
        }
        void Log(string log)
        {
            HataTextBox.Text += "Log:" + log + Environment.NewLine;
        }
        public void openPortToSendData(string port)
        {
            try
            {
                stream_sendDataToHYI.PortName = port;
                stream_sendDataToHYI.BaudRate = HYIBaudRate;
                stream_sendDataToHYI.Parity = Parity.None;
                stream_sendDataToHYI.DataBits = 8;
                stream_sendDataToHYI.StopBits = StopBits.One;
                stream_sendDataToHYI.Open();
            }
            catch (Exception ex) { Log(ex.ToString()); }
        }
        byte calculateCRC(byte[] package)
        {
            int check_sum = 0;
            for (int i = 4; i < 75; i++)
                check_sum += package[i];
            return Convert.ToByte(check_sum % 256);
        }
        private byte[] getBytes(float value)
        {
            var buffer = BitConverter.GetBytes(value);
            //if (!BitConverter.IsLittleEndian) 
            //    return buffer;
            return new[] { buffer[0], buffer[1], buffer[2], buffer[3] };
        }
        private void SerialPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            SerialPort sp = (SerialPort)sender;
            byte[] buffer = new byte[sp.BytesToRead];
            sp.Read(buffer, 0, buffer.Length);
            tempBuffer.AddRange(buffer);
            int ffCount = 0;
            int lastFFIndex = -1;
            for (int i = 0; i < tempBuffer.Count; i++)
            {
                if (tempBuffer[i] == 0xFF)
                {
                    ffCount++;
                    lastFFIndex = i;
                    if (ffCount == 3)
                    {
                        byte[] beforeFFs = tempBuffer.Take(lastFFIndex - 2).ToArray();
                        DisplayReceivedData(beforeFFs);
                        tempBuffer.RemoveRange(0, lastFFIndex + 1);
                        ffCount = 0; // Reset count for next frame
                    }
                }
                else
                    ffCount = 0;
            }
        }
        private void closeButton_Click(object sender, EventArgs e)
        {
            try
            {
                serialPortMain.Close();
                serialPortPayload.Close();
                serialPortHYI.Close();
                HYITimer.Stop();
                CSVOut();
            }
            catch (Exception ex) { Log(ex.ToString()); }
            timer2.Stop();
            try
            {
                closePort_sendData();
            }
            catch (Exception ex) { Log(ex.ToString()); }
        }
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
        private void timer1_Tick(object sender, EventArgs e)//Unity Simulation Into panel_unity Timer
        {
            if (unitywindowtime == -1)
            {
                try//Unity Simulation App Starts And Putting into panel_unity on UI
                {
                    string _3DSimExePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\3D\3D Sim";
                    ProcessStartInfo startInfo = new ProcessStartInfo();
                    startInfo.FileName = _3DSimExePath;
                    startInfo.WindowStyle = ProcessWindowStyle.Maximized;
                    simApplication = Process.Start(startInfo);
                    simApplication.WaitForInputIdle();
                    MoveWindow(simApplication.MainWindowHandle, 0, 0, panel_unity.Width, panel_unity.Height, true);
                    SetParent(simApplication.MainWindowHandle, panel_unity.Handle);
                    MakeExternalWindowBorderless(simApplication.MainWindowHandle);
                    unitywindowtime = 0;
                }
                catch (Exception ex) { Log(ex.ToString()); }
            }
            Log("Timer1" + unitywindowtime.ToString());
            unitywindowtime++;
            if (unitywindowtime <= 5)
            {
                try
                {
                    MoveWindow(simApplication.MainWindowHandle, 0, 0, panel_unity.Width, panel_unity.Height, true);
                    SetParent(simApplication.MainWindowHandle, panel_unity.Handle);
                    MakeExternalWindowBorderless(simApplication.MainWindowHandle);
                }
                catch (Exception) { Log("Hata!"); }
            }
            else
            {
                timer1.Enabled = false;
                timer1.Stop();
                unitywindowtime = null;
                EnableAllControls(this);
            }
        }
        private void DisableAllControls(System.Windows.Forms.Control parent)//Disable All Elements (for Positioning)
        {
            foreach (System.Windows.Forms.Control control in parent.Controls)
            {
                if (control.Name != connectButton.Name) control.Enabled = false;
                if (control.HasChildren)
                    DisableAllControls(control);
            }
        }
        private void EnableAllControls(System.Windows.Forms.Control parent)//Enable All Elements (for Positioning)
        {
            foreach (System.Windows.Forms.Control control in parent.Controls)
            {
                control.Enabled = true;
                if (control.HasChildren)
                    EnableAllControls(control);
            }
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                CSVOut();
                Kill(System.AppDomain.CurrentDomain.FriendlyName);
            }
            catch (Exception ex) { Log(ex.ToString()); }
        }
        private void CSVOut()//CSV Output To The Desktop
        {
            try
            {
                using (StreamWriter sw = new StreamWriter(@Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"/data.csv"))
                {
                    for (int i = 0; i < dataGridView1.Columns.Count; i++)
                    {
                        sw.Write(dataGridView1.Columns[i].Name);
                        if (i < dataGridView1.Columns.Count - 1)
                            sw.Write(",");
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
                                    sw.Write(",");
                            }
                            sw.WriteLine();
                        }
                    }
                }
                Log("CSV dosyası başarıyla oluşturuldu.");
            }
            catch (Exception ex) { Log(ex.ToString()); }
        }
        private void timer2_Tick(object sender, EventArgs e)//Serial Monitor Timer, Always Scrol Down Log TextBox and DataGridView
        {
            try
            {
                HataTextBox.SelectionStart = HataTextBox.Text.Length;
                LogTextBox.SelectionStart = LogTextBox.Text.Length;
                HataTextBox.ScrollToCaret();
                LogTextBox.ScrollToCaret();
                dataGridView1.FirstDisplayedScrollingRowIndex = dataGridView1.Rows.Count - 1;
            }
            catch (Exception ex) { Log(ex.ToString()); }
        }
        public void closePort_sendData()
        {
            if (stream_sendDataToHYI.IsOpen)
            {
                stream_sendDataToHYI.Close();
                sendData = false;
            }
        }
        private void HYI()
        {
            try
            {
                Log("HYI");
                openPortToSendData(HYIComboBox.SelectedItem.ToString().Split('-')[0]);
                sendData = true;
            }
            catch (Exception ex)
            {
                Log(ex.ToString());
                sendData = false;
            }
        }
        private void timer3_Tick(object sender, EventArgs e)//Position All The App(Gmap Connect/Disconnect Bug fixer)
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
        private void button3_Click(object sender, EventArgs e)//Switch Button
        {
            if (!HataTextBox.Visible)
            {
                dataGridView1.Visible = false;
                HataTextBox.Visible = true;
                LogTextBox.Visible = true;
            }
            else
            {
                HataTextBox.Visible = false;
                LogTextBox.Visible = false;
                dataGridView1.Visible = true;
            }
        }
        private void timer4_Tick(object sender, EventArgs e)
        {
            PAKET.Text = paket.ToString();
            paket = 0;
        }
        private void HYITimer_Tick(object sender, EventArgs e)
        {
            if (sendData && PackageCompleted(HYIPackage))
            {
                //crc
                HYIPackage[75] = calculateCRC(HYIPackage);
                stream_sendDataToHYI.Write(HYIPackage, 0, HYIPackage.Length);
                Log("HYI'ya Package Gönderildi: " + PackageCompleted(HYIPackage).ToString());
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try//Unity Simulation App Starts And Putting into panel_unity on UI
            {
                string _3DSimExePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\3D\3D Sim";
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.FileName = _3DSimExePath;
                startInfo.WindowStyle = ProcessWindowStyle.Maximized;
                simApplication = Process.Start(startInfo);
                simApplication.WaitForInputIdle();
                MoveWindow(simApplication.MainWindowHandle, 0, 0, panel_unity.Width, panel_unity.Height, true);
                SetParent(simApplication.MainWindowHandle, panel_unity.Handle);
                MakeExternalWindowBorderless(simApplication.MainWindowHandle);
                unitywindowtime = 0;
                timer1.Enabled = true;
                timer1.Start();
            }
            catch (Exception ex) { Log(ex.ToString()); }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if(xyz==6)
                xyz = 1;
            else
                xyz++;
        }

        private void RefreshButton_Click(object sender, EventArgs e)//Refresh Com Ports Button
        {
            try
            { ListComPorts(); }
            catch (Exception ex) { Log(ex.ToString()); }
        }
    }
}