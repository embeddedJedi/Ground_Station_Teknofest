using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

// Kütüphane //
using System.IO.Ports; // PORT AYARI İÇİN
using ClosedXML.Excel; // Excel dosyası kaydetmek için
using System.IO; // PORT AYARI İÇİN 
using GMap.NET.MapProviders; // Map kontrolu için
using System.Xml.Linq;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Web.UI.WebControls; // WebBrowser için
using DocumentFormat.OpenXml.Vml; // Excel dosyası formatı
using System.Windows.Media.Animation;
using DocumentFormat.OpenXml.Drawing;
using System.Reflection.Emit;
using SharpGL;
using SharpGL.SceneGraph.Assets;
using SharpGL.SceneGraph.Effects;
using SharpGL.SceneGraph.Primitives;
using Org.BouncyCastle.Bcpg;
using DocumentFormat.OpenXml.Drawing.Charts;
using Microsoft.Web.WebView2.WinForms;

// Kütüphane Son//


namespace Yer_İstasyonu_Uygulaması

{

    public partial class Form1 : Form
    {

        private bool stop = true;

        private Texture texture;

        private Timer _timer = new Timer();

        string hyi;
        string[] porthyi;
        string hyip;
        private WebView2 webView2Control;

        public Form1()
        {
            InitializeComponent();
            InitializeWebView2();

            /*_timer.Interval = 3000; // 3 saniye aralıklarla tetiklenir
            _timer.Tick += Timer_Tick;*/
        }
        private async void InitializeWebView2()
        {
            webView2Control = new WebView2
            {
                Dock = DockStyle.Fill
            };
            this.Controls.Add(webView21);

            await webView21.EnsureCoreWebView2Async(null);
            webView21.Source = new Uri("https://www.google.com/maps");
        }

        bool mov; // Pencere haraketi için değişken
        int movx, movy; // Pencere haraketi için değişken

        public string irtifa, uydu, gps_irtifa = "1";
        public string enlem, boylam, x, y = "3";
        string[] value;

        private string data; // data veri

        int durum;

        private SerialPort serialPort;
        private void panel3_Paint(object sender, PaintEventArgs e) { }

        private void Form1_Load(object sender, EventArgs e)
        {

            MessageBox.Show("Magnetar 2023 Yer İstasyonu Uygulaması"); // Mesaj

            /*map.DragButton = MouseButtons.Left; // GMAP HARAKET
            map.MapProvider = GMapProviders.GoogleMap; // GMAP HARAKET */

            string[] ports = SerialPort.GetPortNames(); // Portları al

            foreach (string port in ports)
            {
                comboBox1.Items.Add(port); // Portları ekle
            }

            serialPort1.DataReceived += new SerialDataReceivedEventHandler(SerialPort_data); // Dataları al


            porthyi = SerialPort.GetPortNames(); // Portları al

            foreach (string hyi in porthyi)
            {
                comboBox2.Items.Add(hyi); // Portları ekle
            }

            hyip = hyi;

            _timer.Interval = 3000;
            _timer.Tick += Timer_Tick;

            textBox6.Text = "1"; 

        }

        private void SerialPort_data(object sender, SerialDataReceivedEventArgs e)
    {

            data = serialPort1.ReadLine(); // Sürekli veri okumak için
            this.Invoke(new EventHandler(displaydata));

        }

        private void displaydata(object sender, EventArgs e)
        {
            value = data.Split('/'); // Verileri ayırt etmek için

            textBox10.Text = value[0];
            textBox9.Text = value[1];
            textBox7.Text = value[2];
            textBox1.Text = value[3];
            textBox2.Text = value[4];
            textBox5.Text = value[5];
            textBox6.Text = value[6];

            x = value[0];
            y = value[1];
            irtifa = value[2];
            gps_irtifa = value[3];
            enlem = value[4];   
            boylam = value[5];
            uydu = value[6];

            
        }

        private void button1_Click(object sender, EventArgs e) { }

        private void gMapControl1_Load(object sender, EventArgs e) {}

        private void button1_Click_1(object sender, EventArgs e)
        {
            string enlem = textBox4.Text;
            string boylam = textBox3.Text;

            string konum = "https://www.google.com/maps/search/?api=1&query=" + enlem + "," + boylam; // Google Maps konum bulmak için link oluştur

            webView2Control.Source = new Uri(konum); // WebView2 kontrolüne URL'yi yükle
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Close(); // X butonu kapatmak için
        }

        private void textBox1_TextChanged(object sender, EventArgs e) {}

        private void label4_Click(object sender, EventArgs e) { }

        private void panel1_MouseMove(object sender, MouseEventArgs e)
        {
            if(mov == true)
            {
                this.Left = Cursor.Position.X - movx; // Panele mouse u getince anlaması için
                this.Top = Cursor.Position.Y - movy; // Panele mouse u getince anlaması için

            }
        }

        private void panel1_MouseUp(object sender, MouseEventArgs e)
        {
            mov = false;
        }

        private void label9_Click(object sender, EventArgs e) {}

        private void label11_Click(object sender, EventArgs e) {}

        private void button2_Click(object sender, EventArgs e)
        {

            durum = 0;
            try
            {
                // Portu kapat //

                serialPort1.Close(); 
                button3.Enabled = true; 
                button2.Enabled = false;
                label12.Text = "Bağlantı Yok";
                label12.ForeColor = Color.Red;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, ("Hata!")); // Aksi durumda hata mesajı
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            durum = 1;
            try
            {
                // Portu Aç //

                serialPort1.PortName = comboBox1.Text; // comboBoxdaki seçili Portu açması için
                serialPort1.Open();
                button3.Enabled = false;
                button2.Enabled = true;
                label12.Text = "Bağlandı";
                label12.ForeColor = Color.Green;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, ("Hata!")); // Aksi bi durumdaki hata mesajı
            }
        }

        private void label12_Click(object sender, EventArgs e) { }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            // Eğer portu kapatmadıysak Programı kapatınca kendi otomatik Portu kapatır

            if (serialPort1.IsOpen)
                serialPort1.Close();
        }
        

        private void label7_Click(object sender, EventArgs e) {}

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e) {}

        private void label21_Click(object sender, EventArgs e) {}

        private void label5_Click(object sender, EventArgs e) {}

        private void label10_Click(object sender, EventArgs e) {}

        private void label8_Click(object sender, EventArgs e) {}

        private void label6_Click(object sender, EventArgs e) {}

        private async void Timer_Tick(object sender, EventArgs e)
        {
            int row = 2;
            while (!stop)
            {

                string[] veri = data.Split(','); // Verileri ayırt etmek için

                
                var workbook = new XLWorkbook("C:\\Users\\atzml\\Desktop\\eData.xlsx"); // Excel dosyasını oluştur
                var worksheet = workbook.Worksheet("Data"); // Sayfayı denetle

                // Başlıkları ekle
                worksheet.Cell(1, 1).Value = "X Açı";
                worksheet.Cell(1, 2).Value = "Y Açı";
                worksheet.Cell(1, 3).Value = "İrtifa";
                worksheet.Cell(1, 4).Value = "GPS İrtifa";
                worksheet.Cell(1, 5).Value = "GPS Enlem";
                worksheet.Cell(1, 6).Value = "GPS Boylam";
                worksheet.Cell(1, 7).Value = "GPS Uydu";

                // Verileri satır sayacı ile kaydet
                worksheet.Cell(row, 1).Value = veri[0];
                worksheet.Cell(row, 2).Value = veri[1];
                worksheet.Cell(row, 3).Value = veri[2];
                worksheet.Cell(row, 4).Value = veri[3];
                worksheet.Cell(row, 5).Value = veri[4];
                worksheet.Cell(row, 6).Value = veri[5];
                worksheet.Cell(row, 7).Value = veri[6];

                // Dosyayı kaydet
                workbook.Save();

                await Task.Delay(TimeSpan.FromSeconds(2));

                row += 2;

            }

        }

        private void button9_Click(object sender, EventArgs e) {

            string enlem = textBox4.Text;
            string boylam = textBox3.Text;

            string konum = "https://www.google.com/maps/search/?api=1&query=" + enlem + "," + boylam; // Link oluştur

            webView21.Source = new Uri(konum); // WebView2 kontrolüne URL'yi yükle

        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) { }

        private void button6_Click(object sender, EventArgs e)
        {
            stop = false; // Döngüyü başlat
                          
            Timer timer = new Timer();
            timer.Interval = 3000; // 3 saniye aralık
            timer.Tick += Timer_Tick;
            timer.Start();

            MessageBox.Show("Saving...");

        }

        private void button7_Click(object sender, EventArgs e)
        {
            stop = true; // Döngüyü durdur
            _timer.Stop(); // Timer'ı durdur

            MessageBox.Show("Stopped");
        }

        private void button8_Click(object sender, EventArgs e)
        {
            Form2 form2 = new Form2();
            form2.Show();  // form2 göster diyoruz
            this.Hide();   // bu yani form1 gizle diyoruz
        }

        private void guna2ControlBox1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void guna2HtmlToolTip1_Popup(object sender, PopupEventArgs e)
        {

        }

        private void guna2PictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            string enlem = textBox4.Text;
            string boylam = textBox3.Text;

            string konum = "https://www.google.com/maps/search/?api=1&query=" + enlem + "," + boylam; // Link oluştur

            DialogResult Soru; // Evet & Hayır lı Mesaj kutusu 

            Soru = MessageBox.Show(enlem + " Enlem ve " + boylam + " Boylamına gitmek istermisin? ?", "Uyarı",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);

            if (Soru == DialogResult.Yes)
            {
                System.Diagnostics.Process.Start(konum); // Evet i seçerse Linki açar
            }

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void guna2Shapes1_Click(object sender, EventArgs e)
        {

        }

        private void guna2PictureBox2_Click(object sender, EventArgs e)
        {

        }

        // Union kullanimi ile alakali daha fazla bilgi icin: https://www.learn-c.org/en/Unions
        public struct FLOAT32_UINT8_DONUSTURUCU
        {
            public float sayi;
            public byte[] array;
        }

        // Olusturalacak paketin global tanimlanmasi. Eger farkli bir kullanim planliyorsaniz
        // bu degiskenin tanimlandigi yeri ve ismini degistirebilirsiniz. Bu durumda paket_olustur
        // fonksiyonunu da guncellemeniz gerektigini unutmayin.

        public static byte[] olusturalacak_paket = new byte[78];

        // global olarak alinan olusturalacak_paket array'inin check_sum'ini hesaplar.
        public static byte check_sum_hesapla()
        {
            int check_sum = 0;
            for (int i = 4; i < 75; i++)
            {
                check_sum += olusturalacak_paket[i];
            }
            return (byte)(check_sum % 256);

           
        }

        // olusturalacak_paket array'inin icini gunceller ve check_sum hesaplamasi yapar.
        public void packet()
        {
           
            olusturalacak_paket[0] = 0xFF; // Sabit
            olusturalacak_paket[1] = 0xFF; // Sabit
            olusturalacak_paket[2] = 0x54; // Sabit
            olusturalacak_paket[3] = 0x52; // Sabit

            olusturalacak_paket[4] = 0; // Takim ID = 0
            olusturalacak_paket[5] = 0; // Sayac degeri = 0

            FLOAT32_UINT8_DONUSTURUCU roket_irtifa_irtifa_float32_uint8_donusturucu;
            roket_irtifa_irtifa_float32_uint8_donusturucu.sayi = float.Parse(irtifa); // Roket boylam degerinin atamasini yapiyoruz.
            roket_irtifa_irtifa_float32_uint8_donusturucu.array = BitConverter.GetBytes(roket_irtifa_irtifa_float32_uint8_donusturucu.sayi);
            Array.Copy(roket_irtifa_irtifa_float32_uint8_donusturucu.array, 0, olusturalacak_paket, 6, 4);

            FLOAT32_UINT8_DONUSTURUCU roket_gps_irtifa_float32_uint8_donusturucu;
            roket_gps_irtifa_float32_uint8_donusturucu.sayi = float.Parse(gps_irtifa); // Roket boylam degerinin atamasini yapiyoruz.
            roket_gps_irtifa_float32_uint8_donusturucu.array = BitConverter.GetBytes(roket_gps_irtifa_float32_uint8_donusturucu.sayi);
            Array.Copy(roket_gps_irtifa_float32_uint8_donusturucu.array, 0, olusturalacak_paket, 10, 4);

            FLOAT32_UINT8_DONUSTURUCU roket_enlem_irtifa_float32_uint8_donusturucu;
            roket_enlem_irtifa_float32_uint8_donusturucu.sayi = float.Parse(enlem); // Roket boylam degerinin atamasini yapiyoruz.
            roket_enlem_irtifa_float32_uint8_donusturucu.array = BitConverter.GetBytes(roket_enlem_irtifa_float32_uint8_donusturucu.sayi);
            Array.Copy(roket_enlem_irtifa_float32_uint8_donusturucu.array, 0, olusturalacak_paket, 14, 4);

            FLOAT32_UINT8_DONUSTURUCU roket_boylam_irtifa_float32_uint8_donusturucu;
            roket_boylam_irtifa_float32_uint8_donusturucu.sayi = float.Parse(boylam); // Roket boylam degerinin atamasini yapiyoruz.
            roket_enlem_irtifa_float32_uint8_donusturucu.array = BitConverter.GetBytes(roket_enlem_irtifa_float32_uint8_donusturucu.sayi);
            Array.Copy(roket_enlem_irtifa_float32_uint8_donusturucu.array, 0, olusturalacak_paket, 18, 4);

            olusturalacak_paket[74] = 1; // Durum bilgisi = Iki parasut de tetiklenmedi

            olusturalacak_paket[75] = check_sum_hesapla(); // Check_sum = check_sum_hesapla();
            olusturalacak_paket[76] = 0x0D; // Sabit
            olusturalacak_paket[77] = 0x0A; // Sabit

            for (int i = olusturalacak_paket.Length; i < 78; i++)
            {
                olusturalacak_paket[i] = 0x00;
            }

        }

        private void ConfigureSerialPort()
        {
            serialPort = new SerialPort
            {
                PortName = "COM8", // USB TTL adaptörünün bağlı olduğu seri port
                BaudRate = 19200, // İletişim hızı (baud rate)
                Parity = Parity.None, // Parity bit (None, Odd, Even, Mark, Space)
                DataBits = 8, // Veri bit sayısı
                StopBits = StopBits.One, // Durma bit sayısı
            };

            try
            {
                serialPort.Open();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata: " + ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            checkh();
        }

        private void checkh()
        {
            serialPort.DataReceived += new SerialDataReceivedEventHandler(serialPort_check);
        }

        private void serialPort_check(object sender, SerialDataReceivedEventArgs e)
        {
           
                packet(); // Pakete verileri ekle

                ConfigureSerialPort(); // Seri portu yapılandırın ve bağlantıyı açın
                SendData(olusturalacak_paket); // Veriyi USB TTL üzerinden gönderin

                string hexString = string.Join(" ", olusturalacak_paket.Select(b => "0x" + b.ToString("X2")).ToArray());

                textBox15.Text = hexString;
                serialPort.Close();
            
        }

        private void SendData(byte[] data)
        {
             serialPort.Write(olusturalacak_paket, 0, olusturalacak_paket.Length);
        }


        private void button1_Click_2(object sender, EventArgs e)
        {
            ConfigureSerialPort();
        }

        private void openGLControl1_Load(object sender, EventArgs e)
        {
            
        }

        private void webView21_Click(object sender, EventArgs e)
        {

        }

        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            mov = true;
            movx = Cursor.Position.X - Left;
            movy = Cursor.Position.Y - Top;
        }

       
        private void DrawRocket(OpenGL gl)
        {
            // Roketin gövdesini çiz
            gl.Begin(OpenGL.GL_QUADS);
            gl.Color(1.0f, 0.0f, 0.0f);
            gl.Vertex(-0.2f, -0.5f);
            gl.Vertex(0.2f, -0.5f);
            gl.Vertex(0.2f, 0.5f);
            gl.Vertex(-0.2f, 0.5f);
            gl.End();

            // Roketin yan kanatlarını çiz
            gl.Begin(OpenGL.GL_TRIANGLES);
            gl.Color(1.0f, 1.0f, 1.0f);
            gl.Vertex(-0.2f, -0.5f);
            gl.Vertex(-0.5f, -0.75f);
            gl.Vertex(-0.2f, -0.75f);

            gl.Vertex(0.2f, -0.5f);
            gl.Vertex(0.5f, -0.75f);
            gl.Vertex(0.2f, -0.75f);
            gl.End();

            // Roketin uç kısmını çiz
            gl.Begin(OpenGL.GL_TRIANGLES);
            gl.Color(0.0f, 1.0f, 0.0f);
            gl.Vertex(-0.2f, 0.5f);
            gl.Vertex(0.2f, 0.5f);
            gl.Vertex(0.0f, 0.8f);
            gl.End();
        }

        private void button5_Click(object sender, EventArgs e)
        {
          
        }

        private void openGLControl1_Load_1(object sender, EventArgs e)
        {

        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
           
        }

        private void openGLControl_Load(object sender, EventArgs e)
        {

        }

    }

}
