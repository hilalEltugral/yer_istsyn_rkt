using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MaterialSkin;
using MaterialSkin.Controls;
using System.Data.SqlClient;
using System.IO.Ports;
using System.IO;
using GMap.NET;
using System.Drawing.Imaging;
using System.Data.OleDb;
using System.Threading;

namespace interfaceR
{
    public partial class Form1 : MaterialForm
    {
        long max = 30, min = 0;

        OleDbConnection baglanti = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source =./data/veriler.xlsx; Extended Properties ='Excel 12.0 Xml;'");
        private readonly MaterialSkinManager materialSkinManager;
        private Thread thread;
        private bool runThread = false;
        private bool debugMode = true;
        public Form1()
        {
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;//tam ekran kodu
            this.FormClosed += new FormClosedEventHandler(Form1_Closing);
            InitializeComponent();
            materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.AddFormToManage(this);
            materialSkinManager.Theme = MaterialSkinManager.Themes.LIGHT;
            materialSkinManager.ColorScheme = new ColorScheme(Primary.BlueGrey800, Primary.BlueGrey900, Primary.BlueGrey500, Accent.LightBlue200, TextShade.WHITE);
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            string[] ports = SerialPort.GetPortNames();  //Seri portları diziye ekleme
            foreach (string port in ports)
            {
                materialComboBox1.Items.Add(port);//Seri portları comboBox1'e ekleme
            }
             gMapAktif.MinZoom = 0;
             gMapAktif.MaxZoom = 90;
        }

        private void Form1_Closing(object sender, FormClosedEventArgs e)
        {
            stop();
        }

        private void materialButton1_Click(object sender, EventArgs e)
        {
            try
            {
                if (!debugMode)
                {
                    serialPort1.PortName = materialComboBox1.Text;
                    serialPort1.BaudRate = 9600;
                    serialPort1.Open();
                    serialPort1.DataBits = 8;
                    serialPort1.Parity = Parity.None;
                    serialPort1.StopBits = StopBits.One;
                    serialPort1.Handshake = Handshake.None;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Hata");    //Hata mesajı göster
            }
        }

        private void materialButton2_Click(object sender, EventArgs e)
        {

            baglanti.Open();
            thread = new Thread(new ThreadStart(setUI));
            runThread = true;
            thread.Start();
        }

        private void materialButton3_Click(object sender, EventArgs e)
        {
            Form2 yeniSayfa = new Form2();
            yeniSayfa.ShowDialog();
            this.Show();
        }

        private void stop()
        {
            runThread = false;

            if (!debugMode)
            {
                if (serialPort1.IsOpen) {
                    serialPort1.Close();
                }
            }
            
            baglanti.Close();
            
            if (thread != null)
            {
                if (thread.IsAlive)
                {
                    thread.Abort();
                }  
            }
        }

        private void materialButton4_Click(object sender, EventArgs e)
        {
            stop();
        }

        private void materialLabel20_Click(object sender, EventArgs e)
        {

        }

        private void materialLabel16_Click(object sender, EventArgs e)
        {

        }

        private void materialTextBox17_TextChanged(object sender, EventArgs e)
        {

        }

        private void materialLabel22_Click(object sender, EventArgs e)
        {

        }

        private void materialLabel19_Click(object sender, EventArgs e)
        {

        }

        private void materialTextBox14_TextChanged(object sender, EventArgs e)
        {

        }

        private void materialTextBox16_TextChanged(object sender, EventArgs e)
        {

        }

        private void materialTextBox15_TextChanged(object sender, EventArgs e)
        {

        }

        private void setUI()
        {
            while (runThread)
            {
                string harf = "";
                float x = 0.0f;
                float y = 0.0f;
                float z = 0.0f;
                double sicaklik = 0.0;
                double yukseklik = 0.0;
                double basinc = 0.0;
                double hiz = 0.0;
                double pil = 0.0;
                double enlem = 0.0;
                double boylam = 0.0;

                double roketDurum=0.0;
                double roketDurum2 = 0.0;

                double ySicaklik = 0.0;
                double yEnlem = 0.0;  
                double yBoylam = 0.0;
                double yYukseklik=0.0;

                bool anaVeri = false;
                bool yedekVeri = false;
                
                try
                {
                    string sonuc = "";
                    int say;
                    if (debugMode)
                    {
                        Random random = new Random();
                        say = random.Next(1,3);

                        if (say == 1)
                        {
                            sonuc =
                                "a" + ":" +
                                random.Next(-360, 360) + ":" + // x
                                random.Next(-360, 360) + ":" + // y
                                random.Next(-360, 360) + ":" + // z
                                random.Next(20, 50) + ":" + // temp
                                random.Next(900, 1023) + ":" + // p (basınç)
                                random.Next(-10, 360) + ":" + // yükseklik
                                40.7405 + "" + random.Next(5, 34) + ":" + // enlem
                                30.3355 + "" + random.Next(5, 34) + ":" + // boylam
                                random.Next(-10, 100) + ":" +  // hiz
                                random.Next(5, 12) + ":" + // pil
                                random.Next(0, 5) + ":" + //Roketdurum1
                                random.Next(0, 5);//roketdurum2
                        }
                        else if (say == 2)
                        {
                            sonuc =
                               "y" + ":" +
                               random.Next(5, 10) + ":" + //temp
                               40.7405 + "" + random.Next(5, 34) + ":" + // enlem
                               30.3355 + "" + random.Next(5, 34) + ":" + // boylam
                               random.Next(5, 10);//yukseklik

                           // MessageBox.Show("2");
                        }
                   
                    }
                    else
                    {
                        sonuc = serialPort1.ReadLine();
                    }

                    if (sonuc.Contains(":"))

                    {
                        string[] pot = sonuc.Split(':');
                        harf = pot[0];
                        if (harf == "a") { 
                        x = (float)Convert.ToDouble(pot[1].Replace('.', ','));
                        y = (float)Convert.ToDouble(pot[2].Replace('.', ','));
                        z = (float)Convert.ToDouble(pot[3].Replace('.', ','));
                        sicaklik = Double.Parse(pot[4].Replace('.', ','));
                        basinc = Convert.ToDouble(pot[5].Replace('.', ','));
                        yukseklik = Convert.ToDouble(pot[6].Replace('.', ','));

                            double gelenEnlem = 0.0;
                            double gelenBoylam = 0.0;

                            try {
                                gelenEnlem = Convert.ToDouble(pot[7].Replace('.', ','));
                            } 
                            catch (Exception e)
                            {
                                enlem = 38.387832;
                            } 
                            finally
                            {
                                enlem = gelenEnlem;
                            }

                            try
                            {
                                gelenBoylam = Convert.ToDouble(pot[8].Replace('.', ','));
                            }
                            catch (Exception e)
                            {
                                boylam = 33.741336;
                            }
                            finally
                            {
                                boylam = gelenBoylam;
                            }



                            hiz = Convert.ToDouble(pot[9].Replace('.', ','));
                            pil = Convert.ToDouble(pot[10].Replace('.', ','));
                            roketDurum =  Convert.ToDouble(pot[11].Replace('.', ','));
                            roketDurum2 = Convert.ToDouble(pot[12].Replace('.', ','));

                            anaVeri =true;
                        }
                        else if (harf == "y")
                        {
                            ySicaklik = Convert.ToDouble(pot[1].Replace('.', ','));
                            yEnlem = Convert.ToDouble(pot[2].Replace('.', ','));
                            yBoylam = Convert.ToDouble(pot[3].Replace('.', ','));
                            yYukseklik = Convert.ToDouble(pot[4].Replace('.', ','));

                            yedekVeri = true;

                        }
                    }
                }
                catch (Exception)
                {

                }

                if (runThread)
                {
                    materialTextBox10.Invoke((MethodInvoker)delegate {
                        // Running on the UI thread
                        materialTextBox10.Text = DateTime.Now.ToString();
                        materialTextBox2.Text = "49794";
                    });


                    if (anaVeri) { 
                    chart1.Invoke((MethodInvoker)delegate {
                        chart1.ChartAreas[0].AxisX.Minimum = min;
                        chart1.ChartAreas[0].AxisX.Maximum = max;

                        chart1.ChartAreas[0].AxisY.Maximum = Double.NaN;

                        this.chart1.Series[0].Points.AddXY(max, basinc);
                        chart1.Series[0].IsVisibleInLegend = false;// remove legend
                        chart1.Series["hPa"].BorderWidth = 2;
                    });
                    }

                    if (anaVeri) { 
                    chart2.Invoke((MethodInvoker)delegate
                        {
                            chart2.ChartAreas[0].AxisX.Minimum = min;
                            chart2.ChartAreas[0].AxisX.Maximum = max;

                            chart2.ChartAreas[0].RecalculateAxesScale();

                            this.chart2.Series[0].Points.AddXY(max, hiz);
                            chart2.Series[0].IsVisibleInLegend = false;// remove legend
                            chart2.Series["v"].BorderWidth = 2;
                        });
                    }

                    if (anaVeri)
                    {
                        chart3.Invoke((MethodInvoker)delegate
                        {
                            chart3.ChartAreas[0].AxisX.Minimum = min;
                            chart3.ChartAreas[0].AxisX.Maximum = max;

                            chart3.ChartAreas[0].RecalculateAxesScale();

                            this.chart3.Series[0].Points.AddXY(max, pil);
                            chart3.Series[0].IsVisibleInLegend = false;// remove legend
                            chart3.Series["V"].BorderWidth = 2;
                        });
                    }

                    if (anaVeri)
                    {
                        chart4.Invoke((MethodInvoker)delegate

                        {
                            chart4.ChartAreas[0].AxisX.Minimum = min;
                            chart4.ChartAreas[0].AxisX.Maximum = max;

                            chart4.ChartAreas[0].RecalculateAxesScale();

                            this.chart4.Series[0].Points.AddXY(max, sicaklik);
                            chart4.Series[0].IsVisibleInLegend = false;// remove legend
                            chart4.Series["°C"].BorderWidth = 2;

                        });
                    }

                    if (anaVeri)
                    {
                        chart5.Invoke((MethodInvoker)delegate
                        {
                            chart5.ChartAreas[0].AxisX.Minimum = min;
                            chart5.ChartAreas[0].AxisX.Maximum = max;

                            chart5.ChartAreas[0].RecalculateAxesScale();

                            this.chart5.Series[0].Points.AddXY(max, yukseklik);
                            chart5.Series[0].IsVisibleInLegend = false;// remove legend
                            chart5.Series["m"].BorderWidth = 2;
                        });
                    }
                    
                        materialTextBox10.Invoke((MethodInvoker)delegate {
                            // Running on the UI thread
                            if(anaVeri){
                            materialTextBox13.Text = x + "";
                            materialTextBox12.Text = y + "";
                            materialTextBox11.Text = z + "";
                            materialTextBox5.Text = sicaklik + "";
                            materialTextBox1.Text = basinc + "";
                            materialTextBox4.Text = yukseklik + "";
                            materialTextBox3.Text = hiz + "";
                            materialTextBox6.Text = pil + "";
                            materialTextBox9.Text = enlem + "";
                            materialTextBox8.Text = boylam + "";
                                if (roketDurum == 1)
                                {
                                    materialTextBox19.Text = "Ayrıldı";
                                }
                                if (roketDurum2 == 1)
                                {
                                    materialTextBox7.Text = "Ayrıldı";
                                }
                            }
                            if (yedekVeri) { 
                            materialTextBox15.Text = ySicaklik + "";
                            materialTextBox16.Text = yEnlem + "";
                            materialTextBox17.Text = yBoylam + "";
                            materialTextBox14.Text = yYukseklik + "";
                            }
                        });
                }
                min++;
                max++;

                if (!debugMode)
                {
                    serialPort1.DiscardInBuffer();
                }

                //execele id basma 
                OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [Sayfa1$]", baglanti);
                DataTable dt = new DataTable();
                dt.Clear();
                da.Fill(dt);
                dataGridView2.DataSource = dt;
                //bit

                sayac1.Invoke((MethodInvoker)delegate {
                    // Running on the UI thread
                    sayac1.Text = dataGridView2.RowCount.ToString();

                });

                //excel kayıt alma
                OleDbCommand komut = new OleDbCommand("insert into[Sayfa1$](TAKIMNO,PAKETNUMARASI,GÖNDERMESAATİ,BASINÇ,YÜKSEKLİK,İNİŞHIZI,SICAKLIK,PİLGERİLİMİ,GPSLATİTUDE,GPSLONGİTUDE,PİTCH,ROLL,YAW,ySICAKLIK,yGPSLATİTUDE,yGPSLONGİTUDE,yYUKSEKLIK) values (@p1,@p2,@p3,@p4,@p5,@p6,@p7,@p8,@p9,@p10,@p11,@p12,@p13,@p14,@p15,p16,p17)", baglanti);

                materialTextBox2.Invoke((MethodInvoker)delegate {
                    // Running on the UI thread
                    komut.Parameters.AddWithValue("@p1", materialTextBox2.Text);
                    komut.Parameters.AddWithValue("@p2", sayac1.Text);
                    komut.Parameters.AddWithValue("@p3", materialTextBox10.Text);
                    komut.Parameters.AddWithValue("@p4", materialTextBox1.Text);
                    komut.Parameters.AddWithValue("@p5", materialTextBox4.Text);
                    komut.Parameters.AddWithValue("@p6", materialTextBox3.Text);
                    komut.Parameters.AddWithValue("@p7", materialTextBox5.Text);
                    komut.Parameters.AddWithValue("@p8", materialTextBox6.Text);
                    komut.Parameters.AddWithValue("@p9", materialTextBox9.Text);
                    komut.Parameters.AddWithValue("@p10", materialTextBox8.Text);
                    komut.Parameters.AddWithValue("@p11", materialTextBox13.Text);
                    komut.Parameters.AddWithValue("@p12", materialTextBox11.Text);
                    komut.Parameters.AddWithValue("@p13", materialTextBox12.Text);

                    komut.Parameters.AddWithValue("@p14", materialTextBox15.Text);
                    komut.Parameters.AddWithValue("@p15", materialTextBox16.Text);
                    komut.Parameters.AddWithValue("@p16", materialTextBox17.Text);
                    komut.Parameters.AddWithValue("@p17", materialTextBox14.Text);

                });

                komut.ExecuteNonQuery();

                if (anaVeri) { 
                    // Harita gösterme
                    gMapAktif.Invoke((MethodInvoker)delegate
                    {
                        // Running on the UI thread
                        gMapAktif.MapProvider = GMap.NET.MapProviders.ArcGIS_StreetMap_World_2D_MapProvider.Instance;
                        GMap.NET.GMaps.Instance.Mode = GMap.NET.AccessMode.ServerOnly;
                        gMapAktif.Position = new GMap.NET.PointLatLng(enlem, boylam);
                    });
                }

                Thread.Sleep(1000);
            }
        }

    }
}