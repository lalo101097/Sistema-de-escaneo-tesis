using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using CsvHelper;
using System.IO;
using System.Configuration;

namespace Sistema_de_escaneo
{
    public partial class Form1 : Form
    {
        // 0, 15, 26, 36, 45, 54, 65, 78
        // 0, 17, 25, 32, 38, 43, 50, 56, 63, 68, 73, 81 
        
        string[] let = { "A", "B", "C", "D", "E", "F", "G", "H" };
        
        bool progr = false;
        bool movlib = false;
        
        bool[,] posi = {{ false, false, false, false, false, false, false, false, false, false, false, false },
                        { false, false, false, false, false, false, false, false, false, false, false, false },
                        { false, false, false, false, false, false, false, false, false, false, false, false },
                        { false, false, false, false, false, false, false, false, false, false, false, false },
                        { false, false, false, false, false, false, false, false, false, false, false, false },
                        { false, false, false, false, false, false, false, false, false, false, false, false },
                        { false, false, false, false, false, false, false, false, false, false, false, false },
                        { false, false, false, false, false, false, false, false, false, false, false, false } };
        
        string[] cas = new string[90];
        float[,] mues = new float[4, 90];

        public class sens
        {
            public string Casilla { get; set; }
            public string Rojo { get; set; }
            public string Verde { get; set; }
            public string Azul { get; set; }
            public string Claro { get; set; }
        }

        public Form1()
        {
            InitializeComponent();
        }

        private void casilla(int a, int b)
        {
            if(progr)
            {
                if (posi[a,b] == true)
                {
                    posi[a,b] = false;
                }
                else
                {
                    posi[a,b] = true;
                }
            }
            else if(movlib)
            {
                SpPrin.Write("b" + "," + Properties.Settings.Default.Horizontal[a] + "," + Properties.Settings.Default.Vertical[b] + "," + "\n");
                SpPrin.ReadTimeout = 12000;
                string pass = SpPrin.ReadLine();
                Task.Delay(500).Wait();
                LblCasilla.Text = Convert.ToString(let[a]) + Convert.ToString(b + 1);
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            LblCasilla.Text = "";
            LblEstado.Text = "";
            
            CbxSelec.Items.Clear();
            for (int i = 0; i < 12; i++)
            {
                CbxSelec.Items.Add(i + 1);
            }
            
            GpbEje.Enabled = false;
            GpbRep.Enabled = false;
            NudRepet.Value = Convert.ToDecimal(Properties.Settings.Default.loops);
            CbxSelec.SelectedIndex = 0;
        }

        private void BtnBuscar_Click(object sender, EventArgs e)
        {
            string[] PuertosDisponibles = SerialPort.GetPortNames();

            CbPuertos.Items.Clear();

            foreach (string puerto_simple in PuertosDisponibles)
            {
                CbPuertos.Items.Add(puerto_simple);
            }

            if (CbPuertos.Items.Count > 0)
            {
                CbPuertos.SelectedIndex = 0;
                BtnConcor.Enabled = true;
                BtnConcor.Focus();
            }
            else
            {
                MessageBox.Show("Ningún puerto conectado");
                CbPuertos.Items.Clear();
            }
        }

        private void BtnConcor_Click(object sender, EventArgs e)
        {
            try
            {
                if (BtnConcor.Text == "Conectar")
                {
                    SpPrin.BaudRate = 9600;
                    SpPrin.DataBits = 8;
                    SpPrin.Parity = Parity.None;
                    SpPrin.StopBits = StopBits.One;
                    SpPrin.Handshake = Handshake.None;
                    SpPrin.PortName = CbPuertos.Text;

                    try
                    {
                        SpPrin.Open();
                        Task.Delay(2000).Wait();
                        BtnConcor.Text = "Desconectar";
                        SpPrin.WriteLine("p,");
                        SpPrin.ReadTimeout = 5000;
                        string sp = Convert.ToString(SpPrin.ReadLine());

                        if(sp == "l1\r")
                        {
                            LblEstado.Text = "Conectado";
                            GbCasillas.Enabled = true;
                            GbControl.Enabled = true;
                            GpbRep.Enabled = true;
                            GpbEje.Enabled = true;
                            GpbLuz.Enabled = true;
                        }
                        else if(sp == "l0\r")
                        {
                            LblEstado.Text = "Sen error";
                        }
                        else
                        {
                            LblEstado.Text = "Ardu error";
                            SpPrin.Close();
                            BtnConcor.Text = "Conectar";
                        }
                    }
                    catch (Exception exc)
                    {
                        MessageBox.Show(exc.Message.ToString());
                    }
                }
                else if (BtnConcor.Text == "Desconectar")
                {
                    SpPrin.Close();
                    BtnConcor.Text = "Conectar";
                    GbCasillas.Enabled = false;
                    GbControl.Enabled = false;
                }
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message.ToString());
            }
        }

        private void BtnProg_Click(object sender, EventArgs e)
        {
            if (progr)
            {
                progr = false;
                BtnProg.BackColor = Color.LightGray;
            }
            else
            {
                progr = true;
                BtnProg.BackColor = Color.Red;
            }
        }

        private void BtnMovLib_Click(object sender, EventArgs e)
        {
            if (movlib)
            {
                movlib = false;
                BtnMovLib.BackColor = Color.LightGray;
            }
            else
            {
                movlib = true;
                BtnMovLib.BackColor = Color.Red;
            }
        }
        #region MyRegion

        private void BtnA1_Click(object sender, EventArgs e)
        {
            casilla(0, 0);
            if (progr)
            {
                if (BtnA1.BackColor == Color.LightGray)
                {
                    BtnA1.BackColor = Color.Red;
                }
                else
                {
                    BtnA1.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnA2_Click(object sender, EventArgs e)
        {
            casilla(0, 1);
            if (progr)
            {
                if (BtnA2.BackColor == Color.LightGray)
                {
                    BtnA2.BackColor = Color.Red;
                }
                else
                {
                    BtnA2.BackColor = Color.LightGray;
                }
            }
        }

        private void Btn_Click(object sender, EventArgs e)
        {
            casilla(0, 2);
            if (progr)
            {
                if (BtnA3.BackColor == Color.LightGray)
                {
                    BtnA3.BackColor = Color.Red;
                }
                else
                {
                    BtnA3.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnA4_Click(object sender, EventArgs e)
        {
            casilla(0, 3);
            if (progr)
            {
                if (BtnA4.BackColor == Color.LightGray)
                {
                    BtnA4.BackColor = Color.Red;
                }
                else
                {
                    BtnA4.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnA5_Click(object sender, EventArgs e)
        {
            casilla(0, 4);
            if (progr)
            {
                if (BtnA5.BackColor == Color.LightGray)
                {
                    BtnA5.BackColor = Color.Red;
                }
                else
                {
                    BtnA5.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnA6_Click(object sender, EventArgs e)
        {
            casilla(0, 5);
            if (progr)
            {
                if (BtnA6.BackColor == Color.LightGray)
                {
                    BtnA6.BackColor = Color.Red;
                }
                else
                {
                    BtnA6.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnA7_Click(object sender, EventArgs e)
        {
            casilla(0, 6);
            if (progr)
            {
                if (BtnA7.BackColor == Color.LightGray)
                {
                    BtnA7.BackColor = Color.Red;
                }
                else
                {
                    BtnA7.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnA8_Click(object sender, EventArgs e)
        {
            casilla(0, 7);
            if (progr)
            {
                if (BtnA8.BackColor == Color.LightGray)
                {
                    BtnA8.BackColor = Color.Red;
                }
                else
                {
                    BtnA8.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnA9_Click(object sender, EventArgs e)
        {
            casilla(0, 8);
            if (progr)
            {
                if (BtnA9.BackColor == Color.LightGray)
                {
                    BtnA9.BackColor = Color.Red;
                }
                else
                {
                    BtnA9.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnA10_Click(object sender, EventArgs e)
        {
            casilla(0, 9);
            if (progr)
            {
                if (BtnA10.BackColor == Color.LightGray)
                {
                    BtnA10.BackColor = Color.Red;
                }
                else
                {
                    BtnA10.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnA11_Click(object sender, EventArgs e)
        {
            casilla(0, 10);
            if (progr)
            {
                if (BtnA11.BackColor == Color.LightGray)
                {
                    BtnA11.BackColor = Color.Red;
                }
                else
                {
                    BtnA11.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnA12_Click(object sender, EventArgs e)
        {
            casilla(0, 11);
            if (progr)
            {
                if (BtnA12.BackColor == Color.LightGray)
                {
                    BtnA12.BackColor = Color.Red;
                }
                else
                {
                    BtnA12.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnB1_Click(object sender, EventArgs e)
        {
            casilla(1, 0);
            if (progr)
            {
                if (BtnB1.BackColor == Color.LightGray)
                {
                    BtnB1.BackColor = Color.Red;
                }
                else
                {
                    BtnB1.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnB2_Click(object sender, EventArgs e)
        {
            casilla(1, 1);
            if (progr)
            {
                if (BtnB2.BackColor == Color.LightGray)
                {
                    BtnB2.BackColor = Color.Red;
                }
                else
                {
                    BtnB2.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnB3_Click(object sender, EventArgs e)
        {
            casilla(1, 2);
            if (progr)
            {
                if (BtnB3.BackColor == Color.LightGray)
                {
                    BtnB3.BackColor = Color.Red;
                }
                else
                {
                    BtnB3.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnB4_Click(object sender, EventArgs e)
        {
            casilla(1, 3);
            if (progr)
            {
                if (BtnB4.BackColor == Color.LightGray)
                {
                    BtnB4.BackColor = Color.Red;
                }
                else
                {
                    BtnB4.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnB5_Click(object sender, EventArgs e)
        {
            casilla(1, 4);
            if (progr)
            {
                if (BtnB5.BackColor == Color.LightGray)
                {
                    BtnB5.BackColor = Color.Red;
                }
                else
                {
                    BtnB5.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnB6_Click(object sender, EventArgs e)
        {
            casilla(1, 5);
            if (progr)
            {
                if (BtnB6.BackColor == Color.LightGray)
                {
                    BtnB6.BackColor = Color.Red;
                }
                else
                {
                    BtnB6.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnB7_Click(object sender, EventArgs e)
        {
            casilla(1, 6);
            if (progr)
            {
                if (BtnB7.BackColor == Color.LightGray)
                {
                    BtnB7.BackColor = Color.Red;
                }
                else
                {
                    BtnB7.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnB8_Click(object sender, EventArgs e)
        {
            casilla(1, 7);
            if (progr)
            {
                if (BtnB8.BackColor == Color.LightGray)
                {
                    BtnB8.BackColor = Color.Red;
                }
                else
                {
                    BtnB8.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnB9_Click(object sender, EventArgs e)
        {
            casilla(1, 8);
            if (progr)
            {
                if (BtnB9.BackColor == Color.LightGray)
                {
                    BtnB9.BackColor = Color.Red;
                }
                else
                {
                    BtnB9.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnB10_Click(object sender, EventArgs e)
        {
            casilla(1, 9);
            if (progr)
            {
                if (BtnB10.BackColor == Color.LightGray)
                {
                    BtnB10.BackColor = Color.Red;
                }
                else
                {
                    BtnB10.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnB11_Click(object sender, EventArgs e)
        {
            casilla(1, 10);
            if (progr)
            {
                if (BtnB11.BackColor == Color.LightGray)
                {
                    BtnB11.BackColor = Color.Red;
                }
                else
                {
                    BtnB11.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnB12_Click(object sender, EventArgs e)
        {
            casilla(1, 11);
            if (progr)
            {
                if (BtnB12.BackColor == Color.LightGray)
                {
                    BtnB12.BackColor = Color.Red;
                }
                else
                {
                    BtnB12.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnC1_Click(object sender, EventArgs e)
        {
            casilla(2, 0);
            if (progr)
            {
                if (BtnC1.BackColor == Color.LightGray)
                {
                    BtnC1.BackColor = Color.Red;
                }
                else
                {
                    BtnC1.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnC2_Click(object sender, EventArgs e)
        {
            casilla(2, 1);
            if (progr)
            {
                if (BtnC2.BackColor == Color.LightGray)
                {
                    BtnC2.BackColor = Color.Red;
                }
                else
                {
                    BtnC2.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnC3_Click(object sender, EventArgs e)
        {
            casilla(2, 2);
            if (progr)
            {
                if (BtnC3.BackColor == Color.LightGray)
                {
                    BtnC3.BackColor = Color.Red;
                }
                else
                {
                    BtnC3.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnC4_Click(object sender, EventArgs e)
        {
            casilla(2, 3);
            if (progr)
            {
                if (BtnC4.BackColor == Color.LightGray)
                {
                    BtnC4.BackColor = Color.Red;
                }
                else
                {
                    BtnC4.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnC5_Click(object sender, EventArgs e)
        {
            casilla(2, 4);
            if (progr)
            {
                if (BtnC5.BackColor == Color.LightGray)
                {
                    BtnC5.BackColor = Color.Red;
                }
                else
                {
                    BtnC5.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnC6_Click(object sender, EventArgs e)
        {
            casilla(2, 5);
            if (progr)
            {
                if (BtnC6.BackColor == Color.LightGray)
                {
                    BtnC6.BackColor = Color.Red;
                }
                else
                {
                    BtnC6.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnC7_Click(object sender, EventArgs e)
        {
            casilla(2, 6);
            if (progr)
            {
                if (BtnC7.BackColor == Color.LightGray)
                {
                    BtnC7.BackColor = Color.Red;
                }
                else
                {
                    BtnC7.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnC8_Click(object sender, EventArgs e)
        {
            casilla(2, 7);
            if (progr)
            {
                if (BtnC8.BackColor == Color.LightGray)
                {
                    BtnC8.BackColor = Color.Red;
                }
                else
                {
                    BtnC8.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnC9_Click(object sender, EventArgs e)
        {
            casilla(2, 8);
            if (progr)
            {
                if (BtnC9.BackColor == Color.LightGray)
                {
                    BtnC9.BackColor = Color.Red;
                }
                else
                {
                    BtnC9.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnC10_Click(object sender, EventArgs e)
        {
            casilla(2, 9);
            if (progr)
            {
                if (BtnC10.BackColor == Color.LightGray)
                {
                    BtnC10.BackColor = Color.Red;
                }
                else
                {
                    BtnC10.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnC11_Click(object sender, EventArgs e)
        {
            casilla(2, 10);
            if (progr)
            {
                if (BtnC11.BackColor == Color.LightGray)
                {
                    BtnC11.BackColor = Color.Red;
                }
                else
                {
                    BtnC11.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnC12_Click(object sender, EventArgs e)
        {
            casilla(2, 11);
            if (progr)
            {
                if (BtnC12.BackColor == Color.LightGray)
                {
                    BtnC12.BackColor = Color.Red;
                }
                else
                {
                    BtnC12.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnD1_Click(object sender, EventArgs e)
        {
            casilla(3, 0);
            if (progr)
            {
                if (BtnD1.BackColor == Color.LightGray)
                {
                    BtnD1.BackColor = Color.Red;
                }
                else
                {
                    BtnD1.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnD2_Click(object sender, EventArgs e)
        {
            casilla(3, 1);
            if (progr)
            {
                if (BtnD2.BackColor == Color.LightGray)
                {
                    BtnD2.BackColor = Color.Red;
                }
                else
                {
                    BtnD2.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnD3_Click(object sender, EventArgs e)
        {
            casilla(3, 2);
            if (progr)
            {
                if (BtnD3.BackColor == Color.LightGray)
                {
                    BtnD3.BackColor = Color.Red;
                }
                else
                {
                    BtnD3.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnD4_Click(object sender, EventArgs e)
        {
            casilla(3, 3);
            if (progr)
            {
                if (BtnD4.BackColor == Color.LightGray)
                {
                    BtnD4.BackColor = Color.Red;
                }
                else
                {
                    BtnD4.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnD5_Click(object sender, EventArgs e)
        {
            casilla(3, 4);
            if (progr)
            {
                if (BtnD5.BackColor == Color.LightGray)
                {
                    BtnD5.BackColor = Color.Red;
                }
                else
                {
                    BtnD5.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnD6_Click(object sender, EventArgs e)
        {
            casilla(3, 5);
            if (progr)
            {
                if (BtnD6.BackColor == Color.LightGray)
                {
                    BtnD6.BackColor = Color.Red;
                }
                else
                {
                    BtnD6.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnD7_Click(object sender, EventArgs e)
        {
            casilla(3, 6);
            if (progr)
            {
                if (BtnD7.BackColor == Color.LightGray)
                {
                    BtnD7.BackColor = Color.Red;
                }
                else
                {
                    BtnD7.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnD8_Click(object sender, EventArgs e)
        {
            casilla(3, 7);
            if (progr)
            {
                if (BtnD8.BackColor == Color.LightGray)
                {
                    BtnD8.BackColor = Color.Red;
                }
                else
                {
                    BtnD8.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnD9_Click(object sender, EventArgs e)
        {
            casilla(3, 8);
            if (progr)
            {
                if (BtnD9.BackColor == Color.LightGray)
                {
                    BtnD9.BackColor = Color.Red;
                }
                else
                {
                    BtnD9.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnD10_Click(object sender, EventArgs e)
        {
            casilla(3, 9);
            if (progr)
            {
                if (BtnD10.BackColor == Color.LightGray)
                {
                    BtnD10.BackColor = Color.Red;
                }
                else
                {
                    BtnD10.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnD11_Click(object sender, EventArgs e)
        {
            casilla(3, 10);
            if (progr)
            {
                if (BtnD11.BackColor == Color.LightGray)
                {
                    BtnD11.BackColor = Color.Red;
                }
                else
                {
                    BtnD11.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnD12_Click(object sender, EventArgs e)
        {
            casilla(3, 11);
            if (progr)
            {
                if (BtnD12.BackColor == Color.LightGray)
                {
                    BtnD12.BackColor = Color.Red;
                }
                else
                {
                    BtnD12.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnE1_Click(object sender, EventArgs e)
        {
            casilla(4, 0);
            if (progr)
            {
                if (BtnE1.BackColor == Color.LightGray)
                {
                    BtnE1.BackColor = Color.Red;
                }
                else
                {
                    BtnE1.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnE2_Click(object sender, EventArgs e)
        {
            casilla(4, 1);
            if (progr)
            {
                if (BtnE2.BackColor == Color.LightGray)
                {
                    BtnE2.BackColor = Color.Red;
                }
                else
                {
                    BtnE2.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnE3_Click(object sender, EventArgs e)
        {
            casilla(4, 2);
            if (progr)
            {
                if (BtnE3.BackColor == Color.LightGray)
                {
                    BtnE3.BackColor = Color.Red;
                }
                else
                {
                    BtnE3.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnE4_Click(object sender, EventArgs e)
        {
            casilla(4, 3);
            if (progr)
            {
                if (BtnE4.BackColor == Color.LightGray)
                {
                    BtnE4.BackColor = Color.Red;
                }
                else
                {
                    BtnE4.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnE5_Click(object sender, EventArgs e)
        {
            casilla(4, 4);
            if (progr)
            {
                if (BtnE5.BackColor == Color.LightGray)
                {
                    BtnE5.BackColor = Color.Red;
                }
                else
                {
                    BtnE5.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnE6_Click(object sender, EventArgs e)
        {
            casilla(4, 5);
            if (progr)
            {
                if (BtnE6.BackColor == Color.LightGray)
                {
                    BtnE6.BackColor = Color.Red;
                }
                else
                {
                    BtnE6.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnE7_Click(object sender, EventArgs e)
        {
            casilla(4, 6);
            if (progr)
            {
                if (BtnE7.BackColor == Color.LightGray)
                {
                    BtnE7.BackColor = Color.Red;
                }
                else
                {
                    BtnE7.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnR8_Click(object sender, EventArgs e)
        {
            casilla(4, 7);
            if (progr)
            {
                if (BtnE8.BackColor == Color.LightGray)
                {
                    BtnE8.BackColor = Color.Red;
                }
                else
                {
                    BtnE8.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnE9_Click(object sender, EventArgs e)
        {
            casilla(4, 8);
            if (progr)
            {
                if (BtnE9.BackColor == Color.LightGray)
                {
                    BtnE9.BackColor = Color.Red;
                }
                else
                {
                    BtnE9.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnE10_Click(object sender, EventArgs e)
        {
            casilla(4, 9);
            if (progr)
            {
                if (BtnE10.BackColor == Color.LightGray)
                {
                    BtnE10.BackColor = Color.Red;
                }
                else
                {
                    BtnE10.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnE11_Click(object sender, EventArgs e)
        {
            casilla(4, 10);
            if (progr)
            {
                if (BtnE11.BackColor == Color.LightGray)
                {
                    BtnE11.BackColor = Color.Red;
                }
                else
                {
                    BtnE11.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnE12_Click(object sender, EventArgs e)
        {
            casilla(4, 11);
            if (progr)
            {
                if (BtnE12.BackColor == Color.LightGray)
                {
                    BtnE12.BackColor = Color.Red;
                }
                else
                {
                    BtnE12.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnF1_Click(object sender, EventArgs e)
        {
            casilla(5, 0);
            if (progr)
            {
                if (BtnF1.BackColor == Color.LightGray)
                {
                    BtnF1.BackColor = Color.Red;
                }
                else
                {
                    BtnF1.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnF2_Click(object sender, EventArgs e)
        {
            casilla(5, 1);
            if (progr)
            {
                if (BtnF2.BackColor == Color.LightGray)
                {
                    BtnF2.BackColor = Color.Red;
                }
                else
                {
                    BtnF2.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnF3_Click(object sender, EventArgs e)
        {
            casilla(5, 2);
            if (progr)
            {
                if (BtnF3.BackColor == Color.LightGray)
                {
                    BtnF3.BackColor = Color.Red;
                }
                else
                {
                    BtnF3.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnF4_Click(object sender, EventArgs e)
        {
            casilla(5, 3);
            if (progr)
            {
                if (BtnF4.BackColor == Color.LightGray)
                {
                    BtnF4.BackColor = Color.Red;
                }
                else
                {
                    BtnF4.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnF5_Click(object sender, EventArgs e)
        {
            casilla(5, 4);
            if (progr)
            {
                if (BtnF5.BackColor == Color.LightGray)
                {
                    BtnF5.BackColor = Color.Red;
                }
                else
                {
                    BtnF5.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnF6_Click(object sender, EventArgs e)
        {
            casilla(5, 5);
            if (progr)
            {
                if (BtnF6.BackColor == Color.LightGray)
                {
                    BtnF6.BackColor = Color.Red;
                }
                else
                {
                    BtnF6.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnF7_Click(object sender, EventArgs e)
        {
            casilla(5, 6);
            if (progr)
            {
                if (BtnF7.BackColor == Color.LightGray)
                {
                    BtnF7.BackColor = Color.Red;
                }
                else
                {
                    BtnF7.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnF8_Click(object sender, EventArgs e)
        {
            casilla(5, 7);
            if (progr)
            {
                if (BtnF8.BackColor == Color.LightGray)
                {
                    BtnF8.BackColor = Color.Red;
                }
                else
                {
                    BtnF8.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnF9_Click(object sender, EventArgs e)
        {
            casilla(5, 8);
            if (progr)
            {
                if (BtnF9.BackColor == Color.LightGray)
                {
                    BtnF9.BackColor = Color.Red;
                }
                else
                {
                    BtnF9.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnF10_Click(object sender, EventArgs e)
        {
            casilla(5, 9);
            if (progr)
            {
                if (BtnF10.BackColor == Color.LightGray)
                {
                    BtnF10.BackColor = Color.Red;
                }
                else
                {
                    BtnF10.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnF11_Click(object sender, EventArgs e)
        {
            casilla(5, 10);
            if (progr)
            {
                if (BtnF11.BackColor == Color.LightGray)
                {
                    BtnF11.BackColor = Color.Red;
                }
                else
                {
                    BtnF11.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnF12_Click(object sender, EventArgs e)
        {
            casilla(5, 11);
            if (progr)
            {
                if (BtnF12.BackColor == Color.LightGray)
                {
                    BtnF12.BackColor = Color.Red;
                }
                else
                {
                    BtnF12.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnG1_Click(object sender, EventArgs e)
        {
            casilla(6, 0);
            if (progr)
            {
                if (BtnG1.BackColor == Color.LightGray)
                {
                    BtnG1.BackColor = Color.Red;
                }
                else
                {
                    BtnG1.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnG2_Click(object sender, EventArgs e)
        {
            casilla(6, 1);
            if (progr)
            {
                if (BtnG2.BackColor == Color.LightGray)
                {
                    BtnG2.BackColor = Color.Red;
                }
                else
                {
                    BtnG2.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnG3_Click(object sender, EventArgs e)
        {
            casilla(6, 2);
            if (progr)
            {
                if (BtnG3.BackColor == Color.LightGray)
                {
                    BtnG3.BackColor = Color.Red;
                }
                else
                {
                    BtnG3.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnG4_Click(object sender, EventArgs e)
        {
            casilla(6, 3);
            if (progr)
            {
                if (BtnG4.BackColor == Color.LightGray)
                {
                    BtnG4.BackColor = Color.Red;
                }
                else
                {
                    BtnG4.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnG5_Click(object sender, EventArgs e)
        {
            casilla(6, 4);
            if (progr)
            {
                if (BtnG5.BackColor == Color.LightGray)
                {
                    BtnG5.BackColor = Color.Red;
                }
                else
                {
                    BtnG5.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnG6_Click(object sender, EventArgs e)
        {
            casilla(6, 5);
            if (progr)
            {
                if (BtnG6.BackColor == Color.LightGray)
                {
                    BtnG6.BackColor = Color.Red;
                }
                else
                {
                    BtnG6.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnG7_Click(object sender, EventArgs e)
        {
            casilla(6, 6);
            if (progr)
            {
                if (BtnG7.BackColor == Color.LightGray)
                {
                    BtnG7.BackColor = Color.Red;
                }
                else
                {
                    BtnG7.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnG8_Click(object sender, EventArgs e)
        {
            casilla(6, 7);
            if (progr)
            {
                if (BtnG8.BackColor == Color.LightGray)
                {
                    BtnG8.BackColor = Color.Red;
                }
                else
                {
                    BtnG8.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnG9_Click(object sender, EventArgs e)
        {
            casilla(6, 8);
            if (progr)
            {
                if (BtnG9.BackColor == Color.LightGray)
                {
                    BtnG9.BackColor = Color.Red;
                }
                else
                {
                    BtnG9.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnG10_Click(object sender, EventArgs e)
        {
            casilla(6, 9);
            if (progr)
            {
                if (BtnG10.BackColor == Color.LightGray)
                {
                    BtnG10.BackColor = Color.Red;
                }
                else
                {
                    BtnG10.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnG11_Click(object sender, EventArgs e)
        {
            casilla(6, 10);
            if (progr)
            {
                if (BtnG11.BackColor == Color.LightGray)
                {
                    BtnG11.BackColor = Color.Red;
                }
                else
                {
                    BtnG11.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnG12_Click(object sender, EventArgs e)
        {
            casilla(6, 11);
            if (progr)
            {
                if (BtnG12.BackColor == Color.LightGray)
                {
                    BtnG12.BackColor = Color.Red;
                }
                else
                {
                    BtnG12.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnH1_Click(object sender, EventArgs e)
        {
            casilla(7, 0);
            if (progr)
            {
                if (BtnH1.BackColor == Color.LightGray)
                {
                    BtnH1.BackColor = Color.Red;
                }
                else
                {
                    BtnH1.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnH2_Click(object sender, EventArgs e)
        {
            casilla(7, 1);
            if (progr)
            {
                if (BtnH2.BackColor == Color.LightGray)
                {
                    BtnH2.BackColor = Color.Red;
                }
                else
                {
                    BtnH2.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnH3_Click(object sender, EventArgs e)
        {
            casilla(7, 2);
            if (progr)
            {
                if (BtnH3.BackColor == Color.LightGray)
                {
                    BtnH3.BackColor = Color.Red;
                }
                else
                {
                    BtnH3.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnH4_Click(object sender, EventArgs e)
        {
            casilla(7, 3);
            if (progr)
            {
                if (BtnH4.BackColor == Color.LightGray)
                {
                    BtnH4.BackColor = Color.Red;
                }
                else
                {
                    BtnH4.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnH5_Click(object sender, EventArgs e)
        {
            casilla(7, 4);
            if (progr)
            {
                if (BtnH5.BackColor == Color.LightGray)
                {
                    BtnH5.BackColor = Color.Red;
                }
                else
                {
                    BtnH5.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnH6_Click(object sender, EventArgs e)
        {
            casilla(7, 5);
            if (progr)
            {
                if (BtnH6.BackColor == Color.LightGray)
                {
                    BtnH6.BackColor = Color.Red;
                }
                else
                {
                    BtnH6.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnH7_Click(object sender, EventArgs e)
        {
            casilla(7, 6);
            if (progr)
            {
                if (BtnH7.BackColor == Color.LightGray)
                {
                    BtnH7.BackColor = Color.Red;
                }
                else
                {
                    BtnH7.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnH8_Click(object sender, EventArgs e)
        {
            casilla(7, 7);
            if (progr)
            {
                if (BtnH8.BackColor == Color.LightGray)
                {
                    BtnH8.BackColor = Color.Red;
                }
                else
                {
                    BtnH8.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnH9_Click(object sender, EventArgs e)
        {
            casilla(7, 8);
            if (progr)
            {
                if (BtnH9.BackColor == Color.LightGray)
                {
                    BtnH9.BackColor = Color.Red;
                }
                else
                {
                    BtnH9.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnH10_Click(object sender, EventArgs e)
        {
            casilla(7, 9);
            if (progr)
            {
                if (BtnH10.BackColor == Color.LightGray)
                {
                    BtnH10.BackColor = Color.Red;
                }
                else
                {
                    BtnH10.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnH11_Click(object sender, EventArgs e)
        {
            casilla(7, 10);
            if (progr)
            {
                if (BtnH11.BackColor == Color.LightGray)
                {
                    BtnH11.BackColor = Color.Red;
                }
                else
                {
                    BtnH11.BackColor = Color.LightGray;
                }
            }
        }

        private void BtnH12_Click(object sender, EventArgs e)
        {
            casilla(7, 11);
            if (progr)
            {
                if (BtnH12.BackColor == Color.LightGray)
                {
                    BtnH12.BackColor = Color.Red;
                }
                else
                {
                    BtnH12.BackColor = Color.LightGray;
                }
            }
        }

        private void GpControl_Enter(object sender, EventArgs e)
        {

        }

        private void BtnCalibrar_Click(object sender, EventArgs e)
        {
            SpPrin.WriteLine("c,\n");
            SpPrin.ReadTimeout = 12000;
            string pass = SpPrin.ReadLine();
            Task.Delay(500).Wait();
            LblCasilla.Text = "A1";
            LblEstado.Text = "Calibrado";
        } 
        #endregion

        private void BtnEjecutar_Click(object sender, EventArgs e)
        {
            if(movlib)
            {
                SpPrin.WriteLine("a,");
                SpPrin.ReadTimeout = 7000;
                string nota = SpPrin.ReadLine();
                string[] datos = new string[4];
                datos = nota.Split(',');
                TxtDatos.Text = "R: " + datos[0] + "\r\n" + "G: " + datos[1] + "\r\n" + "B: " + datos[2] + "\r\n" + "C: " + datos[3];
            }
            if (progr)
            {
                List<sens> dato = new List<sens>();
                for (int alfa = 0; alfa < 8; alfa++)
                {
                    for (int beta = 0; beta < 12; beta++)
                    {
                        if (posi[alfa, beta] == true)
                        {
                            LblCasilla.Text = Convert.ToString(let[alfa]) + Convert.ToString(beta + 1);
                            //Task.Delay(3000).Wait();
                            
                            double r = 0, g = 0, b = 0, c = 0;
                            int rep = Convert.ToInt32(NudRepet.Value);
                            for(int fd = 0; fd < rep; fd++)
                            {
                                SpPrin.WriteLine("b," + Properties.Settings.Default.Horizontal[alfa] + "," + Properties.Settings.Default.Vertical[beta] + ",");
                                SpPrin.ReadTimeout = 12000;
                                string pass = SpPrin.ReadLine();
                                Task.Delay(500).Wait();
                                SpPrin.WriteLine("a,");
                                SpPrin.ReadTimeout = 10000;
                                string nota = SpPrin.ReadLine();
                                string[] datos = new string[4];
                                datos = nota.Split(',');
                                r = r + Convert.ToDouble(datos[0]);
                                g = g + Convert.ToDouble(datos[1]);
                                b = b + Convert.ToDouble(datos[2]);
                                c = c + Convert.ToDouble(datos[3]);
                            }
                            r = r / rep;
                            g = g / rep;
                            b = b / rep;
                            c = c / rep;
                            dato.Add(new sens { Casilla = Convert.ToString(let[alfa]) + Convert.ToString(beta + 1), Rojo = Convert.ToString(r), Verde = Convert.ToString(g), Azul = Convert.ToString(b), Claro = Convert.ToString(c) });
                        }
                    }
                }
                string saveFile = "Datos/" + DateTime.Now.ToString("yyyy-dd-M--HH-mm-ss") + ".csv";
                if(!Directory.Exists("Datos"))
                {
                    Directory.CreateDirectory("Datos");
                }
                using (var writer = new StreamWriter(saveFile))
                using (var csv = new CsvWriter(writer, System.Globalization.CultureInfo.InvariantCulture))
                {
                    csv.WriteRecords(dato);
                }
                Console.Beep(3000, 1000);
                MessageBox.Show("El proceso a terminado y los datos han sido guardados.","Término", MessageBoxButtons.OK);
            }
        }

        private void BtnCarpeta_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(@"Datos\");
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void BtnInicio_Click(object sender, EventArgs e)
        {
            SpPrin.Write("b,0,0,");
        }

        private void RbtnVer_CheckedChanged(object sender, EventArgs e)
        {
            if(RbtnVer.Checked == true)
            {
                CbxSelec.Items.Clear();
                for(int i = 0; i < 12; i++)
                {
                    CbxSelec.Items.Add(i+1);
                }
            }
        }

        private void RbtnHor_CheckedChanged(object sender, EventArgs e)
        {
            if(RbtnHor.Checked == true)
            {
                CbxSelec.Items.Clear();
                for(int i = 0; i < 8; i++)
                {
                    CbxSelec.Items.Add(let[i]);
                }
            }
        }

        private void CbxSelec_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(RbtnVer.Checked == true)
            {
                DudGrados.Text = Convert.ToString(Properties.Settings.Default.Vertical[CbxSelec.SelectedIndex]);
            }
            else if(RbtnHor.Checked == true)
            {
                DudGrados.Text = Convert.ToString(Properties.Settings.Default.Horizontal[CbxSelec.SelectedIndex]);
            }
        }

        private void BtnAnt_Click(object sender, EventArgs e)
        {
            if(CbxSelec.SelectedIndex != 0 )
            {
                if(RbtnVer.Checked == true)
                {
                    SpPrin.Write("b,0," + Properties.Settings.Default.Vertical[CbxSelec.SelectedIndex - 1] + "," + "\n");
                }
                else if(RbtnHor.Checked == true)
                {
                    SpPrin.Write("b," + Properties.Settings.Default.Horizontal[CbxSelec.SelectedIndex - 1] + ",0," + "\n");
                }
            }
        }

        private void BtnSig_Click(object sender, EventArgs e)
        {
            if(CbxSelec.SelectedIndex < 11 && RbtnVer.Checked == true)
            {
                SpPrin.Write("b,0," + Properties.Settings.Default.Vertical[CbxSelec.SelectedIndex + 1] + "," + "\n");
            }
            else if (CbxSelec.SelectedIndex < 7 && RbtnHor.Checked == true)
            {
                SpPrin.Write("b," + Properties.Settings.Default.Horizontal[CbxSelec.SelectedIndex + 1] + ",0," + "\n");
            }
        }

        private void BtnAcci_Click(object sender, EventArgs e)
        {
            if (RbtnVer.Checked == true)
            {
                SpPrin.Write("b,0," + DudGrados.Value + "," + "\n");
            }
            else if (RbtnHor.Checked == true)
            {
                SpPrin.Write("b," + DudGrados.Value + ",0," + "\n");
            }
        }

        private void BtnGuardar_Click(object sender, EventArgs e)
        {
            if(RbtnVer.Checked == true)
            {
                Properties.Settings.Default.Vertical[CbxSelec.SelectedIndex] = Convert.ToInt32(DudGrados.Text);
                Properties.Settings.Default.Save();
                Properties.Settings.Default.Upgrade();
            }
            else if(RbtnHor.Checked == true)
            {
                Properties.Settings.Default.Horizontal[CbxSelec.SelectedIndex] = Convert.ToInt32(DudGrados.Text);
                Properties.Settings.Default.Save();
                Properties.Settings.Default.Upgrade();
            }
            Console.Beep(700, 100);
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            
        }

        private void NudRepet_ValueChanged(object sender, EventArgs e)
        {
            
        }

        private void BtnGuardarRep_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.loops = Convert.ToInt32(NudRepet.Value);
            Properties.Settings.Default.Save();
            Properties.Settings.Default.Upgrade();
            Console.Beep(700, 100);
        }

        private void RdbEncLuz_CheckedChanged(object sender, EventArgs e)
        {
            if(RdbEncLuz.Checked == true)
            {
                SpPrin.Write("l,s,\n");
            }
        }

        private void RdbApLuz_CheckedChanged(object sender, EventArgs e)
        {
            if (RdbApLuz.Checked == true)
            {
                SpPrin.Write("l,n,\n");
            }
        }
    }
}
