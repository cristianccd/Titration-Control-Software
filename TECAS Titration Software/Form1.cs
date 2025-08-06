using LabJack.LabJackUD;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Timers;
using System.Windows.Forms;


namespace TECAS_Titration_Software
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        //Variables
        string _LoadCal;
        double SyrCalSlope, SyrCalIntercept, SyrR2, SyrDiameter, SyrCapacity, FirstpHVal, VoltoInf;
        int SyrCOM;
        bool FirstpH=true;
        StreamWriter sw;

        //Labjack Variables (shared)
        private U3 u3;
        LJUD.IO ioType = 0;
        LJUD.CHANNEL channel = 0;
        double dblValue = 0, dummyDouble = 0;
        int dummyInt = 0;

        //Other Vars
        double AccumVolInf, Ticks;
        DateTime ExpStart;
        System.Timers.Timer aTimer;

        //pH Vars
        double pHAccVal, pHAvgVal, pHDeviation, FirstDer, SecDer;
        int PointIndex;
        DateTime InfStart;

        bool Paused = false, Stoped = false;

        //------------------------------------------------------------------------------------
        private void syringeCalibrationFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    _LoadCal = File.ReadAllText(openFileDialog1.FileName);
                    /*COM#Diameter#Capacity#x1#y1#x2#y2#slope#intercept#r2#*/
                    SyrCOM = Convert.ToInt16(_LoadCal.Split('#')[0]);
                    SyrDiameter = Convert.ToDouble(_LoadCal.Split('#')[1]);
                    SyrCapacity = Convert.ToDouble(_LoadCal.Split('#')[2]);
                    SyrCalSlope = Convert.ToDouble(_LoadCal.Split('#')[7]);
                    SyrCalIntercept = Convert.ToDouble(_LoadCal.Split('#')[8]);
                    SyrR2 = Convert.ToDouble(_LoadCal.Split('#')[9]);

                    comboBox3.SelectedIndex = SyrCOM;
                    label19.Text = "Diameter: " + SyrDiameter + " mm";
                    label20.Text = "Capacity: "+ SyrCapacity +" ml";

                    if (SyrCalIntercept >= 0)
                        label17.Text = "Y=" + String.Format("{0:0.000}", SyrCalSlope) + " x+" + String.Format("{0:0.000}", SyrCalIntercept);
                    else
                        label17.Text = "Y=" + String.Format("{0:0.000}", SyrCalSlope) + " x" + String.Format("{0:0.000}", SyrCalIntercept);
                    label21.Text = "R =  " + String.Format("{0:0.000}", SyrR2);
                    checkBox2.Checked = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Calibration", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        double pHCalSlope, pHCalIntercept, pHR2;
        private void electrodeCalibrationFIleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (openFileDialog2.ShowDialog() == DialogResult.OK)
                {
                    _LoadCal = File.ReadAllText(openFileDialog2.FileName);
                    pHCalSlope = Convert.ToDouble(_LoadCal.Split('#')[0]);
                    pHCalIntercept = Convert.ToDouble(_LoadCal.Split('#')[1]);
                    pHR2 = Convert.ToDouble(_LoadCal.Split('#')[2]);
                    if (pHCalIntercept >= 0)
                        label18.Text = "Y=" + String.Format("{0:0.000}", pHCalSlope) + " x+" + String.Format("{0:0.000}", pHCalIntercept);
                    else
                        label18.Text = "Y=" + String.Format("{0:0.000}", pHCalSlope) + " x" + String.Format("{0:0.000}", pHCalIntercept);
                    label24.Text = "R  =" + String.Format("{0:0.000}", pHR2); ;
                    checkBox3.Checked = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Calibration", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
                textBox1.Enabled = true;
            else
                textBox1.Enabled = false;
        }

        double pHEnd, AdditionVol;
        int MixingTime;
        private bool CheckParameters()
        {
            if (checkBox1.Checked == true)
            {
                try
                {
                    pHEnd = Convert.ToDouble(textBox1.Text);
                    if (pHEnd < 0)
                    {
                        MessageBox.Show("pH must be possitive!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }

            try
            {
                MixingTime = Convert.ToInt16(textBox2.Text);
                if (MixingTime < 0)
                {
                    MessageBox.Show("Mixing Time must be possitive!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error Mixing Time!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            try
            {
                AdditionVol = Convert.ToDouble(textBox3.Text);
                if (AdditionVol < 30 )
                {
                    MessageBox.Show("Addition Volume should be over 30ul!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error Addition Volume!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (comboBox3.SelectedIndex == -1)
            {
                MessageBox.Show("Please select a correct COM port", "Error COM!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (checkBox2.Checked == false || checkBox3.Checked == false)
            {
                MessageBox.Show("Calibrations not loaded, please load the files", "Error Calibrations!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return true;        
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!CheckParameters())
                return;
            //Open the serial port
            try
            {
                serialPort1.PortName = "COM" + Convert.ToString(comboBox3.SelectedIndex);
                serialPort1.Open();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Port opening", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //Calculate the real volume to infuse
            VoltoInf = (AdditionVol - SyrCalIntercept) / SyrCalSlope;

            //Configure the pump
            serialPort1.Write("STP\r\n");
            System.Threading.Thread.Sleep(20);
            serialPort1.Write("CLD INF\r\n");
            System.Threading.Thread.Sleep(20);
            serialPort1.Write("CLD WDR\r\n");
            System.Threading.Thread.Sleep(20);
            serialPort1.Write("DIR INF\r\n");
            System.Threading.Thread.Sleep(20);
            serialPort1.Write("DIA " + SyrDiameter + "\r\n");
            System.Threading.Thread.Sleep(20);
            serialPort1.Write("RAT 800 MH\r\n"); //Rate fixed
            System.Threading.Thread.Sleep(20);
            serialPort1.Write("VOL UL\r\n");
            System.Threading.Thread.Sleep(20);
            serialPort1.Write("VOL " + String.Format("{0:000.0}", VoltoInf) + "\r\n");
            System.Threading.Thread.Sleep(20);

            //Configure the DAQ
            try
            {
                if (u3 == null)
                    u3 = new U3(LJUD.CONNECTION.USB, "1", true); // Connection through USB
                LJUD.ePut(u3.ljhandle, LJUD.IO.PIN_CONFIGURATION_RESET, 0, 0, 0);
                LJUD.ePut(u3.ljhandle, LJUD.IO.PUT_ANALOG_ENABLE_PORT, 0, 31, 16);//first 4 FIO analog b0000000000001111
                LJUD.AddRequest(u3.ljhandle, LJUD.IO.GET_AIN_DIFF, 4, 0, 32, 0);//Request FIO4
            }
            catch (LabJackUDException h)
            {
                MessageBox.Show(h.ToString(), "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                serialPort1.Close();
                return;
            }
            
            //Disable the parameters panel
            panel4.Enabled = false;
            //Disable the start button and enable pause and stop
            button1.Enabled = false;
            button2.Enabled = true;
            button3.Enabled = true;
            //Clear Counters and Vars
            AccumVolInf = 0;
            pHAccVal = 0;
            pHAvgVal = 0;
            Ticks = 0;
            FirstpH = true;

            label5.Text="0.000";
            label9.Text="0 ul";
            label13.Text="0.0000 V";
            label11.Text="0 days 00:00:00";
            label7.Text="0.000";
            
            //Clear the graph
            foreach (var series in chart1.Series)
                series.Points.Clear();
            //Get the time
            ExpStart = DateTime.Now;
            InfStart = DateTime.Now;
            aTimer = new System.Timers.Timer(MixingTime*1000);
            aTimer.Elapsed += new ElapsedEventHandler(OnTimedEventA);

            //Export Data
            string _FirstLine = "pH Set: "+ Convert.ToString(pHEnd) +",VOLUME,Volume,pH,,DPH/DV,Volume, dpH/dV,,D(DPH)/D(DV),Volume,d(dpH)/d(dV),,Time\n";
            string[] Content = new string[chart1.Series["Series1"].Points.Count];
            string Path = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\Calcification Experiments";
            try
            {
                if (!System.IO.Directory.Exists(Path))
                    System.IO.Directory.CreateDirectory(Path);
                sw = new StreamWriter(Path + @"\" + DateTime.Now.ToString("yyyy-MM-dd HHmmss") + " Titration.csv");
                sw.Write(_FirstLine);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error Writing File", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            //End Export Data

            aTimer.Enabled = true;
            timer1.Enabled = true;
        }

        //Stop
        private void button3_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Are you sure you want to stop?", "Stop?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;
            Stop();
        }

        //Change graphs series
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBox1.SelectedIndex)
            {
                case -1:
                    chart1.Series["Series1"].Enabled = true;
                    chart1.Series["Series2"].Enabled = false;
                    chart1.Series["Series3"].Enabled = false;
                    break;
                case 0:
                    chart1.Series["Series1"].Enabled = true;
                    chart1.Series["Series2"].Enabled = false;
                    chart1.Series["Series3"].Enabled = false;
                    break;
                case 1:
                    chart1.Series["Series1"].Enabled = false;
                    chart1.Series["Series2"].Enabled = true;
                    chart1.Series["Series3"].Enabled = false;
                    break;
                case 2:
                    chart1.Series["Series1"].Enabled = false;
                    chart1.Series["Series2"].Enabled = false;
                    chart1.Series["Series3"].Enabled = true;
                    break;
            }
            chart1.Update();
        }

        //Add points to the Graph
        private void OnTimedEventA(Object source, ElapsedEventArgs e)
        {
            if (InvokeRequired)
            {
                Invoke(new MethodInvoker(delegate
                {                                       
                    serialPort1.Write("RUN\r\n");
                    System.Threading.Thread.Sleep(20);
                    AccumVolInf = AccumVolInf + AdditionVol;
                    label9.Text = String.Format("{0:00000.00}", AccumVolInf) + " ul";
                    chart1.Series["Series1"].Points.AddXY(AccumVolInf, pHAvgVal);
                    PointIndex = chart1.Series["Series1"].Points.Count;

                    if (PointIndex > 1 && (chart1.Series["Series1"].Points[PointIndex - 1].XValue - chart1.Series["Series1"].Points[PointIndex - 2].XValue) != 0)
                        FirstDer = (chart1.Series["Series1"].Points[PointIndex - 1].YValues[0] - chart1.Series["Series1"].Points[PointIndex - 2].YValues[0])
                                / (chart1.Series["Series1"].Points[PointIndex - 1].XValue - chart1.Series["Series1"].Points[PointIndex - 2].XValue);
                    chart1.Series["Series2"].Points.AddXY(AccumVolInf, FirstDer);
                    
                    if (PointIndex > 2 && (chart1.Series["Series2"].Points[PointIndex - 1].XValue - chart1.Series["Series2"].Points[PointIndex - 2].XValue) != 0)
                        SecDer = (chart1.Series["Series2"].Points[PointIndex - 1].YValues[0] - chart1.Series["Series2"].Points[PointIndex - 2].YValues[0])
                                / (chart1.Series["Series2"].Points[PointIndex - 1].XValue - chart1.Series["Series2"].Points[PointIndex - 2].XValue);
                    chart1.Series["Series3"].Points.AddXY(AccumVolInf, SecDer);

                    sw.Write(",," + chart1.Series["Series1"].Points[PointIndex-1].XValue + "," 
                        + chart1.Series["Series1"].Points[PointIndex - 1].YValues[0] +",,," + chart1.Series["Series2"].Points[PointIndex - 1].XValue 
                        + "," + chart1.Series["Series2"].Points[PointIndex - 1].YValues[0] + ",,," + chart1.Series["Series3"].Points[PointIndex - 1].XValue 
                        + "," + chart1.Series["Series3"].Points[PointIndex - 1].YValues[0] + ",," + label11.Text + "\n");
                }));
            }
        }
        //Get the DAQ Results every 50ms

        private void timer1_Tick(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            //Refresh GUI
            label13.Text = String.Format("{0:0.00000 V}", dblValue);
            label11.Text = String.Format("{0}", (DateTime.Now - ExpStart).Days) + " days " + String.Format("{0:00}", (DateTime.Now - ExpStart).Hours) + ":" + String.Format("{0:00}", (DateTime.Now - ExpStart).Minutes) + ":" + String.Format("{0:00}", (DateTime.Now - ExpStart).Seconds);

            bool requestedExit = false;
            while (!requestedExit)
            {
                try
                {
                    LJUD.GoOne(u3.ljhandle);
                    LJUD.GetFirstResult(u3.ljhandle, ref ioType, ref channel, ref dblValue, ref dummyInt, ref dummyDouble);
                }
                catch (LabJackUDException)
                {
                    MessageBox.Show("Error getting the DAQ results", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                if (ioType == LJUD.IO.GET_AIN_DIFF)
                {
                    Ticks++;
                    pHAccVal = pHAccVal + ((dblValue - pHCalIntercept) / pHCalSlope);
                    if (Ticks >= 10)
                    {
                        pHAvgVal = pHAccVal / 10;
                        if (checkBox1.Checked == true)
                        {
                            pHDeviation = pHEnd - pHAvgVal;
                            label7.Text = String.Format("{0:0.000}", pHDeviation);
                            if (FirstpH)
                            {
                                FirstpHVal = pHAvgVal;
                                FirstpH=false;
                            }
                            //Check for pH start and end
                            else
                            {
                                if (FirstpHVal < pHEnd) //pH is increasing
                                {
                                    if (pHAvgVal >= pHEnd) //stop
                                    {
                                        Stoped = Stop();
                                    }
                                }
                                else //pH decreasing
                                {
                                    if (pHAvgVal <= pHEnd) //stop
                                    {
                                        Stoped = Stop();
                                    }
                                }
                            }

                        }
                        label5.Text = String.Format("{0:0.000}", pHAvgVal);
                        Ticks = 0;
                        pHAccVal = 0;
                    }
                }
                try
                {
                    LJUD.GetNextResult(u3.ljhandle, ref ioType, ref channel, ref dblValue, ref dummyInt, ref dummyDouble);
                }
                catch (LabJackUDException h)
                {
                    if (h.LJUDError == U3.LJUDERROR.NO_MORE_DATA_AVAILABLE)
                        requestedExit = true;//no more data to read
                    else
                        MessageBox.Show(h.ToString(), "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            if (Stoped)
            {
                timer1.Enabled = false;
                Stoped = false;
            }
            else
                timer1.Enabled = true;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (MessageBox.Show("Are you sure you want to close?", "Quit?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                e.Cancel = true;
                return;
            }
            if (serialPort1.IsOpen)
                serialPort1.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if(!Paused)
            {   
                aTimer.Enabled = false;
                button2.Text="Unpause";
                Paused = true;
            }
            else
            {
                aTimer.Enabled = true;
                button2.Text = "Pause";
                Paused = false;
            }
        }

        private bool Stop()
        {
            if (serialPort1.IsOpen)
                serialPort1.Close();
            timer1.Enabled = false;
            aTimer.Enabled = false;
            sw.Close();
            //Enable start again
            button1.Enabled = true;
            button2.Enabled = false;
            button3.Enabled = false;
            panel4.Enabled = true;
            return true;
        }
    }
}
