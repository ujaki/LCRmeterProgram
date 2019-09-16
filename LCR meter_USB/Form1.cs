using System;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Net.Sockets;

//LCR meter 프로그램 LAN
//192.168.0.01:3500
namespace LCR_meter_USB
{

    public partial class ATC_LCRMeter : Form
    {
       private TcpClient LanSocket;   //create TCP socket
       private double timestamp = 0;  //시간 경과 체크용 변수
       private FolderBrowserDialog dialog = new FolderBrowserDialog();
       private string storePath = null; //파일 저장경로
       private string MsgBuf = "";
       private bool isSetCOM = false;    
       private Stopwatch sw = new Stopwatch();
       private int lineWidth = 2;
       private int timePass = 0;
       private DateTime datetime = DateTime.Now;  //acquires date and time
       private string filename;
       private StreamWriter fp;
       private string newLine = System.Environment.NewLine; //\r\n
        

        public ATC_LCRMeter() //생성자
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void button1_Click(object sender, EventArgs e) //start measurement button
        {
            timer1.Start(); //타이머 시작
            timer2.Start();
        }

        private long time_count = 0;
        
        private void timer1_Tick(object sender, EventArgs e)   //타이머 호출 주기 : 10 ms
        {
            int length;
            byte[] ReceiveBuffer = new byte[40];
            try
            {

                //SendMsg(":MEAS?");
                SendMsg(":MEAS?");
                if (LanSocket.GetStream().DataAvailable)
                {

                    length = LanSocket.GetStream().Read(ReceiveBuffer, 0, 40);
                    string data = Encoding.Default.GetString(ReceiveBuffer, 0, length).Replace("\r\n", "");
                    //DateTime dateTime = DateTime.Now;
                    data = data + "," + time_count.ToString() + newLine;
                    time_count += timer1.Interval;
                    fp.Write(data);

                    textBox1.AppendText(data);

                    //MessageBox.Show(sw.ElapsedTicks.ToString());
                    //sw.ElapsedMilliseconds
                    //textBox1.AppendText(length.ToString()+"\r\n");
                    //시간 데이터 저장 코드 
                    //MsgBuf = Encoding.Default.GetString(ReceiveBuffer);
                    // string data = Encoding.Default.GetString(ReceiveBuffer, 0, length).Replace("\r\n", "");
                    //data = data.Split(;
                    //textBox1.AppendText(data);
                    // dataQ.Enqueue(data);
                    //textBox1.AppendText(data);
                    //  break;
                }
                else return;
               
                
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            // double cP;
            //sw.Reset(); //stop watch
            //sw.Start(); //start sw
            //SendQueryMsg(); //query of measurement
            //sw.Stop(); //stop sw
           // MsgBuf = MsgBuf.Replace("\r\n", ""); //개행문자 제거
           // dataQ.Enqueue(MsgBuf + ',' + sw.ElapsedMilliseconds.ToString());  //측정한 값을 큐에 넣음
           // MsgBuf = ""; //MsgBuf reset
        }

        private void StopMeasuring()
        {
            SendMsg(":DISP ON");          //' Display : off
            timer1.Stop(); //측정 타이머 정지
            fp.Close();
            LanSocket.Close();
            button3.Text = "Connect";
            button3.BackColor = Color.Gray;
            button3.ForeColor = Color.Black;
            button3.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)  //stop measurement button
        {
            StopMeasuring();
        }

        private void button3_Click(object sender, EventArgs e) //connection button
        {
            try
            {
                int interval = 5;
                int.TryParse(textBox2.Text, out interval);
                timer1.Interval = interval;
                SetStart_LAN("192.168.0.1", 3500);
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SetStart_LAN(String hostName, UInt16 hostPort)
        {
            LanSocket = new TcpClient(hostName, hostPort);
         
           // SendMsg("*RST");
            SendMsg(":MODE LCR");          //' Mode : LCR
            SendMsg(":AVER 10");          //' Averaging : off
            SendMsg(":DISP OFF");          //' Display : off
                                          //Z : Impedance / Cs, Cp : Capacitance
            //for measuring Rp only
            SendMsg(":MEAS:ITEM 0,8");      //' Measurement Parameter: Z/ Y/ Phase/ Cs/ Cp/ D/ Ls/ Lp/  Q/ Rs/   G/  Rp/ X/ B/ RDC/ T/ OFF
                                            //                         1/ 2/     4/  8/ 16/  32/ 64/128/256/512/1024/2048/ ??? 이렇게 높이??
                                            //                                                       1    2    4    8   16  32  64   128  256
            //SendMsg(":MEAS:ITEM 16,0"); //for measuring Cp only

            SendMsg(":SPEE SLOW");
            SendMsg(":HEAD OFF");          //' Header: OFF
            SendMsg(":LEV V");            //' Signal level: Open-circuit voltage
            SendMsg(":LEV:VOLT 1");      // Signal level: 1V signal level
            SendMsg(":FREQ 1E3");        //' Measurement frequency: 1kHz
            SendMsg(":TRIG INT");        //' Trigger: External trigger
           // SendMsg(":TRIG:DEL 5");        //' Trigger: External trigger

            if (LanSocket.Connected)
            {
                button3.Text = "Connected";
                button3.BackColor = Color.Red;
                button3.ForeColor = Color.White;
                button3.Enabled = false;
                LanSocket.GetStream().Flush();
               
            }
        
            filename = "ATC_" + datetime.ToString("yyyy-MM-dd_HH-mm-ss") + ".csv"; //sets the file name
            fp = new StreamWriter(filename, false, Encoding.GetEncoding("shift_jis")); // file open
            //fp.Write("None,Cp,time" + newLine);
            fp.Write("None,Rp,time" + newLine);
        }

        private void SendMsg(string strMsg)
        {
            byte[] SendBuffer;
            try
            {
               SendBuffer = Encoding.Default.GetBytes(strMsg + newLine);
               LanSocket.GetStream().Write(SendBuffer, 0, SendBuffer.Length); //Send message
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
           // StopMeasuring();
           // timer2.Stop();
        }
    }
}
