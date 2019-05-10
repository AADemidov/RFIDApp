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
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;





namespace RFIDApp
{
    public partial class main_form : Form
    {
        public main_form()
        {
            InitializeComponent();
        }
        public static SerialPort port = new SerialPort();  // Глобальная переменная 
        private void button1_Click(object sender, EventArgs e)
        {
            if (!port.IsOpen)
            {
                try
                {                                            // настройки порта
                    port.PortName = comboBox1.SelectedItem.ToString();
                    port.BaudRate = 38400;
                    port.DataBits = 8;
                    port.Parity = Parity.None;
                    port.StopBits = StopBits.One;
                    port.ReadTimeout = 1000;
                    port.WriteTimeout = 1000;
                    port.Open();
                    SerialDataReceivedEventHandler SerialPort_DataReceived = null;
                    port.DataReceived += SerialPort_DataReceived;
                    label1.Text = "Порт открыт.";
                }
                catch (Exception ex)
                {
                   MessageBox.Show("ERROR: невозможно открыть порт:" + ex.ToString());
                    return;
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            label1.Text = "Порт закрыт.";
            // Получаем список портов
            string[] ports = SerialPort.GetPortNames();
            // Заносим список портов в comboBox
            for (int i = 0; i < ports.Length; i++)
            {
                comboBox1.Items.Add(ports[i].ToString());
            }
            comboBox1.SelectedIndex = 0; //Выбираем COM3 чтобы быстрее было
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (port.IsOpen)
            {
                label1.Text = "Порт закрыт.";
                port.Close();
            }
        }
        public string ReadFromCOM()
        {
            string temp = "";
            if (port.IsOpen)
            {

                try
                {

                    byte[] bMas = { 0xa0, 0x04, 0x01, 0x89, 0x01, 0xd1 };
                    port.Write(bMas, 0, bMas.Length);
                    byte[] data = new byte[Convert.ToInt32(textBox1.Text)];
                    port.Read(data, 0, data.Length);

                    for (int i = 0; i < Convert.ToInt32(textBox1.Text); i++)
                        temp += data[i].ToString() + " ";
                }
                catch (Exception ex)
                {
                    return ("ERROR: невозможно открыть порт:" + ex.ToString());
                }
            }
            else
            {
                temp = "Открой порт сапака!";
            }
            return DateTime.Now.Hour.ToString() + ":" + DateTime.Now.Minute.ToString() + ":"+ DateTime.Now.Second.ToString() + ":" + DateTime.Now.Millisecond.ToString() + ">>" + temp + "\r\n";
        }
        

        private void button3_Click(object sender, EventArgs e)
        {
            richTextBox1.AppendText(ReadFromCOM());
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (timer1.Enabled == true)
                timer1.Stop();
            else
                timer1.Start();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            richTextBox1.AppendText(ReadFromCOM());
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {

            if (port.IsOpen) // если порт открыт
            {

                Excel.Application ex = new Microsoft.Office.Interop.Excel.Application(); //создаем COM-объект Excel

                ex.Visible = true; //делаем объект видимым

                ex.SheetsInNewWorkbook = 1;//количество листов в книге

                ex.Workbooks.Add(Type.Missing); //добавляем книгу

                Excel.Workbook workbook = ex.Workbooks[1]; //Получаем первый лист документа (счет начинается с 1)

                Excel.Worksheet sheet = workbook.Worksheets.get_Item(1); //ссылка на первый лист

                sheet.Name = "Отчет" + DateTime.Now.Hour.ToString(); //Название листа (вкладки снизу)

                // Заполняем столб А
                for (int i = 0; i < 10; i++)
                {
                    sheet.Cells[i, 1].Value = i;
                }

                // Заполняем столб B
                for (int j = 1; j <= 10; j++)
                {
                    sheet.Cells[j, 2].Value = j;
                }


            }

        }
    }
}
