using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.Data.SQLite;
using System.IO;
using System.Data.Common;

using Excel = Microsoft.Office.Interop.Excel;

namespace АнкетаОДеятельностиФСС
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Paint(object sender, PaintEventArgs e)// ----------------  при прорисовке формы --------------------
        {
            string userNameWin, compName, compIP, myHost;
            myHost = System.Net.Dns.GetHostName();// имя хоста
            compIP = System.Net.Dns.GetHostEntry(myHost).AddressList[0].ToString();// IP по имени хоста, выдает список, можно обойти в цикле весь, здесь берется первый адрес
            userNameWin = System.Environment.UserName;
            compName = System.Environment.MachineName;

            textBox1.Text = userNameWin;

            dateTimePicker1.Value = new DateTime(2016, 7, 1);

            if (userNameWin == "РябухинаВВ" || userNameWin == "ЕвстафьеваГС" || userNameWin == "ЗуеваВВ" || userNameWin == "АлексютинаОН")
            {
                comboBox1.Text = "гр_СПР";
                comboBox2.Text = "ФПМ";
                 
                dateTimePicker1.Value = new DateTime(2016, DateTime.Now.Month, 1);

            }

            if (userNameWin == "ЦареваНА" || userNameWin == "РожковаНВ" || userNameWin == "СалийЛВ")
            {
                comboBox1.Text = "специалисты_ОЛКГ";
                comboBox2.Text = "Прием заявок на СКЛ";
            }

           

             
        }//---------------------------------------------------------------------------------------------------------------------

        private void создатьНовуюАнкетуToolStripMenuItem_Click(object sender, EventArgs e)//------ показываем фоорму АНКЕТЫ 
        {
            Form2 form = new Form2();
            form.ShowDialog();//в модальном режиме 
        }

        private void настройкиToolStripMenuItem_Click(object sender, EventArgs e)//----------------показываем фоорму НАСТРОЙКИ 
        {
            Form3 form = new Form3();
            form.ShowDialog();
        }

        private void создатьToolStripButton_Click(object sender, EventArgs e)//-------------------- показываем фоорму АНКЕТЫ 
        {
            Form2 form = new Form2();
            form.ShowDialog();//в модальном режиме 
        }

        private void button1_Click(object sender, EventArgs e)//------------------------------------нажали кнопку ВЫВЕСТИ ОТЧЕТ
        {
            DataTable table = new DataTable();
            DataTable table2 = new DataTable();

            richTextBox1.Clear();

            string databaseName = Directory.GetCurrentDirectory() + "\\Anketa.db";
            SQLiteConnection connection =           new SQLiteConnection(string.Format("Data Source={0};", databaseName));
            connection.Open();
            SQLiteCommand command = new SQLiteCommand("SELECT * FROM 'Anketa' WHERE  Dat >=  '" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "' AND Dat <=  '" + dateTimePicker2.Value.Date.ToString("yyyy-MM-dd") + "' AND Podrazdelenie = '" + comboBox1.Text.ToString() + "' AND q1 = '" + comboBox2.Text.ToString() + "';", connection);
            SQLiteDataReader reader = command.ExecuteReader();
            SQLiteCommand command2 = new SQLiteCommand("SELECT * FROM 'registr' WHERE  Dat >=  '" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "' AND Dat <=  '" + dateTimePicker2.Value.Date.ToString("yyyy-MM-dd") + "' AND Podr = '" + comboBox1.Text.ToString() + "' AND Uslug = '" + comboBox2.Text.ToString() + "';  ", connection);
            SQLiteDataReader reader2 = command2.ExecuteReader();
           
            table.Load(reader);
            table2.Load(reader2);

            connection.Close();

            dataGridView1.DataSource = table;
            dataGridView2.DataSource = table2;
           

            richTextBox1.AppendText("                                        Отчет от " + dateTimePicker3.Value.Date.ToString("dd-MM-yyyy") + "\n");
            richTextBox1.AppendText("                          за период с " + dateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "  по  " + dateTimePicker2.Value.Date.ToString("dd-MM-yyyy") + "\n");

            double sum5=0, sum6=0, sum7=0 , sred5 = 0 , sred6 =0 , sred7=0 ;
            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {
                sum5 += Convert.ToInt32( dataGridView1.Rows[i].Cells[10].Value.ToString() );
                sum6 += Convert.ToInt32( dataGridView1.Rows[i].Cells[11].Value.ToString());
                sum7 += Convert.ToInt32(dataGridView1.Rows[i].Cells[12].Value.ToString());
            }

            sred5 = Math.Round(sum5 / (dataGridView1.RowCount - 1), 2);
            sred6 = Math.Round(sum6 / (dataGridView1.RowCount - 1), 2);
            sred7 = Math.Round(sum7 / (dataGridView1.RowCount - 1), 2);

            richTextBox1.AppendText("\nКол-во АНКЕТ - " + (dataGridView1.RowCount - 1).ToString() + "\n\n");
            richTextBox1.AppendText("_________________________________________________|__сумма__|__Среднее\n");
            richTextBox1.AppendText("Вопрос5 Сколько раз Вы лично пришли ...............____|__  " + sum5 + "      __|__    " + sred5 + " \n");
            richTextBox1.AppendText("Вопрос6 В очереди на подачу заяв-я ждал(а)минут__|__  " + sum6 + "      __|__    " + sred6 + " \n");
            richTextBox1.AppendText("Вопрос7 Фактического получ........... ждал(а) дней____|__  " + sum7 + "      __|__    " + sred7 + " \n");
            //for (int i = 0; i < dataGridView1.RowCount-1; i++ ) 
            //{
            //    richTextBox1.AppendText("---" + (dataGridView1.RowCount - 1).ToString() + "\n");
            //}

            // выводим в Ексель
            ToExcel(comboBox1.Text.ToString(), comboBox2.Text.ToString());
           

        }//----------------------------------------------------------------------------------------------------------
          
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text.ToString() == "гр_СПР")
            {
                comboBox2.Text = "ФПМ";
                comboBox2.Items.Clear();
                comboBox2.Items.Add("Обеспечение  по обязательному страхованию от несчастных случаев на произв. и проф. заб.");
                comboBox2.Items.Add("Назначение ЕДВ и ЕЖВ");
                comboBox2.Items.Add("Скидки");
                comboBox2.Items.Add("Дополнительные расходы");
                comboBox2.Items.Add("Подтверждение ОКВЭД");
                comboBox2.Items.Add("ФПМ");

                dateTimePicker1.Value = new DateTime(2016, DateTime.Now.Month, 1);
            }

            if (comboBox1.Text.ToString() == "отдел_АиСВ")
            {
                comboBox2.Text = "Прием расчета Форма-4 ФСС";
                comboBox2.Items.Clear();
                comboBox2.Items.Add("Прием расчета Форма-4 ФСС");
                comboBox2.Items.Add("Регистрация и снятие с регистрационного учета");

            }

            if (comboBox1.Text.ToString() == "специалисты_ОЛКГ")
            {
                comboBox2.Text = "Прием заявок на СКЛ";
                comboBox2.Items.Clear();
                comboBox2.Items.Add("Прием заявок на СКЛ");
                comboBox2.Items.Add("Обеспечение инвалидов техническими средствами реабилитации и (или) услугами");

            }





        }//-----------------------------------------------------------------------------------------------------------------------
         
        void ToExcel(string podr, string uslug) //---------------------- ФУНКЦИЯ  в  ЕКСЕЛЬ-----------------------------------------
        {
            //int sum5 = 0, sum6 = 0, sum7 = 0, sred5 = 0, sred6 = 0, sred7 = 0;
            double sum5 = 0, sum6 = 0, sum7 = 0, sred5 = 0, sred6 = 0, sred7 = 0;
            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {
                sum5 += Convert.ToInt32(dataGridView1.Rows[i].Cells[10].Value.ToString());
                sum6 += Convert.ToInt32(dataGridView1.Rows[i].Cells[11].Value.ToString());
                sum7 += Convert.ToInt32(dataGridView1.Rows[i].Cells[12].Value.ToString());
            }

            sred5 = Math.Round(sum5 / (dataGridView1.RowCount - 1),2);
            sred6 = Math.Round(sum6 / (dataGridView1.RowCount - 1),2);
            sred7 = Math.Round(sum7 / (dataGridView1.RowCount - 1), 2);


            DataTable table = new DataTable();
 
            Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(Application.StartupPath + "\\" + podr + "\\" + uslug + ".xlsx");//   открываем книгу 
            Excel.Worksheet ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];//Таблица.

            ObjWorkSheet.Cells[2, 2] = "Отчет от " + dateTimePicker3.Value.Date.ToString("dd-MM-yyyy") + "за период с " + dateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "  по  " + dateTimePicker2.Value.Date.ToString("dd-MM-yyyy");

             
            //Заполняем значениями Zapros(NomVopr,NomOt)  Ексель[строка,столбец]

            //if (comboBox2.Text == "Прием расчета Форма-4 ФСС" || comboBox2.Text == "Регистрация и снятие с регистрационного учета")
            //{
                ObjWorkSheet.Cells[6, 3] = Zapros("1", "1"); ObjWorkSheet.Cells[6, 4] = Zapros("1", "2"); ObjWorkSheet.Cells[6, 5] = Zapros("1", "3");

                ObjWorkSheet.Cells[7, 3] = Zapros("2", "1"); ObjWorkSheet.Cells[7, 4] = Zapros("2", "2"); ObjWorkSheet.Cells[7, 5] = Zapros("2", "3"); ObjWorkSheet.Cells[7, 6] = Zapros("2", "4");
                ObjWorkSheet.Cells[8, 3] = Zapros("3", "1"); ObjWorkSheet.Cells[8, 4] = Zapros("3", "2"); ObjWorkSheet.Cells[8, 5] = Zapros("3", "3"); ObjWorkSheet.Cells[8, 6] = Zapros("3", "4");

                ObjWorkSheet.Cells[9, 3] = Zapros("4", "1"); ObjWorkSheet.Cells[9, 4] = Zapros("4", "2"); ObjWorkSheet.Cells[9, 5] = Zapros("4", "3"); ObjWorkSheet.Cells[9, 6] = Zapros("4", "4");

                ObjWorkSheet.Cells[10, 16] = (dataGridView1.RowCount - 1).ToString(); ObjWorkSheet.Cells[10, 3] = "Сумма " + sum5.ToString() + " ;  Среднее " + sred5.ToString() + " РАЗ";
                ObjWorkSheet.Cells[11, 16] = (dataGridView1.RowCount - 1).ToString(); ObjWorkSheet.Cells[11, 3] = "Сумма " + sum6.ToString() + " ;  Среднее " + sred6.ToString() + " МИНУТ";
                ObjWorkSheet.Cells[12, 16] = (dataGridView1.RowCount - 1).ToString(); ObjWorkSheet.Cells[12, 3] = "Сумма " + sum7.ToString() + " ;  Среднее " + sred7.ToString() + " ДНЕЙ";


                ObjWorkSheet.Cells[13, 3] = Zapros("8", "1"); ObjWorkSheet.Cells[13, 4] = Zapros("8", "2"); ObjWorkSheet.Cells[13, 5] = Zapros("8", "3"); ObjWorkSheet.Cells[13, 6] = Zapros("8", "4"); ObjWorkSheet.Cells[13, 7] = Zapros("8", "5");
                ObjWorkSheet.Cells[14, 3] = Zapros("9", "1"); ObjWorkSheet.Cells[14, 4] = Zapros("9", "2");

                ObjWorkSheet.Cells[15, 3] = Zapros("10", "1");  ObjWorkSheet.Cells[15, 4] = Zapros("10", "2");

                ObjWorkSheet.Cells[16, 3] = Zapros("111", "1"); ObjWorkSheet.Cells[16, 4] = Zapros("111", "2"); ObjWorkSheet.Cells[16, 5] = Zapros("111", "3"); ObjWorkSheet.Cells[16, 6] = Zapros("111", "4"); ObjWorkSheet.Cells[16, 7] = Zapros("111", "5");
                ObjWorkSheet.Cells[17, 3] = Zapros("112", "1"); ObjWorkSheet.Cells[17, 4] = Zapros("112", "2"); ObjWorkSheet.Cells[17, 5] = Zapros("112", "3"); ObjWorkSheet.Cells[17, 6] = Zapros("112", "4"); ObjWorkSheet.Cells[17, 7] = Zapros("112", "5");
                ObjWorkSheet.Cells[18, 3] = Zapros("113", "1"); ObjWorkSheet.Cells[18, 4] = Zapros("113", "2"); ObjWorkSheet.Cells[18, 5] = Zapros("113", "3"); ObjWorkSheet.Cells[18, 6] = Zapros("113", "4"); ObjWorkSheet.Cells[18, 7] = Zapros("113", "5");
                ObjWorkSheet.Cells[19, 3] = Zapros("114", "1"); ObjWorkSheet.Cells[19, 4] = Zapros("114", "2"); ObjWorkSheet.Cells[19, 5] = Zapros("114", "3"); ObjWorkSheet.Cells[19, 6] = Zapros("114", "4"); ObjWorkSheet.Cells[19, 7] = Zapros("114", "5");

                ObjWorkSheet.Cells[20, 3] = Zapros("121", "1"); ObjWorkSheet.Cells[20, 4] = Zapros("121", "2"); ObjWorkSheet.Cells[20, 5] = Zapros("121", "3"); ObjWorkSheet.Cells[20, 6] = Zapros("121", "4");
                ObjWorkSheet.Cells[21, 3] = Zapros("122", "1"); ObjWorkSheet.Cells[21, 4] = Zapros("122", "2"); ObjWorkSheet.Cells[21, 5] = Zapros("122", "3"); ObjWorkSheet.Cells[21, 6] = Zapros("122", "4");
                ObjWorkSheet.Cells[22, 3] = Zapros("123", "1"); ObjWorkSheet.Cells[22, 4] = Zapros("123", "2"); ObjWorkSheet.Cells[22, 5] = Zapros("123", "3"); ObjWorkSheet.Cells[22, 6] = Zapros("123", "4");
                ObjWorkSheet.Cells[23, 3] = Zapros("124", "1"); ObjWorkSheet.Cells[23, 4] = Zapros("124", "2"); ObjWorkSheet.Cells[23, 5] = Zapros("124", "3"); ObjWorkSheet.Cells[23, 6] = Zapros("124", "4");

                ObjWorkSheet.Cells[24, 3] = Zapros("13", "1"); ObjWorkSheet.Cells[24, 4] = Zapros("13", "2"); ObjWorkSheet.Cells[24, 5] = Zapros("13", "3");

                ObjWorkSheet.Cells[25, 3] = Zapros("141", "1"); ObjWorkSheet.Cells[25, 4] = Zapros("142", "2"); ObjWorkSheet.Cells[25, 5] = Zapros("143", "3"); ObjWorkSheet.Cells[25, 6] = Zapros("144", "4"); ObjWorkSheet.Cells[25, 7] = Zapros("145", "5"); ObjWorkSheet.Cells[25, 8] = Zapros("146", "6"); ObjWorkSheet.Cells[25, 9] = Zapros("147", "7");

                ObjWorkSheet.Cells[26, 3] = Zapros("1513", "1"); ObjWorkSheet.Cells[26, 4] = Zapros("1513", "2"); ObjWorkSheet.Cells[26, 5] = Zapros("1513", "3"); ObjWorkSheet.Cells[26, 6] = Zapros("1513", "4"); ObjWorkSheet.Cells[26, 7] = Zapros("1513", "5"); ObjWorkSheet.Cells[26, 8] = Zapros("1513", "6"); ObjWorkSheet.Cells[26, 9] = Zapros("1513", "7"); ObjWorkSheet.Cells[26, 10] = Zapros("1513", "8"); ObjWorkSheet.Cells[26, 11] = Zapros("159", "9"); ObjWorkSheet.Cells[26, 12] = Zapros("1513", "10"); ObjWorkSheet.Cells[26, 13] = Zapros("1513", "11"); ObjWorkSheet.Cells[26, 13] = Zapros("1513", "12"); ObjWorkSheet.Cells[26, 15] = Zapros("1513", "13");
                ObjWorkSheet.Cells[27, 16] = Zapros("16", "1");
            //}


            //if (comboBox2.Text == "ФПМ")
            //{
            //    ObjWorkSheet.Cells[6, 5] = Zapros("1", "3");

            //    ObjWorkSheet.Cells[7, 4] = Zapros("2", "2"); ObjWorkSheet.Cells[7, 5] = Zapros("2", "3");  
            //    ObjWorkSheet.Cells[8, 4] = Zapros("3", "2"); ObjWorkSheet.Cells[8, 5] = Zapros("3", "3");

            //    ObjWorkSheet.Cells[9, 3] = Zapros("4", "1"); ObjWorkSheet.Cells[9, 6] = Zapros("4", "4");

            //    ObjWorkSheet.Cells[10, 16] = (dataGridView1.RowCount - 1).ToString();
            //    ObjWorkSheet.Cells[11, 16] = (dataGridView1.RowCount - 1).ToString();
            //    ObjWorkSheet.Cells[12, 16] = (dataGridView1.RowCount - 1).ToString();


            //    ObjWorkSheet.Cells[13, 3] = Zapros("8", "1"); ObjWorkSheet.Cells[13, 4] = Zapros("8", "2");
            //    ObjWorkSheet.Cells[14, 3] = Zapros("9", "1"); ObjWorkSheet.Cells[14, 4] = Zapros("9", "2");

            //    ObjWorkSheet.Cells[15, 3] = Zapros("10", "1");

            //    ObjWorkSheet.Cells[16, 6] = Zapros("111", "4"); ObjWorkSheet.Cells[16, 7] = Zapros("111", "5");
            //    ObjWorkSheet.Cells[17, 6] = Zapros("112", "4"); ObjWorkSheet.Cells[17, 7] = Zapros("112", "5");
            //    ObjWorkSheet.Cells[18, 6] = Zapros("113", "4"); ObjWorkSheet.Cells[18, 7] = Zapros("113", "5");
            //    ObjWorkSheet.Cells[19, 6] = Zapros("114", "4"); ObjWorkSheet.Cells[19, 7] = Zapros("114", "5");

            //    ObjWorkSheet.Cells[20, 5] = Zapros("121", "3"); ObjWorkSheet.Cells[20, 6] = Zapros("121", "4");
            //    ObjWorkSheet.Cells[21, 5] = Zapros("122", "3"); ObjWorkSheet.Cells[21, 6] = Zapros("122", "4");
            //    ObjWorkSheet.Cells[22, 5] = Zapros("123", "3"); ObjWorkSheet.Cells[22, 6] = Zapros("123", "4");
            //    ObjWorkSheet.Cells[23, 5] = Zapros("124", "3"); ObjWorkSheet.Cells[23, 6] = Zapros("124", "4");
            //    ObjWorkSheet.Cells[24, 3] = Zapros("13", "1");

            //    ObjWorkSheet.Cells[25, 3] = Zapros("141", "1"); ObjWorkSheet.Cells[25, 4] = Zapros("142", "2"); ObjWorkSheet.Cells[25, 5] = Zapros("143", "3"); ObjWorkSheet.Cells[25, 6] = Zapros("144", "4"); ObjWorkSheet.Cells[25, 7] = Zapros("145", "5"); ObjWorkSheet.Cells[25, 8] = Zapros("146", "6");

            //    ObjWorkSheet.Cells[26, 11] = Zapros("159", "9"); ObjWorkSheet.Cells[26, 15] =   Zapros("1513", "13");

            //}



            //if (comboBox2.Text == "Назначение ЕДВ и ЕЖВ" || comboBox2.Text == "Дополнительные расходы" )   
            //{
            //    ObjWorkSheet.Cells[6, 4] = Zapros("1", "3");

            //    ObjWorkSheet.Cells[7, 4] = Zapros("2", "2"); ObjWorkSheet.Cells[7, 5] = Zapros("2", "3");  
            //    ObjWorkSheet.Cells[8, 4] = Zapros("3", "2"); ObjWorkSheet.Cells[8, 5] = Zapros("3", "3");

            //    ObjWorkSheet.Cells[9, 3] = Zapros("4", "1"); ObjWorkSheet.Cells[9, 6] = Zapros("4", "4");

            //    ObjWorkSheet.Cells[10, 16] = (dataGridView1.RowCount - 1).ToString();
            //    ObjWorkSheet.Cells[11, 16] = (dataGridView1.RowCount - 1).ToString();
            //    ObjWorkSheet.Cells[12, 16] = (dataGridView1.RowCount - 1).ToString();

            //    ObjWorkSheet.Cells[13, 3] = Zapros("8", "1"); ObjWorkSheet.Cells[13, 4] = Zapros("8", "2");
            //    ObjWorkSheet.Cells[14, 3] = Zapros("9", "1"); ObjWorkSheet.Cells[14, 4] = Zapros("9", "2");

            //    ObjWorkSheet.Cells[15, 3] = Zapros("10", "1");

            //    ObjWorkSheet.Cells[16, 6] = Zapros("111", "4"); ObjWorkSheet.Cells[16, 7] = Zapros("111", "5");
            //    ObjWorkSheet.Cells[17, 6] = Zapros("112", "4"); ObjWorkSheet.Cells[17, 7] = Zapros("112", "5");
            //    ObjWorkSheet.Cells[18, 6] = Zapros("113", "4"); ObjWorkSheet.Cells[18, 7] = Zapros("113", "5");
            //    ObjWorkSheet.Cells[19, 6] = Zapros("114", "4"); ObjWorkSheet.Cells[19, 7] = Zapros("114", "5");

            //    ObjWorkSheet.Cells[20, 5] = Zapros("121", "3"); ObjWorkSheet.Cells[20, 6] = Zapros("121", "4");
            //    ObjWorkSheet.Cells[21, 5] = Zapros("122", "3"); ObjWorkSheet.Cells[21, 6] = Zapros("122", "4");
            //    ObjWorkSheet.Cells[22, 5] = Zapros("123", "3"); ObjWorkSheet.Cells[22, 6] = Zapros("123", "4");
            //    ObjWorkSheet.Cells[23, 5] = Zapros("124", "3"); ObjWorkSheet.Cells[23, 6] = Zapros("124", "4");
            //    ObjWorkSheet.Cells[24, 3] = Zapros("13", "1");

            //    ObjWorkSheet.Cells[25, 3] = Zapros("141", "1"); ObjWorkSheet.Cells[25, 4] = Zapros("142", "2"); ObjWorkSheet.Cells[25, 5] = Zapros("143", "3"); ObjWorkSheet.Cells[25, 6] = Zapros("144", "4"); ObjWorkSheet.Cells[25, 7] = Zapros("145", "5"); ObjWorkSheet.Cells[25, 8] = Zapros("146", "6");
               
            //    ObjWorkSheet.Cells[26, 11] = Zapros("159", "1");       ObjWorkSheet.Cells[26, 15] = Zapros("1513", "2");

            //}

            //if (comboBox2.Text == "Прием заявок на СКЛ")
            //{
            //    ObjWorkSheet.Cells[6, 5] = Zapros("1", "3");

            //    ObjWorkSheet.Cells[7, 4] = Zapros("2", "2"); ObjWorkSheet.Cells[7, 5] = Zapros("2", "3");
            //    ObjWorkSheet.Cells[8, 4] = Zapros("3", "2"); ObjWorkSheet.Cells[8, 5] = Zapros("3", "3");

            //    ObjWorkSheet.Cells[9, 3] = Zapros("4", "1");  

            //    ObjWorkSheet.Cells[10, 16] = (dataGridView1.RowCount - 1).ToString();
            //    ObjWorkSheet.Cells[11, 16] = (dataGridView1.RowCount - 1).ToString();
            //    ObjWorkSheet.Cells[12, 16] = (dataGridView1.RowCount - 1).ToString();


            //    ObjWorkSheet.Cells[13, 3] = Zapros("8", "1"); ObjWorkSheet.Cells[13, 4] = Zapros("8", "2");
            //    ObjWorkSheet.Cells[14, 3] = Zapros("9", "1"); ObjWorkSheet.Cells[14, 4] = Zapros("9", "2");

            //    ObjWorkSheet.Cells[15, 3] = Zapros("10", "1"); ObjWorkSheet.Cells[15, 4] = Zapros("10", "2");

            //    ObjWorkSheet.Cells[16, 6] = Zapros("111", "4"); ObjWorkSheet.Cells[16, 7] = Zapros("111", "5");
            //    ObjWorkSheet.Cells[17, 6] = Zapros("112", "4"); ObjWorkSheet.Cells[17, 7] = Zapros("112", "5");
            //    ObjWorkSheet.Cells[18, 6] = Zapros("113", "4"); ObjWorkSheet.Cells[18, 7] = Zapros("113", "5");
            //    ObjWorkSheet.Cells[19, 6] = Zapros("114", "4"); ObjWorkSheet.Cells[19, 7] = Zapros("114", "5");

            //    ObjWorkSheet.Cells[20, 5] = Zapros("121", "3"); ObjWorkSheet.Cells[20, 6] = Zapros("121", "4");
            //    ObjWorkSheet.Cells[21, 5] = Zapros("122", "3"); ObjWorkSheet.Cells[21, 6] = Zapros("122", "4");
            //    ObjWorkSheet.Cells[22, 5] = Zapros("123", "3"); ObjWorkSheet.Cells[22, 6] = Zapros("123", "4");
            //    ObjWorkSheet.Cells[23, 5] = Zapros("124", "3"); ObjWorkSheet.Cells[23, 6] = Zapros("124", "4");
            //    ObjWorkSheet.Cells[24, 3] = Zapros("13", "1");

            //    ObjWorkSheet.Cells[25, 3] = Zapros("141", "1"); ObjWorkSheet.Cells[25, 4] = Zapros("142", "2"); ObjWorkSheet.Cells[25, 5] = Zapros("143", "3"); ObjWorkSheet.Cells[25, 6] = Zapros("144", "4"); ObjWorkSheet.Cells[25, 7] = Zapros("145", "5"); ObjWorkSheet.Cells[25, 8] = Zapros("146", "6");

            //    ObjWorkSheet.Cells[26, 11] = Zapros("159", "9"); ObjWorkSheet.Cells[26, 15] = Zapros("1513", "13");

            //}

            //if (comboBox2.Text == "Обеспечение инвалидов техническими средствами реабилитации и (или) услугами")
            //{
            //    ObjWorkSheet.Cells[6, 3] = Zapros("1", "1");

            //    ObjWorkSheet.Cells[7, 3] = Zapros("2", "1"); ObjWorkSheet.Cells[7, 4] = Zapros("2", "2"); ObjWorkSheet.Cells[7, 5] = Zapros("2", "3");
            //    ObjWorkSheet.Cells[8, 3] = Zapros("3", "1"); ObjWorkSheet.Cells[8, 4] = Zapros("3", "2"); ObjWorkSheet.Cells[8, 5] = Zapros("3", "3");

            //    ObjWorkSheet.Cells[9, 3] = Zapros("4", "1"); ObjWorkSheet.Cells[9, 5] = Zapros("4", "3"); ObjWorkSheet.Cells[9, 6] = Zapros("4", "4");

            //    ObjWorkSheet.Cells[10, 16] = (dataGridView1.RowCount - 1).ToString();
            //    ObjWorkSheet.Cells[11, 16] = (dataGridView1.RowCount - 1).ToString();
            //    ObjWorkSheet.Cells[12, 16] = (dataGridView1.RowCount - 1).ToString();


            //    ObjWorkSheet.Cells[13, 3] = Zapros("8", "1"); ObjWorkSheet.Cells[13, 4] = Zapros("8", "2"); ObjWorkSheet.Cells[13, 7] = Zapros("8", "5");
            //    ObjWorkSheet.Cells[14, 3] = Zapros("9", "1"); ObjWorkSheet.Cells[14, 4] = Zapros("9", "2");

            //    ObjWorkSheet.Cells[15, 3] = Zapros("10", "1"); ObjWorkSheet.Cells[15, 4] = Zapros("10", "2");

            //    ObjWorkSheet.Cells[16, 6] = Zapros("111", "4"); ObjWorkSheet.Cells[16, 7] = Zapros("111", "5");
            //    ObjWorkSheet.Cells[17, 6] = Zapros("112", "4"); ObjWorkSheet.Cells[17, 7] = Zapros("112", "5");
            //    ObjWorkSheet.Cells[18, 6] = Zapros("113", "4"); ObjWorkSheet.Cells[18, 7] = Zapros("113", "5");
            //    ObjWorkSheet.Cells[19, 6] = Zapros("114", "4"); ObjWorkSheet.Cells[19, 7] = Zapros("114", "5");

            //    ObjWorkSheet.Cells[20, 5] = Zapros("121", "3"); ObjWorkSheet.Cells[20, 6] = Zapros("121", "4");
            //    ObjWorkSheet.Cells[21, 5] = Zapros("122", "3"); ObjWorkSheet.Cells[21, 6] = Zapros("122", "4");
            //    ObjWorkSheet.Cells[22, 5] = Zapros("123", "3"); ObjWorkSheet.Cells[22, 6] = Zapros("123", "4");
            //    ObjWorkSheet.Cells[23, 5] = Zapros("124", "3"); ObjWorkSheet.Cells[23, 6] = Zapros("124", "4");
            //    ObjWorkSheet.Cells[24, 3] = Zapros("13", "1");

            //    ObjWorkSheet.Cells[25, 3] = Zapros("141", "1"); ObjWorkSheet.Cells[25, 4] = Zapros("142", "2"); ObjWorkSheet.Cells[25, 5] = Zapros("143", "3"); ObjWorkSheet.Cells[25, 6] = Zapros("144", "4"); ObjWorkSheet.Cells[25, 7] = Zapros("145", "5"); ObjWorkSheet.Cells[25, 8] = Zapros("146", "6");

            //    ObjWorkSheet.Cells[26, 11] = Zapros("159", "9"); ObjWorkSheet.Cells[26, 15] = Zapros("1513", "13");

            //}

            ObjExcel.Visible = true;//В итоге, делаем созданную эксельку видимой 
            ObjExcel.UserControl = true; // и доступной!
     
        }//--------------------------------------------------------------------------------------------------------------

        string Zapros(string NomVopr, string NomOt)//--------------  запрос ----------------------------------------
        {
            DataTable table = new DataTable();
             
            string databaseName = Directory.GetCurrentDirectory() + "\\Anketa.db";
            SQLiteConnection connection = new SQLiteConnection(string.Format("Data Source={0};", databaseName));
            connection.Open();
            
            SQLiteCommand command = new SQLiteCommand("SELECT SUM(ot"+NomOt+") FROM 'registr' WHERE  Dat >=  '" 
                + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") 
                + "' AND Dat <=  '" 
                + dateTimePicker2.Value.Date.ToString("yyyy-MM-dd")
                + "' AND Podr = '" 
                + comboBox1.Text.ToString()
                + "' AND Uslug = '" 
                + comboBox2.Text.ToString()
                + "' AND  NomVopr= '" + NomVopr  + "';", connection);
          
            SQLiteDataReader reader = command.ExecuteReader();
            table.Load(reader);
            connection.Close();
             
             return   table.Rows[0][0].ToString()    ;
        }//---------------------------------------------------------------------------------------------------------------------



    }
}
