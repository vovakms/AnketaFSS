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

using System.DirectoryServices ;
 

namespace АнкетаОДеятельностиФСС
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            if (checkBox3.Checked == true) { textBox5.Text = checkBox3.Text.ToString();   }
            if (checkBox4.Checked == true) { textBox5.Text = checkBox4.Text.ToString();   }
            if (checkBox5.Checked == true) { textBox5.Text = comboBox2.Text.ToString();   }

        }

        private void Form2_Paint(object sender, PaintEventArgs e)//-----------------------------------------------------------
        {
            DataTable table = new DataTable();

            string databaseName = Directory.GetCurrentDirectory() + "\\Anketa.db";
            SQLiteConnection connection = new SQLiteConnection(string.Format("Data Source={0};", databaseName));
            connection.Open();
            SQLiteCommand command = new SQLiteCommand("SELECT MAX(NomAnket) FROM 'Anketa' WHERE Podrazdelenie = '" + comboBox1.Text.ToString() + "' ;", connection);
            SQLiteDataReader reader = command.ExecuteReader();

            table.Load(reader);
            connection.Close();

            textBox1.Text = (Convert.ToInt32(table.Rows[0][0].ToString()) + 1).ToString();


            string userNameWin, compName, compIP, myHost;
            myHost = System.Net.Dns.GetHostName();// имя хоста
            compIP = System.Net.Dns.GetHostEntry(myHost).AddressList[0].ToString();// IP по имени хоста, выдает список, можно обойти в цикле весь, здесь берется первый адрес
            userNameWin = System.Environment.UserName;
            compName = System.Environment.MachineName;

            textBox3.Text = userNameWin;


            if (userNameWin == "РябухинаВВ" || userNameWin == "ЕвстафьеваГС" )
            {
                comboBox1.Text = "гр_СПР";
                comboBox2.Text = "ФПМ";
            }

            if (userNameWin == "ЦареваНА" || userNameWin == "РожковаНВ" || userNameWin == "СалийЛВ")
            {
                comboBox1.Text = "специалисты_ОЛКГ";
                comboBox2.Text = "Прием заявок на СКЛ";
            }



        }//---------------------------------------------------------------------------------------------------------------------------------------
         
        private void label17_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)//--------------------------------------нажали кнопку СОХРАНИТЬ -----------------------------
        {
            string ot1 = "", ot2 = "", ot3 = "", ot4 = "", ot5 = "", ot6 = "", ot7 = "", ot8 = "", ot9 = "", ot10 = "", ot111 = "", ot112 = "", ot113 = "", ot114 = "", ot121 = "", ot122 = "", ot123 = "", ot124 = "", ot13 = "", ot14 = "", ot15 = "", ot16="";

            if (checkBox6.Checked == false && checkBox7.Checked == false && checkBox8.Checked == false && checkBox9.Checked == false  )
            {
             MessageBox.Show("Вы не ответили на второй вопрос", "Второй вопрос", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
             return  ;
            }
            if (checkBox10.Checked == false && checkBox11.Checked == false && checkBox12.Checked == false && checkBox13.Checked == false)
            {
                MessageBox.Show("Вы не ответили на третий вопрос", "Третий вопрос", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                return;
            }
            if (checkBox14.Checked == false && checkBox15.Checked == false && checkBox16.Checked == false && checkBox17.Checked == false)
            {
                MessageBox.Show("Вы не ответили на четвертый вопрос", "Четвертый вопрос", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                return;
            }
            if (checkBox18.Checked == false && checkBox19.Checked == false && checkBox20.Checked == false && checkBox21.Checked == false && checkBox86.Checked == false)
            {
                MessageBox.Show("Вы не ответили на восьмой вопрос", "Восьмой вопрос", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                return;
            }

            if (checkBox26.Checked==false && checkBox27.Checked==false && checkBox28.Checked==false && checkBox29.Checked==false && checkBox30.Checked==false)
            {
                MessageBox.Show("Вы не ответили на вопрос 11.1", "Вопрос 11.1", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                return;
            }

            if (checkBox31.Checked == false && checkBox32.Checked == false && checkBox33.Checked == false && checkBox34.Checked == false && checkBox35.Checked == false)
            {
                MessageBox.Show("Вы не ответили на вопрос 11.2", "Вопрос 11.2", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                return;
            }

            if (checkBox36.Checked == false && checkBox37.Checked == false && checkBox38.Checked == false && checkBox39.Checked == false && checkBox40.Checked == false)
            {
                MessageBox.Show("Вы не ответили на вопрос 11.3", "Вопрос 11.3", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                return;
            }

            if (checkBox41.Checked == false && checkBox42.Checked == false && checkBox43.Checked == false && checkBox44.Checked == false && checkBox45.Checked == false)
            {
                MessageBox.Show("Вы не ответили на вопрос 11.4", "Вопрос 11.4", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                return;
            }

            if (checkBox46.Checked == false && checkBox47.Checked == false && checkBox48.Checked == false   && checkBox49.Checked == false)
            {
                MessageBox.Show("Вы не ответили на вопрос 12.1", "Вопрос 12.1", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                return;
            }

            if (checkBox50.Checked == false && checkBox51.Checked == false && checkBox52.Checked == false && checkBox53.Checked == false)
            {
                MessageBox.Show("Вы не ответили на вопрос 12.2", "Вопрос 12.2", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                return;
            }

            if (checkBox54.Checked == false && checkBox55.Checked == false && checkBox56.Checked == false && checkBox57.Checked == false)
            {
                MessageBox.Show("Вы не ответили на вопрос 12.3", "Вопрос 12.3", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                return;
            }

            if (checkBox58.Checked == false && checkBox59.Checked == false && checkBox60.Checked == false && checkBox61.Checked == false)
            {
                MessageBox.Show("Вы не ответили на вопрос 12.4", "Вопрос 12.4", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                return;
            }

            if (checkBox65.Checked == false && checkBox66.Checked == false && checkBox67.Checked == false && checkBox68.Checked == false && checkBox69.Checked == false && checkBox70.Checked == false && checkBox71.Checked == false )
            {
                MessageBox.Show("Вы не ответили на вопрос 14", "Вопрос 14", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                return;
            }

            string DatbaseName = Directory.GetCurrentDirectory() + "\\Anketa.db";
            SQLiteConnection connection = new SQLiteConnection(string.Format("Data Source={0};", DatbaseName));
            connection.Open();
            SQLiteCommand command ;

            if (checkBox3.Checked == true) { textBox5.Text = checkBox3.Text.ToString(); ot1 = checkBox3.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot1, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + checkBox3.Text.ToString() + "', '1', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox4.Checked == true) { textBox5.Text = checkBox4.Text.ToString(); ot1 = checkBox4.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot2, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + checkBox4.Text.ToString() + "', '1', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox5.Checked == true) { textBox5.Text = comboBox2.Text.ToString(); ot1 = comboBox2.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot3, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + comboBox2.Text.ToString() + "', '1', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }

            if (checkBox6.Checked == true) { ot2 = checkBox6.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot1, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '2', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox7.Checked == true) { ot2 = checkBox7.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot2, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '2', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox8.Checked == true) { ot2 = checkBox8.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot3, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '2', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox9.Checked == true) { ot2 = checkBox9.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot4, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '2', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }

            if (checkBox10.Checked == true) { ot3 = checkBox10.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot1, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '3', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox11.Checked == true) { ot3 = checkBox11.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot2, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '3', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox12.Checked == true) { ot3 = checkBox12.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot3, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '3', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox13.Checked == true) { ot3 = checkBox13.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot4, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '3', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }

            if (checkBox14.Checked == true) { ot4 = checkBox14.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot1, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '4', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox15.Checked == true) { ot4 = checkBox15.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot2, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '4', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox16.Checked == true) { ot4 = checkBox16.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot3, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '4', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox17.Checked == true) { ot4 = checkBox17.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot4, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '4', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }

            if (numericUpDown2.Value > 0) { ot5 = numericUpDown2.Value.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot14, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '5', '" + numericUpDown2.Value.ToString() + "' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (numericUpDown3.Value > 0) { ot6 = numericUpDown3.Value.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot15, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '6', '" + numericUpDown3.Value.ToString() + "' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (numericUpDown4.Value > 0) { ot7 = numericUpDown4.Value.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot16, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '7', '" + numericUpDown4.Value.ToString() + "' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }


            if (checkBox18.Checked == true) { ot8 = checkBox18.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot1, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '8', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox19.Checked == true) { ot8 = checkBox19.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot2, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '8', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox20.Checked == true) { ot8 = checkBox20.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot3, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '8', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox21.Checked == true) { ot8 = checkBox21.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot4, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '8', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox86.Checked == true) { ot8 = checkBox86.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot5, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '8', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }

            if (checkBox22.Checked == true) { ot9 = checkBox22.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot1, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '9', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox23.Checked == true) { ot9 = checkBox23.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot2, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '9', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }

            if (checkBox24.Checked == true) { ot10 = checkBox24.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot1, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '10', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox25.Checked == true) { ot10 = checkBox25.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot2, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '10', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }

            if (checkBox26.Checked == true) { ot111 = "Очень плохо"; command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot1, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '111', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox27.Checked == true) { ot111 = "Плохо"; command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot2, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '111', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox28.Checked == true) { ot111 = "Средне"; command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot3, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '111', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox29.Checked == true) { ot111 = "Хорошо"; command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot4, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '111', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox30.Checked == true) { ot111 = "Отлично"; command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot5, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '111', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }

            if (checkBox31.Checked == true) { ot112 = "Очень плохо"; command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot1, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '112', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox32.Checked == true) { ot112 = "Плохо"; command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot2, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '112', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox33.Checked == true) { ot112 = "Средне"; command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot3, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '112', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox34.Checked == true) { ot112 = "Хорошо"; command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot4, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '112', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox35.Checked == true) { ot112 = "Отлично"; command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot5, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '112', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }

            if (checkBox36.Checked == true) { ot113 = "Очень плохо"; command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot1, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '113', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox37.Checked == true) { ot113 = "Плохо"; command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot2, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '113', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox38.Checked == true) { ot113 = "Средне"; command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot3, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '113', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox39.Checked == true) { ot113 = "Отлично"; command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot4, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '113', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox40.Checked == true) { ot113 = "Отлично"; command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot5, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '113', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }

            if (checkBox41.Checked == true) { ot114 = "Очень плохо"; command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot1, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '114', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox42.Checked == true) { ot114 = "Плохо"; command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot2, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '114', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox43.Checked == true) { ot114 = "Средне"; command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot3, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '114', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox44.Checked == true) { ot114 = "Отлично"; command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot4, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '114', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox45.Checked == true) { ot114 = "Отлично"; command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot5, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '114', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }

            if (checkBox46.Checked == true) { ot121 = "Очень плохо"; command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot1, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '121', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox47.Checked == true) { ot121 = "Плохо"; command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot2, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '121', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox48.Checked == true) { ot121 = "Средне"; command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot3, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '121', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox49.Checked == true) { ot121 = "Хорошо"; command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot4, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '121', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }

            if (checkBox50.Checked == true) { ot122 = "Очень плохо"; command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot1, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '122', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox51.Checked == true) { ot122 = "Плохо"; command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot2, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '122', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox52.Checked == true) { ot122 = "Средне"; command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot3, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '122', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox53.Checked == true) { ot122 = "Хорошо"; command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot4, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '122', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }

            if (checkBox54.Checked == true) { ot123 = "Очень плохо"; command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot1, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '123', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox55.Checked == true) { ot123 = "Плохо"; command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot2, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '123', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox56.Checked == true) { ot123 = "Средне"; command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot3, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '123', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox57.Checked == true) { ot123 = "Хорошо"; command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot4, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '123', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }

            if (checkBox58.Checked == true) { ot124 = "Очень плохо"; command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot1, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '124', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox59.Checked == true) { ot124 = "Плохо"; command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot2, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '124', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox60.Checked == true) { ot124 = "Средне"; command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot3, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '124', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox61.Checked == true) { ot124 = "Хорошо"; command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot4, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '124', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }

            if (checkBox62.Checked == true) { ot13 = checkBox62.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat', 'Podr', 'Uslug', 'NomVopr', 'ot1', User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '13', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox63.Checked == true) { ot13 = checkBox63.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', 'ot2', User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '13', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox64.Checked == true) { ot13 = checkBox64.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', 'ot3', User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '13', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }

            if (checkBox65.Checked == true) { ot14 = checkBox65.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot1, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '141', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox66.Checked == true) { ot14 = checkBox66.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot2, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '142', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox67.Checked == true) { ot14 = checkBox67.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot3, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '143', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox68.Checked == true) { ot14 = checkBox68.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot4, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '144', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox69.Checked == true) { ot14 = checkBox69.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot5, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '145', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox70.Checked == true) { ot14 = checkBox70.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot6, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '146', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox71.Checked == true) { ot14 = textBox2.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot7, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '147', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }

            if (checkBox76.Checked == true) { ot15 = checkBox76.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot1, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '151', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox77.Checked == true) { ot15 = checkBox77.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot2, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '152', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox78.Checked == true) { ot15 = checkBox78.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot3, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '153', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox73.Checked == true) { ot15 = checkBox73.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot4, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '154', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox74.Checked == true) { ot15 = checkBox74.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot5, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '155', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox75.Checked == true) { ot15 = checkBox75.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot6, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '156', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox72.Checked == true) { ot15 = checkBox72.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot7, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '157', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox82.Checked == true) { ot15 = checkBox82.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot8, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '158', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox83.Checked == true) { ot15 = checkBox83.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot9, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '159', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox84.Checked == true) { ot15 = checkBox84.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot10, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '1510', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox79.Checked == true) { ot15 = checkBox79.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot11, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '1511', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox80.Checked == true) { ot15 = textBox4.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot12, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '1512', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }
            if (checkBox81.Checked == true) { ot15 = checkBox81.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot13, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '1513', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }

            if (comboBox3.Text.ToString() != "") { ot16 = comboBox3.Text.ToString(); command = new SQLiteCommand("INSERT INTO 'registr' ('NomAnket', 'Dat','Podr','Uslug','NomVopr', ot1, User ) VALUES ('" + textBox1.Text.ToString() + "' ,'" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "','" + comboBox1.Text.ToString() + "','" + textBox5.Text.ToString() + "', '16', '1' , '" + textBox3.Text.ToString() + "');", connection); command.ExecuteNonQuery(); }


            connection.Close();
 
            connection.Open();
            command = new SQLiteCommand("INSERT INTO 'Anketa' ('Dat','NomAnket', 'Pol','Vozrast','Podrazdelenie','q1',q2,q3,q4,q5,q6,q7,q8,q9,q10,q111,q112,q113,q114,q121,q122,q123,q124,q13,q14,q15,q16,User  ) VALUES ('"
                                                         + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "', '"
                                                         + textBox1.Text.ToString() + "', '"
                                                         + (checkBox1.Checked == true ? "Муж" : "Жен") + "' ,'"
                                                         + numericUpDown1.Value.ToString() + "' ,'" 
                                                         + comboBox1.Text.ToString() + "' ,'"
                                                        
                                                         + ot1   + "' ,'"
                                                         + ot2   + "' ,'"
                                                         + ot3   + "' ,'"
                                                         + ot4   + "' ,'"
                                                         + ot5   + "' ,'"
                                                         + ot6   + "' ,'"
                                                         + ot7   + "' ,'"
                                                         + ot8   + "' ,'"
                                                         + ot9   + "' ,'"
                                                         + ot10  + "' ,'"
                                                         + ot111 + "' ,'"
                                                         + ot112 + "' ,'"
                                                         + ot113 + "' ,'"
                                                         + ot114 + "' ,'"
                                                         + ot121 + "' ,'"
                                                         + ot122 + "' ,'"
                                                         + ot123 + "' ,'"
                                                         + ot124 + "' ,'"
                                                         + ot13  + "' ,'"
                                                         + ot14  + "' ,'"
                                                         + ot15  + "' ,'"
                                                         + comboBox3.Text.ToString() + "' ,'"
                                                         + textBox3.Text.ToString() + "' );", connection);
            command.ExecuteNonQuery();
            connection.Close();
             
            Close();



        }//----------------------------------------------------------------------------------------------------------------------------------------

        private void checkBox1_CheckedChanged(object sender, EventArgs e) // ----  выбираем пол  Муж
        {
            if (checkBox2.Checked == true)
                checkBox2.Checked = false;
             
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e) // ----  выбираем пол  Жен
        { 
            if (checkBox1.Checked == true)
                checkBox1.Checked = false;
            
        }

        private void checkBox14_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox14.Checked == true)
            {
                checkBox15.Checked = false;
                checkBox16.Checked = false;
                checkBox17.Checked = false;
            }
        }

        

       

        private void checkBox36_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)// ---------  выбор подразделения ----------------------------
        {
            if (comboBox1.Text.ToString() == "гр_СПР")
            {
                comboBox2.Text = "ФПМ";
                comboBox2.Items.Clear();
                comboBox2.Items.Add("Назначение ЕДВ и ЕЖВ");
                comboBox2.Items.Add("Скидки");
                comboBox2.Items.Add("Дополнительные расходы");
                comboBox2.Items.Add("Подтверждение ОКВЭД");
                comboBox2.Items.Add("ФПМ");
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
                

            }




            DataTable table = new DataTable();

            string databaseName = Directory.GetCurrentDirectory() + "\\Anketa.db";
            SQLiteConnection connection = new SQLiteConnection(string.Format("Data Source={0};", databaseName));
            connection.Open();
            SQLiteCommand command = new SQLiteCommand("SELECT MAX(NomAnket) FROM 'Anketa' WHERE Podrazdelenie = '" + comboBox1.Text.ToString() + "' ;", connection);
            SQLiteDataReader reader = command.ExecuteReader();

            table.Load(reader);
            connection.Close();

            textBox1.Text = (Convert.ToInt32(table.Rows[0][0].ToString()) + 1).ToString();
  

        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked == true)
            {
                checkBox4.Checked = false;
                checkBox5.Checked = false;
            }

             textBox5.Text = checkBox3.Text.ToString();  
             
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked == true)
            {
                checkBox3.Checked = false;
                checkBox5.Checked = false;
            }

            
             textBox5.Text = checkBox4.Text.ToString();  
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked == true)
            {
                checkBox4.Checked = false;
                checkBox3.Checked = false;
            }

             textBox5.Text = comboBox2.Text.ToString();  

        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox6.Checked == true)
            {
                checkBox7.Checked = false;
                checkBox8.Checked = false;
                checkBox9.Checked = false;
            }
        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox7.Checked == true)
            {
                checkBox6.Checked = false;
                checkBox8.Checked = false;
                checkBox9.Checked = false;
            }
        }

        private void checkBox8_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox8.Checked == true)
            {
                checkBox7.Checked = false;
                checkBox6.Checked = false;
                checkBox9.Checked = false;
            }
        }

        private void checkBox9_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox9.Checked == true)
            {
                checkBox7.Checked = false;
                checkBox8.Checked = false;
                checkBox6.Checked = false;
            }
        }

        private void checkBox10_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox10.Checked == true)
            {
                checkBox11.Checked = false;
                checkBox12.Checked = false;
                checkBox13.Checked = false;
            }
        }

        private void checkBox11_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox11.Checked == true)
            {
                checkBox10.Checked = false;
                checkBox12.Checked = false;
                checkBox13.Checked = false;
            }
        }

        private void checkBox12_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox12.Checked == true)
            {
                checkBox11.Checked = false;
                checkBox10.Checked = false;
                checkBox13.Checked = false;
            }
        }

        private void checkBox13_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox13.Checked == true)
            {
                checkBox11.Checked = false;
                checkBox12.Checked = false;
                checkBox10.Checked = false;
            }
        }

        private void checkBox15_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox15.Checked == true)
            {
                checkBox14.Checked = false;
                checkBox16.Checked = false;
                checkBox17.Checked = false;
            }
        }

        private void checkBox16_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox16.Checked == true)
            {
                checkBox15.Checked = false;
                checkBox14.Checked = false;
                checkBox17.Checked = false;
            }
        }

        private void checkBox17_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox17.Checked == true)
            {
                checkBox15.Checked = false;
                checkBox16.Checked = false;
                checkBox14.Checked = false;
            }
        }

        private void checkBox18_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox18.Checked == true)
            {
                checkBox19.Checked = false;
                checkBox20.Checked = false;
                checkBox21.Checked = false;
                checkBox86.Checked = false;
            }
        }

        private void checkBox19_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox19.Checked == true)
            {
                checkBox18.Checked = false;
                checkBox20.Checked = false;
                checkBox21.Checked = false;
                checkBox86.Checked = false;
            }
        }

        private void checkBox20_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox20.Checked == true)
            {
                checkBox19.Checked = false;
                checkBox18.Checked = false;
                checkBox21.Checked = false;
                checkBox86.Checked = false;
            }
        }

        private void checkBox21_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox21.Checked == true)
            {
                checkBox19.Checked = false;
                checkBox20.Checked = false;
                checkBox18.Checked = false;
                checkBox86.Checked = false;
            }
        }

        private void checkBox86_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox86.Checked == true)
            {
                checkBox19.Checked = false;
                checkBox20.Checked = false;
                checkBox21.Checked = false;
                checkBox18.Checked = false;
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            checkBox5.Checked = true;
            textBox5.Text = comboBox2.Text.ToString(); 
        }

        private void checkBox76_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox76.Checked == true )   checkBox81.Checked = false; 
        }

        private void checkBox77_CheckedChanged(object sender, EventArgs e)
        {
            if (  checkBox77.Checked == true )  checkBox81.Checked = false; 
        }

        private void checkBox78_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox78.Checked == true) checkBox81.Checked = false; 
        }

        private void checkBox73_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox73.Checked == true) checkBox81.Checked = false; 
        }

        private void checkBox74_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox74.Checked == true) checkBox81.Checked = false; 
        }

        private void checkBox75_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox75.Checked == true) checkBox81.Checked = false; 
        }

        private void checkBox72_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox72.Checked == true) checkBox81.Checked = false; 
        }

        private void checkBox82_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox82.Checked == true) checkBox81.Checked = false; 
        }

        private void checkBox83_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox83.Checked == true) checkBox81.Checked = false; 
        }

        private void checkBox84_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox84.Checked == true) checkBox81.Checked = false; 
        }

        private void checkBox79_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox79.Checked == true) checkBox81.Checked = false; 
        }

        private void checkBox80_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox80.Checked == true) checkBox81.Checked = false; 
        }

        private void checkBox81_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox81.Checked == true)
            {
                checkBox76.Checked = false;
                checkBox77.Checked = false;
                checkBox78.Checked = false;
                checkBox73.Checked = false;
                checkBox74.Checked = false;
                checkBox75.Checked = false;
                checkBox72.Checked = false;
                checkBox82.Checked = false;
                checkBox83.Checked = false;
                checkBox84.Checked = false;
                checkBox79.Checked = false;
                checkBox80.Checked = false;
             }
        }
 



    }
}

























