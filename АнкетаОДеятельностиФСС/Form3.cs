using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.IO;
using System.Data.SQLite;

namespace АнкетаОДеятельностиФСС
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string path = Directory.GetCurrentDirectory();

            textBox1.Text = path.ToString();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "SQLite files(*.db)|*.db" ;
            saveFileDialog1.InitialDirectory = Directory.GetCurrentDirectory();
            //saveFileDialog1.DefaultExt = "db";
            saveFileDialog1.ShowDialog();

            if (saveFileDialog1.FileName != "")
            {
                textBox2.Text = saveFileDialog1.FileName.ToString()  ;
                string databaseName = saveFileDialog1.FileName;
                SQLiteConnection.CreateFile(databaseName);

                //SQLiteConnection connection = new SQLiteConnection(string.Format("Data Source={0};", databaseName));
                //SQLiteCommand command = new SQLiteCommand("CREATE TABLE example (id INTEGER PRIMARY KEY, value TEXT);", connection);
                //connection.Open();
                //command.ExecuteNonQuery();
                //connection.Close();
                 

                label2.Text = File.Exists(databaseName) ? "База данных создана" : "Возникла ошиюка при создании базы данных" ;
    

            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            textBox3.Text = dateTimePicker1.Value.Date.ToString("yyyy-MM-dd");
        }





        //--------------------------------------------------------------
    }
}
