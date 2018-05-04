using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        SqlConnection database;
        public Form1()
        {
            InitializeComponent();
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            database = new SqlConnection(@"Data Source=DESKTOP-jtdifdm;Initial Catalog=ws;Integrated Security=True");
            database.Open();
            //Опустошаем поток для чтения
            SqlDataReader sqlReader = null;
            //Прописываем команду
            SqlCommand sqlcom = new SqlCommand("SELECT dbo.Runner.*FROM dbo.Runner", database);
            //Выводим данные в listbox
            
            string fileName = "D:\\provest visual\\WindowsFormsApp1\\WindowsFormsApp1\\bin\\Debug\\Лист Microsoft Excel.xlsx"; //имя Excel файла  
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWb = xlApp.Workbooks.Open(fileName); //открываем Excel файл
            Excel.Worksheet xlSht = xlWb.Sheets[1]; //первый лист в файле
            int iLastRow = xlSht.Cells[xlSht.Rows.Count, "A"].End[Excel.XlDirection.xlUp].Row;  //последняя заполненная строка в столбце А
            try
            {
                sqlReader = sqlcom.ExecuteReader();
                while (sqlReader.Read())
                {
                    iLastRow++;
                    label1.Text= sqlReader["Email"].ToString();
                    xlSht.Cells[iLastRow, "A"].Value=sqlReader["Email"].ToString() ;
                }
            }
            catch { }


           /* for (int i = 1; i < 51; i++)
            {
                iLastRow++;
                 = i.ToString();
            }*/
            //xlApp.Visible = true;
            xlWb.Close(true); //закрыть и сохранить книгу
            xlApp.Quit();
            MessageBox.Show("Файл успешно сохранён!");
        }
    }
}
