using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//using Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using ExcelObj = Microsoft.Office.Interop.Excel;

namespace FraseV1
{
    public partial class Form1 : Form
    {
       
        public Form1()
        {
            InitializeComponent();
        }
        private void DataGridViewChecked(DataGridView dgv)
        {
            try
            {
                int n = dgv.RowCount;
                for (int i = n-2; i >= 0; i--)
                {
                    if (Convert.ToBoolean(dgv.Rows[i].Cells[0].Value) == false)
                    {
                        dgv.Rows.RemoveAt(i);
                        //dataGridView1.Refresh();
                    }
                }
                dgv.Refresh();
            }
            catch(System.InvalidOperationException ex)
            {
                MessageBox.Show("Возникла ошибка: " + ex.Message);
            }
            catch (System.ArgumentOutOfRangeException ex)
            {
                MessageBox.Show("Возникла ошибка: " + ex.Message);
            }
        }
        private static void OpenTable(OpenFileDialog ofd, DataGridView dataGridView1, int n)
        {
            ExcelObj.Application app = new ExcelObj.Application();
            ExcelObj.Workbook workbook;//книги
            ExcelObj.Worksheet NwSheet;//страницы
            ExcelObj.Range ShtRange;//ячейки
            DataTable dt = new DataTable();//таблица, куда заносятся данные

                //открытие файла
                /*
                workbook = app.Workbooks.Open(ofd.FileName, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing);
                */
            workbook = app.Workbooks.Open(ofd.FileName);



            //Устанавливаем номер листа из котрого будут извлекаться данные
            //Листы нумеруются от 1
            NwSheet = (ExcelObj.Worksheet)workbook.Sheets.get_Item(n);
                ShtRange = NwSheet.UsedRange;//все ячейки, содержащие значения на данный момент
                //имя колонок
                for (int Cnum = 1; Cnum <= ShtRange.Columns.Count; Cnum++) 
                    dt.Columns.Add(new DataColumn((ShtRange.Cells[1, Cnum] as ExcelObj.Range).Value2.ToString()));
                dt.AcceptChanges();


                string columnNames;
            try
            {
                columnNames = dt.Columns[0].ColumnName;
            }
            catch (Exception ex)
            {

                MessageBox.Show("Ошибка: " + ex.Message);
            }
                    

                /*
                for (int Rnum = 2; Rnum <= ShtRange.Rows.Count; Rnum++)
                {
                    DataRow dr = dt.NewRow();
                    for (int Cnum = 1; Cnum <= ShtRange.Columns.Count; Cnum++)
                        if ((ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).Value2 != null)
                            dr[Cnum - 1] = (ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).Value2.ToString();
                    dt.Rows.Add(dr);
                    dt.AcceptChanges();
                }*/

                dataGridView1.DataSource = dt;            
                app.Quit();
            
        }
        public void Save()
        {
            FileStream fs = new FileStream("база.txt", FileMode.Append);
            StreamWriter streamWriter = new StreamWriter(fs, Encoding.Unicode);
            try
            {
                streamWriter.WriteLine(dataGridView1.Columns[1].HeaderCell.Value);
                for (int j = 0; j < dataGridView1.Rows.Count; j++)
                        streamWriter.WriteLine(dataGridView1.Rows[j].Cells[1].Value);
                streamWriter.Close();
                fs.Close();

                MessageBox.Show("Файл успешно сохранен");
            }
            catch
            {
                MessageBox.Show("Ошибка при сохранении файла!");
            }
        }
        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "*.xls;*.xlsx";
            ofd.Filter = "Excel Sheet(*.xlsx)|*.xlsx";
            ofd.Title = "Выберите документ для загрузки данных";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                OpenTable(ofd, dataGridView1, 1);
                //   OpenTable(ofd, dataGridView2, 2);
            }
        }

        private void сохранитьКакToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Save();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DataGridViewChecked(dataGridView1);
        }
    }

    }
