namespace FraseV1
{
    using System;
    using System.IO;
    using System.Data;
    using System.Text;
    using System.Drawing;
    using System.Windows.Forms;
    using Excel = Microsoft.Office.Interop.Excel;

    public partial class Form1 : Form
    {
        string folderName = "temp";

        public Form1()
        {
            InitializeComponent();
        }

        private void OpenToolStripMenuItemClick(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.DefaultExt = "*.xls;*.xlsx";
            openFileDialog.Filter = "Excel Sheet(*.xlsx)|*.xlsx";
            openFileDialog.Title = "Выберите документ для загрузки данных";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                progressBar1.Increment(progressBar1.Maximum / 100);

                // подзагрузка элементов
                ReadDataExcel(openFileDialog.FileName, comboBox1);

                progressBar1.Value = progressBar1.Maximum;

                MessageBox.Show("Данные были загружены!");
                progressBar1.Value = progressBar1.Minimum;
                if (comboBox1.Items.Count > 0)
                {
                    comboBox1.SelectedIndex = 0;
                }
            }
        }

        private void SaveAsToolStripMenuItemClick(object sender, EventArgs e)
        {
            SaveDataAsFile();
        }

        private void SelectFromDataGrid(object sender, EventArgs e)
        {
            DataGridViewChecked(dataGridView1);
        }

        private int GetPages(string fileName)
        {
            try
            {
                Excel.Application excelApp = new Excel.Application();
                var excelBooks = excelApp.Workbooks.Open(fileName);
                var count = excelBooks.Sheets.Count;

                excelBooks.Close();
                excelApp?.Quit();
                return count;
            }
            catch { }
            return 0;
        }

        private void ReadDataExcel(string fileName, ComboBox comboBox)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelBooks;
            Excel.Worksheet excelPages;
            Excel.Range excelCells;


            var maxColls = 1;
            try
            {
                //открытие файла
                excelBooks = excelApp.Workbooks.Open(fileName);
                var excelMaxPages = excelBooks.Sheets.Count;
                progressBar1.Step = progressBar1.Maximum / excelMaxPages;
                if (excelMaxPages > 0)
                {
                    for (int page = 1; page <= excelMaxPages; page++)
                    {
                        try
                        {
                            var dataTable = new DataTable();
                            //Устанавливаем номер листа из котрого будут извлекаться данные (Листы нумеруются от 1)
                            excelPages = (Excel.Worksheet)excelBooks.Sheets.get_Item(page);
                            excelCells = excelPages.UsedRange; //все ячейки, содержащие значения на данный момент
                                                               //имя колонок
                            var titleRow = 1;
                            var titleColl = 1;
                            for (int coll = 1; coll <= maxColls; coll++)
                            {
                                dataTable.Columns.Add(new DataColumn((excelCells.Cells[titleRow, titleColl] as Excel.Range).Value2.ToString()));
                            }

                            dataTable.AcceptChanges();

                            var columnName = String.Empty;
                            columnName = dataTable.Columns[0].ColumnName;
                            // добавляем заголовки таблиц в выпадающий список
                            comboBox.Items.Add(columnName);
                            for (var row = 2; row <= excelCells.Rows.Count; row++)
                            {
                                var dataRow = dataTable.NewRow();
                                for (int coll = 1; coll <= maxColls; coll++)
                                {
                                    if ((excelCells.Cells[row, coll] as Excel.Range).Value2 != null)
                                    {
                                        dataRow[coll - 1] = (excelCells.Cells[row, coll] as Excel.Range).Value2.ToString();
                                    }
                                }
                                // Добавляем в таблицу строку данных
                                dataTable.Rows.Add(dataRow);
                                dataTable.AcceptChanges();
                            }

                            // Запись страницы в txt
                            ConvertTableToTxt(dataTable, page);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Ошибка: " + ex.Message);
                        }
                        progressBar1.PerformStep();
                    }

                }
                else
                {
                    MessageBox.Show("Файл пуст или не существует.");
                }


                excelBooks.Close();
                excelApp?.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
            }
        }

        private void ConvertTableToTxt(DataTable dataTable, int page)
        {
            var dataLines = new string[dataTable.Rows.Count + 1];
            if (dataTable.Columns.Count > 0)
            {
                dataLines[0] = dataTable.Columns[0].ToString();               
            }
            if (dataTable.Rows.Count > 0)
            {
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    dataLines[i + 1] = dataTable.Rows[i][0].ToString();
                }
            }
            Directory.CreateDirectory(folderName);
            File.WriteAllLines($"{folderName}\\page{page}.txt", dataLines);
        }

        private void SetDataToGrid(DataGridView dataGridView, int page)
        {
            try
            {
                if (Directory.Exists(folderName))
                {
                    if (File.Exists($"{folderName}\\page{page}.txt"))
                    {
                        var dataLines = File.ReadAllLines($"{folderName}\\page{page}.txt");
                        dataGridView.Rows.Clear();
                        dataGridView.RowCount = dataLines.Length - 1;
                        dataGridView.ColumnCount = 2;
                        if (dataLines.Length > 0)
                        {
                            dataGridView.Columns[1].HeaderText = dataLines[0];

                        }

                        var spaceLineCount = 0; 
                        if (dataGridView.RowCount > 0)
                        {
                            for (int i = 0; i < dataGridView.RowCount; i++)
                            {
                                var line = dataLines[i + 1];
                                if (line.Trim(' ') != string.Empty)
                                {
                                    dataGridView.Rows[i - spaceLineCount].Cells[1].Value = line;
                                }
                                else
                                {
                                    spaceLineCount++;
                                }
                            }
                        }

                        dataGridView.RowCount -= spaceLineCount;
                    }
                    else
                    {
                        MessageBox.Show("Данные отсутствуют. Проверте файл Excel и перезапустите приложение, не вмешиваясь в временные файлы.");
                    }
                }
                else
                {
                    MessageBox.Show("Данные отсутствуют. Проверте файл Excel и перезапустите приложение, не вмешиваясь в временные файлы.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Возникла ошибка: {ex.Message}");
            }
        }

        private void UploadTableFromList(object sender, EventArgs e)
        {
            var index = comboBox1.SelectedIndex;
            if (index >= 0)
            {
                //Тут выбираем страницу которая будет отображаться в dataGridView1
                SetDataToGrid(dataGridView1, index + 1);
            }
            else
            {
                MessageBox.Show("Такой таблицы нет. Возможно данные не загрузились.");
            }
        }

        public void SaveDataAsFile()
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.DefaultExt = "*.txt";
            saveFileDialog.Filter = "Text File(*.txt)|*.txt";
            saveFileDialog.Title = "Выберите место для сохранения";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                FileStream fs = new FileStream(saveFileDialog.FileName, FileMode.Append);
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

        }

        private void DataGridViewChecked(DataGridView dataGridView)
        {
            try
            {
                int rowCount = dataGridView.RowCount;
                int newRowCount = 0;
                for (int i = 0; i < rowCount; i++)
                {

                    if (Convert.ToBoolean(dataGridView.Rows[i].Cells[0].Value) == true)
                    {
                        newRowCount++;
                    }
                }

                var index = 0;
                var newRows = new string[newRowCount];
                for (int i = 0; i < rowCount; i++)
                {
                    if (dataGridView.Rows[i].Cells[0].Value != null)
                    {
                        if (Convert.ToBoolean(dataGridView.Rows[i].Cells[0].Value) == true)
                        {
                            newRows[index] = dataGridView.Rows[i].Cells[1].Value.ToString();
                            index++;
                        }
                    }
                }
            
            dataGridView.Rows.Clear();
            dataGridView.RowCount = newRowCount;
                for (int i = 0; i < newRowCount; i++)
                {
                    dataGridView.Rows[i].Cells[1].Value = newRows[i];
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Возникла ошибка: {ex.Message}");
            }
        }


    }

}
