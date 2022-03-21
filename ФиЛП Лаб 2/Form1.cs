using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ФиЛП_Лаб_2
{
    public partial class Form1 : Form
    {
        DataTableCollection tablesCollection;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog dialog = new OpenFileDialog()
            { Filter = "Excel Workbook|*.xlsx|Excel 97-2003 Workbook|*.xls", ValidateNames = true })
            {
                dialog.InitialDirectory = @"C:\Users\kirill\Desktop\Учеба\Семестр 6\Функциональное програмирование\Лаб 2";
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    using (var fileStream = File.Open(dialog.FileName, FileMode.Open, FileAccess.Read))
                    {
                        using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(fileStream))
                        {
                            DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable=(_)=>new ExcelDataTableConfiguration() { UseHeaderRow = true}
                            });
                            tablesCollection = result.Tables;
                            dataGridView1.DataSource = tablesCollection[0];
                            dataGridView2.DataSource = tablesCollection[0];
                        }
                    }
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if(tablesCollection !=null)
            {
                DataTable table0 = tablesCollection[0];
                DataTable table1 = tablesCollection[1];
                DataTable table2 = tablesCollection[2];
                DataTable table3 = tablesCollection[3];

                switch (comboBox1.SelectedIndex)
                {
                    case 0:
                        dataGridView1.DataSource = null;
                        dataGridView1.Rows.Clear();
                        dataGridView1.ColumnCount = 2;
                        dataGridView1.Columns[0].Name = "Страна";
                        dataGridView1.Columns[1].Name = "Площадь";

                        double max = (double)table0.Rows[0][2];
                        string name = (string)table0.Rows[0][1];
                        for (int i = 1; i < table0.Rows.Count; ++i)
                            if (max <= (double)table0.Rows[i][2])
                            {
                                max = (double)table0.Rows[i][2];
                                name = (string)table0.Rows[i][1];
                            }

                        dataGridView1.Rows.Add(name, max);
                        break;

                    case 1:
                        dataGridView1.DataSource = null;
                        dataGridView1.Rows.Clear();
                        dataGridView1.ColumnCount = 1;
                        dataGridView1.Columns[0].Name = "Страна";

                        string value1 = textBox1.Text;

                        for (int i = 0; i < table1.Rows.Count; ++i)
                            if (value1 == table1.Rows[i][5].ToString())
                            {
                                int index = Convert.ToInt32(table1.Rows[i][0]);
                                for (int j = 1; j < table0.Rows.Count; ++j)
                                    if (index == Convert.ToInt32(table0.Rows[j][0]))
                                        dataGridView1.Rows.Add(table0.Rows[j][1]);
                            }
                        break;

                    case 2:
                        dataGridView1.DataSource = null;
                        dataGridView1.Rows.Clear();
                        dataGridView1.ColumnCount = 2;
                        dataGridView1.Columns[0].Name = "Страна";
                        dataGridView1.Columns[1].Name = "Кол-во национальностей";

                        try
                        {
                            int value2 = Convert.ToInt32(textBox1.Text);
                            int[] ncount = new int[table3.Rows.Count];

                            for (int i = 0; i < table3.Rows.Count; ++i)
                                if (table3.Rows[i][1].ToString() != "")
                                    ncount[Convert.ToInt32(table3.Rows[i][0])]++;


                            for (int i = 1; i < ncount.Length; ++i)
                                if (ncount[i] > value2)
                                    dataGridView1.Rows.Add(table0.Rows[i][1], ncount[i]);
                        }
                        catch
                        {
                            MessageBox.Show("Ввили не инт", "Беда");
                        }
                        break;

                    case 3:
                        dataGridView1.DataSource = null;
                        dataGridView1.Rows.Clear();
                        dataGridView1.ColumnCount = 1;
                        dataGridView1.Columns[0].Name = "Горные хребты";

                        string value3 = textBox1.Text;
                        int index3 = -1;

                        for (int i = 1; i < table0.Rows.Count; ++i)
                            if (table0.Rows[i][1].ToString() == value3)
                                index3 = Convert.ToInt32(table0.Rows[i][0]);

                        for (int i = 1; i < table1.Rows.Count; ++i)
                            if (Convert.ToInt32(table1.Rows[i][0]) == index3)
                                if (table1.Rows[i][5].ToString() != "")
                                    dataGridView1.Rows.Add(table1.Rows[i][5]);

                        break;

                    case 4:
                        dataGridView1.DataSource = null;
                        dataGridView1.Rows.Clear();
                        dataGridView1.ColumnCount = 2;
                        dataGridView1.Columns[0].Name = "Страна";
                        dataGridView1.Columns[1].Name = "Население";

                        try
                        {
                            int value4 = Convert.ToInt32(textBox1.Text);

                            for (int i = 0; i < table2.Rows.Count; ++i)
                                if ((double)table2.Rows[i][1] < value4)
                                    dataGridView1.Rows.Add(table0.Rows[i][1], table2.Rows[i][1]);
                        }
                        catch
                        {
                            MessageBox.Show("Ввили не инт", "Беда");
                        }
                        break;

                    default:
                        MessageBox.Show("Опция не выбрана", "Беда");
                        break;
                }
            }
            else
                MessageBox.Show("База данных не открыта", "Беда");
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch(comboBox1.SelectedIndex)
            {
                case 0:
                    textBox1.Enabled = false;
                    textBox1.Visible = false;
                    label2.Visible = false;
                    break;

                case 1:
                    textBox1.Enabled = true;
                    textBox1.Visible = true;
                    label2.Visible = true;
                    label2.Text = "Укажите название горного хребта:";
                    break;

                case 2:
                    textBox1.Enabled = true;
                    textBox1.Visible = true;
                    label2.Visible = true;
                    label2.Text = "Введите число национальностей:";
                    break;

                case 3:
                    textBox1.Enabled = true;
                    textBox1.Visible = true;
                    label2.Visible = true;
                    label2.Text = "Укажите страну:";
                    break;

                case 4:
                    textBox1.Enabled = true;
                    textBox1.Visible = true;
                    label2.Visible = true;
                    label2.Text = "Введите численность населения:";
                    break;

                default:
                    textBox1.Enabled = false;
                    textBox1.Visible = false;
                    label2.Visible = false;
                    break;
            }
        }
    }
}
