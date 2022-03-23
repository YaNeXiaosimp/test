using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml;
using System.Data;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace WpfApp1
{
    public partial class MainWindow : System.Windows.Window
    {

        public MainWindow()
        {
            InitializeComponent();
            string path1 = @"C:\OOOSealReceipt.xml";
            string path2 = @"C:\OOOSealShipment.xml";

            DataSet ds = new DataSet();
            DataSet ds2 = new DataSet();

            ds.ReadXml(path1);
            ds2.ReadXml(path2);

            rename(ds);
            rename(ds2);

            DataView dataView = new DataView(ds.Tables[1]);
            DataView dataView2 = new DataView(ds2.Tables[1]);

            grid1.ItemsSource = dataView;
            grid2.ItemsSource = dataView2;

        }
        public void rename(DataSet ds)
        {
            ds.Tables[1].Columns["product_name"].ColumnName = "Наименование";
            ds.Tables[1].Columns["count"].ColumnName = "Количество";
            ds.Tables[1].Columns["m"].ColumnName = "Масса";
            ds.Tables[1].Columns["fragile"].ColumnName = "Хрупкое";
            ds.Tables[1].Columns["storage_id"].ColumnName = "Номер скалада";
            ds.Tables[1].Columns["date"].ColumnName = "Дата";
        }
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            string string_date = textbox1.Text;
            int count_A = 0, count_B = 0, count_C = 0;
            double mass_A = 0, mass_B = 0, mass_C = 0;
            DateTime date_time1 = DateTime.Parse(string_date);

            System.Data.DataTable table1 = ((DataView)grid1.ItemsSource).ToTable();
            System.Data.DataTable table2 = ((DataView)grid2.ItemsSource).ToTable();

            for (int i = 0; i < table1.Rows.Count; i++)
            {
                DateTime curren_time = DateTime.Parse(table1.Rows[i][4].ToString());
                if (DateTime.Compare(curren_time, date_time1) <= 0)
                {
                    if (table1.Rows[i][0].ToString() == "Товар A")
                    {
                        count_A = count_A + int.Parse(table1.Rows[i][1].ToString());
                    }
                    if (table1.Rows[i][0].ToString() == "Товар B")
                    {
                        count_B = count_B + int.Parse(table1.Rows[i][1].ToString());
                    }
                    if (table1.Rows[i][0].ToString() == "Товар C")
                    {
                        count_C = count_C + int.Parse(table1.Rows[i][1].ToString());
                    }
                }
            }

            for (int i = 0; i < table2.Rows.Count; i++)
            {
                DateTime curren_time = DateTime.Parse(table2.Rows[i][4].ToString());
                if (DateTime.Compare(curren_time, date_time1) <= 0)
                {
                    if (table2.Rows[i][0].ToString() == "Товар A")
                    {
                        count_A = count_A - int.Parse(table2.Rows[i][1].ToString());
                    }
                    if (table2.Rows[i][0].ToString() == "Товар B")
                    {
                        count_B = count_B - int.Parse(table2.Rows[i][1].ToString());
                    }
                    if (table2.Rows[i][0].ToString() == "Товар C")
                    {
                        count_C = count_C - int.Parse(table2.Rows[i][1].ToString());
                    }
                }
            }
            mass_A = count_A * 0.3;
            mass_B = count_B * 0.96;
            mass_C = count_C * 4;

            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("Товар");
            dt.Columns.Add("Количество");
            dt.Columns.Add("Масса");
            dt.Rows.Add("Товар А", count_A, mass_A);
            dt.Rows.Add("Товар B", count_B, mass_B);
            dt.Rows.Add("Товар C", count_C, mass_C);
            dt.Rows.Add("Итого", count_A + count_B + count_C, mass_A + mass_B + mass_C);

            DataView dv = new DataView(dt);
            grid3.ItemsSource = dv;
        }

        private void download_btn_Click(object sender, RoutedEventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            excel.Visible = true; 
            Workbook workbook = excel.Workbooks.Add(System.Reflection.Missing.Value);
            Worksheet sheet1 = (Worksheet)workbook.Sheets[1];

            for (int j = 0; j < grid3.Columns.Count; j++) 
            {
                Range myRange = (Range)sheet1.Cells[1, j + 1];
                sheet1.Cells[1, j + 1].Font.Bold = true; 
                sheet1.Columns[j + 1].ColumnWidth = 15; 
                myRange.Value2 = grid3.Columns[j].Header;
            }
            for (int i = 0; i < grid3.Columns.Count; i++)
            {
                for (int j = 0; j < grid3.Items.Count; j++)
                {
                    TextBlock b = grid3.Columns[i].GetCellContent(grid3.Items[j]) as TextBlock;
                    Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[j + 2, i + 1];
                    myRange.Value2 = b.Text;
                }
            }
        }
    }
}
