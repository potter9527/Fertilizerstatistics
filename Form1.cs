using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;





namespace Fertilizerstatistics3
{
    public partial class Form1 : Form
    {
        int index = 1;

        public class DBConfig
        {
            //log.db要放在【bin\Debug底下】      
            public static string dbFile = Application.StartupPath + @"\log.db";

            public static string dbPath = "Data source=" + dbFile;

            public static SQLiteConnection sqlite_connect;
            public static SQLiteCommand sqlite_cmd;
            public static SQLiteDataReader sqlite_datareader;
        }

        private void Load_DB()
        {
            DBConfig.sqlite_connect = new SQLiteConnection(DBConfig.dbPath);
            DBConfig.sqlite_connect.Open();// Open

        }

        private void Show_DB()
        {
            this.dataGridView1.Rows.Clear();

            string sql = @"SELECT * from record;";
            DBConfig.sqlite_cmd = new SQLiteCommand(sql, DBConfig.sqlite_connect);
            DBConfig.sqlite_datareader = DBConfig.sqlite_cmd.ExecuteReader();

            if (DBConfig.sqlite_datareader.HasRows)
            {
                while (DBConfig.sqlite_datareader.Read()) //read every data
                {
                    int _serial = Convert.ToInt32(DBConfig.sqlite_datareader["serial"]);
                    int _date = Convert.ToInt32(DBConfig.sqlite_datareader["date"]);
                    int _type = Convert.ToInt32(DBConfig.sqlite_datareader["type"]);
                    string _name = Convert.ToString(DBConfig.sqlite_datareader["name"]);
                    double _price = Convert.ToDouble(DBConfig.sqlite_datareader["price"]);
                    double _number = Convert.ToDouble(DBConfig.sqlite_datareader["number"]);
                    double _total = _price * _number;

                    string _date_str = DateTimeOffset.FromUnixTimeSeconds(_date).ToString("yy-MM-dd hh:mm:ss");

                    string _type_str = "";
                    if (_type == 0)
                    { _type_str = "進貨"; }
                    else { _type_str = "出貨"; }

                    index = _serial;
                    DataGridViewRowCollection rows = dataGridView1.Rows;
                    rows.Add(new Object[] { index, _date_str, _type_str, _name, _price, _number
                                               , _total });
                }
                DBConfig.sqlite_datareader.Close();
            }
        }

        public Form1()
        {
            InitializeComponent();
            Load_DB();
            updateChart();

            Show_DB();
            this.label5.Text = index.ToString();

        }

        private void dataGridView1_CellDoubleClick(object sender
                                                                        , DataGridViewCellEventArgs e)
        {

            DataGridViewCellCollection selRowData = dataGridView1.Rows[e.RowIndex].Cells;
 
      string _type = "";
      _type = Convert.ToString(selRowData[2].Value);
 
      if (_type.Equals("進貨"))
      {
        radioButton1.Checked = true;
      }
      else
      {
        radioButton2.Checked = true;
      }
 
      this.comboBox1.Text = Convert.ToString(selRowData[3].Value);
      this.textBox1.Text = Convert.ToString(selRowData[4].Value);
      this.textBox2.Text = Convert.ToString(selRowData[5].Value);
      this.label5.Text = Convert.ToString(selRowData[0].Value);
 
    }




        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

            string _name = "";
            long _date = 0;
            int _stock_type = 0;
            double _price = 0;
            double _number = 0;
            double _sum = 0;

            // 抓取textbox的資料
            _name = comboBox1.Text;
            _price = Convert.ToDouble(textBox1.Text);
            _number = Convert.ToDouble(textBox2.Text);

            _sum = _price * _number;
            _date = DateTimeOffset.Now.ToUnixTimeSeconds();
            if (radioButton1.Checked == true)
            {
                _stock_type = 0;
            }
            else
            {
                _stock_type = 1;
            }
            // update
            this.index = this.index + 1;

            // add item into database

            string sql = @"INSERT INTO record (date, type, name,price,number)
                VALUES( "
                       + " '" + _date.ToString() + "' , "
                       + " '" + _stock_type.ToString() + "' , "
                       + " '" + _name.ToString() + "' , "
                       + " '" + _price.ToString() + "' , "
                       + " '" + _number.ToString() + "'   "
                      + ");";
            DBConfig.sqlite_cmd = new SQLiteCommand(sql, DBConfig.sqlite_connect);
            DBConfig.sqlite_cmd.ExecuteNonQuery();

            // show database in the gui
            Show_DB();
            updateChart();


        }






        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {

        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This is about");

        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            string _name = "";
            int _serial = 0;
            int _stock_type = 0;
            double _price = 0;
            double _number = 0;

            if (radioButton1.Checked == true)
            {
                _stock_type = 0;
            }
            else
            {
                _stock_type = 1;
            }

            // 抓取textbox的資料
            _name = comboBox1.Text;


            _price = Convert.ToDouble(textBox1.Text);
            _number = Convert.ToDouble(textBox2.Text);
            _serial = Convert.ToInt32(label5.Text);


            string sql = @"UPDATE record " +
                      " SET name = '" + _name + "',"
                        + " type = '" + _stock_type.ToString() + "' , "
                        + " price = '" + _price.ToString() + "',"
                        + " number = '" + _number.ToString() + "' "
                        + "   where serial = " + _serial.ToString() + ";";


            DBConfig.sqlite_cmd = new SQLiteCommand(sql, DBConfig.sqlite_connect);
            DBConfig.sqlite_cmd.ExecuteNonQuery();
            Show_DB();
            updateChart();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            save.FileName = "Export_Data";
            save.Filter = "*.xlsx|*.xlsx";
            if (save.ShowDialog() != DialogResult.OK) return;

            // Excel 物件
            Microsoft.Office.Interop.Excel.Application xls = null;
            try
            {
                xls = new Microsoft.Office.Interop.Excel.Application();
                // Excel WorkBook
                Microsoft.Office.Interop.Excel.Workbook book = xls.Workbooks.Add();
                //Excel.Worksheet Sheet = (Excel.Worksheet)book.Worksheets[1];
                Microsoft.Office.Interop.Excel.Worksheet Sheet = xls.ActiveSheet;

                // 把 DataGridView 資料塞進 Excel 內

                // DataGridView 標題
                for (int k = 0; k < this.dataGridView1.Columns.Count; k++)
                {
                    Sheet.Cells[1, k + 1] = this.dataGridView1.Columns[k].HeaderText.ToString();
                }
                // DataGridView 內容
                for (int i = 0; i < this.dataGridView1.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < this.dataGridView1.Columns.Count; j++)
                    {
                        string value = dataGridView1.Rows[i].Cells[j].Value.ToString();
                        Sheet.Cells[i + 2, j + 1] = value;
                    }
                }

                // 儲存檔案
                book.SaveAs(save.FileName);
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                xls.Quit();
            }
        }

        private void chart1_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            save.FileName = "Export_bar_Chart_Data";
            save.Filter = "*.xlsx|*.xlsx";
            if (save.ShowDialog() != DialogResult.OK) return;

            // Excel 物件

            Microsoft.Office.Interop.Excel.Application xls = null;
            try
            {
                xls = new Microsoft.Office.Interop.Excel.Application();
                // Excel WorkBook
                Microsoft.Office.Interop.Excel.Workbook book = xls.Workbooks.Add();
                //Excel.Worksheet Sheet = (Excel.Worksheet)book.Worksheets[1];
                Microsoft.Office.Interop.Excel.Worksheet Sheet = xls.ActiveSheet;

                // 把資料塞進 Excel 內

                // 標題
                Sheet.Cells[1, 1] = "標籤";
                Sheet.Cells[1, 2] = "數量";

                // 內容
                for (int k = 0; k < this.chart1.Series["stocks"].Points.Count; k++)
                {
                    Sheet.Cells[k + 2, 1] = this.chart1.Series["stocks"].Points[k].AxisLabel.ToString();
                    Sheet.Cells[k + 2, 2] = this.chart1.Series["stocks"].Points[k].YValues[0].ToString();
                }


                // 儲存檔案
                book.SaveAs(save.FileName);
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                xls.Quit();
            }

        }

        private void button6_Click(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            save.FileName = "Export_bar_Chart1_JPG";
            save.Filter = "*.jpg|*.jpg";
            if (save.ShowDialog() != DialogResult.OK) return;

            chart1.SaveImage(save.FileName, System.Drawing.Imaging.ImageFormat.Jpeg);

        }

        private void button8_Click(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            save.FileName = "Export_pie_Chart_Data";
            save.Filter = "*.xlsx|*.xlsx";
            if (save.ShowDialog() != DialogResult.OK) return;

            // Excel 物件
            Microsoft.Office.Interop.Excel.Application xls = null;
            try
            {
                xls = new Microsoft.Office.Interop.Excel.Application();
                // Excel WorkBook
                Microsoft.Office.Interop.Excel.Workbook book = xls.Workbooks.Add();
                //Excel.Worksheet Sheet = (Excel.Worksheet)book.Worksheets[1];
                Microsoft.Office.Interop.Excel.Worksheet Sheet = xls.ActiveSheet;

                // 把資料塞進 Excel 內

                // 標題
                Sheet.Cells[1, 1] = "標籤";
                Sheet.Cells[1, 2] = "數量";

                // 內容
                for (int k = 0; k < this.chart2.Series["stocks"].Points.Count; k++)
                {
                    Sheet.Cells[k + 2, 1] = this.chart2.Series["stocks"].Points[k].AxisLabel.ToString();
                    Sheet.Cells[k + 2, 2] = this.chart2.Series["stocks"].Points[k].YValues[0].ToString();
                }


                // 儲存檔案
                book.SaveAs(save.FileName);
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                xls.Quit();
            }

        }

        private void button7_Click(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            save.FileName = "Export_Chart_pie_JPG";
            save.Filter = "*.jpg|*.jpg";
            if (save.ShowDialog() != DialogResult.OK) return;

            chart2.SaveImage(save.FileName, System.Drawing.Imaging.ImageFormat.Jpeg);

        }
        public void updateChart()
        {
            // 1. 進貨統計

            string sql = @"SELECT * from record where type=1;";
            DBConfig.sqlite_cmd = new SQLiteCommand(sql, DBConfig.sqlite_connect);
            DBConfig.sqlite_datareader = DBConfig.sqlite_cmd.ExecuteReader();

            Dictionary<string, double> _stocks_bar_out = new Dictionary<string, double>();
            Dictionary<string, double> _stocks_bar_out_sum = new Dictionary<string, double>();

            if (DBConfig.sqlite_datareader.HasRows)
            {
                while (DBConfig.sqlite_datareader.Read()) //read every data
                {
                    string _name = Convert.ToString(DBConfig.sqlite_datareader["name"]);
                    double _price = Convert.ToDouble(DBConfig.sqlite_datareader["price"]);
                    double _number = Convert.ToDouble(DBConfig.sqlite_datareader["number"]);
                    if (!_stocks_bar_out.ContainsKey(_name))
                    {
                        _stocks_bar_out.Add(_name, 0);
                        _stocks_bar_out_sum.Add(_name, 0);
                    }
                    _stocks_bar_out[_name] = _stocks_bar_out[_name] + _number;
                    _stocks_bar_out_sum[_name] = _stocks_bar_out_sum[_name] + _number * _price;
                }
                DBConfig.sqlite_datareader.Close();
            }



            this.chart1.Series["stocks"].Points.Clear();
            foreach (var OneItem in _stocks_bar_out)
            {
                this.chart1.Series["stocks"].Points.AddXY(OneItem.Key, OneItem.Value);
            }

            this.chart2.Series["stocks"].Points.Clear();
            foreach (var OneItem in _stocks_bar_out_sum)
            {
                this.chart2.Series["stocks"].Points.AddXY(OneItem.Key, OneItem.Value);
            }

        }


    }
}








