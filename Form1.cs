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

using iText.IO.Font;
using iText.IO.Image;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Pdf.Extgstate;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using iText.Kernel.Pdf.Canvas.Draw;
using System.Drawing.Imaging;
using System.IO;

using MathNet.Numerics.Statistics;
using System.Data.SQLite;





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
            Show_Statistic();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            PrintPDF();
        }

        void PrintPDF()
        {
            // Set the output dir and file name
            // string directory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            SaveFileDialog save = new SaveFileDialog();
            save.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            save.FileName = "Export_PDF";
            save.Filter = "*.pdf|*.pdf";
            if (save.ShowDialog() != DialogResult.OK) return;

            manipulatePdf(save.FileName);
        }
        public byte[] BmpToBytes(Bitmap bmp)
        {
            System.IO.MemoryStream ms = new System.IO.MemoryStream();
            bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
            byte[] b = ms.GetBuffer();
            return b;
        }
        void manipulatePdf(String src)
        {

            String src_tmp = src + "_tmp.pdf";
            src = src + ".pdf";

            // 1. create pdf
            PdfWriter writer = new PdfWriter(src_tmp);
            PdfDocument pdf = new PdfDocument(writer);
            Document document = new Document(pdf, PageSize.A4.Rotate());
            document.SetMargins(40, 40, 40, 40);
            PdfFont font = PdfFontFactory.CreateFont(@"c:/Windows/fonts/kaiu.ttf", PdfEncodings.IDENTITY_H);

            // 2. create content
            // 2.1. add header
            Paragraph header_1 = new Paragraph("北護 藥物倉儲系統 統計報表")
               .SetTextAlignment(TextAlignment.CENTER)
               .SetFontSize(36)
               .SetFont(font);
            document.Add(header_1);

            // Line separator
            LineSeparator ls = new LineSeparator(new SolidLine());
            document.Add(ls);

            Paragraph space_1 = new Paragraph(" ")
               .SetTextAlignment(TextAlignment.CENTER)
               .SetFontSize(10)
               .SetFont(font);
            document.Add(space_1);

            /////////////////////////////////
            // table chart 1

            Paragraph header_2 = new Paragraph("出貨各類別圖表\n")
               .SetTextAlignment(TextAlignment.CENTER)
               .SetFontSize(24)
               .SetFont(font);
            document.Add(header_2);

            Table table_ch_1 = new Table(2, true);
            table_ch_1.SetFont(font);

            table_ch_1.AddHeaderCell("出貨各類別數量");
            table_ch_1.AddHeaderCell("出貨各類別總量比例");

            // 出貨各類別數量
            Bitmap bitmap_ch1_1 = new Bitmap(chart1.Width, chart1.Height, PixelFormat.Format24bppRgb);
            chart1.DrawToBitmap(bitmap_ch1_1, new System.Drawing.Rectangle(0, 0, chart1.Width, chart1.Height));
            ImageData imageData_ch1_1 = ImageDataFactory.Create(BmpToBytes(bitmap_ch1_1));
            iText.Layout.Element.Image image_ch1_1 = new iText.Layout.Element.Image(imageData_ch1_1);
            image_ch1_1.SetAutoScale(true);
            table_ch_1.AddCell(image_ch1_1);

            // 出貨各類別數量
            Bitmap bitmap_ch1_2 = new Bitmap(chart2.Width, chart2.Height, PixelFormat.Format24bppRgb);
            chart2.DrawToBitmap(bitmap_ch1_2, new System.Drawing.Rectangle(0, 0, chart2.Width, chart2.Height));
            ImageData imageData_ch1_2 = ImageDataFactory.Create(BmpToBytes(bitmap_ch1_2));
            iText.Layout.Element.Image image_ch1_2 = new iText.Layout.Element.Image(imageData_ch1_2);
            image_ch1_2.SetAutoScale(true);
            table_ch_1.AddCell(image_ch1_2);

            document.Add(table_ch_1);
            table_ch_1.Complete();

            // move to next page
            // Creating an Area Break
            AreaBreak a_ch_1 = new AreaBreak();

            // Adding area break to the PDF       
            document.Add(a_ch_1);

            /////////////////////////////////
            // table chart 2

            Paragraph header_3 = new Paragraph("出貨統計\n")
               .SetTextAlignment(TextAlignment.CENTER)
               .SetFontSize(24)
               .SetFont(font);
            document.Add(header_2);

            Table table_ch_2 = new Table(2, true);
            table_ch_2.SetFont(font);

            table_ch_2.AddHeaderCell("出貨統計資料");
            table_ch_2.AddHeaderCell("出貨統計資料 QR Code");

            // 出貨統計資料

            table_ch_2.AddCell(new Paragraph(richTextBox2.Text));

            // 出貨統計資料 QR Code
            System.Drawing.Bitmap bitmap_2 = get_qrcode(richTextBox2.Text, pictureBox2.Width, pictureBox2.Height);
            ImageData imageData_2 = ImageDataFactory.Create(BmpToBytes(bitmap_2));
            iText.Layout.Element.Image image_ch2 = new iText.Layout.Element.Image(imageData_2);
            image_ch2.SetAutoScale(true);
            table_ch_2.AddCell(image_ch2);

            document.Add(table_ch_2);
            table_ch_2.Complete();

            // move to next page
            // Creating an Area Break          
            AreaBreak a_ch_2 = new AreaBreak();

            // Adding area break to the PDF       
            document.Add(a_ch_2);

            ///////////////////////////
            // 2.3 table

            int _table_num = dataGridView1.ColumnCount;
            Table table = new Table(_table_num, true);
            table.SetFont(font);

            // add header
            for (int i = 0; i < _table_num; i++)
            {
                table.AddHeaderCell(new Paragraph(dataGridView1.Columns[i].HeaderText));
            }

            // add content
            for (int row = 0; row < dataGridView1.Rows.Count; row++)
            {
                for (int col = 0; col < dataGridView1.Rows[row].Cells.Count; col++)
                {
                    if (dataGridView1.Rows[row].Cells[col].Value != null)
                    {
                        string _tmp = dataGridView1.Rows[row].Cells[col].Value.ToString();
                        table.AddCell(new Paragraph(_tmp));
                    }
                }
            }

            document.Add(table);
            table.Complete();

            // 3. close pdf

            document.Close();


            // 4. edit existed pdf
            PdfReader reader2 = new PdfReader(src_tmp);
            PdfWriter writer2 = new PdfWriter(src);
            PdfDocument pdfDoc2 = new PdfDocument(reader2, writer2);
            Document document2 = new Document(pdfDoc2);

            // 5. add Page numbers
            draw_header(pdfDoc2, document2);
            document2.Close();
            File.Delete(src_tmp);
        }
        // 畫pdf的頁碼
        void draw_simple_page_num(PdfDocument pdfDoc, Document document)
        {
            int n = pdfDoc.GetNumberOfPages();
            for (int i = 1; i <= n; i++)
            {
                document.ShowTextAligned(new Paragraph(String
                     .Format("page" + i + " of " + n)),
                      806, 595, i, TextAlignment.RIGHT,
                    VerticalAlignment.TOP, 0);
            }
        }

        // 畫pdf的頁首頁尾
        void draw_header(PdfDocument pdfDoc, Document document)
        {
            PdfFont font = PdfFontFactory.CreateFont(@"c:/Windows/fonts/kaiu.ttf", PdfEncodings.IDENTITY_H);
            iText.Kernel.Geom.Rectangle pageSize;
            PdfCanvas canvas;
            int n = pdfDoc.GetNumberOfPages();
            for (int i = 1; i <= n; i++)
            {
                PdfPage page = pdfDoc.GetPage(i);
                pageSize = page.GetPageSize();
                canvas = new PdfCanvas(page);




                //Draw header text
                canvas.BeginText()
                    .SetFontAndSize(font, 15)
                    .MoveText(pageSize.GetWidth() / 2 - 54, pageSize.GetHeight() - 20)
                    .ShowText("北護藥局倉儲系統")
                    .EndText();

                //Draw footer line
                iText.Kernel.Colors.Color bgColour = new DeviceRgb(0, 0, 0);
                canvas.SetStrokeColor(bgColour)
                    .SetLineWidth(2.2f)
                    .MoveTo(pageSize.GetWidth() / 2 - 30, 20)
                    .LineTo(pageSize.GetWidth() / 2 + 30, 20)
                    .Stroke();

                //Draw page number
                canvas.BeginText()
                    .SetFontAndSize(font, 7)
                    .MoveText(pageSize.GetWidth() / 2 - 7, 10)
                    .ShowText(i.ToString())
                    .ShowText(" of ")
                    .ShowText(n.ToString())
                    .EndText();
                //Draw watermark
                Paragraph p = new Paragraph("極  機  密 \n Confidential").SetFont(font).SetFontSize(60);
                canvas.SaveState();
                PdfExtGState gs1 = new PdfExtGState().SetFillOpacity(0.2f);
                canvas.SetExtGState(gs1);
                document.ShowTextAligned(p, pageSize.GetWidth() / 2, pageSize.GetHeight() / 2, pdfDoc.GetPageNumber(page), TextAlignment.CENTER, VerticalAlignment.MIDDLE, 45);
                canvas.RestoreState();
            }
        }


        public Bitmap get_qrcode(string log, int i_width, int i_height)
        {
            System.Drawing.Bitmap bitmap = null;
            //let string to qr-code
            string strQrCodeContent = log;

            ZXing.BarcodeWriter writer = new ZXing.BarcodeWriter
            {
                Format = ZXing.BarcodeFormat.QR_CODE,
                Options = new ZXing.QrCode.QrCodeEncodingOptions
                {
                    //Create Photo 
                    Height = i_width,
                    Width = i_height,
                    CharacterSet = "UTF-8",

                    //錯誤修正容量
                    //L水平    7%的字碼可被修正
                    //M水平    15%的字碼可被修正
                    //Q水平    25%的字碼可被修正
                    //H水平    30%的字碼可被修正
                    ErrorCorrection = ZXing.QrCode.Internal.ErrorCorrectionLevel.H
                }

            };
            //Create Qr-code , use input string
            bitmap = writer.Write(strQrCodeContent);
            /*
            string strDir;
            strDir = Directory.GetCurrentDirectory();
            strDir += "\\temp.jpg";
            bitmap.Save(strDir, System.Drawing.Imaging.ImageFormat.Jpeg);
            */
            return bitmap;
        }


        private void Show_Statistic()
        {
            // 1. 進貨統計

            string sql = @"SELECT * from record where type=0;";
            DBConfig.sqlite_cmd = new SQLiteCommand(sql, DBConfig.sqlite_connect);
            DBConfig.sqlite_datareader = DBConfig.sqlite_cmd.ExecuteReader();


            // 2. 出貨統計
            sql = @"SELECT * from record where type=1;";
            DBConfig.sqlite_cmd = new SQLiteCommand(sql, DBConfig.sqlite_connect);
            DBConfig.sqlite_datareader = DBConfig.sqlite_cmd.ExecuteReader();

            List<double> stock_out = new List<double>();

            if (DBConfig.sqlite_datareader.HasRows)
            {
                while (DBConfig.sqlite_datareader.Read()) //read every data
                {
                    string _name = Convert.ToString(DBConfig.sqlite_datareader["name"]);
                    double _price = Convert.ToDouble(DBConfig.sqlite_datareader["price"]);
                    double _number = Convert.ToDouble(DBConfig.sqlite_datareader["number"]);
                    stock_out.Add(_price * _number);
                }
                DBConfig.sqlite_datareader.Close();
            }
            // 3. get statistic data
            string _log = "出貨統計\n\n" + statistic(stock_out);
            richTextBox2.Text = _log;

            System.Drawing.Bitmap qr_code = get_qrcode(richTextBox2.Text, pictureBox2.Width, pictureBox2.Height);
            pictureBox2.Image = qr_code;
        }
        public string statistic(List<double> data)
        {
            double mean = Statistics.Mean(data);
            double stddiv = Statistics.StandardDeviation(data);
            double pstddiv = Statistics.PopulationStandardDeviation(data);
            double variance = Statistics.Variance(data);
            double median = Statistics.Median(data);
            double lowerQuartile = Statistics.LowerQuartile(data);
            double upperQuartile = Statistics.UpperQuartile(data);
            double interQuartileRange = Statistics.InterquartileRange(data);
            double min = Statistics.Minimum(data);
            double max = Statistics.Maximum(data);

            string _log = "";
            _log = string.Format("平均值: {0}\n" +
                "標準差: {1}\n" +
                "變異數: {2}\n" +
                "中位數: {3}\n" +
                "最小值: {4}\n" +
                "最大值: {5}",
                mean, stddiv, variance, median, min, max);

            return _log;
        }

    }
}








