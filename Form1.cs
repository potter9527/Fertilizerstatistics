using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Fertilizerstatistics3
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

       

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //計算
            string i_input_price = textBox1.Text;
            string i_input_num = textBox2.Text;
            double _price = Convert.ToDouble(i_input_price);
            double _num = Convert.ToDouble(i_input_num);
            richTextBox1.Text = (_price * _num).ToString();

            //進出貨選擇
            string _radiobutton_log = "";
            if (radioButton1.Checked == true)
            {
                _radiobutton_log = "進貨";
            }
            else
            {
                _radiobutton_log = "出貨";
            }

            richTextBox1.Text = String.Format("{0} : {1} "
                          , _price * _num, _radiobutton_log);

            //品項輸入
            string _combobox_log = comboBox1.SelectedItem.ToString();

            richTextBox1.Text = String.Format("{0} : {1} {2} "
                , _price * _num, _radiobutton_log, _combobox_log);

            //匯出
            DataGridViewRowCollection rows = dataGridView1.Rows;
            DateTime date = DateTime.Now; // 現在時間
            rows.Add(new Object[] { "", date.ToString("yyyy/MM/dd HH:mm:ss")
                  , _radiobutton_log, _combobox_log, _price, _num, _price * _num });


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

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
