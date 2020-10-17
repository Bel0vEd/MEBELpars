using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ParserMEBEL
{
    public partial class Form1 : Form
    {
        int stroka = 0;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            throw new NotImplementedException();
        }
        int sdelano = 0;
        private void button1_Click(object sender, EventArgs e)
        {
            Thread thread = new Thread(Start) { IsBackground = true };
            thread.Start();
        }

        public void ProgressCallBack()
        {
            progressBar1.Invoke(new Action(() => progressBar1.Value += 1));
        }
        public void ProgressMax(int maxValue)
        {
            progressBar1.Invoke(new Action(() => progressBar1.Maximum = maxValue));
        }
        private void Start()
        {
            if (radioButton1.Checked == true)
            {
                button1.Invoke(new Action(() => button1.Enabled = false));
                MessageBox.Show("Парсинг начался!" + DateTime.Now.ToString("hh.mm"));
                if (textBox3.Text != "")
                {
                    sdelano = Int32.Parse(Regex.Match(textBox3.Text, "(?<=\\s)(\\d)+(?=\\.xlsx)").Value);
                    ExcelPackage package = new ExcelPackage(new FileInfo(textBox3.Text));
                    ExcelWorksheet sheet = package.Workbook.Worksheets[1];
                    stroka = sheet.Dimension.Rows-1;
                }
                rp.getexc(textBox1.Text, textBox3.Text);
                rp.GetObj(textBox1.Text, ProgressCallBack, ProgressMax, sdelano, textBox2.Text, stroka);
                if (progressBar1.Value < progressBar1.Maximum - 1)
                {
                    sdelano = progressBar1.Value;
                    MessageBox.Show("Пропал интернет");
                }
                MessageBox.Show("Парсинг завершён!" + DateTime.Now.ToString("hh.mm"));
                button1.Invoke(new Action(() => button1.Enabled = true));
            }
            if (radioButton2.Checked == true)
            {
                button1.Invoke(new Action(() => button1.Enabled = false));
                MessageBox.Show("Парсинг начался!" + DateTime.Now.ToString("hh.mm"));
                klenmarket.getexc(textBox1.Text);
                try
                {
                    klenmarket.GetObj(textBox1.Text, ProgressCallBack, ProgressMax, sdelano);
                }
                catch (WebException)
                {
                    sdelano = progressBar1.Value;
                    MessageBox.Show("Пропал интернет");
                }
                MessageBox.Show("Парсинг завершён!" + DateTime.Now.ToString("hh.mm"));
                button1.Invoke(new Action(() => button1.Enabled = true));
            }
            if (radioButton3.Checked == true)
            {
                button1.Invoke(new Action(() => button1.Enabled = false));
                MessageBox.Show("Парсинг начался!" + DateTime.Now.ToString("hh.mm"));
                if (textBox3.Text != "")
                {
                    sdelano = Int32.Parse(Regex.Match(textBox3.Text, "(?<=\\s)(\\d)+(?=\\.xlsx)").Value);
                    ExcelPackage package = new ExcelPackage(new FileInfo(textBox3.Text));
                    ExcelWorksheet sheet = package.Workbook.Worksheets[1];
                    stroka = sheet.Dimension.Rows-1;
                }
                entero.getexc(textBox1.Text, textBox3.Text);
                try
                {
                    entero.GetObj(textBox1.Text, ProgressCallBack, ProgressMax,sdelano, stroka);
                }
                catch (WebException)
                {
                    sdelano = progressBar1.Value;
                    MessageBox.Show("Пропал интернет");
                }
                MessageBox.Show("Парсинг завершён!" + DateTime.Now.ToString("hh.mm"));
                button1.Invoke(new Action(() => button1.Enabled = true));
            }
        }
        public void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = dialog.SelectedPath;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "TXT|*.txt";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string path = dialog.FileName;
                textBox2.Text = path;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "EXCEL|*.xlsx";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string path = dialog.FileName;
                textBox3.Text = path;
            }
        }
    }
}
