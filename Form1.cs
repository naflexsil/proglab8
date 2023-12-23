using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;

namespace proglab8 {
    public partial class Form1 : Form {
        public Form1() {
            InitializeComponent();
        }

        private void InitializeDataGridView() {
            dataGridView1.ColumnCount = 2;
            dataGridView1.Columns[0].Name = "X";
            dataGridView1.Columns[1].Name = "Y";
        }


        public struct Dots {
            public double x;
            public double y;
            public Dots(double myX, double myY) {
                x = myX;
                y = myY;
            }
        }
        List<Dots> dots = new List<Dots>();
        void DrowDots() {
            for (int turn = 0; turn < dots.Count; ++turn) {
                this.chart1.Series[1].Points.AddXY(dots[turn].x, dots[turn].y);
            }
        }


        double FindMin() {
            double min = double.MaxValue;
            for (int turn = 0; turn < dots.Count; ++turn) {
                if (dots[turn].x < min) min = dots[turn].x;
            }
            return min;
        }
        double FindMax() {
            double max = double.MinValue;
            for (int turn = 0; turn < dots.Count; ++turn) {
                if (dots[turn].x > max) max = dots[turn].x;
            }
            return max;
        }



        string Func2() {
            double sumXY = 0; double sumX = 0; double sumY = 0; double sumPowX = 0;
            for (int turn = 0; turn < dots.Count; ++turn) {
                sumXY += dots[turn].x * dots[turn].y;
                sumX += dots[turn].x;
                sumY += dots[turn].y;
                sumPowX += dots[turn].x * dots[turn].x;
            }
            double a = (dots.Count * sumXY - sumX * sumY) / (dots.Count * sumPowX - sumX * sumX);
            double b = (sumY - a * sumX) / dots.Count;
            for (int turn = Convert.ToInt32(FindMin()); turn <= Convert.ToInt32(FindMax()); ++turn) {
                double thisY = a * turn + b;
                this.chart1.Series[0].Points.AddXY(turn, thisY);
            }
            String answer = "";
            if (b >= 0) {
                answer = "A = " + Convert.ToString(Math.Round(a, 3)) + "\n" + "B = " + Convert.ToString(Math.Round(b, 3)) + "\n" +
                    "Y = " + Convert.ToString(Math.Round(a, 3)) + "X + " + Convert.ToString(Math.Round(b, 3)) + "\n";
            }
            else {
                answer = "A = " + Convert.ToString(Math.Round(a, 3)) + "\n" + "B = " + Convert.ToString(Math.Round(b, 3)) + "\n" +
                    "Y = " + Convert.ToString(Math.Round(a, 3)) + "X " + Convert.ToString(Math.Round(b, 3)) + "\n";
            }
            return answer;
        }




        private void рассчитатьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                this.chart1.Series[0].Points.Clear();
                this.chart1.Series[1].Points.Clear();
                this.chart1.Series[2].Points.Clear();
                for (int turn = 0; turn < dataGridView1.RowCount - 1; ++turn)
                {
                    dots.Add(new Dots(Convert.ToDouble(dataGridView1[0, turn].Value), Convert.ToDouble(dataGridView1[1, turn].Value)));
                }
                DrowDots();
                label1.Text = "Линейная регрессия:\n";
                label1.Text += Func2();
                //label2.Text += Func3();
            }
            catch
            {
                Mistake();
            }

        }


            private void рандомТочкиToolStripMenuItem_Click(object sender, EventArgs e) {
            var rand = new Random();
            string userInput = randBox.Text;
            if (int.TryParse(userInput, out int number) && number > 0)
            {
                dataGridView1.Rows.Clear(); 
                dots.Clear(); 

                if (dataGridView1.Columns.Count == 0)
                {
                    InitializeDataGridView();
                }

                for (int i = 0; i < number; i++)
                {
                    double x = Convert.ToDouble(rand.Next(-100, 100));
                    double y = Convert.ToDouble(rand.Next(-100, 100));
                    dots.Add(new Dots(x, y));
                    dataGridView1.Rows.Add(new object[] { x, y });
                }
            }
            else
            {
                MessageBox.Show("Введите положительное целое число для количества точек.");
            }
        }



        void Mistake()
        {
            label1.Text = "Ошибка";
        }



        private void загрузитьИзExcelToolStripMenuItem_Click(object sender, EventArgs e) {
            string filePath = @"C:\prog.xlsx";
            LoadExcelFile(filePath);
        }


        private void LoadExcelFile(string filePath)
        {
            FileInfo fileInfo = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                DataTable dt = new DataTable();
                foreach (var firstRowCell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
                {
                    dt.Columns.Add(firstRowCell.Text);
                }
                for (int rowNum = 2; rowNum <= worksheet.Dimension.End.Row; rowNum++)
                {
                    var wsRow = worksheet.Cells[rowNum, 1, rowNum, worksheet.Dimension.End.Column];
                    DataRow row = dt.Rows.Add();
                    foreach (var cell in wsRow)
                    {
                        row[cell.Start.Column - 1] = cell.Text;
                    }
                }
                dataGridView1.DataSource = dt;
            }
        }



        private void очиститьToolStripMenuItem_Click(object sender, EventArgs e) {
            label1.Text = "";
            this.chart1.Series[0].Points.Clear();
            this.chart1.Series[1].Points.Clear();
            this.chart1.Series[2].Points.Clear();
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            dots.Clear();
        }

       
    }
}
