using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using CorrelationAnalysis;

namespace CorrelationAnalysis
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            CorrelationAnalysis corr = new CorrelationAnalysis(71, 10);
            corr.Execute();
            Output(dataGridView5, corr.DispMatr);
            Output(dataGridView1, corr.StandartMatr, 71, 10);
            Output(dataGridView2, corr.CovarMatr, 10, 10, true);
            Output(dataGridView3, corr.CorrMatr, 10, 10, true);
            Output(dataGridView4, corr.HyptMatr, 10, 10);

        }

        private void Output( DataGridView dataGrid, double[,] matr, int row, int column, bool toPaint = false)
        {
            for (int i = 0; i < row; i++)
            {
                dataGrid.Rows.Add();
                for (int j = 0; j < column; j++)
                {
                    dataGrid.Rows[i].Cells[j].Value = String.Format("{0:f3}", matr[i, j]);
                    dataGrid.AutoResizeColumn(j);
                    if (toPaint && i == j){
                        dataGrid.Rows[i].Cells[j].Style.BackColor = Color.Bisque;
                    }
                }
            }
        }

        private void Output(DataGridView dataGrid, int[,] matr, int row, int column)
        {
            for (int i = 0; i < row; i++)
            {
                dataGrid.Rows.Add();
                for (int j = 0; j < column; j++)
                {
                    if (matr[i, j] == 0)
                        dataGrid.Rows[i].Cells[j].Value = "  H0  ";
                    else
                        dataGrid.Rows[i].Cells[j].Value = "  H1  ";
                    if (i == j)
                        dataGrid.Rows[i].Cells[j].Value = "  -   ";
                    dataGrid.AutoResizeColumn(j);
                }
            }
        }

        private void Output(DataGridView dataGrid, double [] arr)
        {
            dataGrid.Rows.Add();
            for(int i = 0; i < arr.Length; i++)
            {
                dataGrid.Rows[0].Cells[i].Value = String.Format("{0:f3}", arr[i]);
                dataGrid.AutoResizeColumn(i);
            }
        }
    }
}
