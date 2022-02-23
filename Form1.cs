﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using file = Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using LicenseContext = OfficeOpenXml.LicenseContext;

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

        }

        public void Output( DataGridView dataGrid, double[,] matr, int row, int column, bool toPaint = false)
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

        public void Output(DataGridView dataGrid, double [] arr)
        {
            dataGrid.Rows.Add();
            for(int i = 0; i < arr.Length; i++)
            {
                dataGrid.Rows[0].Cells[i].Value = String.Format("{0:f3}", arr[i]);
                dataGrid.AutoResizeColumn(i);
            }
        }
    }

    class CorrelationAnalysis
    {
        int column;
        int row;
        double[,] matr;
        double[] averValues;
        double[] dispEstMatr;
        double[,] standartMatr;
        double[,] covarMatr;
        double[,] corrMatr;

        public CorrelationAnalysis(int row, int column)
        {
            this.row = row;
            this.column = column;
            matr = new double[row, column];
            averValues = new double[column];
            dispEstMatr = new double[column];
            standartMatr = new double[row, column];
            covarMatr = new double[column, column];
            corrMatr = new double[column, column];
        }

        public double[,] StandartMatr { get => standartMatr; }
        
        public double[,] CovarMatr { get => covarMatr; }

        public double[,] CorrMatr { get => corrMatr; }

        public double[] DispMatr { get => dispEstMatr; }

        public void Execute()
        {
            ExelWork exelObj = new ExelWork(row, column);
            matr = exelObj.ExportFile();

            FindAverage(matr);
            FindVarianceEstimate(matr);
            StandartizedMatrix(matr);
            CorrelationMatrix();
            CovariationMatrix(matr);
        }

        private void FindAverage(double[,] matr)
        {
            for (int i = 0; i < column; i++)
            {
                double sum = 0;
                for (int j = 0; j < row; j++)
                {
                    sum += matr[j, i];
                }
                averValues[i] = sum / row;
            }
        }

        private void FindVarianceEstimate(double[,] matr)
        {
            for (int i = 0; i < column; i++)
            {
                double sum = 0;
                for (int j = 0; j < row; j++)
                {
                    sum += Math.Pow(matr[j, i] - averValues[i], 2);
                }
                dispEstMatr[i] = sum / row;
            }
        }

        public void StandartizedMatrix(double[,] matr)
        {
            for (int i = 0; i < row; i++)
            {
                for (int j = 0; j < column; j++)
                {
                    standartMatr[i, j] = (matr[i, j] - averValues[j]) / Math.Sqrt(dispEstMatr[j]);
                }
            }
        }

        public void CovariationMatrix(double[,] matr)
        {
            for (int i = 0; i < column; i++)
            {
                for (int j = 0; j < column; j++)
                {
                    double sum = 0;
                    for (int k = 0; k < row; k++)
                    {
                        sum += (matr[k, i] - averValues[i]) * (matr[k, j] - averValues[j]);
                    }
                    covarMatr[i, j] = sum / row;
                }
            }
        }

        public void CorrelationMatrix()
        {
            for (int i = 0; i < column; i++)
            {
                for (int j = 0; j < column; j++)
                {
                    double sum = 0;
                    for (int k = 0; k < row; k++)
                    {
                        sum += standartMatr[k, i] * standartMatr[k, j];
                    }
                    corrMatr[i, j] = sum / row;
                }
            }
        }
    }

    //class of expoerting data from exel
    class ExelWork
    {
        int column;
        int row;
        double[,] matr;

        public ExelWork(int row, int column)
        {
            this.row = row;
            this.column = column;
            matr = new double[row, column];
        }

        public double[,] ExportFile()
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Title = "Выбор документа";
            fileDialog.DefaultExt = "*.xls;*.xlsx";

            if (!(fileDialog.ShowDialog() == DialogResult.OK))
                return matr;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var package = new ExcelPackage(fileDialog.FileName);
            ExcelWorksheet sheet = package.Workbook.Worksheets[0];

            for (int i = 0; i < row; i++)
            {
                for (int j = 0; j < column; j++)
                {
                    matr[i, j] = (double)sheet.Cells[i + 1, j + 1].Value;
                }
            }

            return matr;
        }
    }
}
