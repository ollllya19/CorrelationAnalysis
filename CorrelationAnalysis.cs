using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace CorrelationAnalysis
{
    public class CorrelationAnalysis
    {
        int column;
        int row;
        double[,] matr;
        double[] averValues;
        double[] dispEstMatr;
        double[,] standartMatr;
        double[,] covarMatr;
        double[,] corrMatr;
        int[,] hyptmatr;
        const double tTable = 1.994954;

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
            hyptmatr = new int[column, column];
        }

        public double[,] StandartMatr { get => standartMatr; }

        public double[,] CovarMatr { get => covarMatr; }

        public double[,] CorrMatr { get => corrMatr; }

        public double[] DispMatr { get => dispEstMatr; }

        public int [,] HyptMatr { get => hyptmatr; }

        public void Execute()
        {
            ExelWork exelObj = new ExelWork(row, column);
            matr = exelObj.ExportFile();

            FindAverage(matr);
            FindVarianceEstimate(matr);
            StandartizedMatrix(matr);
            CorrelationMatrix();
            CovariationMatrix(matr);
            HypoteticMatr();
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

        public void HypoteticMatr()
        {
            for (int i = 0; i < column; i++)
            {
                for (int j = 0; j < column; j++)
                {
                    if (Math.Abs(corrMatr[i, j] * Math.Sqrt(row - 2)
                        / Math.Sqrt(1 - Math.Pow(corrMatr[i, j], 2))) < tTable)
                        hyptmatr[i, j] = 0;
                    else
                        hyptmatr[i, j] = 1;
                }
            }

        }
    }

    //class of exporting data from exel
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
