using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace parser
{
    class ExcelHelper : IDisposable
    {
        private Excel.Application _excel;
        private Workbook _workbook;
        private string _filePath;

        public ExcelHelper()
        {
            _excel = new Excel.Application();
        }

        public void Dispose()
        {
            try
            {
                _workbook.Close();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message,"Вышло плохо с закрытием");
            }
        }

        internal bool Open(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    _workbook = _excel.Workbooks.Open(filePath);
                    _filePath = filePath;
                    
                }
               
                else
                {
                    _workbook = _excel.Workbooks.Add();
                    _filePath = filePath;
                }
                
                return true; 
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,"Вышло плохо с открытием");
                return false;
            }
        }

        internal void Set(string column, int row, string data)
        {
            try
            {
                ((Excel.Worksheet)_excel.ActiveSheet).Cells[row, column] = data;
            }
            catch
            {
                MessageBox.Show("Вышло плохо в Сете");
            }
        }

        internal void Save()
        {
            if (!string.IsNullOrEmpty(_filePath))
            {
                _workbook.SaveAs(_filePath);
            }
            else
            {
                _workbook.SaveAs(_filePath);
            }
        }
    }
}
