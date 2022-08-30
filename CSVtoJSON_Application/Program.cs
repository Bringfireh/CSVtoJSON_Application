using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using IronXL;

namespace CSVtoJSON_Application
{
    class Program
    {
        static void Main(string[] args)
        {
            //var csvFilereader = new DataTable();
            //csvFilereader = ReadExcel(csvFileName);
        }
        private DataTable ReadExcel(string fileName)
        {
            WorkBook workbook = WorkBook.Load(fileName);
            
            WorkSheet sheet = workbook.DefaultWorkSheet;
            
            return sheet.ToDataTable(true);
        }
        public void ReadCSVData(string csvFileName)
        {
            var csvFilereader = new DataTable();
            csvFilereader = ReadExcel(csvFileName);
        }
    }
}
