using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
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
            string basedir = AppDomain.CurrentDomain.BaseDirectory;
            string[] realpath = basedir.Split(new string[] { "bin", "Debug" }, StringSplitOptions.None);
            string apppath = realpath[0];

            //get the file name
            string csvFileName = Path.Combine(apppath + "Data\\CSV", "input.csv");

            //Read data into datatable;
            DataTable csvFileReader = new DataTable();
            WorkBook workbook = WorkBook.Load(csvFileName);

            WorkSheet sheet = workbook.DefaultWorkSheet;

             csvFileReader = sheet.ToDataTable(true);
        }
        //public DataTable ReadExcel(string fileName)
        //{
        //    WorkBook workbook = WorkBook.Load(fileName);
            
        //    WorkSheet sheet = workbook.DefaultWorkSheet;
            
        //    return sheet.ToDataTable(true);
        //}
        //public void ReadCSVData(string csvFileName)
        //{
        //    var csvFilereader = new DataTable();
        //    csvFilereader = ReadExcel(csvFileName);
        //}
    }
}
