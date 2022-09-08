using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


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
            string csvFileName = Path.Combine(apppath + "Data\\CSV", "input.xlsx");

            ////Read data into datatable;
            DataTable csvFileReader = new DataTable();
            //WorkBook workbook = WorkBook.Load(csvFileName);

            //WorkSheet sheet = workbook.DefaultWorkSheet;

            // csvFileReader = sheet.ToDataTable(true);
            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + csvFileName + ";Extended Properties='Excel 12.0 Xml; HDR = YES; IMEX = 1'");
            OleDbCommand cmd = new OleDbCommand("select * from [input$]", conn);
            DataTable tbl = new DataTable();
            
            conn.Open();
            tbl.Load(cmd.ExecuteReader());
            conn.Close();
            csvFileReader = tbl;

            //Read all values into a List of csvfilefields
            List<CSVFileFields> csvFileFields = new List<CSVFileFields>();
            for(int i =0; i < csvFileReader.Rows.Count; i++)
            {
                csvFileFields.Add(new CSVFileFields { firstName = csvFileReader.Rows[i][0].ToString(), Surname = csvFileReader.Rows[i][1].ToString(), Username = csvFileReader.Rows[i][2].ToString(), Password=csvFileReader.Rows[i][3].ToString(), Roles = csvFileReader.Rows[i][4].ToString(), Groups = csvFileReader.Rows[i][5].ToString(), organisationUnits = csvFileReader.Rows[i][6].ToString(), OUCapture = csvFileReader.Rows[i][7].ToString() });
            }

            //Write the JSON Output file
            string jsonFileName = Path.Combine(apppath + "Data\\JSON", "output.txt");
            using (StreamWriter file = new StreamWriter(jsonFileName, append: false))
            {
                //file.WriteLine("[");
                int count = 0;
                foreach (var fields in csvFileFields)
                {
                    count++;
                    file.WriteLine("{");
                    string id = Guid.NewGuid().ToString().Substring(0, 12);
                    string idL ="  " + "\"id\"" + ":" + "\"" + id + "\"";
                    string firstName ="  "+ "\"firstName\"" + ": " + "\"" + fields.firstName + "\"";
                    string surname = "  " + "\"surname\"" + ": " + "\"" + fields.Surname + "\"";
                    string email = "  " + "\"email\"" + ": " + "\"" + fields.Email + "\"";
                    string userCredentials = "  " + "\"userCredentials\"" + ": " + "{" + "";
                    string userInfo ="    " + "\"userInfo\"" + ": { " + "\"id\""+ ": "+ "\"" + id + "\" }";
                    string username = "    " + "\"username\"" + ": " + "\"" + fields.Username + "\"";
                    string password = "    " + "\"password\"" + ": " + "\"" + fields.Password + "\"";
                    string userRolesS = "    " + "\"userRoles\"" + ": [ { " ;
                    string userRolesM = "      " + "\"id\"" + ": " + "\"" + fields.Roles + "\"";
                    string userRolesE = "    " + "} ]";
                    string uCE = "  " + "}";
                    string organisationUnitsS = "  " + "\"organisationUnits\"" + ": [ { ";
                    string organisationUnitsM = "    " + "\"id\"" + ": " + "\"" + fields.organisationUnits + "\"";
                    string organisationUnitsE = "  " + "} ]";
                    string userGroupsS = "  " + "\"userGroups\"" + ": [ { ";
                    string userGroupsM = "    " + "\"id\"" + ": " + "\"" + fields.Groups + "\"";
                    string userGroupsE = "  " + "} ]";


                    file.WriteLine(idL + ",");
                    file.WriteLine(firstName + ",");
                    file.WriteLine(surname + ",");
                    file.WriteLine(email + ",");
                    file.WriteLine(userCredentials);
                    file.WriteLine(userInfo + ",");
                    file.WriteLine(username + ",");
                    file.WriteLine(password + ",");
                    file.WriteLine(userRolesS);
                    file.WriteLine(userRolesM);
                    file.WriteLine(userRolesE);
                    file.WriteLine(uCE + ",");
                    file.WriteLine(organisationUnitsS);
                    file.WriteLine(organisationUnitsM);
                    file.WriteLine(organisationUnitsE + ",");
                    file.WriteLine(userGroupsS);
                    file.WriteLine(userGroupsM);
                    file.WriteLine(userGroupsE);

                    string d = (csvFileFields.Count() == count) ? "" : ",";
                    file.WriteLine("}" + d);
                    file.WriteLine(" ");
                }
                //file.WriteLine("]");
            }
        }
       
    }
}
