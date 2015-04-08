using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Collections;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data.OleDb;
using System.Data;



namespace CustomTools.ExtractXML
{
    /// <summary>
    /// Extract XML is used to take XML information inside excel workbook cells, and deposit the info into a collection of .xml files
    /// that the user selects. Created as a tool to use during a client project.
    /// </summary>
    /// <example>
    /// You have an excel spreadsheet. In column 1 (0-indexed) there's XML information. In column 3, there's a description of the module that 
    /// the xml information belongs to. When prompted, enter the name of the file (in same directory as the executable), then the number 1, 
    /// then the number 3. The extractor will put the xml files in folders with the same name as what's in the third column, as children
    /// of the folder "Extracted Files" (in the same directory as the executable).
    /// </example>
    /// <featuresneeded>
    /// I originally wrote part of the code to work with modern .xlsx files, and it should be fairly easy, but the need has not arisen.
    /// </featuresneeded>
    class ExtractXML
    {        
        static void Main(string[] args)
        {
            int columnWithCode;
            Console.Write("Enter Excel File Name. Must:\n");
            Console.Write("\t* Be In operating directory of the executable \n\t* A 97-03 files \n\t* Include extension\n");
            Console.Write("File Name: ");
            string xlsFileLocation = Console.ReadLine();
            Console.Write("Enter 0-indexed column number that contains data to be extracted:");
            Int32.TryParse(Console.ReadLine(), out columnWithCode);
            Console.Write("Enter 0-indexed column number that contains a folder differentiator");
            Console.Write("(or '-' for default folder):");
            string folderDifferentiator = Console.ReadLine();
            Console.Write("Enter Extension to write (or '-' to use .xml): ");
            string fileExtension = Console.ReadLine();
            if (fileExtension == "-") fileExtension = ".xml";

            if (!Directory.Exists("ExtractedFiles"))
            {
                //create default folder
                Directory.CreateDirectory("ExtractedFiles");
            }

            try
            {
                ReadExcelFile(xlsFileLocation,columnWithCode,folderDifferentiator,fileExtension);
            }
            catch (Exception e)
            {
                Console.Write("Operation Failed (Did you enter correct data?). Error: \n");
                Console.Write(e);
            }
            Console.Write("\nOperation Completed. Press Enter to quit.\n");
            Console.ReadLine();
        }
        

        /// <summary>
        /// Connects to an excel file, converts it into a table, and then writes the files into
        /// </summary>
        /// <param name="fileName">Name of the file if in .exe directory</param>
        /// <param name="columnWithCode">Column containing XML data in Woorkbook</param>
        /// <param name="folderDifferentiator"></param>
        /// <param name="fileExtension"></param>
        static void ReadExcelFile(string fileName, int columnWithCode, string folderDifferentiator, string fileExtension)
        {
            OleDbConnection conn;
            DataTable data;
            Console.Write("Loading Excel File...\n");
            fileName = ConnectToFile(fileName, out conn, out data);
            Console.Write("Loaded.\n");

            string folderName = "";
            int i = 0;
            Console.Write("Processing...\n");
            foreach(DataRow row in data.Rows) {
                if (i % (data.Rows.Count / 10) == 0) //10% progress marker
                {
                    Console.Write(" -+- ");
                }
                folderName = WriteFile(columnWithCode, folderDifferentiator, fileExtension, folderName, i, row);
                i++;
            }
            Console.Write("\nProcessed.\n");
            conn.Close();            
        }

        /// <summary>
        /// Connects to the excel file
        /// </summary>
        /// <param name="fileName">Name of the excel file in the current directory</param>
        /// <param name="conn"></param>
        /// <param name="data"></param>
        /// <returns>The name of the file</returns>
        private static string ConnectToFile(string fileName, out OleDbConnection conn, out DataTable data)
        {
            fileName = string.Format("{0}\\" + fileName, Directory.GetCurrentDirectory());
            var connectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0; data source={0}; Extended Properties=Excel 8.0;", fileName);

            conn = new OleDbConnection(connectionString);
            conn.Open();
            //get all sheet names
            DataTable sheetsName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "Table" });
            //Get the First Sheet Name
            string firstSheetName = sheetsName.Rows[0][2].ToString();
            //Query String 
            string sql = string.Format("SELECT * FROM [{0}]", firstSheetName);

            var adapter = new OleDbDataAdapter(sql, connectionString);
            var ds = new DataSet();
            adapter.Fill(ds, "1");
            data = ds.Tables["1"];//data is excel spreadsheet
            return fileName;
        }

        /// <summary>
        /// Writes the actual file that we're extracting from a spreadsheet.
        /// </summary>
        /// <param name="columnWithCode">The column containing file contents</param>
        /// <param name="folderDifferentiator">the value of the column making a folder name</param>
        /// <param name="fileExtension">Extension of files to write</param>
        /// <param name="folderName">Name of folder we're writing to</param>
        /// <param name="i">The row number that we're pulling from</param>
        /// <param name="row">The row with the information we need</param>
        /// <returns></returns>
        private static string WriteFile(int columnWithCode, string folderDifferentiator, string fileExtension, string folderName, int i, DataRow row)
        {
            if (folderDifferentiator == "-")//all in one folder
            {
                folderName = "ExtractedFiles";
                if (!Directory.Exists("ExtractedFiles\\ExtractedFiles"))
                {
                    //create default folder
                    Directory.CreateDirectory("ExtractedFiles\\ExtractedFiles");
                }
            }
            else //Using folder differentiator
            {
                folderName = row.ItemArray[Convert.ToInt32(folderDifferentiator)].ToString();
                //Create specific folder, if it does not exist
                if (!Directory.Exists("ExtractedFiles\\" + folderName))
                {
                    Directory.CreateDirectory("ExtractedFiles\\" + folderName);
                }
            }
            string text = row.ItemArray[columnWithCode].ToString();
            //determine filename from text
            string fileName = GetFileName(text);
            StreamWriter fileWriter = new StreamWriter("ExtractedFiles\\" + folderName + "\\" + fileName + i + fileExtension);
            fileWriter.Write(text);
            fileWriter.Close();
            return folderName;
        }


        /// <summary>
        /// This is a special purpose method. It uses the first child element of the xml code to determine a filename.
        /// </summary>
        /// <TODO>
        /// While this works for This project, this method should be optional or should allow the child to be defined through
        /// looping or recursion
        /// </TODO>
        /// <param name="text">The XML Code</param>
        /// <returns>The name of the first child in the XML Code</returns>
        private static string GetFileName(string text)
        {
            int index = text.IndexOf(">");
            string fileName = text.Substring(index+2);
            index = fileName.IndexOf(" ");
            fileName = fileName.Substring(0, index);
            return fileName;
        }
    }
}
