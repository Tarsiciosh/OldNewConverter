using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace OldNewConverter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void readTxtFileButton_Click(object sender, EventArgs e)
        {
            //in the real implementation this information comes from the excel file
            string originFolderString = originFolderTextBox.Text;
            string destinationFolderString = destinationFolderTextBox.Text;
            string destinationFilePath;

            System.IO.StreamReader originStreamReader;
            System.IO.StreamWriter destinationStreamWriter;

            char[] originBuffer = new char[100000];

            IEnumerable<string> filePaths = System.IO.Directory.EnumerateFiles(originFolderString, "*.json", System.IO.SearchOption.AllDirectories);

            foreach (string originFilePath in filePaths) //all the .json files in that folder and subfolders
            {
                int index, startIndex, endIndex;
                string torqueString, angleString;

                // open and read data from the origin file
                originStreamReader = System.IO.File.OpenText(originFilePath); //open file
                originStreamReader.Read(originBuffer, 0, originBuffer.Length - 1); // read file
                string originFileString = new string(originBuffer); // convert to string 

                index = originFileString.LastIndexOf("\"torque\":");
                startIndex = originFileString.IndexOf(":", index);
                endIndex = originFileString.IndexOf(",", index);       
                torqueString = new String(originBuffer, startIndex + 2, endIndex - startIndex -2);

                index = originFileString.LastIndexOf("\"angle\":");
                angleString = new String(originBuffer, index + 9, 11);

                MessageBox.Show(torqueString);
                MessageBox.Show(angleString);

                // Writes data to the destination file
                destinationFilePath = System.IO.Path.Combine(destinationFolderString, "teste.txt");
                destinationStreamWriter = System.IO.File.CreateText(destinationFilePath);

                destinationStreamWriter.Write("torque= " + torqueString + "  angle= " + angleString);
                destinationStreamWriter.Flush();

                originStreamReader.Close();
                destinationStreamWriter.Close();

                System.IO.File.Delete(originFilePath);
            }
        }

        private void readExcelFileButton_Click(object sender, EventArgs e)
        {
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;
            //Excel.Range oRng;

            try
            {
                // to read the path of the file from the text box
                string path = configPathTextBox.Text;
                path = path.Replace(@"\", @"\\");

                // to hardcode the path to a string
                path = "C:\\OldNewGateway\\config\\stations.xlsx";

                //start excel and get application object.
                oXL = new Excel.Application();
                oXL.Visible = true;

                //open a workbook.
                oWB = (Excel._Workbook)(oXL.Workbooks.Open(@path));
                oSheet = (Excel._Worksheet)oWB.ActiveSheet; // .Sheets.Item()

                oSheet.Cells[1, 2] = "Station Name";

                //originFolderTextBox.Text = oSheet.get_Range("C2", "C2").Value2;
                //destinationFolderTextBox.Text = oSheet.get_Range("D2", "D2").Value2;

                originFolderTextBox.Text = oSheet.Cells[2,3].Value2; // then change the 2 with a variable to iterate al the file
                destinationFolderTextBox.Text = oSheet.Cells[2,4].Value2;

                oXL.Visible = false;
                oXL.UserControl = false;

                oWB.Save();
                oWB.Close();

            } catch (Exception theException) //catch and report the error if there is any
            {
                string errorMessage;
                errorMessage = "Error:";
                errorMessage = String.Concat( errorMessage, theException.Message);
                errorMessage = String.Concat( errorMessage, "Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");
            }

        }
    }
}



/*
           string configPath = configPathTextBox.Text;  //string configPath = "C:\OldNewGateway\config";
           string stationID, stationName, originPath, destinationPath;

           int startIndex, endIndex;
           string delimiter = "*";

           char[] configBuffer = new char[1024];
           System.IO.StreamReader configStreamReader;

           configStreamReader = System.IO.File.OpenText(configPath);
           configStreamReader.Read(configBuffer, 0, configBuffer.Length - 1);

           string configString = new string(configBuffer);

           endIndex = 0;


           for (int i=0; i<3; i++) {
               startIndex = endIndex;
               endIndex = configString.IndexOf(delimiter, startIndex + 1); 
               stationID = new string(configBuffer, startIndex + 1, endIndex - startIndex - 1); 
               MessageBox.Show("station ID: " + stationID); 

               startIndex = endIndex;
               endIndex = configString.IndexOf(delimiter, startIndex + 1);
               stationName = new string(configBuffer, startIndex + 1, endIndex - startIndex - 1); 
               MessageBox.Show("name: " + stationName);

               startIndex = endIndex;
               endIndex = configString.IndexOf(delimiter, startIndex + 1);
               originPath = new string(configBuffer, startIndex + 1, endIndex - startIndex - 1); 
               MessageBox.Show("origin: " + originPath);

               startIndex = endIndex;
               endIndex = configString.IndexOf(delimiter, startIndex + 1);
               destinationPath = new string(configBuffer, startIndex + 1, endIndex - startIndex - 1); 
               MessageBox.Show("destination: " + destinationPath);
           } 
           */
