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

            System.IO.StreamReader originFile;
            System.IO.StreamReader modelFile;
            System.IO.StreamWriter destinationFile;
            
            char[] originBuffer = new char[100000];
            char[] modelBuffer = new char[100000];
            char[] destinationBuffer = new char[10];

            IEnumerable<string> filePaths = System.IO.Directory.EnumerateFiles(originFolderString, "*.json", System.IO.SearchOption.AllDirectories);

            foreach (string originFilePath in filePaths) //all the .json files in that folder and subfolders
            {
                int index, startIndex, endIndex;
                string date, Tmin, T, Tmax, Amin, A, Amax, t, tmax;

                // open and read data from the origin file
                originFile = System.IO.File.OpenText(originFilePath); // open 
                originFile.Read(originBuffer, 0, originBuffer.Length - 1); // read
                string originFileString = new string(originBuffer); // convert to string 

                index = originFileString.IndexOf("\"prg date\":");
                startIndex = originFileString.IndexOf(":", index);
                endIndex = originFileString.IndexOf(",", startIndex+1);
                date = new String(originBuffer, startIndex + 3, endIndex - startIndex - 4);
                //T = T.Substring(0, 5); //normalize data

                index = originFileString.LastIndexOf("\"torque\":");
                startIndex = originFileString.IndexOf(":", index);
                endIndex = originFileString.IndexOf(",", startIndex+1);       
                T = new String(originBuffer, startIndex + 2, endIndex - startIndex -2);
                T = T.Substring(0, 5); //normalize data
               
                index = originFileString.LastIndexOf("\"angle\":");
                startIndex = originFileString.IndexOf(":", index);
                endIndex = originFileString.IndexOf(",", startIndex+1);
                A = new String(originBuffer, startIndex + 2, endIndex - startIndex - 2);
                A = A.Substring(0, 8); // normalize data

                index = originFileString.LastIndexOf("\"duration\":");
                startIndex = originFileString.IndexOf(":", index);
                endIndex = originFileString.IndexOf(",", startIndex+1);
                t = new String(originBuffer, startIndex + 2, endIndex - startIndex - 2);
                t = t.Substring(0, 6); // normalize data


                //MessageBox.Show("torque: " + torque);
                //MessageBox.Show("angle: " + angle);

                // reead model fiel and converted into a string
                modelFile = System.IO.File.OpenText("C:\\OldNewGateway\\file models\\model1.txt"); // open
                modelFile.Read(modelBuffer, 0, modelBuffer.Length - 1); // read
                
                /*
                byte myByte = 0x02; modelBuffer[0] = myByte.ToString().ToCharArray()[0]; //STX
                myByte = 0x20; modelBuffer[1] = myByte.ToString().ToCharArray()[0]; //_
                myByte = 0x0E; modelBuffer[2] = myByte.ToString().ToCharArray()[0]; //SO
                myByte = 0x0D; modelBuffer[3] = myByte.ToString().ToCharArray()[0]; //CR
                myByte = 0x0A; modelBuffer[4] = myByte.ToString().ToCharArray()[0]; //CR
                */

                string modelString = new string(modelBuffer); // convert to string

                // Add data to the model string - all data
                
                modelString = modelString.Insert(20, T);
                modelString = modelString.Insert(34, A);
                modelString = modelString.Insert(43, t);
                modelString = modelString.Insert(52, date);

                //MessageBox.Show(modelString);

                // Create file and copy the information from the buffer

                destinationFilePath = System.IO.Path.Combine(destinationFolderString, "test-result.txt");
                destinationFile = System.IO.File.CreateText(destinationFilePath);

                destinationBuffer = modelString.ToCharArray();

                destinationFile.Write(destinationBuffer);

                destinationFile.Flush();

                originFile.Close();
                destinationFile.Close();

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
                //to read the path of the file from the text box
                //string path = configPathTextBox.Text;
                //path = path.Replace(@"\", @"\\");

                // to hardcode the path to a string
                string path = "C:\\OldNewGateway\\config\\stations.xlsx";

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

     // To test the funtion with strings
                //destinationFile.Write("torque= " + torqueString + "  angle= " + angleString);
           */
