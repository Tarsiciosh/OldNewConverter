﻿using System;
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

        struct Station
        {
            public string name;
            public string ip;
            public string originPath;
            public string destinationPath;

            /*
            public Station (string name, string ip, string originPath, string destinationPath)
            {
                this.name = name;
                this.ip = ip;
                this.originPath = originPath;
                this.destinationPath = destinationPath;
            }*/
        }

        public enum DataType
        {
            Text = 0,
            Number = 1, 
        }

        private void readTxtFileButton_Click(object sender, EventArgs e)
        {

            //in the real implementation this information comes from the excel file
            //string originFolderString = "C:\\FTP Server Results\\Origin"; //originFolderTextBox.Text;
            //string destinationFolderString = "C:\\FTP Server Results\\Destination"; //destinationFolderTextBox.Text;
            string destinationFilePath;

            System.IO.StreamReader originFile;
            System.IO.StreamReader modelFile;
            System.IO.StreamWriter destinationFile;

            char[] destinationBuffer;

            int maxStationNumber = 5;
            Station[] stations = new Station[maxStationNumber];


            readExcelFile();

            stations[0].name = "OP100";
            stations[0].ip = "172.16.1.23";
            stations[0].originPath = "C:\\FTP Server Results\\Origin";
            stations[0].destinationPath = "C:\\FTP Server Results\\Destination";

            stations[1].name = "OP200";
            stations[1].ip = "172.16.1.24";
            stations[1].originPath = "C:\\FTP Server Results\\Origin_2";
            stations[1].destinationPath = "C:\\FTP Server Results\\Destination_2";

            int i = 0;
            while ( ! String.IsNullOrEmpty (stations[i].name))
            {
                IEnumerable<string> filePaths = System.IO.Directory.EnumerateFiles(stations[i].originPath, "*.json", System.IO.SearchOption.AllDirectories);
                foreach (string originFilePath in filePaths) //all the .json files in that folder and subfolders
                {
                    int index, startIndex, endIndex;
                    string result, prg, cycle, date, id, qc, Tmin, T, Tmax, Amin, A, Amax, t, tmax;
                    string originString, destinationString;

                    // READ ORIGIN FILE
                    originFile = System.IO.File.OpenText(originFilePath);
                    originString = originFile.ReadToEnd();

                    result = getData(originString, "result", DataType.Text);

                    prg = getData(originString, "prg nr", DataType.Number);
                    prg = expandAndShift(prg, 2); 

                    cycle = getData(originString, "cycle", DataType.Number);
                    cycle = expandAndShift(cycle, 7);

                    date = getData(originString, "date", DataType.Text);
                    date = date.Insert(11, "H ");

                    id = getData(originString, "id code", DataType.Text);
                    id = id + "_xxx"; 

                    // qc

                    // last result

                    // row

                    // column    

                    T = getLastData(originString, "torque", DataType.Number);
                    T = cutAndShift(T, 5); 

                    A = getLastData(originString, "angle", DataType.Number);
                    A = cutAndShift(A, 8);

                    Tmin = getLastDataWithSubname(originString, "MF TorqueMin", "act", DataType.Number);
                    Tmin = cutAndShift(Tmin,5);

                    Tmax = "     ";

                    Amin = "        ";

                    Amax = "        ";

                    // READ MODEL FILE
                    modelFile = System.IO.File.OpenText("C:\\OldNewGateway\\file models\\model.txt");
                    destinationString = modelFile.ReadToEnd(); // read as string

                    //ID code souce and ID code
                    destinationString = destinationString.Insert(12-1, id); 

                    index = destinationString.IndexOf('\x0A'); 

                    index = destinationString.IndexOf('\x0A', index + 1); // date, time    
                    destinationString = destinationString.Insert(index + 3, date); 

                    index = destinationString.IndexOf('\x0A', index + 1); // measured values with result
                    destinationString = destinationString.Insert(index + 6, T);
                    destinationString = destinationString.Insert(index + 14, A);
                    destinationString = destinationString.Insert(index + 28, "     "); // G gradient
                    destinationString = destinationString.Insert(index + 34, result);

                    index = destinationString.IndexOf('\x0A', index + 1); // redundancy values (optional)
                    destinationString = destinationString.Insert(index + 6, "     "); // 5 spaces 
                    destinationString = destinationString.Insert(index + 14, "        "); // 8 spaces
                    destinationString = destinationString.Insert(index + 26, " 0"); // 2 spaces

                    index = destinationString.IndexOf('\x0A', index + 1); // angle limits
                    destinationString = destinationString.Insert(index + 3, Amin); 
                    destinationString = destinationString.Insert(index + 14, Amax); 

                    index = destinationString.IndexOf('\x0A', index + 1); // torque limits
                    destinationString = destinationString.Insert(index + 6, Tmin);
                    destinationString = destinationString.Insert(index + 17, Tmax);


                    index = destinationString.IndexOf('\x0A', index + 1); // gradient limits

                    index = destinationString.IndexOf('\x0A', index + 1); // step, quality code, stopped by

                    index = destinationString.IndexOf('\x0A', index + 1); // consecutive no. and program no.
                    destinationString = destinationString.Insert(index + 3, cycle);
                    destinationString = destinationString.Insert(index + 13, prg);


                    index = destinationString.IndexOf('\x0A', index + 1); // hardware ID and channel no.

                    // Add data to the model buffer
                    // modelBuffer[1] = 0x20;

                    //MessageBox.Show($"len: {modelFileString.Length.ToString()}");

                    //byte stx = 0x03;    
                    //destinationBuffer[0] = Convert.ToChar(stx);

                    //destinationBuffer[0] = '\x02'; // write direct byte type values

                    //modelBuffer[2] = '\x03';
                    //Add data to the model string - all data
                    //modelString = modelString.Insert(20, T);
                    //modelString = modelString.Insert(34, A);
                    //modelString = modelString.Insert(43, t); 
                    //modelString = modelString.Insert(12-1, id);
                    //MessageBox.Show(modelString);

                    // Create file and copy the information from the buffer



                    destinationFilePath = System.IO.Path.Combine(stations[i].destinationPath, "test-result.txt");
                    destinationFile = System.IO.File.CreateText(destinationFilePath);

                    destinationBuffer = destinationString.ToCharArray(); // convert to char array    
                    destinationFile.Write(destinationBuffer);

                    destinationFile.Flush();

                    originFile.Close();
                    destinationFile.Close();

                    //System.IO.File.Delete(originFilePath); UNCOMMENT IN REAL SCENARIO
                }
                i++; if (i >= maxStationNumber) break;
            }       
        }


        private string getData(string source, string name, DataType t)
        {
            int index, startIndex, endIndex, endIndexComma, endIndexSpace ;
            string result = "";
            
            index = source.IndexOf("\"" + name + "\":");
            startIndex = source.IndexOf(":", index); // (string , start index)

            endIndex = source.IndexOf(",", startIndex + 1);
            /*
            endIndexComma = source.IndexOf(",", startIndex + 1);
            endIndexSpace = source.IndexOf(" ", startIndex + 1);
            if (endIndexSpace < endIndexComma)
                endIndex = endIndexSpace;
            else
                endIndex = endIndexComma;
            */

            if (t == DataType.Text) 
                result = source.Substring(startIndex + 3, endIndex - startIndex - 4); // string type with "" 

            if (t == DataType.Number)
                result = source.Substring(startIndex + 2, endIndex - startIndex - 2); // number type without ""

            return result;
        }

        private string getLastData(string source, string name, DataType t)
        {
            int index, startIndex, endIndex, endIndexComma, endIndexSpace;
            string result = "";

            index = source.LastIndexOf("\"" + name + "\":");
            startIndex = source.IndexOf(":", index); // (string , start index)
   
            endIndex = source.IndexOf(",", index);
            /*
            endIndexComma = source.IndexOf(",", startIndex + 1);
            endIndexSpace = source.IndexOf(" ", startIndex + 1);
            if (endIndexSpace < endIndexComma)
                endIndex = endIndexSpace;
            else
                endIndex = endIndexComma;
            */

            if (t == DataType.Text)
                result = source.Substring(startIndex + 3, endIndex - startIndex - 4); // string type with "" 

            if (t == DataType.Number)
                result = source.Substring(startIndex + 2, endIndex - startIndex - 2); // number type without ""

            return result;
        }

        private string getLastDataWithSubname(string source, string name, string subName,  DataType t)
        {
            int index, startIndex, endIndex, endIndexComma, endIndexSpace;
            string result = "";

            index = source.LastIndexOf("\"" + name + "\":");
            index = source.IndexOf("\"" + subName + "\":", index);
            startIndex = source.IndexOf(":", index); // (string , start index)

            endIndex = source.IndexOf(" ", startIndex + 1);
            /*
            endIndexComma = source.IndexOf(",", startIndex + 1);
            endIndexSpace = source.IndexOf(" ", startIndex + 1);
            if (endIndexSpace < endIndexComma)
                endIndex = endIndexSpace;
            else
                endIndex = endIndexComma;
            */

            if (t == DataType.Text)
                result = source.Substring(startIndex + 3, endIndex - startIndex - 4); // string type with "" 

            if (t == DataType.Number)
                result = source.Substring(startIndex + 2, endIndex - startIndex - 2); // number type without ""

            return result;
        }


        private string cutAndShift(string s, int n)
        {
            try
            {
                if (n > s.Length) return null;
                int indexOfPoint = s.IndexOf(".");
                char[] charArray = s.ToCharArray();
                for (int i = 0; i < (n - indexOfPoint - 3) ; i++) // round in 2 decimals
                {  
                  s = s.Insert(0, " ");   
                }
                s = s.Substring(0, n); // last cut
                return s;
            }
            catch (Exception theException) //catch and report the error if there is any
            {
                string errorMessage;
                errorMessage = "Error:";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, "Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");
                return null;
            }
        }

        private string expandAndShift(string s, int n)
        {
            try
            {
                if (n < s.Length) return null;

       
                char[] charArray = s.ToCharArray();

                int len = s.Length;

                for (int i = 0; i < (n - len) ; i++)
                {
                    s = s.Insert(0, " ");
                }
                return s;
            }

            catch (Exception theException) //catch and report the error if there is any
            {
                string errorMessage;
                errorMessage = "Error:";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, "Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");
                return null;
            }
        }


        private void readExcelFile()
        {
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;
            //Excel.Range oRng;

            try
            {
                // hardcode the path of the file
                string path = "C:\\OldNewGateway\\config\\stations.xlsx";

                //start excel and get application object.
                oXL = new Excel.Application();
                oXL.Visible = true;

                //open a workbook.
                oWB = (Excel._Workbook)(oXL.Workbooks.Open(@path));
                oSheet = (Excel._Worksheet)oWB.ActiveSheet; // .Sheets.Item()

                oSheet.Cells[1, 2] = "Station Name";

                originFolderTextBox.Text = oSheet.Cells[2, 3].Value2; // then change the 2 with a variable to iterate al the file
                destinationFolderTextBox.Text = oSheet.Cells[2, 4].Value2;

                oXL.Visible = false;
                oXL.UserControl = false;

                oWB.Save();
                oWB.Close();

            }
            catch (Exception theException) //catch and report the error if there is any
            {
                string errorMessage;
                errorMessage = "Error:";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, "Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");
            }
        }

        private void readExcelFileButton_Click(object sender, EventArgs e)
        {
        }
    }
};




/*
    destinationFile.Write(destinationBuffer, 0, 10); //(index,count)
    modelFile.Read(modelBuffer, 0, modelBuffer.Length - 1); // read
    string modelString = new string(modelBuffer); // convert to string

    originFile.Read(originBuffer, 0, originBuffer.Length - 1); // read
    string originFileString = new string(originBuffer); // convert to string 


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
           
       //to read the path of the file from the text box
                //string path = configPathTextBox.Text;
                //path = path.Replace(@"\", @"\\");
     
     
                byte myByte = 0x02; modelBuffer[0] = myByte.ToString().ToCharArray()[0]; //STX
                myByte = 0x20; modelBuffer[1] = myByte.ToString().ToCharArray()[0]; //_
                myByte = 0x0E; modelBuffer[2] = myByte.ToString().ToCharArray()[0]; //SO
                myByte = 0x0D; modelBuffer[3] = myByte.ToString().ToCharArray()[0]; //CR
                myByte = 0x0A; modelBuffer[4] = myByte.ToString().ToCharArray()[0]; //CR
*/


