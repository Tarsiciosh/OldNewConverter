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

        struct Station
        {
            public string name;
            public string ip;
            public string originPath;
            public string destinationPath;
      
            public Station (string name, string ip, string originPath, string destinationPath)
            {
                this.name = name;
                this.ip = ip;
                this.originPath = originPath;
                this.destinationPath = destinationPath;
            }
        }

        public enum SearchType
        {
            FirstOcurrence = 0,
            LastOcurrence = 1
        }

        private void startStopButton_Click(object sender, EventArgs e)
        {

            //in thfe real implementation this information comes from the excel file
            //string originFolderString = "C:\\FTP Server Results\\Origin"; //originFolderTextBox.Text;
            //string destinationFolderString = "C:\\FTP Server Results\\Destination"; //destinationFolderTextBox.Text;
            string destinationFilePath;

            System.IO.StreamReader originFile;
            System.IO.StreamReader modelFile;
            System.IO.StreamWriter destinationFile;

            char[] destinationBuffer;

            int maxStationNumber = 5;
            Station[] stations = new Station[maxStationNumber];

            //stations = readExcelFile(); // 

            stations[0].name = "Prueba";
            stations[0].ip = "ip prueba";
            stations[0].originPath = originTextBox.Text; // "C:\\FTP Server Results\\Origin";
            stations[0].destinationPath = destinationTextBox.Text;  //"C:\\FTP Server Results\\Destination";

            /*
            stations[1].name = "OP200";
            stations[1].ip = "172.16.1.24";
            stations[1].originPath = "C:\\FTP Server Results\\Origin_2";
            stations[1].destinationPath = "C:\\FTP Server Results\\Destination_2";
            */
            int i = 0;
            while ( ! String.IsNullOrEmpty (stations[i].name))
            {
                IEnumerable<string> filePaths = System.IO.Directory.EnumerateFiles(stations[i].originPath, "*.json", System.IO.SearchOption.AllDirectories);
                foreach (string originFilePath in filePaths) //all the .json files in that folder and subfolders
                {
                    int index;
                    string result, prg, cycle, date, id, qc, row, column, step, sb, Tmin, T, Tmax, Amin, A, Amax;
                    string originString, destinationString;

                    // READ ORIGIN FILE
                    originFile = System.IO.File.OpenText(originFilePath);
                    originString = originFile.ReadToEnd();
       
                    result = getData(originString, "result", 0, SearchType.FirstOcurrence);

                    prg = getData(originString, "prg nr", 0, SearchType.FirstOcurrence);
                    prg = expandAndShift(prg, 2);

                    cycle = getData(originString, "cycle", 0, SearchType.FirstOcurrence);
                    cycle = expandAndShift(cycle, 7);

                    date = getData(originString, "date", 0, SearchType.FirstOcurrence);
                    date = date.Insert(11, "H ");
                    
                    id = getData(originString, "id code", 0, SearchType.FirstOcurrence);
                    id = id + "_xxx";

                    qc = getData(originString, "quality code", 0, SearchType.FirstOcurrence);
                    qc = expandAndShift(qc, 3);

                    // ... last result

                    row = getData(originString, "row", 0, SearchType.LastOcurrence);
                    row = expandAndShift(row, 2);
                    column = getData(originString, "column", 0, SearchType.LastOcurrence);
                    step = row.Insert(row.Length, column);

                    T = getData(originString, "torque", 0, SearchType.LastOcurrence);
                    T = cutAndShift(T, 5);

                    A = getData(originString, "angle", 0, SearchType.LastOcurrence);
                    A = cutAndShift(A, 8);

                    index = originString.LastIndexOf("MF TorqueMin");
                    Tmin = getData(originString, "nom", index, SearchType.FirstOcurrence);
                    Tmin = cutAndShift(Tmin, 5);
                  
                    index = originString.LastIndexOf("MFs TorqueMax");
                    Tmax = getData(originString, "nom", index, SearchType.FirstOcurrence);
                    Tmax = cutAndShift(Tmax, 5);

                    index = originString.LastIndexOf("MF AngleMin");
                    Amin = getData(originString, "nom", index, SearchType.FirstOcurrence);
                    Amin = expandAndShift(Amin, 8);

                    index = originString.LastIndexOf("MFs AngleMax");
                    Amax = getData(originString, "nom", index, SearchType.FirstOcurrence);
                    Amax = expandAndShift(Amax, 8);
                  
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
                    destinationString = destinationString.Insert(index + 6, "     "); // MR: 5 spaces 
                    destinationString = destinationString.Insert(index + 14, "        "); // WR: 8 spaces
                    destinationString = destinationString.Insert(index + 26, " 0"); // QR: " 0"

                    index = destinationString.IndexOf('\x0A', index + 1); // angle limits
                    destinationString = destinationString.Insert(index + 3, Amin); 
                    destinationString = destinationString.Insert(index + 14, Amax); 

                    index = destinationString.IndexOf('\x0A', index + 1); // torque limits
                    destinationString = destinationString.Insert(index + 6, Tmin);
                    destinationString = destinationString.Insert(index + 17, Tmax);

                    index = destinationString.IndexOf('\x0A', index + 1); // gradient limits
                    destinationString = destinationString.Insert(index + 5, "      ");
                    destinationString = destinationString.Insert(index + 17, "     ");

                    index = destinationString.IndexOf('\x0A', index + 1); // step, quality code, stopped by
                    destinationString = destinationString.Insert(index + 3, step);
                    destinationString = destinationString.Insert(index + 11, qc);
                    destinationString = destinationString.Insert(index + 18, " 3");

                    index = destinationString.IndexOf('\x0A', index + 1); // consecutive no. and program no.
                    destinationString = destinationString.Insert(index + 3, cycle);
                    destinationString = destinationString.Insert(index + 13, prg);

                    index = destinationString.IndexOf('\x0A', index + 1); // hardware ID and channel no.

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


        private string getData(string source, string name, int fromIndex, SearchType t)
        {
            int index = 0, i;
            string result = "";

            char[] charArray = source.ToCharArray();


            if (t == SearchType.FirstOcurrence)
                index = source.IndexOf("\"" + name + "\":", fromIndex);

            if (t == SearchType.LastOcurrence)
                index = source.LastIndexOf("\"" + name + "\":");

            index = index + name.Length + 4; // two quotation marks, one colon and a space

            if (charArray[index] == '"') // string case
            {
                i = 1; // offset of the quotation mark
                while (charArray[i + index] != '"')
                {
                    result = result.Insert(result.Length, charArray[i + index].ToString());
                    i++;
                }
            }
            else // number case
            {
                i = 0; // no offset
                while (charArray[i + index] != ',' && charArray[i + index] != ' ')
                {
                    result = result.Insert(result.Length, charArray[i + index].ToString());
                    i++;
                }
            }
            return result;
        }

        private string cutAndShift(string s, int n)
        {
            try
            {
                if (n > s.Length) return null;
                int indexOfPoint = s.IndexOf(".");
                char[] charArray = s.ToCharArray();
                for (int i = 0; i < (n - indexOfPoint - 3); i++) // round in 2 decimals
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

        private void readExcelFile() // todo - resturn an array of stations (readed from each row)
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

                originTextBox.Text = oSheet.Cells[2, 3].Value2; // then change the 2 with a variable to iterate al the file
                destinationTextBox.Text = oSheet.Cells[2, 4].Value2;

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

  
    }
};




/*
 
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


     private string getLastData(string source, string name, DataType t)
        {
            int index, startIndex, endIndex, endIndexComma, endIndexSpace;
            string result = "";

            index = source.LastIndexOf("\"" + name + "\":");
            startIndex = source.IndexOf(":", index); // (string , start index)
   
            endIndex = source.IndexOf(",", startIndex + 1);
            

            if (t == DataType.Text)
                result = source.Substring(startIndex + 3, endIndex - startIndex - 4); // string type with "" 

            if (t == DataType.Number)
                result = source.Substring(startIndex + 2, endIndex - startIndex - 2); // number type without ""

            return result;
        }

  
   private string getData(string source, string name, DataType t)
        {
            int index, startIndex, endIndex, endIndexComma, endIndexSpace ;
            string result = "";
            
            index = source.IndexOf("\"" + name + "\":");
            startIndex = source.IndexOf(":", index); // (string , start index)

            endIndex = source.IndexOf(",", startIndex + 1);
           
            endIndexComma = source.IndexOf(",", startIndex + 1);
            endIndexSpace = source.IndexOf(" ", startIndex + 1);
            if (endIndexSpace < endIndexComma)
                endIndex = endIndexSpace;
            else
                endIndex = endIndexComma;
          

            if (t == DataType.Text) 
                result = source.Substring(startIndex + 3, endIndex - startIndex - 4); // string type with "" 

            if (t == DataType.Number)
                result = source.Substring(startIndex + 2, endIndex - startIndex - 2); // number type without ""

            return result;
        }
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


