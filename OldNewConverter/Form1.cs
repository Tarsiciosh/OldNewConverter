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

            /*
            public Station (string name, string ip, string originPath, string destinationPath)
            {
                this.name = name;
                this.ip = ip;
                this.originPath = originPath;
                this.destinationPath = destinationPath;
            }*/
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
                    string date, id, Tmin, T, Tmax, Amin, A, Amax, t, tmax;
                    string originFileString, modelFileString;

                    // READ ORIGIN FILE
                    originFile = System.IO.File.OpenText(originFilePath);
                    originFileString = originFile.ReadToEnd();

                    index = originFileString.IndexOf("\"prg date\":"); 
                    startIndex = originFileString.IndexOf(":", index); // (string , start index)
                    endIndex = originFileString.IndexOf(",", startIndex + 1);  
                    date = originFileString.Substring(startIndex + 3, endIndex - startIndex - 4); // string type with "" (start index, length)
                    date = date.Insert(11, "H "); //normalize data

                    index = originFileString.IndexOf("\"id code\":");
                    startIndex = originFileString.IndexOf(":", index); 
                    endIndex = originFileString.IndexOf(",", startIndex + 1);
                    id = originFileString.Substring(startIndex + 3, endIndex - startIndex - 4);
                    id = id + "_01"; //normalize data

                    index = originFileString.LastIndexOf("\"torque\":");
                    startIndex = originFileString.IndexOf(":", index);
                    endIndex = originFileString.IndexOf(",", startIndex + 1);
                    T = originFileString.Substring(startIndex + 2, endIndex - startIndex - 2); // number type without ""
                    //T = "7.200"; 
                    T = cutAndShift(T, 5); //normalize data

                    index = originFileString.LastIndexOf("\"angle\":");
                    startIndex = originFileString.IndexOf(":", index);
                    endIndex = originFileString.IndexOf(",", startIndex + 1);
                    A = originFileString.Substring(startIndex + 2, endIndex - startIndex - 2);
                    A = cutAndShift(A, 8); // normalize data

                    index = originFileString.LastIndexOf("\"duration\":");
                    startIndex = originFileString.IndexOf(":", index);
                    endIndex = originFileString.IndexOf(",", startIndex + 1);
                    t = originFileString.Substring(startIndex + 2, endIndex - startIndex - 2);
                    t = cutAndShift(t, 6); // normalize data

                    // MessageBox.Show("date: " + date);
                    // MessageBox.Show("id: " + id);
                    MessageBox.Show("T: " + T);
                    // MessageBox.Show("A: " + A);

                    // READ MODEL FILE
                    modelFile = System.IO.File.OpenText("C:\\OldNewGateway\\file models\\model.txt");
                    modelFileString = modelFile.ReadToEnd(); // read as string

                    modelFileString = modelFileString.Insert(12-1, id);

                    index = modelFileString.IndexOf('\x0A'); // first line feed

                    index = modelFileString.IndexOf('\x0A', index + 1); // next line feed    
                    modelFileString = modelFileString.Insert(index + 3, date); //to-do: change insert with "replace"

                    index = modelFileString.IndexOf('\x0A', index + 1); // next line feed
                    modelFileString = modelFileString.Insert(index + 6, T);


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

                    destinationBuffer = modelFileString.ToCharArray(); // convert to char array    
                    destinationFile.Write(destinationBuffer);

                    destinationFile.Flush();

                    originFile.Close();
                    destinationFile.Close();

                    //System.IO.File.Delete(originFilePath); UNCOMMENT IN REAL SCENARIO
                }
                i++; if (i >= maxStationNumber) break;
            }       
        }

        private string cutAndShift(string s, int n)
        {
            try
            {
                if (n > s.Length) return null;
                 
                s = s.Substring(0, n); // first cut

                bool stillSearch = true;
                char[] charArray = s.ToCharArray();
                for (int i = s.Length - 1; i >= 0; i--)
                {
                    if (charArray[i] == '0' && stillSearch == true)
                    {
                        s = s.Insert(0, " ");
                    }
                    else
                    {
                        stillSearch = false;
                    }
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


