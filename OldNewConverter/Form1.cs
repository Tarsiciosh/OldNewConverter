﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Globalization;


namespace OldNewConverter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            SetupDataGridView();
        }

    
        public enum SearchType
        {
            FirstOcurrence = 0,
            LastOcurrence = 1
        }

        struct Station
        {
            public string name;
            public string ip;
            public string originPath;
            public string destinationPath;
            public string lastActivityDate;

            public Station(string name, string ip, string originPath, string destinationPath, string lastActivityDate)
            {
                this.name = name;
                this.ip = ip;
                this.originPath = originPath;
                this.destinationPath = destinationPath;
                this.lastActivityDate = lastActivityDate;
            }
        }

        static int maxStationNumber = 5;
        Station[] stations = new Station[maxStationNumber];

        private void startStopButton_Click(object sender, EventArgs e)
        {
            System.Timers.Timer myTimer = new System.Timers.Timer();

            if (startStopButton.Text == "Start")
            { 
                // send the "start" message to the running service 
            
                getStationInfo(); // (SERVICE START)

                foreach (Station station in stations) //creates the grid rows  (SERVICE START)
                {
                    string[] row = {station.name, station.ip};
                    stationsDataGridView.Rows.Add(row);
                }

                // setup the timer (SERVICE START)
                myTimer.Interval = 5000; // 1 second
                myTimer.Elapsed += new ElapsedEventHandler(this.OnTimer);
                myTimer.Start();

                // if all other steps where ok display "Started"
                statusLabel.Text = "Started";
                statusLabel.ForeColor = Color.Green;
                startStopButton.Text = "Stop";
            }
            else
            {
                // send "stop" message to the running service 
                // erase the content of the data grid view
                myTimer.Stop();

                int count = stationsDataGridView.RowCount;
                for (int i = 0; i < count; i++)
                {
                    stationsDataGridView.Rows.RemoveAt(0);
                }

                // if all other steps where ok display "Stopped" 
                statusLabel.Text = "Stopped";
                statusLabel.ForeColor = Color.Red;
                startStopButton.Text = "Start";
            }
        }

        private void getStationInfo()
        {
            System.IO.StreamReader configFile;
            string configString;
       
            string [] lines;
            string [] fields;
         
            configFile = System.IO.File.OpenText("C:\\OldNewGateway\\config\\stations.csv");
            configString = configFile.ReadToEnd();
            lines = configString.Split(new char[]{'\x0D','\x0A'}, StringSplitOptions.RemoveEmptyEntries);

            int i = -1;
            foreach (string line in lines)
            {
                if (i == -1) i = 0; else // skip the first line
                {
                    fields = line.Split(';');
                    stations[i].name = fields[0];
                    stations[i].ip = fields[1];
                    stations[i].originPath = fields[2];
                    stations[i].destinationPath = fields[3];
                    i++;
                }
            }
            configFile.Close(); 
        }

        private void OnTimer(object sender, ElapsedEventArgs args)
        {
            readAndWriteStationData();
            updateGridData();
        }
  
        private void readAndWriteStationData()
        {
            try
            {
                System.IO.StreamReader originFile;
                System.IO.StreamReader modelFile;
                System.IO.StreamWriter destinationFile;

                int i = 0;
                while (!String.IsNullOrEmpty(stations[i].name))
                {
                    IEnumerable<string> filePaths = System.IO.Directory.EnumerateFiles(stations[i].originPath, "*.json", System.IO.SearchOption.AllDirectories);
                    foreach (string originFilePath in filePaths) //all the .json files in that folder and subfolders
                    {
                        int index;
                        string result, prg, cycle, date, id, qc, row, column, step, Tmin, T, Tmax, Amin, A, Amax;
                        string originString, destinationString;
                        string destinationFilePath;

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
                        destinationString = destinationString.Insert(12 - 1, id);

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

                        destinationFile.Write(destinationString.ToCharArray());

                        destinationFile.Flush();

                        originFile.Close();
                        destinationFile.Close();

                        System.IO.File.Delete(originFilePath);

                        DateTime localDate = DateTime.Now;
                        var culture = new CultureInfo("en-GB");
                        stations[i].lastActivityDate = localDate.ToString(culture);

                    }
                    i++; if (i >= maxStationNumber) break;
                }
            }
            catch (Exception theException) //catch and report the error if there is any
            {
                string errorMessage;
                errorMessage = "Error:";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, "Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);
                //MessageBox.Show(errorMessage, "Error");
                System.Windows.Forms.Application.Exit();
            }          
        }

        private void updateGridData()
        {
            if (stationsDataGridView.RowCount != 0)
            {
                for (int i = 0; i < stations.Length; i++)
                {
                    stationsDataGridView.Rows[i].Cells[0].Value = stations[i].name;
                    stationsDataGridView.Rows[i].Cells[1].Value = stations[i].ip;
                    stationsDataGridView.Rows[i].Cells[2].Value = stations[i].lastActivityDate;
                }
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

            if (charArray[index] == '"') // STRING CASE!
            {
                i = 1; // offset of the quotation mark
                while (charArray[i + index] != '"')
                {
                    result = result.Insert(result.Length, charArray[i + index].ToString());
                    i++;
                }
            }
            else // NUMBER CASE!
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
                //if (n > s.Length) return null;
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
                //if (n < s.Length) return null;
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

                //originTextBox.Text = oSheet.Cells[2, 3].Value2; // then change the 2 with a variable to iterate al the file
                //destinationTextBox.Text = oSheet.Cells[2, 4].Value2;

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

        private void SetupDataGridView()
        {
            //stationsDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            stationsDataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill; //Tar
            stationsDataGridView.ColumnCount = 4;

            //stationsDataGridView.ColumnHeadersDefaultCellStyle.BackColor = Color.Red;
            //stationsDataGridView.ColumnHeadersDefaultCellStyle.ForeColor = Color.Red;
            stationsDataGridView.ColumnHeadersDefaultCellStyle.Font = new Font(stationsDataGridView.Font, FontStyle.Bold);

            //stationsDataGridView.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCellsExceptHeaders;
            
            //stationsDataGridView.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
            //stationsDataGridView.CellBorderStyle = DataGridViewCellBorderStyle.Single;

            stationsDataGridView.GridColor = Color.Black;
            stationsDataGridView.RowHeadersVisible = false;

            stationsDataGridView.Columns[0].Name = "Station Name";
            stationsDataGridView.Columns[1].Name = "IP Address";
            stationsDataGridView.Columns[2].Name = "Last Activity";
            stationsDataGridView.Columns[3].Name = "Status";

            stationsDataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            stationsDataGridView.MultiSelect = false;

            stationsDataGridView.AllowUserToAddRows = false;
            //stationsDataGridView.Dock = DockStyle.Fill;  

        }
    }
};








/*
 
     //System.Diagnostics.Debug.WriteLine($"readed lines: {lines.Length}");
           
                DataSet dataSet = new DataSet("Suppliers");
                dataSet.Tables[0].Rows[0][0] = "Hello";
                stationsDataGridView.DataSource = dataSet;
                stationsDataGridView.DataMember = "hola";
                
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


