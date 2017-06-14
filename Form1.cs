using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Drawing;
using System.Threading;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Data;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Drawing.Imaging;

namespace imageeditor
{
    public partial class Form1 : Form
    {
        string folder = string.Empty;
        string disptext = string.Empty;
        string[] files;
        string[] filecount;
        System.Data.DataTable excelTable = new System.Data.DataTable();
        System.Data.DataSet ds1 = new System.Data.DataSet();
        private StringBuilder errorMessages;
        
        public Form1()
        {            
            InitializeComponent();
            //// To report progress from the background worker we need to set this property
            //backgroundWorker1.WorkerReportsProgress = true;           
        }
        public StringBuilder ErrorMessages
        {
            get { return errorMessages; }
            set { errorMessages = value; }
        }        
        public System.Data.DataTable XLStoDTusingInterOp(string FilePath)
        {
            #region Excel important Note.
            /*
             * Excel creates XLS and XLSX files. These files are hard to read in C# programs. 
             * They are handled with the Microsoft.Office.Interop.Excel assembly. 
             * This assembly sometimes creates performance issues. Step-by-step instructions are helpful.
             * 
             * Add the Microsoft.Office.Interop.Excel assembly by going to Project -> Add Reference.
             */
            #endregion

            Microsoft.Office.Interop.Excel.Application excelApp = null;
            Microsoft.Office.Interop.Excel.Workbook workbook = null;


            System.Data.DataTable dt = new System.Data.DataTable(); //Creating datatable to read the content of the Sheet in File.

            try
            {

                excelApp = new Microsoft.Office.Interop.Excel.Application(); // Initialize a new Excel reader. Must be integrated with an Excel interface object.

                //Opening Excel file(myData.xlsx)
                workbook = excelApp.Workbooks.Open(FilePath, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

                Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets.get_Item(1);

                Microsoft.Office.Interop.Excel.Range excelRange = ws.UsedRange; //gives the used cells in sheet

                ws = null; // now No need of this so should expire.

                //Reading Excel file.               
                object[,] valueArray = (object[,])excelRange.get_Value(Microsoft.Office.Interop.Excel.XlRangeValueDataType.xlRangeValueDefault);

                excelRange = null; // you don't need to do any more Interop. Now No need of this so should expire.

                dt = ProcessObjects(valueArray);

            }
            catch (Exception ex)
            {
                ErrorMessages.Append(ex.Message);
            }
            finally
            {
                #region Clean Up                
                if (workbook != null)
                {
                    #region Clean Up Close the workbook and release all the memory.
                    workbook.Close(false, FilePath, Missing.Value);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                    #endregion
                }
                workbook = null;

                if (excelApp != null)
                {
                    excelApp.Quit();
                }
                excelApp = null;

                #endregion
            }
            return (dt);
        }

        /// <summary>
        /// Scan the selected Excel workbook and store the information in the cells
        /// for this workbook in an object[,] array. Then, call another method
        /// to process the data.
        /// </summary>
        private void ExcelScanIntenal(Microsoft.Office.Interop.Excel.Workbook workBookIn)
        {
            //
            // Get sheet Count and store the number of sheets.
            //
            int numSheets = workBookIn.Sheets.Count;

            //
            // Iterate through the sheets. They are indexed starting at 1.
            //
            for (int sheetNum = 1; sheetNum < numSheets + 1; sheetNum++)
            {
                Worksheet sheet = (Worksheet)workBookIn.Sheets[sheetNum];

                //
                // Take the used range of the sheet. Finally, get an object array of all
                // of the cells in the sheet (their values). You can do things with those
                // values. See notes about compatibility.
                //
                Range excelRange = sheet.UsedRange;
                object[,] valueArray = (object[,])excelRange.get_Value(XlRangeValueDataType.xlRangeValueDefault);

                //
                // Do something with the data in the array with a custom method.
                //
                ProcessObjects(valueArray);
            }
        }
        private System.Data.DataTable ProcessObjects(object[,] valueArray)
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            #region Get the COLUMN names

            for (int k = 1; k <= valueArray.GetLength(1); k++)
            {
                dt.Columns.Add((string)valueArray[1, k]);  //add columns to the data table.
            }
            #endregion

            #region Load Excel SHEET DATA into data table

            object[] singleDValue = new object[valueArray.GetLength(1)];
            //value array first row contains column names. so loop starts from 2 instead of 1
            for (int i = 2; i <= valueArray.GetLength(0); i++)
            {
                for (int j = 0; j < valueArray.GetLength(1); j++)
                {
                    if (valueArray[i, j + 1] != null)
                    {
                        singleDValue[j] = valueArray[i, j + 1].ToString();
                    }
                    else
                    {
                        singleDValue[j] = valueArray[i, j + 1];
                    }
                }
                dt.LoadDataRow(singleDValue, System.Data.LoadOption.PreserveChanges);
            }
            #endregion


            return (dt);
        }
        private void btn_excel_Click(object sender, EventArgs e)
        {
            
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = @"C:\";
            openFileDialog1.Title = "Browse Excel Files";
            openFileDialog1.CheckFileExists = true;
            openFileDialog1.CheckPathExists = true;
            openFileDialog1.DefaultExt = "txt";
            openFileDialog1.Filter = "Excel|*.xls|Excel 2010|*.xlsx";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.ReadOnlyChecked = true;
            openFileDialog1.ShowReadOnly = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string selectedFile = openFileDialog1.FileName;
                if (string.IsNullOrEmpty(selectedFile) || selectedFile.Contains(".lnk"))
                {
                    MessageBox.Show("Please select a valid Excel File");
                    return;
                }
                else
                {
                    txtexcelfile.Text = openFileDialog1.FileName;
                   
                }
            }
            else if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
            {
                return;
            }   
        }       
        private void btn_folder_Click(object sender, EventArgs e)
        {
            
            ChooseFolder();
        }
        public void ChooseFolder()
        {
         
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
               
                    txtfolder.Text = folderBrowserDialog1.SelectedPath;
                    folder = folderBrowserDialog1.SelectedPath;
                    //files = Directory.GetFiles(folderBrowserDialog1.SelectedPath);
                string[] files1 = Directory.GetFiles(folderBrowserDialog1.SelectedPath);
                var files = files1.Where(name => !name.EndsWith(".xlsx"));
                //cmbfiles.Items.AddRange(files.Select((string filePath) => Path.GetFileName(filePath)).ToArray());
                lblmessage.Text = "Files found: " + files.Count().ToString();
                
            }
        }
        private void btn_folderpath_Click(object sender, EventArgs e)
        {
          
            //create a folder path,if path is not given then create a temp path and add folder
            SaveFileDialog savefiledialog1 = new SaveFileDialog();
            if (savefiledialog1.ShowDialog() == DialogResult.OK)
            {
                //savefiledialog1.InitialDirectory = @"c:\";
                txtfolderpath.Text = savefiledialog1.FileName;
                if (Directory.Exists(txtfolderpath.Text))
                {
                    MessageBox.Show("Folder Already Exists,Please Select Another Folder Name");
                }
                else
                {
                    Directory.CreateDirectory(txtfolderpath.Text);
                }
            }
            else
            {
                String Todaysdate = DateTime.Now.ToString("dd-MMM-yyyy");
                if (!Directory.Exists("c:\\editortool\\editortoollist\\files\\" + Todaysdate))
                {
                    Directory.CreateDirectory("c:\\editortool\\editortoollist\\files\\" + Todaysdate);
                }
                txtfolderpath.Text = savefiledialog1.FileName;
                //DirectoryInfo LocalDirectory = Directory.CreateDirectory(string.Format("C:\\test\\finaltest\\snaps\\{0}-{1}-{2}", DateTime.Now.Day, DateTime.Now.Month, DateTime.Now.Year));
            }

        }         
        private void btb_start_Click(object sender, EventArgs e)
        {

            if (txtexcelfile.Text=="")
            {
                MessageBox.Show("Please Select Excel File");
            }
           else if(txtfolder.Text=="")
            {
                MessageBox.Show("Please Select Folder Path");
            }
           else if(txtfolderpath.Text=="")
            {
                MessageBox.Show("Please Select Desired Folder Path To Save The Files");
            }
           else
            {

                btb_start.Enabled = false;
                btn_stop.Enabled = true;
               
                lblmessage.Refresh();
                lblmessage.Text = "Excel values are copying into application";
                string excelfile = string.Empty;
                excelfile = txtexcelfile.Text;
                System.Data.DataTable dt = XLStoDTusingInterOp(excelfile);
                string rel = "Released";
                dt.DefaultView.RowFilter = "Dateiname like '%.tif'";

                if (radioButton1.Checked== true)
                {
                    excelTable = dt;
                }
                else
                {
                    excelTable = dt.AsEnumerable()
                  .Where(row => row.Field<String>("LifeCycle State ") == rel)
                  .CopyToDataTable();
                }
                
                int excelcount = excelTable.Rows.Count;
                progressBar1.Maximum = excelcount;
                lblmessage.Refresh();
                lblmessage.Text = Convert.ToString(0);
                //progressBar1.Maximum = excelcount;
                ThreadStart theprogress = new ThreadStart(mainlogic);
                // Now the thread which we create using the delegate
                //Thread.CurrentThread.Priority = ThreadPriority.Lowest;
              

                Thread startprogress = new Thread(theprogress);
                startprogress.Priority = ThreadPriority.Lowest;
            // We can give it a name (optional) 
            startprogress.Name = "Update ProgressBar";                           
            // Start the execution 
            startprogress.Start();
                progressBar1.Visible = true;
            }
        }
        public delegate void updatebar();
        public int count { get; set; }
        private void UpdateProgress()
        {
            lblext.Visible = true;
            progressBar1.Value += 1;
            // Here we are just updating a label to show the progressbar value
            count = Convert.ToInt32(lblmessage.Text) + 1;
            lblmessage.Text = Convert.ToString(count); 

        }


    public void mainlogic()
        {
            
            excelclass c = new excelclass();
            try
            {
                Bitmap b;
                ImageCodecInfo myImageCodecInfo;
                System.Drawing.Imaging.Encoder myEncoder;
                EncoderParameter myEncoderParameter;
                EncoderParameters myEncoderParameters;

                string[] files1 = Directory.GetFiles(folderBrowserDialog1.SelectedPath);
                var files=files1.Where(name => !name.EndsWith(".xlsx"));
                string searchPattern = "*.*";
                var resultData = Directory.GetFiles(folderBrowserDialog1.SelectedPath, searchPattern, SearchOption.AllDirectories)
                    .Select(x => new { FileName = Path.GetFileName(x), FilePath = x });
                c.WriteToLog("check wether data table has data or not", "pass", "pass");
                //check wether data table has data or not, if exists then get all the list of files into a data table
                if (excelTable.Rows.Count != 0)
                {

                    System.Data.DataTable FilesTable = new System.Data.DataTable();
                    FilesTable.TableName = "FileList";
                    FilesTable.Columns.Add("Dateiname");
                    
                    DataRow dRow;

                    foreach (var item in resultData)
                    {
                        dRow = FilesTable.NewRow();
                        dRow["Dateiname"] = item.FileName;
                        FilesTable.Rows.Add(dRow);
                    }
                    //compare two data tables by file names
                    //if file list matches then create a copy of the file in the selected folder and add the text based on filename and ID                
                    foreach (DataRow rowMasterItems in excelTable.Rows)
                    {
                        int k = excelTable.Rows.Count;
                       
                        c.WriteToLog("if file list matches then create a copy of the file in the selected folder and add the text based on filename and ID", "pass", "don");
                        System.Windows.Forms.Application.DoEvents();
                        foreach (DataRow rowItems in FilesTable.Rows)
                        {
                            System.Windows.Forms.Application.DoEvents();
                            if (rowMasterItems["Dateiname"].ToString().Equals(rowItems["Dateiname"].ToString()))
                            {
                                System.Windows.Forms.Application.DoEvents();
                                string nameoffile = string.Empty;
                                nameoffile = Convert.ToString(rowItems["Dateiname"]);
                                string foldername = string.Empty;
                                foldername = txtfolderpath.Text;
                                string folderpath = string.Empty;
                                folderpath = Convert.ToString(txtfolder.Text);


                                //save the copy of the file into the selected folder
                                System.IO.File.Copy(folderpath + "//" + nameoffile, foldername + "//" + nameoffile);
                                //check wether the file is tiff or pdf
                                string path = folderpath + "//" + nameoffile;
                                string ext = Path.GetExtension(path);
                                string displaytext = string.Empty;

                                displaytext = "Footer Text";
                                if ((ext == ".tif") || (ext == ".tiff") || (ext == ".TIF") || (ext == ".TIFF"))
                                {
                                    //System.IO.File.Copy(folderpath + "//" + nameoffile, foldername + "//" + nameoffile);
                                    c.WriteToLog("file copied", "tiff", "only");
                                    // //code to add text into the tiff image
                                    FileStream fs = new FileStream(foldername + "//" + nameoffile, FileMode.Open, FileAccess.Read);
                                    System.Drawing.Image image = System.Drawing.Image.FromStream(fs);
                                    fs.Flush();
                                    fs.Close();
                                    b = new Bitmap(image);
                                    Graphics graphics = Graphics.FromImage(b);
                                    //check the file type
                                    int w = b.Width;
                                    int h = b.Height;
                                    //float width = b.PhysicalDimension.Width;
                                    //float height = b.PhysicalDimension.Height;
                                    //float hresolution = b.HorizontalResolution;
                                    //float vresolution = b.VerticalResolution;

                                    int fz = 0;
                                    w = b.Width - 2000;
                                    if (b.Width < 10000)
                                    {
                                        h = b.Height - 50;
                                        fz = 30;
                                    }
                                    else if (b.Width > 10000)
                                    {
                                        h = b.Height - 84;
                                        fz = 40;
                                    }
                                    // Get an ImageCodecInfo object that represents the TIFF codec.
                                    myImageCodecInfo = GetEncoderInfo("image/tiff");

                                    // Create an Encoder object based on the GUID
                                    // for the Compression parameter category.
                                    myEncoder = System.Drawing.Imaging.Encoder.Compression;

                                    // Create an EncoderParameters object.
                                    // An EncoderParameters object has an array of EncoderParameter
                                    // objects. In this case, there is only one
                                    // EncoderParameter object in the array.
                                    myEncoderParameters = new EncoderParameters(1);

                                    // Save the bitmap as a TIFF file with LZW compression.
                                    myEncoderParameter = new EncoderParameter(
                                        myEncoder,
                                        (long)EncoderValue.CompressionCCITT4);
                                    myEncoderParameters.Param[0] = myEncoderParameter;

                                    graphics.DrawString(displaytext, new System.Drawing.Font("Arial", fz, FontStyle.Regular), SystemBrushes.WindowText, new System.Drawing.Point(w, h));

                                    b.Save(foldername + "//" + nameoffile, myImageCodecInfo, myEncoderParameters);
                                    b.Dispose();
                                    image.Dispose();
                                    b.Dispose();
                                    progressBar1.Invoke(new updatebar(this.UpdateProgress));

                                }
                                else if ((ext == ".pdf") || (ext == ".PDF"))
                                {

                                    progressBar1.Invoke(new updatebar(this.UpdateProgress));
                                    excelclass pdffuncton = new excelclass();
                                    pdffuncton.AddPageNumber(foldername + "//" + nameoffile, displaytext);
                                    c.WriteToLog("pdf addedd successfully", "task", "pdf files");

                                    //byte[] bytes = File.ReadAllBytes(foldername + "//" + nameoffile);

                                    //iTextSharp.text.Font blackFont = FontFactory.GetFont("Arial", 5, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);

                                    //using (MemoryStream stream = new MemoryStream())

                                    //{

                                    //    PdfReader reader = new PdfReader(bytes);
                                    //    PdfReader.unethicalreading = true;
                                    //    using (PdfStamper stamper = new PdfStamper(reader, stream))

                                    //    {

                                    //        int pages = reader.NumberOfPages;

                                    //        for (int i = 1; i <= pages; i++)

                                    //        {
                                    //            var size = reader.GetPageSize(i);
                                    //            float w = size.Width;
                                    //            float h = size.Height;
                                    //               ColumnText.ShowTextAligned(stamper.GetUnderContent(i), Element.ALIGN_BASELINE, new Phrase(displaytext.ToString(), blackFont), 250f, 5f, 0);

                                    //        }

                                    //    }

                                    //    bytes = stream.ToArray();

                                    //}

                                    //File.WriteAllBytes(foldername + "//" + nameoffile, bytes);


                                }

                            }

                        }

                    }
                    filecount = Directory.GetFiles(txtfolderpath.Text);
                    
                    MessageBox.Show("Files found in the folder: " + files.Count().ToString() + Environment.NewLine + "Files found in Released Mode: " + excelTable.Rows.Count.ToString() + Environment.NewLine + "Files created in the folder: " + filecount.Length.ToString() + Environment.NewLine + "Files were Added Sucessfully");
                   
                    //th.Abort();

                    System.Windows.Forms.Application.Exit();
                }
            }
            catch(Exception ex)
            {
                c.WriteToLog("error", ex.Message, "error in reading excel file");
                MessageBox.Show(ex.Message);
                
            }
            finally
            {
                c.WriteToLog("success", "reading of excel done successfully", "success");
            }
        }
        private static ImageCodecInfo GetEncoderInfo(String mimeType)
        {
            int j;
            ImageCodecInfo[] encoders;
            encoders = ImageCodecInfo.GetImageEncoders();
            for (j = 0; j < encoders.Length; ++j)
            {
                if (encoders[j].MimeType == mimeType)
                    return encoders[j];
            }
            return null;
        }

        private void btn_stop_Click(object sender, EventArgs e)
        {            
            if (!backgroundWorker1.IsBusy)
                backgroundWorker1.CancelAsync(); 
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            

        }
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {

            for (int i = 1; i <= 100; i++)
            {
                if (backgroundWorker1.CancellationPending)
                {
                    e.Cancel = true;
                }
                else
                {
                   
                }
            }

            
                //int percentage = (progressBar1.Value / progressBar1.Maximum) * 100;
                //lblmessage.Text = "Current progress: " + percentage.ToString() + "%";

            
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
            
        }
        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {         

            if (e.Cancelled) MessageBox.Show("Operation was canceled");
            else if (e.Error != null) MessageBox.Show(e.Error.Message);
            if (progressBar1.Style == ProgressBarStyle.Marquee)
            {
                progressBar1.Visible = false;
            }
            progressBar1.Visible = false;           
            txtexcelfile.Text = "";
            txtfolder.Text = "";
            txtfolderpath.Text = "";
            System.Windows.Forms.Application.Exit();
        }
        private void loading()
        {
            for (int i = 0; i < 1000; i++)
            {
                if (progressBar1.InvokeRequired)
                    progressBar1.Invoke(new System.Action(loading));
                else
                    progressBar1.Value = i;
                //int percentage = (progressBar1.Value / progressBar1.Maximum) * 100;
                //lblmessage.Text = "Current progress: " + percentage.ToString() + "%";
                int percent = (int)(((double)progressBar1.Value / (double)progressBar1.Maximum) * 100);
                //progressBar1.Refresh();
                progressBar1.CreateGraphics().DrawString(percent.ToString() + "%",
                    new System.Drawing.Font("Arial", (float)8.25, FontStyle.Regular),
                    Brushes.Black,
                    new PointF(progressBar1.Width / 2 - 10, progressBar1.Height / 2 - 7));
            }
            
        }

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }
    }
}
