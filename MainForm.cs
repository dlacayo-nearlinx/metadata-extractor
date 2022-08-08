using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using MetadataExtractor.Common.Models;
using MetadataExtractor.Common.Utilities;
using NLog;


namespace MetadataExtractor
{
    public partial class MainForm : Form
    {

        private string _folderPath;
        private string _outputFile;
        private Logger _logger = LogManager.GetLogger("MetadataExtractor");

        public MainForm()
        {
            InitializeComponent();
            UpdateButtonStatus();
            lblProgress.Text = "";
        }        

        private void btnChooseDir_Click(object sender, EventArgs e)
        {
            var result = folderDialog.ShowDialog();
            if (result == DialogResult.OK) 
            {                
                _folderPath = folderDialog.SelectedPath;
                txtPath.Text = _folderPath;
            }
            UpdateButtonStatus();
        }

       

        private void btnProcessFiles_Click(object sender, EventArgs e)
        {
            btnProcessFiles.Enabled = false;

            try
            {
                lblProgress.Text = "Getting file and directory information";
                lblProgress.Refresh();

                Stopwatch sw = Stopwatch.StartNew();
                var processSubDirs = chkRecursive.Checked ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;


                var directoryCount = chkRecursive.Checked
                    ? Directory.GetDirectories(_folderPath, "*", processSubDirs).Count()
                    : 0;
                var directoryInfo = new DirectoryInfo(_folderPath);
                var files = directoryInfo.GetFiles("*.*", processSubDirs);
                var fileCount = files.Count();

                progressBar.Minimum = 0;
                progressBar.Maximum = fileCount;
                progressBar.Value = 0;

                _logger.Info(string.Format("Processing {0} files across {1} directories.\n\n", fileCount, directoryCount));

                var fileProgressCounter = 1;
                var filesData = new List<FileMetadata>();
                foreach (var file in files)
                {
                    lblProgress.Text = string.Format("Processing file {0} of {1} across {2} directories",fileProgressCounter,fileCount,directoryCount);
                    lblProgress.Refresh();
                    fileProgressCounter++;

                    var fileData = MetaDataExtractor.GetFileMetadata(file);
                    filesData.Add(fileData);
                    progressBar.Value++;
                    //txtInfo.Text += fileData.ToString();
                    _logger.Info(fileData.ToString());
                    //txtInfo.Text += "\n\n";
                }

                _logger.Info("Exporting data to Excel File");

                lblProgress.Text = "Exporting data to Excel File";
                lblProgress.Refresh();
                ExcelFileCreator.ExportDataToExcel(filesData, _outputFile);
                
                sw.Stop();

                lblProgress.Text = string.Format("Processing Finished! Time was {0} Hours {1} Minutes {2} Seconds", sw.Elapsed.Hours, sw.Elapsed.Minutes, sw.Elapsed.Seconds); ;
                lblProgress.Refresh();
                var message = string.Format("Processing Finished!\n{0} Hours\n{1} Minutes \n{2} Seconds", sw.Elapsed.Hours, sw.Elapsed.Minutes, sw.Elapsed.Seconds);
                _logger.Info(message);                

            }
            catch (Exception ex)
            {
                _logger.Error(ex);
                MessageBox.Show(ex.Message);
            }
            finally
            {
                btnProcessFiles.Enabled = true;
                
            }
        }

        private void btnOutputFile_Click(object sender, EventArgs e)
        {
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                var outputFile = fileDialog.FileName;
                txtOutputFile.Text = outputFile;
                _outputFile = outputFile;
            }
            UpdateButtonStatus();
        }

        private void UpdateButtonStatus() 
        {
            btnProcessFiles.Enabled = string.IsNullOrWhiteSpace(_folderPath) == false && string.IsNullOrWhiteSpace(_outputFile) == false;
        }

    }
}
