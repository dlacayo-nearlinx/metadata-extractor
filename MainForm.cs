using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;


namespace WindowsFormsADSTest
{
    public partial class MainForm : Form
    {

        private string _folderPath;
        private string _outputFile;

        public MainForm()
        {
            InitializeComponent();
            UpdateButtonStatus();
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
            var processSubDirs = chkRecursive.Checked ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
            
            
            var directoryCount = chkRecursive.Checked ? Directory.GetDirectories(_folderPath, "*", processSubDirs).Count() : 0;
            var directoryInfo = new DirectoryInfo(_folderPath);
            var files = directoryInfo.GetFiles("*.*",processSubDirs);

            txtInfo.Text = "";
            txtInfo.Text += string.Format("Processing {0} files across {1} directories.\n\n",files.Count(),directoryCount);

            var filesData = new List<FileMetadata>();
            foreach (var file in files)
            {
                var fileData = MetaDataExtractor.GetFileMetadata(file);
                filesData.Add(fileData);
                txtInfo.Text += fileData.ToString();                
                txtInfo.Text += "\n\n";
            }

            ExcelFileCreator.ExportDataToExcel(filesData, _outputFile);            
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
