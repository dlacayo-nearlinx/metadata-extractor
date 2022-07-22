using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;


namespace WindowsFormsADSTest
{    
    public partial class MainForm : Form
    {

        private string _folderPath;

        public MainForm()
        {
            InitializeComponent();
        }        

        private void btnChooseDir_Click(object sender, EventArgs e)
        {
            var result = folderDialog.ShowDialog();
            if (result == DialogResult.OK) 
            {                
                _folderPath = folderDialog.SelectedPath;
                txtPath.Text = _folderPath;
            }
        }

        private void btnProcessFiles_Click(object sender, EventArgs e)
        {
            txtInfo.Text = "";
            // list all files
            DirectoryInfo info = new DirectoryInfo(_folderPath);
            var files = info.GetFiles();
            // process all files
            foreach (var file in files)
            {
                //txtInfo.Text += MetaDataExtractor.GetExtendedPropertiesFromFolderData(file);
                //txtInfo.Text += MetaDataExtractor.GetExtendedProperties(file);
                txtInfo.Text += MetaDataExtractor.GetKeywordsUsingCodeFluent(file);
                txtInfo.Text += "\n\n";
            }
        }        
    }
}
