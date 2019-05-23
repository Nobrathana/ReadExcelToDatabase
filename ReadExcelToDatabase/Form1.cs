using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace ReadExcelToDatabase
{
    public partial class Form1 : Form
    {
        private ImportExcel fileImportHandler;
        public Form1()
        {
            InitializeComponent();
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            if (fileImportHandler.ImportToDB())
            {
                MessageBox.Show("Done!!!", "Message");
            }
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            var x = openFileDialog1.ShowDialog();
            if(x == DialogResult.OK)
            {
                fileImportHandler = new ImportExcel(openFileDialog1.FileName, openFileDialog1.SafeFileName);
                
            }
        }
    }
}
