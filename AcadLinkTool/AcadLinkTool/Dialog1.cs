using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AcadLinkTool
{
    public partial class Dialog1 : Form
    {
        //Private Dialog Object
        private static Dialog1 instance;

        private void Dialog1_Load(object sender, EventArgs e)
        {
        }
        private void button1_Click(object sender, EventArgs e)
        {

            FolderBrowserDialog FolderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            string myFolderPath = "";

            // Then use the following code to create the Dialog window
            // Change the .SelectedPath property to the default location
            var _with1 = FolderBrowserDialog1;

            // Desktop is the root folder in the dialog.
            _with1.RootFolder = Environment.SpecialFolder.Desktop;

            // Select the C:\Windows directory on entry.
            _with1.SelectedPath = "G:\\";

            // Prompt the user with a custom message.
            _with1.Description = "Select the directory of P&ID Drawings to hyperlink";

            DialogResult result1 = _with1.ShowDialog();

            if (result1 == System.Windows.Forms.DialogResult.OK)
            {
                myFolderPath = _with1.SelectedPath;

                txtFolderPath.Text = _with1.SelectedPath;

                switch (MessageBox.Show("Do you want to process Sub Directories?", "Sub Directories?", MessageBoxButtons.YesNo))
                {
                    case DialogResult.Yes:
                        LinkCode.LoopThroughFilesInDirectory(myFolderPath, SearchOption.AllDirectories);
                        break;
                    case DialogResult.No:
                        LinkCode.LoopThroughFilesInDirectory(myFolderPath, SearchOption.TopDirectoryOnly);
                        break;
                }
            }
            else
            {
                return;
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            this.Reset();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            this.CopyToClipboard();
        }

        //Public References
        public static Dialog1 Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new Dialog1();
                }
                return instance;
            }
        }

        public Dialog1()
        {
            InitializeComponent();
        }

        //Folder Path
        public string folderSelectedPath;

        //Success List
        //private List<string> listFilesSucceded;
        public List<string> Succeded_FileList
        {
            get {
                return listbxProcessed.Items.Cast<string>().ToList();
            }
            set
            {
                //listFilesSucceded = value;
                listbxProcessed.Items.AddRange(value.ToArray());
            }

        }
        public void fileSucceded(string filename)
        {
            listbxProcessed.Items.Add(filename);
            listbxProcessed.Refresh();
        }

        //Error List
       // private List<string> listFilesErrored;
        public List<string> Errored_FileList
        {
            get {
                 return listbxErrors.Items.Cast<string>().ToList();
            }
            set
            {
                //listFilesErrored = value;
                listbxErrors.Items.AddRange(value.ToArray());
            }

        }
        public void fileErrored(string filename)
        {
            listbxErrors.Items.Add(filename);
            listbxErrors.Refresh();
        }

        //Progress Bar Methods & Properties
        public void Progress()
        {
            //Catch High Numbers
            if (progressBar1.Value == progressBar1.Maximum) { progressBar1.Value = 0; }

            progressBar1.Value++;
            ProgressBar1Label.Text  = string.Format("Processing {0} of {1}", progressBar1.Value, progressBar1.Maximum);
            ProgressBar1Label.Refresh();
        }

        private int hCount = 0;
        public void ProgressHyperlinks()
        {
             hCount++;
             lblHyperlinkCount.Text = string.Format("Hyperlinks : {0}", hCount);
             lblHyperlinkCount.Refresh();
        }

        public void ShowProgressCtrl()
        {

        }
        public void Complete()
        {
            progressBar1.ForeColor = Color.LightGray;
            hCount = 0;
        }
        public void resetMaxFileCount(int value)
        {
                progressBar1.Value = 0;
                progressBar1.Maximum = value;
        }
        public void Reset()
        {
            txtFolderPath.Text = "";

            progressBar1.Value = 0;
            progressBar1.Maximum = 100;

            lblHyperlinkCount.Text = "0";

            Succeded_FileList.Clear();
            listbxErrors.Items.Clear();

            Errored_FileList.Clear();
            listbxProcessed.Items.Clear();
        }
        public void CopyToClipboard()
        {
            LinkCode.ed.WriteMessage("Errored Files copied to Clipboard \n");

            string s = string.Join(Environment.NewLine, Errored_FileList.ToArray());
            Clipboard.SetText(s);
        }


    }
}
