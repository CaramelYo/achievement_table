using System;
using System.Windows.Forms;
using System.IO;

namespace achievement_table
{
    public partial class Form1 : Form
    {
        FolderBrowserDialog folder_dialog_0;
        string base_dir_path;

        public Form1()
        {
            InitializeComponent();

            log_tbx.Text = "";
            base_dir_path = "";

            folder_dialog_0 = new FolderBrowserDialog()
            {
                SelectedPath = AppDomain.CurrentDomain.BaseDirectory
            };

            system_log_sw = new StreamWriter(system_log_sw_path, true, encoding: System.Text.Encoding.Default);
            used_log = false;
        }

        ~Form1()
        {
            system_log_sw.Close();
        }
    }
}
