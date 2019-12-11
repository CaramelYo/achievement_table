using System;
using System.Windows.Forms;
using System.IO;

namespace achievement_table
{
    public partial class Form1 : Form
    {
        StreamWriter log_sw, system_log_sw;
        string system_log_sw_path = "system_log.txt";
        string error_message_start = "!!!!! ", error_message_end = " !!!!!";

        bool used_log;
        
        void set_log_file(string p)
        {
            used_log = true;
            log_sw = new StreamWriter(p, true, encoding: System.Text.Encoding.Default);
        }

        void close_log_file()
        {
            used_log = false;
            log_sw.Close();
        }

        void log(string m, bool with_time = false, bool error = false)
        {
            string t = error ? error_message_start + m + error_message_end : m;

            string t_with_time = DateTime.Now + " : " + t;
            log_tbx.Text += (with_time ? t_with_time : t) + Environment.NewLine;

            system_log_sw.WriteLine(t_with_time);
            system_log_sw.Flush();

            if (used_log)
            {
                log_sw.WriteLine(with_time ? t_with_time : t);
                log_sw.Flush();
            }
        }
    }
}