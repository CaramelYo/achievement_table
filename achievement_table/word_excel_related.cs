using System.Windows.Forms;

// for excel access
using Excel = Microsoft.Office.Interop.Excel;

// for word access
using Word = Microsoft.Office.Interop.Word;

namespace achievement_table
{
    public partial class Form1 : Form
    {
        Excel.Application excel_app = null;
        Word.Application word_app = null;
        
        int notice_row_interval = 38, notice_column_interval = 10;
        int notice_license_plate_row_start = 11, notice_license_plate_column_start = 3;
        int notice_text_row_offset = 1, notice_text_column_offset = -1;
        int notice_data_table_row_start = 3;
        int notice_data_table_serial_number_column = 1;

        int total_table_row_start = 3;

        int before_strongly_execute_row_start = 2;
        
        int personal_birthday_row_start = 3, personal_birthday_column_start = 1;
        int personal_name_row_start = 2, personal_name_column_start = 2;
        int personal_serial_number_row_start = 2, personal_serial_number_column_start = 1;
        int personal_address_row_start = 9, personal_address_column_start = 1;

        void excel_app_open()
        {
            excel_app = new Excel.Application();
            //app.Visible = true;
            excel_app.DisplayAlerts = false;
        }

        void excel_app_close()
        {
            excel_app.Quit();
            excel_app = null;
        }

        void excel_open_book_and_sheets(string path, out Excel._Workbook book, out Excel.Sheets sheets)
        {
            if (excel_app == null)
            {
                log("錯誤!並未開啟excel應用程式", error: true);
                excel_app_open();
            }

            book = excel_app.Workbooks.Open(path, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true);
            sheets = book.Worksheets;
        }

        void word_app_open()
        {
            word_app = new Word.Application();
            //word_app.Visible = true;
            word_app.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
        }

        void word_app_close()
        {
            word_app.Quit();
            word_app = null;
        }

        void word_open_doc(string path, out Word.Document doc)
        {
            if (word_app == null)
            {
                log("錯誤!並未開啟word應用程式", error: true);
                word_app_open();
            }

            doc = word_app.Documents.Open(path, ReadOnly: true);
        }
    }
}