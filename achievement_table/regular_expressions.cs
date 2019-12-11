using System;
using System.IO;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace achievement_table
{
    public partial class Form1 : Form
    {
        // for weekly achievement table
        static string exposure_count_table_pattern = @"(\d+)_舉發件數統計表\(員警法條車種別\)\.xls";
        static Regex exposure_count_table_regex = new Regex(exposure_count_table_pattern, RegexOptions.Compiled);
        //static string exposure_count_table_pattern = @"(\d+)_舉發件數統計表\(員警法條別\)\.xls";
        //static Regex exposure_count_table_regex = new Regex(exposure_count_table_pattern, RegexOptions.Compiled);
        static string exposure_organization_pattern = @"舉發單位：交通隊";
        static Regex exposure_organization_regex = new Regex(exposure_organization_pattern, RegexOptions.Compiled);
        static string exposure_police_name_pattern = @"舉發員警：(\w+)";
        static Regex exposure_police_name_regex = new Regex(exposure_police_name_pattern, RegexOptions.Compiled);
        static string exposure_amount_pattern = @"員警合計：";
        static Regex exposure_amount_regex = new Regex(exposure_amount_pattern, RegexOptions.Compiled);


        static string motorcycle_table_pattern = @"(\d+)_.xls";
        static Regex motorcycle_table_regex = new Regex(motorcycle_table_pattern, RegexOptions.Compiled);
        static string pernicious_ban_table_pattern = @"(\d+)_取締惡性違規績效評核一覽表.xls";
        static Regex pernicious_ban_table_regex = new Regex(pernicious_ban_table_pattern, RegexOptions.Compiled);

        //static string achievement_table_pattern = @"(\d{3})(\d{4})績效表\.xls";
        //static Regex achievement_table_regex = new Regex(achievement_table_pattern, RegexOptions.Compiled);
        static string achievement_table_pattern = @"(\d+)_績效表\.xls";
        static Regex achievement_table_regex = new Regex(achievement_table_pattern, RegexOptions.Compiled);
        static string achievement_table_police_pattern = @"警員";
        static Regex achievement_table_police_regex = new Regex(achievement_table_police_pattern, RegexOptions.Compiled);

        static string monthly_dir_pattern = @"(\d+)月份";
        static Regex monthly_dir_regex = new Regex(monthly_dir_pattern, RegexOptions.Compiled);


        static string serial_number_range_pattern = @"(\d+)-(\d+)";
        static Regex serial_number_range_regex = new Regex(serial_number_range_pattern, RegexOptions.Compiled);

        // why "\w" could detect some other words like characters ??
        //static string license_plate_pattern = @"\w+-\w+";
        static string license_plate_pattern = @"[A-Za-z0-9_]+-[A-Za-z0-9_]+";
        static Regex license_plate_regex = new Regex(license_plate_pattern, RegexOptions.Compiled);
        //static string serial_number_license_plate_pattern = @"(\d+)\.(\w+-\w+)";
        static string serial_number_license_plate_pattern = @"(\d+)\.(" + license_plate_pattern + @")";
        static Regex serial_number_license_plate_regex = new Regex(serial_number_license_plate_pattern, RegexOptions.Compiled);
        static string excel_extension_pattern = @"xlsx?";
        static Regex excel_extension_regex = new Regex(excel_extension_pattern, RegexOptions.Compiled);
        static string xml_extension_pattern = @"[xX][mM][lL]";
        static Regex xml_extension_regex = new Regex(xml_extension_pattern, RegexOptions.Compiled);
        static string digital_number_pattern = @"(\d+,)*(\d+)";
        static Regex digital_number_regex = new Regex(digital_number_pattern, RegexOptions.Compiled);
        static string notice_name_pattern = @"繳款通知書";
        static Regex notice_name_regex = new Regex(notice_name_pattern, RegexOptions.Compiled);
        static string word_extension_pattern = @"docx?";
        static Regex word_extension_regex = new Regex(word_extension_pattern, RegexOptions.Compiled);

        static string total_table_name_pattern = @"總表";
        static Regex total_table_name_regex = new Regex(total_table_name_pattern, RegexOptions.Compiled);
        static string chinese_birthday_pattern = @"民國(\d+)年(\d+)月(\d+)日";
        static Regex chinese_birthday_regex = new Regex(chinese_birthday_pattern, RegexOptions.Compiled);
        static string chinese_name_pattern = @"姓名：(\w+)";
        static Regex chinese_name_regex = new Regex(chinese_name_pattern, RegexOptions.Compiled);
        static string chinese_personal_serial_number_pattern = @"統號：([A-Z]\d{9})";
        static Regex chinese_personal_serial_number_regex = new Regex(chinese_personal_serial_number_pattern, RegexOptions.Compiled);
        static string chinese_personal_address_pattern = @"地址：(\w+)";
        static Regex chinese_personal_address_regex = new Regex(chinese_personal_address_pattern, RegexOptions.Compiled);
        static string chinese_company_name_pattern = @"(\w+) .Google搜尋";
        static Regex chinese_company_name_regex = new Regex(chinese_company_name_pattern, RegexOptions.Compiled);
        //static string chinese_company_address_pattern = @"([\w（）]+) .電子地圖";
        static string chinese_company_address_pattern = @"([\w（）]+)( .電子地圖)?";
        static Regex chinese_company_address_regex = new Regex(chinese_company_address_pattern, RegexOptions.Compiled);
        static string chinese_address_pattern = @"(\d+)?(臺灣省)?(\w+?[縣市])(\w+?[鄉鎮區市])";
        static Regex chinese_address_regex = new Regex(chinese_address_pattern, RegexOptions.Compiled);
        static string chinese_household_office_pattern = @"戶政事務所";
        static Regex chinese_household_office_regex = new Regex(chinese_household_office_pattern, RegexOptions.Compiled);
        static string chinese_dead_ke_word_pattern = @"死亡";
        static Regex chinese_dead_ke_word_regex = new Regex(chinese_dead_ke_word_pattern, RegexOptions.Compiled);

        static string chinese_address_title_pattern_0 = @"公司所在地";
        static Regex chinese_address_title_regex_0 = new Regex(chinese_address_title_pattern_0, RegexOptions.Compiled);
        static string chinese_address_title_pattern_1 = @"地址";
        static Regex chinese_address_title_regex_1 = new Regex(chinese_address_title_pattern_1, RegexOptions.Compiled);
        static string chinese_company_name_title_pattern_0 = @"公司名稱";
        static Regex chinese_company_name_title_regex_0 = new Regex(chinese_company_name_title_pattern_0, RegexOptions.Compiled);
        static string chinese_company_name_title_pattern_1 = @"商業名稱";
        static Regex chinese_company_name_title_regex_1 = new Regex(chinese_company_name_title_pattern_1, RegexOptions.Compiled);

        static string chinese_company_serial_number_title_pattern = @"統一編號";
        static Regex chinese_company_serial_number_title_regex = new Regex(chinese_company_serial_number_title_pattern, RegexOptions.Compiled);

        static string chinese_before_strongly_execute_pattern = @"強執前";
        static Regex chinese_before_strongly_execute_regex = new Regex(chinese_before_strongly_execute_pattern, RegexOptions.Compiled);

        //static string serial_number_license_plate_optional_pattern = @"(\d+)\.?[" + license_plate_pattern + @"]?";
        //static Regex serial_number_license_plate_optional_regex = new Regex(serial_number_license_plate_optional_pattern, RegexOptions.Compiled);

        static string serial_number_dash_license_plate_pattern = @"(\d+)_(" + license_plate_pattern + @")";
        static Regex serial_number_dash_license_plate_regex = new Regex(serial_number_dash_license_plate_pattern, RegexOptions.Compiled);

        static string inventory_fee_infomation_pattern = @"尚欠   催繳工本費：計 (\d+)筆(\d+)元 總計：(\d+)元";
        static Regex inventory_fee_infomation_regex = new Regex(inventory_fee_infomation_pattern, RegexOptions.Compiled);

        bool digital_number_string_to_int(GroupCollection groups, out int n)
        {
            n = 0;

            if (groups[1].Value.Trim() != "")
            {
                // there are some ',' in this string

                int temp;
                if (!int.TryParse(groups[groups.Count - 1].Value.Trim(), out temp))
                {
                    log("錯誤!將最末碼阿拉伯數字從文字轉換成數字時 發生問題" + Environment.NewLine + "該阿拉伯數字為 : " + groups[0].Value.Trim(), error: true);
                    return false;
                }
                n += temp;

                for (int i = groups.Count - 2, th = 1000; i >= 1; --i, th *= 1000)
                {
                    string s = groups[i].Value.Trim();
                    s = s.Remove(s.Length - 1, 1);
                    if (!int.TryParse(s, out temp))
                    {
                        log("錯誤!將阿拉伯數字從文字轉換成數字時 發生問題" + Environment.NewLine + "該阿拉伯數字為 : " + groups[0].Value.Trim(), error: true);
                        return false;
                    }

                    n += temp * th;
                }
            }
            else if (!int.TryParse(groups[0].Value.Trim(), out n))
            {
                log("錯誤!將完整的阿拉伯數字從文字轉換成數字時 發生問題" + Environment.NewLine + "該阿拉伯數字為 : " + groups[0].Value.Trim(), error: true);
                return false;
            }

            return true;
        }

        void digital_number_int_to_string(ref int n, out string s)
        {
            s = n.ToString();
            int l = s.Length;

            while (l > 3)
            {
                l -= 3;
                s = s.Insert(l, ",");
            }
        }

        bool match_number_range(ref string s, out int n0, out int n1)
        {
            n0 = -1;
            n1 = -1;

            // match serial number range
            MatchCollection matches = serial_number_range_regex.Matches(s);
            if (matches.Count != 1)
            {
                log("錯誤!尋找編號範圍時 發生問題" + Environment.NewLine + "該編號範圍為 : " + s, error: true);
                return false;
            }

            foreach (Match match in matches)
            {
                GroupCollection groups = match.Groups;
                if (groups.Count != 3)
                {
                    log("錯誤!解析編號範圍時 發生問題" + Environment.NewLine + "該編號範圍為 : " + s, error: true);
                    return false;
                }

                if (!int.TryParse(groups[1].Value.Trim(), out n0))
                {
                    log("錯誤!解析起始之編號時 發生問題" + Environment.NewLine + "該起始之編號為 : " + groups[1].Value.Trim(), error: true);
                    return false;
                }
                if (!int.TryParse(groups[2].Value.Trim(), out n1))
                {
                    log("錯誤!解析起始之編號時 發生問題" + Environment.NewLine + "該起始之編號為 : " + groups[2].Value.Trim(), error: true);
                    return false;
                }
                if (n0 >= n1)
                {
                    log("錯誤!該編號範圍不合邏輯" + Environment.NewLine + "該編號範圍為 : " + s, error: true);
                    return false;
                }
            }

            return true;
        }

        bool serial_number_license_plate_match(ref string s, out int n, out string l_p)
        {
            n = -1;
            l_p = "";

            MatchCollection matches = serial_number_license_plate_regex.Matches(Path.GetFileName(s));
            if (matches.Count != 1 || matches[0].Groups.Count != 3)
            {
                log("錯誤!解析 序號.車牌 時 發生問題" + Environment.NewLine + "該 序號.車牌 為 : " + Path.GetFileName(s), error: true);
                return false;
            }

            GroupCollection groups = matches[0].Groups;
            if (!int.TryParse(groups[1].Value.Trim(), out n))
            {
                log("錯誤!解析 序號.車牌 之序號時 發生問題" + Environment.NewLine + "該序號為 : " + groups[1].Value.Trim(), error: true);
                return false;
            }
            l_p = groups[2].Value.Trim();

            log("序號為 : " + n.ToString() + " 車牌為 : " + l_p);
            return true;
        }

        bool match_addresses(ref string address_0, ref string address_1, ref bool found_error, ref string error_info)
        {
            // match the difference between excel personal address & personal address
            {
                MatchCollection matches_0 = chinese_address_regex.Matches(address_0);
                MatchCollection matches_1 = chinese_address_regex.Matches(address_1);
                if (matches_0.Count == 0 || matches_1.Count == 0)
                {
                    log("錯誤!地址深入解析時 發生問題!" + Environment.NewLine + "該地址為 : " + address_0 + " & " + address_1, error: true);

                    log("matches_0 count == " + matches_0.Count.ToString() + " matches_1 count == " + matches_1.Count.ToString());

                    found_error = true;
                    error_info += "_地址深入解析";

                    return true;
                }

                if (matches_0.Count != matches_1.Count)
                {
                    log("matches_0 count == " + matches_0.Count.ToString() + " matches_1 count == " + matches_1.Count.ToString());
                    found_error = true;
                    error_info += "_地址深入解析數量不合";
                    return false;
                }

                //log("matches_0 count == " + matches_0.Count + " matches_1 count == " + matches_1.Count);
                //log(" matches_0[0].Groups.Count == " + matches_0[0].Groups.Count + "  matches_1[0].Groups.Count == " + matches_1[0].Groups.Count);

                GroupCollection groups_0 = matches_0[0].Groups;
                //for (int i = 0; i < groups_0.Count; ++i)
                //{
                //    log("groups_0[i] == " + groups_0[i].Value);
                //}
                string n_0 = groups_0[1].Value.Trim(), level_0_0 = groups_0[2].Value.Trim(), level_1_0 = groups_0[3].Value.Trim(), level_2_0 = groups_0[4].Value.Trim();

                GroupCollection groups_1 = matches_1[0].Groups;
                //for (int i = 0; i < groups_1.Count; ++i)
                //{
                //    log("groups_1[i] == " + groups_1[i].Value);
                //}
                string n_1 = groups_1[1].Value.Trim(), level_0_1 = groups_1[2].Value.Trim(), level_1_1 = groups_1[3].Value.Trim(), level_2_1 = groups_1[4].Value.Trim();

                if (level_1_0 != level_1_1 || level_2_0 != level_2_1)
                {
                    log("地址深入解析 發現地址不同" + Environment.NewLine + "excel 地址為 : " + address_0 + " 而 word 地址為 : " + address_1);
                    log("level_1_0 == " + level_1_0 + " level_1_1 == " + level_1_1 + " level_2_0 == " + level_2_0 + " level_2_1 == " + level_2_1);
                    return false;
                }

                return true;
            }
        }

        bool match_single_special_word(string s, ref Regex regex)
        {
            MatchCollection matches = regex.Matches(s);
            //return matches.Count == 1 && matches[0].Groups.Count == 1;
            return matches.Count == 1;
        }

        bool find_string_in_strings(string[] ss, out string out_s, ref Regex regex, string error_message = "錯誤!")
        {
            out_s = "";
            
            foreach (string s in ss)
            {
                if (match_single_special_word(s, ref regex))
                {
                    out_s = s;
                    return true;
                }
            }
            
            if (error_message != "")
                log(error_message, error: true);

            return false;
        }
    }
}