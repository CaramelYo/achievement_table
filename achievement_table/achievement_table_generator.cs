using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;

// for excel access
using Excel = Microsoft.Office.Interop.Excel;

namespace achievement_table
{
    public partial class Form1 : Form
    {
        int exposure_count_table_row_start = 11;
        int achievement_table_row_start = 5;

        void achievement_table_generator()
        {
            // find file path
            string exposure_count_table_path = "";
            string motorcycle_table_path = "";
            string achievement_table_path = "";
            {
                string[] base_dir_file_paths = Directory.GetFiles(base_dir_path);

                if (!find_string_in_strings(base_dir_file_paths, out exposure_count_table_path, ref exposure_count_table_regex, "錯誤!找無 舉發件數統計表(員警法條車種別) 之檔案"))
                    return;

                if (!find_string_in_strings(base_dir_file_paths, out motorcycle_table_path, ref motorcycle_table_regex, "錯誤!找無 機車表格 之檔案"))
                    return;

                if (!find_string_in_strings(base_dir_file_paths, out achievement_table_path, ref achievement_table_regex, "錯誤!找無 績效表 之檔案"))
                    return;
            }

            excel_app_open();

            // get all of needed law serial number regex
            List<Regex> law_regexes = new List<Regex>();
            {
                law_regexes.Add(Achievement.DWI_regex);
                law_regexes.Add(Achievement.running_the_red_light_stopping_regex);
                law_regexes.Add(Achievement.license_plate_related_regex);
                law_regexes.Add(Achievement.motorcyle_in_wrong_way_regex);
                law_regexes.Add(Achievement.reverse_driving_regex);
                law_regexes.Add(Achievement.unused_headlight_regex);
                law_regexes.Add(Achievement.violating_pedestrian_right_regex);
                law_regexes.Add(Achievement.unused_safety_helmet_regex);
                law_regexes.Add(Achievement.dangerously_driving_regex);
                law_regexes.Add(Achievement.gravel_truck_related_regex);
            }

            Excel._Workbook exposure_count_table_book;
            Excel.Sheets exposure_count_table_sheets;
            excel_open_book_and_sheets(exposure_count_table_path, out exposure_count_table_book, out exposure_count_table_sheets);
            Excel._Worksheet exposure_count_table_sheet = (Excel.Worksheet)exposure_count_table_sheets.Item[1];

            List<Achievement> achievement_list = new List<Achievement>();
            int null_counter = 0;


            for (int r = exposure_count_table_row_start; r <= exposure_count_table_sheet.Rows.Count; ++r)
            {
                if (++null_counter > 10)
                {
                    // end of data
                    break;
                }

                if (exposure_count_table_sheet.Cells[r, 1].Value == null || !match_single_special_word(exposure_count_table_sheet.Cells[r, 1].Value.Trim(), ref exposure_organization_regex))
                    continue;

                null_counter = 0;

                int temp_r = r;

                string police_name = "";
                {
                    ++temp_r;

                    MatchCollection matches = exposure_police_name_regex.Matches(exposure_count_table_sheet.Cells[temp_r, 1].Value.Trim());
                    if (matches.Count != 1 || matches[0].Groups.Count != 2)
                    {
                        log("錯誤!舉發員警之格式有誤!" + Environment.NewLine + "該句為 : " + exposure_count_table_sheet.Cells[temp_r, 1].Value.Trim());
                        continue;
                    }

                    police_name = matches[0].Groups[1].Value.Trim();
                }

                Achievement achievement = new Achievement(police_name);
                log("開始統計警員 : " + police_name + " 之績效" + Environment.NewLine);

                temp_r += 2;

                for (; temp_r <= exposure_count_table_sheet.Rows.Count; ++temp_r)
                {
                    if (exposure_count_table_sheet.Cells[temp_r, 1].Value != null && exposure_count_table_sheet.Cells[temp_r, 1].Value.GetType() == typeof(string) && match_single_special_word(exposure_count_table_sheet.Cells[temp_r, 1].Value.Trim(), ref exposure_amount_regex))
                    {
                        // end of data
                        break;
                    }

                    string law_serial_number = exposure_count_table_sheet.Cells[temp_r, 1].Value.ToString();
                    string law_content = exposure_count_table_sheet.Cells[temp_r, 2].Value.Trim();
                    double stopping_count = exposure_count_table_sheet.Cells[temp_r, 3].Value;
                    double exposure_count = exposure_count_table_sheet.Cells[temp_r, 4].Value;
                    double amount = exposure_count_table_sheet.Cells[temp_r, 5].Value;

                    if (match_single_special_word(law_serial_number, ref Achievement.DWI_regex))
                    {
                        // found
                        achievement.DWI_score += amount;
                        log("該法條代碼 " + law_serial_number + " 為 酒駕 件數為 : " + amount.ToString() + Environment.NewLine + "違規法條名稱為 : " + law_content);
                    }
                    else if (match_single_special_word(law_serial_number, ref Achievement.running_the_red_light_stopping_regex))
                    {
                        achievement.running_the_red_light_stopping_score += stopping_count;
                        log("該法條代碼 " + law_serial_number + " 為 闖紅燈_攔停 件數為 : " + stopping_count.ToString() + Environment.NewLine + "違規法條名稱為 : " + law_content);
                    }
                    else if (match_single_special_word(law_serial_number, ref Achievement.license_plate_related_regex))
                    {
                        achievement.license_plate_related_score += amount;
                        log("該法條代碼 " + law_serial_number + " 為 牌照 件數為 : " + amount.ToString() + Environment.NewLine + "違規法條名稱為 : " + law_content);
                    }
                    else if (match_single_special_word(law_serial_number, ref Achievement.motorcyle_in_wrong_way_regex))
                    {
                        achievement.motorcyle_in_wrong_way_score += amount;
                        log("該法條代碼 " + law_serial_number + " 為 機車行駛禁行機車道 件數為 : " + amount.ToString() + Environment.NewLine + "違規法條名稱為 : " + law_content);
                    }
                    else if (match_single_special_word(law_serial_number, ref Achievement.reverse_driving_regex))
                    {
                        achievement.reverse_driving_score += amount;
                        log("該法條代碼 " + law_serial_number + " 為 逆向行駛 件數為 : " + amount.ToString() + Environment.NewLine + "違規法條名稱為 : " + law_content);
                    }
                    else if (match_single_special_word(law_serial_number, ref Achievement.unused_headlight_regex))
                    {
                        achievement.unused_headlight_score += amount;
                        log("該法條代碼 " + law_serial_number + " 為 未開大燈 件數為 : " + amount.ToString() + Environment.NewLine + "違規法條名稱為 : " + law_content);
                    }
                    else if (match_single_special_word(law_serial_number, ref Achievement.violating_pedestrian_right_regex))
                    {
                        achievement.violating_pedestrian_right_score += amount;
                        log("該法條代碼 " + law_serial_number + " 為 汽機車違反行人路權 件數為 : " + amount.ToString() + Environment.NewLine + "違規法條名稱為 : " + law_content);
                    }
                    else if (match_single_special_word(law_serial_number, ref Achievement.unused_safety_helmet_regex))
                    {
                        achievement.unused_safety_helmet_score += amount;
                        log("該法條代碼 " + law_serial_number + " 為 未戴安全帽 件數為 : " + amount.ToString() + Environment.NewLine + "違規法條名稱為 : " + law_content);
                    }
                    else if (match_single_special_word(law_serial_number, ref Achievement.dangerously_driving_regex))
                    {
                        achievement.dangerously_driving_score += amount;
                        log("該法條代碼 " + law_serial_number + " 為 危駕 件數為 : " + amount.ToString() + Environment.NewLine + "違規法條名稱為 : " + law_content);
                    }
                    else if (match_single_special_word(law_serial_number, ref Achievement.gravel_truck_related_regex))
                    {
                        achievement.gravel_truck_related_score += amount;
                        log("該法條代碼 " + law_serial_number + " 為 砂石車 件數為 : " + amount.ToString() + Environment.NewLine + "違規法條名稱為 : " + law_content);
                    }
                    else
                    {
                        log("該法條代碼 " + law_serial_number + " 找無對應之績效");
                    }

                    log("");
                }

                achievement_list.Add(achievement);
                r = temp_r + 1;

                log("完成!" + Environment.NewLine);

                //return;
            }

            exposure_count_table_book.Close();
            exposure_count_table_book = null;
            exposure_count_table_sheets = null;
            exposure_count_table_sheet = null;

            log("舉發件數統計表 已完成!" + Environment.NewLine);

            Excel._Workbook motorcycle_table_book;
            Excel.Sheets motorcycle_table_sheets;
            excel_open_book_and_sheets(motorcycle_table_path, out motorcycle_table_book, out motorcycle_table_sheets);
            Excel._Worksheet motorcycle_table_sheet = (Excel.Worksheet)motorcycle_table_sheets.Item[1];

            Dictionary<string, double> motorcycle_police_achievements = new Dictionary<string, double>();

            for (int r = motorcycle_table_row_start; r <= motorcycle_table_sheet.Rows.Count; ++r)
            {
                if (motorcycle_table_sheet.Cells[r, 3].Value == null)
                {
                    // end of data
                    break;
                }

                string law_serial_number = motorcycle_table_sheet.Cells[r, 3].Value.ToString();
                {
                    if (!match_single_special_word(law_serial_number, ref Achievement.car_turning_regex))
                    {
                        log("該法條代碼 " + law_serial_number + " 不為 轉彎未依規定(汽車)");
                        continue;
                    }
                }

                List<string> polices = new List<string>();
                for (int c = 0; c < 4; ++c)
                {
                    if (motorcycle_table_sheet.Cells[r, 5 + c].Value == null)
                        break;

                    polices.Add(motorcycle_table_sheet.Cells[r, 5 + c].Value.Trim());
                }

                for (int p = 0; p < polices.Count; ++p)
                {
                    string n = polices[p];
                    if (motorcycle_police_achievements.ContainsKey(n))
                    {
                        motorcycle_police_achievements[n] += 1.0 / polices.Count;
                    }
                    else
                    {
                        motorcycle_police_achievements[n] = 1.0 / polices.Count;
                    }
                }
            }

            foreach (KeyValuePair<string, double> pair in motorcycle_police_achievements)
            {
                for (int i = 0; i < achievement_list.Count; ++i)
                {
                    if (achievement_list[i].name == pair.Key)
                    {
                        // found
                        //achievement_list[i].car_turning_score = achievement_list[i].car_turning_score - pair.Value;
                        achievement_list[i].motorcycle_two_way_turning_left_score = pair.Value;
                        break;
                    }
                }
            }

            motorcycle_table_book.Close();
            motorcycle_table_book = null;
            motorcycle_table_sheets = null;
            motorcycle_table_sheet = null;

            Excel._Workbook achievement_table_book;
            Excel.Sheets achievement_table_sheets;
            excel_open_book_and_sheets(achievement_table_path, out achievement_table_book, out achievement_table_sheets);
            Excel._Worksheet achievement_table_sheet = (Excel.Worksheet)achievement_table_sheets.Item[2];

            ////108 年 10月28日 - 11月3日 交通隊分隊部 舉發交通違規評比項目達成率總表
            //// 108 年 5月 交通隊分隊部 舉發交通違規評比項目達成率總表      
            //{
            //    string title = "108 年 " + 12.ToString("D2") + "月 交通隊分隊部 舉發交通違規評比項目達成率總表";
            //    achievement_table_sheet.Cells[1, 6] = title;
            //}

            for (int r = achievement_table_row_start; r <= achievement_table_sheet.Rows.Count; ++r)
            {
                if (!match_single_special_word(achievement_table_sheet.Cells[r, 6].Value.Trim(), ref achievement_table_police_regex))
                {
                    // end of data
                    log("績效表完成!" + Environment.NewLine);
                    break;
                }

                string name = achievement_table_sheet.Cells[r, 7].Value.Trim();

                bool found = false;
                for (int i = 0; i < achievement_list.Count; ++i)
                {
                    if (achievement_list[i].name == name)
                    {
                        // match
                        found = true;

                        achievement_table_sheet.Cells[r, 9] = achievement_list[i].DWI_score;
                        achievement_table_sheet.Cells[r, 15] = achievement_list[i].running_the_red_light_stopping_score;
                        achievement_table_sheet.Cells[r, 21] = achievement_list[i].license_plate_related_score;
                        achievement_table_sheet.Cells[r, 27] = achievement_list[i].motorcyle_in_wrong_way_score;
                        achievement_table_sheet.Cells[r, 33] = achievement_list[i].motorcycle_two_way_turning_left_score;

                        // 外送
                        //achievement_table_sheet.Cells[r, 39] = achievement_list[i].running_the_red_light_exposure_score;

                        achievement_table_sheet.Cells[r, 45] = achievement_list[i].reverse_driving_score;
                        achievement_table_sheet.Cells[r, 51] = achievement_list[i].unused_headlight_score;
                        achievement_table_sheet.Cells[r, 57] = achievement_list[i].violating_pedestrian_right_score;
                        achievement_table_sheet.Cells[r, 63] = achievement_list[i].unused_safety_helmet_score;
                        achievement_table_sheet.Cells[r, 69] = achievement_list[i].dangerously_driving_score;
                        achievement_table_sheet.Cells[r, 75] = achievement_list[i].gravel_truck_related_score;

                        achievement_table_book.Save();

                        log("警員 : " + name + " 績效完成!" + Environment.NewLine);

                        achievement_list.RemoveAt(i);
                        break;
                    }
                }

                if (!found)
                {
                    achievement_table_sheet.Cells[r, 9] = 0;
                    achievement_table_sheet.Cells[r, 15] = 0;
                    achievement_table_sheet.Cells[r, 21] = 0;
                    achievement_table_sheet.Cells[r, 27] = 0;
                    achievement_table_sheet.Cells[r, 33] = 0;
                    // 外送
                    //achievement_table_sheet.Cells[r, 39] = 0;

                    achievement_table_sheet.Cells[r, 45] = 0;
                    achievement_table_sheet.Cells[r, 51] = 0;
                    achievement_table_sheet.Cells[r, 57] = 0;
                    achievement_table_sheet.Cells[r, 63] = 0;
                    achievement_table_sheet.Cells[r, 69] = 0;
                    achievement_table_sheet.Cells[r, 75] = 0;

                    achievement_table_book.Save();

                    log("警員 : " + name + " 找無績效!" + Environment.NewLine);
                }
            }

            achievement_table_book.Close();
            achievement_table_book = null;
            achievement_table_sheets = null;
            achievement_table_sheet = null;

            excel_app_close();
        }

        private void achievement_table_generator_tsmi_Click(object sender, EventArgs e)
        {
            if (folder_dialog_0.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    log("根目錄為 : " + folder_dialog_0.SelectedPath + Environment.NewLine);
                    base_dir_path = folder_dialog_0.SelectedPath;
                }
                catch (Exception ex)
                {
                    log("錯誤!錯誤訊息為 : " + ex.Message + Environment.NewLine + "ex.StackTrace 為 : " + ex.StackTrace);
                }

                achievement_table_generator();
            }
        }
    }
}
