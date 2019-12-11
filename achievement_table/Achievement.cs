using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

namespace achievement_table
{
    class Achievement
    {
        static string DWI_pattern = @"(^35\d+)|(^731\d+)";
        public static Regex DWI_regex = new Regex(DWI_pattern, RegexOptions.Compiled);
        static string running_the_red_light_stopping_pattern = @"^53\d+";
        public static Regex running_the_red_light_stopping_regex = new Regex(running_the_red_light_stopping_pattern, RegexOptions.Compiled);
        static string license_plate_related_pattern = @"(^12\d+)|(^13\d+)";
        public static Regex license_plate_related_regex = new Regex(license_plate_related_pattern, RegexOptions.Compiled);
        static string car_turning_pattern = @"^4810201";
        public static Regex car_turning_regex = new Regex(car_turning_pattern, RegexOptions.Compiled);
        static string motorcyle_in_wrong_way_pattern = @"^4511301";
        public static Regex motorcyle_in_wrong_way_regex = new Regex(motorcyle_in_wrong_way_pattern, RegexOptions.Compiled);

        static string reverse_driving_pattern = @"(^4510301\d*)|(^4510101)";
        public static Regex reverse_driving_regex = new Regex(reverse_driving_pattern, RegexOptions.Compiled);
        static string unused_headlight_pattern = @"^42\d+";
        public static Regex unused_headlight_regex = new Regex(unused_headlight_pattern, RegexOptions.Compiled);
        static string violating_pedestrian_right_pattern = @"^482\d+";
        public static Regex violating_pedestrian_right_regex = new Regex(violating_pedestrian_right_pattern, RegexOptions.Compiled);
        static string unused_safety_helmet_pattern = @"(^3160001\d*)|(^3160002)|(^734\d+)";
        public static Regex unused_safety_helmet_regex = new Regex(unused_safety_helmet_pattern, RegexOptions.Compiled);
        static string dangerously_driving_pattern = @"(^16\d+)|(^18\d+)|(^21\d+)|(^22\d+)";
        public static Regex dangerously_driving_regex = new Regex(dangerously_driving_pattern, RegexOptions.Compiled);
        static string gravel_truck_related_pattern = @"(^29\d+)|(^30\d+)|(^60\d+)";
        public static Regex gravel_truck_related_regex = new Regex(gravel_truck_related_pattern, RegexOptions.Compiled);

        public Achievement(string n)
        {
            name = n;

            DWI_score = 0;
            running_the_red_light_stopping_score = 0;
            license_plate_related_score = 0;
            car_turning_score = 0;
            motorcycle_two_way_turning_left_score = 0;
            reverse_driving_score = 0;
            unused_headlight_score = 0;
            violating_pedestrian_right_score = 0;
            unused_safety_helmet_score = 0;
            dangerously_driving_score = 0;
            gravel_truck_related_score = 0;
        }

        //public bool check_law(ref string serial_number, ref string content, ref double stopping_count, ref double exposure_count, ref double amount)
        //{
        //    if ()
        //}

        public string name = "";

        public double DWI_score, running_the_red_light_stopping_score, license_plate_related_score;
        public double car_turning_score, motorcycle_two_way_turning_left_score, reverse_driving_score;
        public double unused_headlight_score, violating_pedestrian_right_score;
        public double unused_safety_helmet_score, dangerously_driving_score, gravel_truck_related_score;

        public double motorcyle_in_wrong_way_score;
        public double running_the_red_light_exposure_score;
    }
}
