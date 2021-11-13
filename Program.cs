using System.Text;
using System.Net;
using System.IO;
using System;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using Microsoft.CSharp;

namespace coderandom
{
    class Program
    {
        static Random random = new Random();
        static string[] sname = { ", " };
        static string[] fnames = File.ReadAllText(@"D:\211\DB\lab\BTL2\coderandom\ten.txt").Split(sname, StringSplitOptions.RemoveEmptyEntries);
        static string[] lnames = File.ReadAllText(@"D:\211\DB\lab\BTL2\coderandom\ho.txt").Split(sname, StringSplitOptions.RemoveEmptyEntries);
        private static readonly string[] VietnameseSigns = File.ReadAllText(@"D:\211\DB\lab\BTL2\coderandom\codau.txt").Split(sname, StringSplitOptions.RemoveEmptyEntries);
        public static string RemoveSign4VietnameseString(string str)
        {
            for (int i = 1; i < VietnameseSigns.Length; i++)
            {
                for (int j = 0; j < VietnameseSigns[i].Length; j++)
                    str = str.Replace(VietnameseSigns[i][j], VietnameseSigns[0][i]);
            }
            return str;
        }
        static string randomLname()
        {
            int n = random.Next(lnames.Length);
            return lnames[n];
        }
        static string randomFname()
        {
            int n = random.Next(fnames.Length);
            return fnames[n];
        }
        static string randomAcademicYear()
        {
            return random.Next(19, 22).ToString();
        }
        static string randomStatus()
        {
            int r = random.Next(0, 101);
            if (r == 0)
                return "Bảo lưu";
            return "Đang học";
        }
        static string randomDateOfBirth(string aYear)
        {
            string y = "", m = "", d = "";
            if (aYear == "19")
                y = "2001";
            else if (aYear == "20")
                y = "2002";
            else if (aYear == "21")
                y = "2003";

            int mm = random.Next(1, 13);
            int dd = 0;
            if (mm == 2)
                dd = random.Next(1, 29);
            else if (mm == 4 || mm == 6 || mm == 9 || mm == 11)
                dd = random.Next(1, 31);
            else
                dd = random.Next(1, 32);
            if (mm < 10)
                m = "0" + mm.ToString();
            else
                m = mm.ToString();

            if (dd < 10)
                d = "0" + dd.ToString();
            else
                d = dd.ToString();

            return y + "-" + m + "-" + d;
        }
        static string randomSex()
        {
            int r = random.Next(0, 3);
            if (r == 0)
                return "Nam";
            else if (r == 1)
                return "Nữ";
            return "Khác";
        }
        static string randomCMND()
        {
            string s = "";
            for (int i = 0; i < 9; i++)
                s += random.Next(0, 9).ToString();
            return "'" + s;
        }
        static string randomEmail1(string fname, string lname)
        {
            string ss = RemoveSign4VietnameseString(fname).ToLower() + "." + RemoveSign4VietnameseString(lname).ToLower();
            int n = random.Next(1, 10);
            for (int i = 0; i <= n; i++)
            {
                int x = random.Next(97, 123);
                ss += ((char)x);
            }
            return ss + "@hcmut.edu.vn";
        }
        static string randomEmail2()
        {
            string ss = "";
            int n = random.Next(1, 10);
            for (int i = 0; i <= n; i++)
            {
                int x = random.Next(97, 123);
                ss += ((char)x);
            }
            return ss + "@gmail.com";
        }
        static string randomAddress1()
        {
            string pp = "", dd = "", ww = "";
            try
            {
                dynamic listP = JsonConvert.DeserializeObject(File.ReadAllText(@"D:\211\DB\lab\BTL2\coderandom\local.json"));
                int n = random.Next(0, 63);
                dynamic p = listP[n];
                pp = p["name"];
                n = random.Next(0, p["districts"].Count);

                dynamic d = p["districts"][n];

                dd = d["name"];

                n = random.Next(0, d["wards"].Count);

                ww = d["wards"][n]["prefix"] + " " + d["wards"][n]["name"];

            }
            catch { }
            return ww + " - " + dd + " - " + pp;
        }
        static string randomAddress2()
        {
            string pp = "", dd = "", ww = "";
            try
            {
                dynamic listP = JsonConvert.DeserializeObject(File.ReadAllText(@"D:\211\DB\lab\BTL2\coderandom\local.json"));
                int n;
                dynamic p = listP[0];
                pp = p["name"];
                n = random.Next(0, p["districts"].Count);

                dynamic d = p["districts"][n];

                dd = d["name"];

                n = random.Next(0, d["wards"].Count);

                ww = d["wards"][n]["prefix"] + " " + d["wards"][n]["name"];
            }
            catch { }
            return ww + " - " + dd + " - " + pp;
        }
        static void randomSinhvien()
        {
            object misvalue = System.Reflection.Missing.Value;
            Application oXl = new Application();
            // Workbook oWb = oXl.Workbooks.Add();
            Workbook oWb = oXl.Workbooks.Open(@"D:\211\DB\lab\BTL2\coderandom\Sinhvien.xlsx");
            Worksheet oWs = (Worksheet)oWb.Worksheets.Item[1];

            // oWs.Cells[1, 1] = "MSSV";
            // oWs.Cells[1, 2] = "Họ và tên lót";
            // oWs.Cells[1, 3] = "Tên";
            // oWs.Cells[1, 4] = "Khóa";
            // oWs.Cells[1, 5] = "Tình trạng";
            // oWs.Cells[1, 6] = "Ngày sinh";
            // oWs.Cells[1, 7] = "Giới tính";
            // oWs.Cells[1, 8] = "CMND/CCCD";
            // oWs.Cells[1, 9] = "Địa chỉ tạm trú";
            // oWs.Cells[1, 10] = "Hộ khẩu";
            // oWs.Cells[1, 11] = "Email trường";
            // oWs.Cells[1, 12] = "Email cá nhân";
            // oWs.Cells[1, 13] = "Tên lớp chủ nhiệm";
            // oWs.Cells[1, 14] = "Mã khoa";

            for (int i = 2; i < 502; i++)
            {
                int a = Convert.ToInt32((oWs.Cells[i, 4] as Range).Value2.ToString());
                oWs.Cells[i, 1] = (a * 100000 + i - 1).ToString();

                // oWs.Cells[i, 2] = randomLname(); // random
                // oWs.Cells[i, 3] = randomFname(); // random
                // oWs.Cells[i, 4] = randomAcademicYear(); // random
                // oWs.Cells[i, 5] = randomStatus(); //
                // oWs.Cells[i, 6] = randomDateOfBirth((oWs.Cells[i, 4] as Range).Value2.ToString());
                // oWs.Cells[i, 7] = randomSex();
                // oWs.Cells[i, 8] = randomCMND();
                // oWs.Cells[i, 9] = randomAddress2();
                // oWs.Cells[i, 10] = randomAddress1();
                Console.WriteLine(i);
                // oWs.Cells[i, 11] = randomEmail1((oWs.Cells[i, 3] as Range).Value2.ToString(), (oWs.Cells[i, 2] as Range).Value2.ToString());
                // oWs.Cells[i, 12] = randomEmail2();
                // oWs.Cells[i, 13] = "NULL";
                // oWs.Cells[i, 14] = "NULL";
            }

            // oWb.SaveAs(@"D:\211\DB\lab\BTL2\coderandom\Sinhvien.xlsx", XlFileFormat.xlWorkbookDefault, misvalue, misvalue, misvalue, misvalue, XlSaveAsAccessMode.xlShared, misvalue, misvalue, misvalue, misvalue, misvalue);
            oWb.Save();
            oWb.Close(true, misvalue, misvalue);
            oXl.Quit();
        }

        static void randomPN()
        {
            object misvalue = System.Reflection.Missing.Value;
            Application oXl = new Application();
            // Workbook oWb = oXl.Workbooks.Add();
            Workbook oWb = oXl.Workbooks.Open(@"D:\211\DB\lab\BTL2\coderandom\Sinhvien.xlsx");
            Worksheet oWs = (Worksheet)oWb.Worksheets.Item[1];
            Worksheet oWs2 = (Worksheet)oWb.Worksheets.Item[2];


            oWb.Save();
            oWb.Close(true, misvalue, misvalue);
            oXl.Quit();
        }
        static void Main(string[] args)
        {
            // randomSinhvien();
            randomPN();
        }
    }
}
