using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;


using Excel = Microsoft.Office.Interop.Excel;

namespace CS_AddressTest
{
    class Program
    {
        static void Main(string[] args)
        {
            //string path = @"E:\VS_Projects\CS_AddressTest\CS_AddressTest\법정동코드 전체자료_utf8.txt";
            string sCur = System.Environment.CurrentDirectory;
            string path = sCur + "\\20210118_단지_기본정보.xls";

            //string TxtPth = sCur + "\\AptCode_AddressTable.txt";
            //string textVal = File.ReadAllText(TxtPth);

            //for(int i = 0; i<textVal.Length; i++)
            //{
            //    string tee = textVal[i];
            //    int reare = 1;
            //}

            Excel.Application ExcelApp = new Excel.Application();
            Excel.Workbook workbook = ExcelApp.Workbooks.Open(path);
            Excel.Worksheet worksheet1 = workbook.Worksheets["Sheet 1"];
            ExcelApp.Visible = true;

            Excel.Range AptCodeCol = worksheet1.get_Range("E:E");
            Excel.Range AptNameCol = worksheet1.get_Range("F:F");
            Excel.Range AdminAddressCol = worksheet1.get_Range("H:H");
            Excel.Range RoadAddressCol = worksheet1.get_Range("I:I");

            string RetString = null;
            string fPath = "AptCode_AddressTable.txt";
            List<string> tempRetList = new List<string>();

            for (int RowIdx = 3; ; RowIdx++)
            {
                Excel.Range rAptCode = AptCodeCol.Cells[RowIdx, 1];
                string sAptCode = rAptCode.Value;

                Excel.Range rAptName = AptNameCol.Cells[RowIdx, 1];
                string sAptName = rAptName.Value;

                Excel.Range rAdminAddr = AdminAddressCol.Cells[RowIdx, 1];
                string sAdminAddr = rAdminAddr.Value;

                Excel.Range rRoadAddr = RoadAddressCol.Cells[RowIdx, 1];
                string sRoadAddr = rRoadAddr.Value;


                if (sAptCode == null)
                    break;
                
                //if (sAdminAddr == null || sRoadAddr == null)
                //    continue; 

                Debug.WriteLine("{0}, {1}, {2}, {3} / {4}", sAptCode, sAptName, sAdminAddr, sRoadAddr, RowIdx);
                RetString = sAdminAddr + "\t" + sRoadAddr + "\t" + sAptName + "\t" + sAptCode;
                tempRetList.Add(RetString);
            }

            tempRetList.Sort();
            File.WriteAllLines(fPath, tempRetList);
        }
    }
}
