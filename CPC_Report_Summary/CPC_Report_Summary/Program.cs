using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using Application2 = System.Windows.Forms;
using System.Net.Mail;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace CPC_Report_Summary
{
    class Program
    {
        private static string _subject = "";
        public static string subject { get { return _subject; } set { _subject = value; } }

        public static string projectCode = "CD24";
        public static string programMode = "1";
        public static string weekPeriod = "5";
        public static string currentWeek = "1";
        public static string mReportDetailEmail = "";
        public static string mReportDetailLog = "";
        public static string mCodingDetailEmail = "";
        public static string mCodingDetailLog = "";
        public static string mSampleFileDetailEmail = "";
        public static string mSampleFileDetailLog = "";
        public static string DialogLogDetails = "";
        public static string FilePath = "";
        public static string FileName = "";

        public static bool isCodingError = false;

        public static int[] missingAgentRow_Initial_List = new int[10000];
        public static int missingAgentRow_Initial_Index = 0;
        public static string[] InitialSurveyId_List = new string[10000];
        public static int InitialSurveyId_Index = 0;

        public static int[] missingAgentRow_Transferred_List = new int[10000];
        public static int missingAgentRow_Transferred_Index = 0;
        public static string[] TransferredSurveyId_List = new string[10000];
        public static int TransferredSurveyId_Index = 0;

        public static int OthersQ3, OthersQ3A, OthersQ3B, OthersQ3C, OthersQ3D, OthersQ3E, OthersQ4 = 0;
        public static string SaveFilePathName = "";

        public static string qid = "";

        public static string[] questionList = new string[10000];
        public static int questionIndex = 0;
        public static string[] headerList = new string[10000];
        public static int headerIndex = 0;

        public static string[] filePathToAttach = new string[10];
        public static int fileID = 0;

        private static int _lastRow = 1;
        public static int lastRow { get { return _lastRow; } set { _lastRow = value; } }

        private static string _SampleFilePath = "";
        public static string SampleFilePath { get { return _SampleFilePath; } set { _SampleFilePath = value; } }

        delegate void MessageDelegate1(MailMessage mailMessage);
        public static void MessageDisplay(MailMessage mailMsg, bool showBody)
        {
            MessageDelegate1 md1 = new MessageDelegate1(ShowMessageSubject);
            md1 += ShowMessageBody;
            if (!showBody)
                md1 -= ShowMessageBody;
            md1(mailMsg);
        }

        /*  METHOD: SHOW MESSAGE SUBJECT CALLBACK   */
        public static void ShowMessageSubject(MailMessage msg)
        {
            System.Diagnostics.Debug.WriteLine("message subject: " + msg.Subject + "\n");
        }

        /*  METHOD: SHOW MESSAGE BODY CALLBACK   */
        public static void ShowMessageBody(MailMessage msg)
        {
            System.Diagnostics.Debug.WriteLine("message body: " + msg.Body + "\n");
        }

        /* METHOD: RETURNS LAST ROW NUMBER THAT HAS VALUES */
        public static int getLastRowNumberWithValues(Worksheet sheet, string ColumnRange)
        {
            Microsoft.Office.Interop.Excel.Range endRange = sheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            Microsoft.Office.Interop.Excel.Range range = sheet.get_Range("A1", endRange);

            lastRow = endRange.Row;
            string cellVal = "";
            for (int i = 1; i < lastRow; i++)
            {
                try
                {
                    cellVal = sheet.Range[ColumnRange + i.ToString() + ":" + ColumnRange + i.ToString()].Value.ToString();
                }
                catch
                {
                    cellVal = "";
                }
                if (cellVal == "")
                {
                    System.Diagnostics.Debug.WriteLine("Last Row of " + sheet.Name + " is:\t" + (i - 1).ToString() + "\n");
                    Console.WriteLine("Last Row of " + sheet.Name + " is:\t" + (i - 1).ToString() + "\n");
                    return (i - 1);
                }
            }
            System.Diagnostics.Debug.WriteLine("Last Row of_ " + sheet.Name + " is:\t" + lastRow + "\n");
            Console.WriteLine("Last Row of_ " + sheet.Name + " is:\t" + lastRow + "\n");
            return lastRow;
        }


        /* METHOD: RETURNS LAST ROW NUMBER OF GIVEN EXCEL FILE */
        public static void getLastRowNumber(Worksheet sheet, string SampleFilePath)
        {
            Microsoft.Office.Interop.Excel.Range endRange = sheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            Microsoft.Office.Interop.Excel.Range range = sheet.get_Range("A1", endRange);

            lastRow = endRange.Row;

            //Store Range values
            int dictionaryID = 1;
            int startRow = 2;
            int currentRow = startRow;

            //STORE QIDS IN DICTIONARY
            while (currentRow <= lastRow)
            {
                currentRow++;
                dictionaryID++;
            }
            System.Diagnostics.Debug.WriteLine("Last Row of " + sheet.Name + " is:\t" + lastRow + "\n");
            Console.WriteLine("Last Row of " + sheet.Name + " is:\t" + lastRow + "\n");
        }



        /* METHOD: SETUP CODING FILE TO IMPORT (SINGLE MENTION TYPES ONLY) */
        public static void CPC_Report_Summary(Microsoft.Office.Interop.Excel.Application xlApp, string mySampleFilePath)
        {
            DateTimeOffset dateToday;
            string dateTodayString = "";
            dateToday = DateTimeOffset.Now;
            dateTodayString = dateToday.ToString("yyyyMMdd");


            SaveFilePathName = @"\\CC3-MGMT2\EXPORT\" + projectCode + @"\Weekly Data Report Summary " + projectCode + "_" + dateTodayString + ".xlsx";
            System.IO.File.Delete(SaveFilePathName);

            int integerValue = 0;
            string stringValue = "";
            int cCnt;
            int rCnt;
            int rw = 0;
            int cl = 0;
            float progress = 0.00f;

            //OPEN UP (NECESSARY FILES) AND OBTAIN LAST ROW NUMBER USED RESPECTIVELY
            SampleFilePath = mySampleFilePath;
            Microsoft.Office.Interop.Excel.Workbook DataFile_WB = xlApp.Workbooks.Open(SampleFilePath);
            Worksheet DataFile_SHEET = (Worksheet)DataFile_WB.Sheets[1];
            getLastRowNumber(DataFile_SHEET, SampleFilePath);

            //READ SAMPLEFILE DATA FOR GIVEN "RESPONDENT ROW"
            Microsoft.Office.Interop.Excel.Range range = (Microsoft.Office.Interop.Excel.Range)DataFile_SHEET.UsedRange;
            range = DataFile_SHEET.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;

            for (rCnt = 1; rCnt <= rw; rCnt++)
            {
                progress = (float)rCnt / (float)rw * 100;
                Console.WriteLine("Progress: " + progress + "%");

                for (cCnt = 1; cCnt <= cl; cCnt++)
                {
                    try
                    {
                        integerValue = (int)(range.Cells[rCnt, cCnt].Value2);
                        stringValue = integerValue.ToString();
                    }
                    catch
                    {
                        stringValue = (string)(range.Cells[rCnt, cCnt].Value2);
                    }
                    if (rCnt == 1)
                    {
                        headerList[cCnt] = stringValue;
                    }

                    //ACTUAL DATA (WITHOUT HEADER)
                    if (rCnt > 1)
                    {
                        questionList[cCnt] = stringValue;
                        //1) MISSING AGENT USERS (INITIAL CALL)
                        if (cCnt == 22)
                        {
                            if (questionList[cCnt] == null || questionList[cCnt] == "vacant?" || questionList[cCnt] == "XXXX")
                            {
                                Console.WriteLine("val: " + DataFile_SHEET.Range["D" + rCnt].Value);
                                missingAgentRow_Initial_List[missingAgentRow_Initial_Index] = rCnt;
                                InitialSurveyId_List[InitialSurveyId_Index] = (string)DataFile_SHEET.Range["D" + rCnt].Value;
                                missingAgentRow_Initial_Index++;
                                InitialSurveyId_Index++;
                            }
                        }
                        //2) MISSING AGENT USERS (TRANSFERRED CALL)
                        if (cCnt == 31)
                        {
                            if (questionList[3] == "Yes" && (questionList[cCnt] == null || questionList[cCnt] == "vacant?" || questionList[cCnt] == "XXXX"))
                            {
                                missingAgentRow_Transferred_List[missingAgentRow_Transferred_Index] = rCnt;
                                TransferredSurveyId_List[TransferredSurveyId_Index] = (string)DataFile_SHEET.Range["D" + rCnt].Value;
                                missingAgentRow_Transferred_Index++;
                                TransferredSurveyId_Index++;
                            }
                        }
                    }
                }
            }

            //COLUMN WIDTHS
            DataFile_SHEET.Columns["A:BS"].ColumnWidth = 20.71f;
            //Save RedAlertFile
            DataFile_WB.SaveAs(SaveFilePathName, XlFileFormat.xlOpenXMLWorkbook, Missing.Value, Missing.Value, false, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlUserResolution, true, Missing.Value, Missing.Value, Missing.Value);
            //Close all excel workbooks
            DataFile_WB.Close(false);
            //AttachFile
            filePathToAttach[fileID] = SaveFilePathName;
            fileID++;
        }

        public static void AnalyzeDailyLoadedPerQuota(Microsoft.Office.Interop.Excel.Application xlApp, string mySampleFilePath)
        {
            int ESCA_Loaded = 0, ESCC_Loaded = 0, ESCH_Loaded = 0, PHIL_Loaded = 0, HPDK_Loaded = 0, CMBC_Loaded = 0, BUSN_Loaded = 0, CONSA_Loaded = 0, CONSH_Loaded = 0, CONSC_Loaded = 0, EPST_Loaded = 0, CM12_Loaded = 0, CML3_Loaded = 0, CML4_Loaded = 0, CML5_Loaded = 0;

            DateTimeOffset dateToday;
            string dateTodayString = "";
            dateToday = DateTimeOffset.Now;
            dateTodayString = dateToday.ToString("yyyyMMdd");

            SaveFilePathName = @"\\CC3-MGMT2\EXPORT\" + projectCode + @"\" + projectCode + "_Loaded_Per_Quota_Report.xlsx";
            string AltSavePathName = @"\\QNAP\CATI\CPC\Cumulative Reports\" + projectCode + "_Loaded_Per_Quota_Report.xlsx";
            try { System.IO.File.Delete(SaveFilePathName); } catch { }
            try { System.IO.File.Delete(AltSavePathName); } catch { }

            Console.WriteLine("Analyzing Loaded Sample Counts for:\t" + mySampleFilePath);

            int integerValue = 0;
            string stringValue = "";

            //OPEN UP (NECESSARY FILES) AND OBTAIN LAST ROW NUMBER USED RESPECTIVELY
            SampleFilePath = mySampleFilePath;
            Microsoft.Office.Interop.Excel.Workbook DataFile_WB = xlApp.Workbooks.Open(SampleFilePath);
            Worksheet DataFile_SHEET = (Worksheet)DataFile_WB.Sheets[1];
            getLastRowNumber(DataFile_SHEET, SampleFilePath);


            //INITIAL CALCULATION TO DETERMINE IF SAMPLE IS AN: INITIAL OR TRANSFERRED CALL TYPE RECORD
            DataFile_SHEET.Range["AAA2"].Formula = "=IF(CY2=\"\",\"INITIAL: \" &U2,\"TRANSFERRED: \"&U2)";
            DataFile_SHEET.Range["AAA2"].AutoFill(DataFile_SHEET.Range["AAA2:AAA" + lastRow]);

            //Make the summary (DEDUPE IMPORT DATES)
            DataFile_SHEET.Range["H1:H" + lastRow].Copy();
            Worksheet All_Loaded_Summary_Sheet = xlApp.Worksheets.Add();
            System.Drawing.Color currentColor = System.Drawing.Color.FromArgb(35, 45, 55);
            All_Loaded_Summary_Sheet.Range["A:XFD"].Interior.Color = currentColor;
            All_Loaded_Summary_Sheet.Name = "CPC - Daily Import Summary";
            All_Loaded_Summary_Sheet.Activate();
            All_Loaded_Summary_Sheet.Range["A2"].PasteSpecial();
            object cols = new object[] { 1 };
            All_Loaded_Summary_Sheet.Range["A2:A" + (lastRow + 1)].RemoveDuplicates(cols, XlYesNoGuess.xlYes);

            //Header
            All_Loaded_Summary_Sheet.Range["B1:P1"].MergeCells = true;
            All_Loaded_Summary_Sheet.Range["Q1:AE1"].MergeCells = true;
            All_Loaded_Summary_Sheet.Range["A1:AE1"].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            All_Loaded_Summary_Sheet.Range["A1:AE1"].VerticalAlignment = XlVAlign.xlVAlignCenter;
            All_Loaded_Summary_Sheet.Range["A1"].Value = "Last Run: " + dateTodayString;
            All_Loaded_Summary_Sheet.Range["A1"].WrapText = true;
            All_Loaded_Summary_Sheet.Range["B1:P1"].Value = "[INITIAL CALL] - LOADED (DATE OF IMPORT BY QUOTAS)";
            All_Loaded_Summary_Sheet.Range["Q1:AE1"].Value = "[TRANSFERRED CALL] - LOADED (DATE OF IMPORT BY QUOTAS)";
            All_Loaded_Summary_Sheet.Range["B2"].Value = "BUSN";
            All_Loaded_Summary_Sheet.Range["C2"].Value = "CM12";
            All_Loaded_Summary_Sheet.Range["D2"].Value = "CML3";
            All_Loaded_Summary_Sheet.Range["E2"].Value = "CML4";
            All_Loaded_Summary_Sheet.Range["F2"].Value = "CML5";
            All_Loaded_Summary_Sheet.Range["G2"].Value = "CMBC";
            All_Loaded_Summary_Sheet.Range["H2"].Value = "CONA";
            All_Loaded_Summary_Sheet.Range["I2"].Value = "CONC";
            All_Loaded_Summary_Sheet.Range["J2"].Value = "CONH";
            All_Loaded_Summary_Sheet.Range["K2"].Value = "EPST";
            All_Loaded_Summary_Sheet.Range["L2"].Value = "ESCA";
            All_Loaded_Summary_Sheet.Range["M2"].Value = "ESCC";
            All_Loaded_Summary_Sheet.Range["N2"].Value = "ESCH";
            All_Loaded_Summary_Sheet.Range["O2"].Value = "HPDK";
            All_Loaded_Summary_Sheet.Range["P2"].Value = "PHIL";

            All_Loaded_Summary_Sheet.Range["Q2"].Value = "BUSN";
            All_Loaded_Summary_Sheet.Range["R2"].Value = "CM12";
            All_Loaded_Summary_Sheet.Range["S2"].Value = "CML3";
            All_Loaded_Summary_Sheet.Range["T2"].Value = "CML4";
            All_Loaded_Summary_Sheet.Range["U2"].Value = "CML5";
            All_Loaded_Summary_Sheet.Range["V2"].Value = "CMBC";
            All_Loaded_Summary_Sheet.Range["W2"].Value = "CONA";
            All_Loaded_Summary_Sheet.Range["X2"].Value = "CONC";
            All_Loaded_Summary_Sheet.Range["Y2"].Value = "CONH";
            All_Loaded_Summary_Sheet.Range["Z2"].Value = "EPST";
            All_Loaded_Summary_Sheet.Range["AA2"].Value = "ESCA";
            All_Loaded_Summary_Sheet.Range["AB2"].Value = "ESCC";
            All_Loaded_Summary_Sheet.Range["AC2"].Value = "ESCH";
            All_Loaded_Summary_Sheet.Range["AD2"].Value = "HPDK";
            All_Loaded_Summary_Sheet.Range["AE2"].Value = "PHIL";


            All_Loaded_Summary_Sheet.Range["B1:AE2"].Interior.Color = currentColor;
            All_Loaded_Summary_Sheet.Range["A2"].Interior.Color = currentColor;
            currentColor = System.Drawing.Color.FromArgb(250, 50, 100);
            All_Loaded_Summary_Sheet.Range["A1"].Interior.Color = currentColor;
            All_Loaded_Summary_Sheet.Range["B1:AE2"].Font.Bold = true;
            All_Loaded_Summary_Sheet.Range["A2"].Font.Bold = true;
            All_Loaded_Summary_Sheet.Range["A1"].ColumnWidth = "10";
            All_Loaded_Summary_Sheet.Range["B2:AE2"].ColumnWidth = "7.5";
            All_Loaded_Summary_Sheet.Range["A:AE"].Font.Name = "Verdana";
            All_Loaded_Summary_Sheet.Range["A:AE"].Font.Size = "8";
            currentColor = System.Drawing.Color.FromArgb(100, 255, 255);
            All_Loaded_Summary_Sheet.Range["A1:AE2"].Font.Color = currentColor;
            //update last row number
            int summaryLastRow = 0;
            summaryLastRow = getLastRowNumberWithValues(All_Loaded_Summary_Sheet, "A");
            //SORT ASCENDING
            if (summaryLastRow > 3)
            {
                All_Loaded_Summary_Sheet.Range["A3:A" + summaryLastRow].Sort(All_Loaded_Summary_Sheet.Range["A3:A" + summaryLastRow], XlSortOrder.xlAscending, Type.Missing, Type.Missing, XlSortOrder.xlAscending, Type.Missing, XlSortOrder.xlAscending, XlYesNoGuess.xlYes, Type.Missing, Type.Missing, XlSortOrientation.xlSortColumns, XlSortMethod.xlPinYin, XlSortDataOption.xlSortNormal, XlSortDataOption.xlSortNormal, XlSortDataOption.xlSortNormal);
            }


            //FORMULAS
            All_Loaded_Summary_Sheet.Range["B3"].Formula = "=COUNTIFS('" + DataFile_SHEET.Name + "'!H:H,$A3,'" + DataFile_SHEET.Name + "'!AAA:AAA,\"INITIAL: \"&B$2)";
            All_Loaded_Summary_Sheet.Range["C3"].Formula = "=COUNTIFS('" + DataFile_SHEET.Name + "'!H:H,$A3,'" + DataFile_SHEET.Name + "'!AAA:AAA,\"INITIAL: \"&C$2)";
            All_Loaded_Summary_Sheet.Range["D3"].Formula = "=COUNTIFS('" + DataFile_SHEET.Name + "'!H:H,$A3,'" + DataFile_SHEET.Name + "'!AAA:AAA,\"INITIAL: \"&D$2)";
            All_Loaded_Summary_Sheet.Range["E3"].Formula = "=COUNTIFS('" + DataFile_SHEET.Name + "'!H:H,$A3,'" + DataFile_SHEET.Name + "'!AAA:AAA,\"INITIAL: \"&E$2)";
            All_Loaded_Summary_Sheet.Range["F3"].Formula = "=COUNTIFS('" + DataFile_SHEET.Name + "'!H:H,$A3,'" + DataFile_SHEET.Name + "'!AAA:AAA,\"INITIAL: \"&F$2)";
            All_Loaded_Summary_Sheet.Range["G3"].Formula = "=COUNTIFS('" + DataFile_SHEET.Name + "'!H:H,$A3,'" + DataFile_SHEET.Name + "'!AAA:AAA,\"INITIAL: \"&G$2)";
            All_Loaded_Summary_Sheet.Range["H3"].Formula = "=COUNTIFS('" + DataFile_SHEET.Name + "'!H:H,$A3,'" + DataFile_SHEET.Name + "'!AAA:AAA,\"INITIAL: \"&H$2)";
            All_Loaded_Summary_Sheet.Range["I3"].Formula = "=COUNTIFS('" + DataFile_SHEET.Name + "'!H:H,$A3,'" + DataFile_SHEET.Name + "'!AAA:AAA,\"INITIAL: \"&I$2)";
            All_Loaded_Summary_Sheet.Range["J3"].Formula = "=COUNTIFS('" + DataFile_SHEET.Name + "'!H:H,$A3,'" + DataFile_SHEET.Name + "'!AAA:AAA,\"INITIAL: \"&J$2)";
            All_Loaded_Summary_Sheet.Range["K3"].Formula = "=COUNTIFS('" + DataFile_SHEET.Name + "'!H:H,$A3,'" + DataFile_SHEET.Name + "'!AAA:AAA,\"INITIAL: \"&K$2)";
            All_Loaded_Summary_Sheet.Range["L3"].Formula = "=COUNTIFS('" + DataFile_SHEET.Name + "'!H:H,$A3,'" + DataFile_SHEET.Name + "'!AAA:AAA,\"INITIAL: \"&L$2)";
            All_Loaded_Summary_Sheet.Range["M3"].Formula = "=COUNTIFS('" + DataFile_SHEET.Name + "'!H:H,$A3,'" + DataFile_SHEET.Name + "'!AAA:AAA,\"INITIAL: \"&M$2)";
            All_Loaded_Summary_Sheet.Range["N3"].Formula = "=COUNTIFS('" + DataFile_SHEET.Name + "'!H:H,$A3,'" + DataFile_SHEET.Name + "'!AAA:AAA,\"INITIAL: \"&N$2)";
            All_Loaded_Summary_Sheet.Range["O3"].Formula = "=COUNTIFS('" + DataFile_SHEET.Name + "'!H:H,$A3,'" + DataFile_SHEET.Name + "'!AAA:AAA,\"INITIAL: \"&O$2)";
            All_Loaded_Summary_Sheet.Range["P3"].Formula = "=COUNTIFS('" + DataFile_SHEET.Name + "'!H:H,$A3,'" + DataFile_SHEET.Name + "'!AAA:AAA,\"INITIAL: \"&P$2)";

            All_Loaded_Summary_Sheet.Range["Q3"].Formula = "=COUNTIFS('" + DataFile_SHEET.Name + "'!H:H,$A3,'" + DataFile_SHEET.Name + "'!AAA:AAA,\"TRANSFERRED: \"&Q$2)";
            All_Loaded_Summary_Sheet.Range["R3"].Formula = "=COUNTIFS('" + DataFile_SHEET.Name + "'!H:H,$A3,'" + DataFile_SHEET.Name + "'!AAA:AAA,\"TRANSFERRED: \"&R$2)";
            All_Loaded_Summary_Sheet.Range["S3"].Formula = "=COUNTIFS('" + DataFile_SHEET.Name + "'!H:H,$A3,'" + DataFile_SHEET.Name + "'!AAA:AAA,\"TRANSFERRED: \"&S$2)";
            All_Loaded_Summary_Sheet.Range["T3"].Formula = "=COUNTIFS('" + DataFile_SHEET.Name + "'!H:H,$A3,'" + DataFile_SHEET.Name + "'!AAA:AAA,\"TRANSFERRED: \"&T$2)";
            All_Loaded_Summary_Sheet.Range["U3"].Formula = "=COUNTIFS('" + DataFile_SHEET.Name + "'!H:H,$A3,'" + DataFile_SHEET.Name + "'!AAA:AAA,\"TRANSFERRED: \"&U$2)";
            All_Loaded_Summary_Sheet.Range["V3"].Formula = "=COUNTIFS('" + DataFile_SHEET.Name + "'!H:H,$A3,'" + DataFile_SHEET.Name + "'!AAA:AAA,\"TRANSFERRED: \"&V$2)";
            All_Loaded_Summary_Sheet.Range["W3"].Formula = "=COUNTIFS('" + DataFile_SHEET.Name + "'!H:H,$A3,'" + DataFile_SHEET.Name + "'!AAA:AAA,\"TRANSFERRED: \"&W$2)";
            All_Loaded_Summary_Sheet.Range["X3"].Formula = "=COUNTIFS('" + DataFile_SHEET.Name + "'!H:H,$A3,'" + DataFile_SHEET.Name + "'!AAA:AAA,\"TRANSFERRED: \"&X$2)";
            All_Loaded_Summary_Sheet.Range["Y3"].Formula = "=COUNTIFS('" + DataFile_SHEET.Name + "'!H:H,$A3,'" + DataFile_SHEET.Name + "'!AAA:AAA,\"TRANSFERRED: \"&Y$2)";
            All_Loaded_Summary_Sheet.Range["Z3"].Formula = "=COUNTIFS('" + DataFile_SHEET.Name + "'!H:H,$A3,'" + DataFile_SHEET.Name + "'!AAA:AAA,\"TRANSFERRED: \"&Z$2)";
            All_Loaded_Summary_Sheet.Range["AA3"].Formula = "=COUNTIFS('" + DataFile_SHEET.Name + "'!H:H,$A3,'" + DataFile_SHEET.Name + "'!AAA:AAA,\"TRANSFERRED: \"&AA$2)";
            All_Loaded_Summary_Sheet.Range["AB3"].Formula = "=COUNTIFS('" + DataFile_SHEET.Name + "'!H:H,$A3,'" + DataFile_SHEET.Name + "'!AAA:AAA,\"TRANSFERRED: \"&AB$2)";
            All_Loaded_Summary_Sheet.Range["AC3"].Formula = "=COUNTIFS('" + DataFile_SHEET.Name + "'!H:H,$A3,'" + DataFile_SHEET.Name + "'!AAA:AAA,\"TRANSFERRED: \"&AC$2)";
            All_Loaded_Summary_Sheet.Range["AD3"].Formula = "=COUNTIFS('" + DataFile_SHEET.Name + "'!H:H,$A3,'" + DataFile_SHEET.Name + "'!AAA:AAA,\"TRANSFERRED: \"&AD$2)";
            All_Loaded_Summary_Sheet.Range["AE3"].Formula = "=COUNTIFS('" + DataFile_SHEET.Name + "'!H:H,$A3,'" + DataFile_SHEET.Name + "'!AAA:AAA,\"TRANSFERRED: \"&AE$2)";
            //AUTOFILL
            if (summaryLastRow > 3)
            {
                All_Loaded_Summary_Sheet.Range["B3"].AutoFill(All_Loaded_Summary_Sheet.Range["B3:B" + summaryLastRow]);
                All_Loaded_Summary_Sheet.Range["C3"].AutoFill(All_Loaded_Summary_Sheet.Range["C3:C" + summaryLastRow]);
                All_Loaded_Summary_Sheet.Range["D3"].AutoFill(All_Loaded_Summary_Sheet.Range["D3:D" + summaryLastRow]);
                All_Loaded_Summary_Sheet.Range["E3"].AutoFill(All_Loaded_Summary_Sheet.Range["E3:E" + summaryLastRow]);
                All_Loaded_Summary_Sheet.Range["F3"].AutoFill(All_Loaded_Summary_Sheet.Range["F3:F" + summaryLastRow]);
                All_Loaded_Summary_Sheet.Range["G3"].AutoFill(All_Loaded_Summary_Sheet.Range["G3:G" + summaryLastRow]);
                All_Loaded_Summary_Sheet.Range["H3"].AutoFill(All_Loaded_Summary_Sheet.Range["H3:H" + summaryLastRow]);
                All_Loaded_Summary_Sheet.Range["I3"].AutoFill(All_Loaded_Summary_Sheet.Range["I3:I" + summaryLastRow]);
                All_Loaded_Summary_Sheet.Range["J3"].AutoFill(All_Loaded_Summary_Sheet.Range["J3:J" + summaryLastRow]);
                All_Loaded_Summary_Sheet.Range["K3"].AutoFill(All_Loaded_Summary_Sheet.Range["K3:K" + summaryLastRow]);
                All_Loaded_Summary_Sheet.Range["L3"].AutoFill(All_Loaded_Summary_Sheet.Range["L3:L" + summaryLastRow]);
                All_Loaded_Summary_Sheet.Range["M3"].AutoFill(All_Loaded_Summary_Sheet.Range["M3:M" + summaryLastRow]);
                All_Loaded_Summary_Sheet.Range["N3"].AutoFill(All_Loaded_Summary_Sheet.Range["N3:N" + summaryLastRow]);
                All_Loaded_Summary_Sheet.Range["O3"].AutoFill(All_Loaded_Summary_Sheet.Range["O3:O" + summaryLastRow]);
                All_Loaded_Summary_Sheet.Range["P3"].AutoFill(All_Loaded_Summary_Sheet.Range["P3:P" + summaryLastRow]);

                All_Loaded_Summary_Sheet.Range["Q3"].AutoFill(All_Loaded_Summary_Sheet.Range["Q3:Q" + summaryLastRow]);
                All_Loaded_Summary_Sheet.Range["R3"].AutoFill(All_Loaded_Summary_Sheet.Range["R3:R" + summaryLastRow]);
                All_Loaded_Summary_Sheet.Range["S3"].AutoFill(All_Loaded_Summary_Sheet.Range["S3:S" + summaryLastRow]);
                All_Loaded_Summary_Sheet.Range["T3"].AutoFill(All_Loaded_Summary_Sheet.Range["T3:T" + summaryLastRow]);
                All_Loaded_Summary_Sheet.Range["U3"].AutoFill(All_Loaded_Summary_Sheet.Range["U3:U" + summaryLastRow]);
                All_Loaded_Summary_Sheet.Range["V3"].AutoFill(All_Loaded_Summary_Sheet.Range["V3:V" + summaryLastRow]);
                All_Loaded_Summary_Sheet.Range["W3"].AutoFill(All_Loaded_Summary_Sheet.Range["W3:W" + summaryLastRow]);
                All_Loaded_Summary_Sheet.Range["X3"].AutoFill(All_Loaded_Summary_Sheet.Range["X3:X" + summaryLastRow]);
                All_Loaded_Summary_Sheet.Range["Y3"].AutoFill(All_Loaded_Summary_Sheet.Range["Y3:Y" + summaryLastRow]);
                All_Loaded_Summary_Sheet.Range["Z3"].AutoFill(All_Loaded_Summary_Sheet.Range["Z3:Z" + summaryLastRow]);
                All_Loaded_Summary_Sheet.Range["AA3"].AutoFill(All_Loaded_Summary_Sheet.Range["AA3:AA" + summaryLastRow]);
                All_Loaded_Summary_Sheet.Range["AB3"].AutoFill(All_Loaded_Summary_Sheet.Range["AB3:AB" + summaryLastRow]);
                All_Loaded_Summary_Sheet.Range["AC3"].AutoFill(All_Loaded_Summary_Sheet.Range["AC3:AC" + summaryLastRow]);
                All_Loaded_Summary_Sheet.Range["AD3"].AutoFill(All_Loaded_Summary_Sheet.Range["AD3:AD" + summaryLastRow]);
                All_Loaded_Summary_Sheet.Range["AE3"].AutoFill(All_Loaded_Summary_Sheet.Range["AE3:AE" + summaryLastRow]);
            }
            All_Loaded_Summary_Sheet.Range["A2:AE" + (summaryLastRow + 3)].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            All_Loaded_Summary_Sheet.Range["A2:AE" + (summaryLastRow + 3)].VerticalAlignment = XlVAlign.xlVAlignCenter;
            currentColor = System.Drawing.Color.FromArgb(50, 200, 200);
            All_Loaded_Summary_Sheet.Range["A3:P" + summaryLastRow].Interior.Color = currentColor;
            currentColor = System.Drawing.Color.FromArgb(0, 200, 150);
            All_Loaded_Summary_Sheet.Range["Q3:AE" + summaryLastRow].Interior.Color = currentColor;
            currentColor = System.Drawing.Color.FromArgb(250, 50, 150);
            All_Loaded_Summary_Sheet.Range["A3:AE" + summaryLastRow].Font.Color = currentColor;

            All_Loaded_Summary_Sheet.Range["A" + (summaryLastRow + 1)].Value = "Total";

            All_Loaded_Summary_Sheet.Range["B" + (summaryLastRow + 1)].Formula = "=SUM(B3:B" + summaryLastRow + ")";
            All_Loaded_Summary_Sheet.Range["C" + (summaryLastRow + 1)].Formula = "=SUM(C3:C" + summaryLastRow + ")";
            All_Loaded_Summary_Sheet.Range["D" + (summaryLastRow + 1)].Formula = "=SUM(D3:D" + summaryLastRow + ")";
            All_Loaded_Summary_Sheet.Range["E" + (summaryLastRow + 1)].Formula = "=SUM(E3:E" + summaryLastRow + ")";
            All_Loaded_Summary_Sheet.Range["F" + (summaryLastRow + 1)].Formula = "=SUM(F3:F" + summaryLastRow + ")";
            All_Loaded_Summary_Sheet.Range["G" + (summaryLastRow + 1)].Formula = "=SUM(G3:G" + summaryLastRow + ")";
            All_Loaded_Summary_Sheet.Range["H" + (summaryLastRow + 1)].Formula = "=SUM(H3:H" + summaryLastRow + ")";
            All_Loaded_Summary_Sheet.Range["I" + (summaryLastRow + 1)].Formula = "=SUM(I3:I" + summaryLastRow + ")";
            All_Loaded_Summary_Sheet.Range["J" + (summaryLastRow + 1)].Formula = "=SUM(J3:J" + summaryLastRow + ")";
            All_Loaded_Summary_Sheet.Range["K" + (summaryLastRow + 1)].Formula = "=SUM(K3:K" + summaryLastRow + ")";
            All_Loaded_Summary_Sheet.Range["L" + (summaryLastRow + 1)].Formula = "=SUM(L3:L" + summaryLastRow + ")";
            All_Loaded_Summary_Sheet.Range["M" + (summaryLastRow + 1)].Formula = "=SUM(M3:M" + summaryLastRow + ")";
            All_Loaded_Summary_Sheet.Range["N" + (summaryLastRow + 1)].Formula = "=SUM(N3:N" + summaryLastRow + ")";
            All_Loaded_Summary_Sheet.Range["O" + (summaryLastRow + 1)].Formula = "=SUM(O3:O" + summaryLastRow + ")";
            All_Loaded_Summary_Sheet.Range["P" + (summaryLastRow + 1)].Formula = "=SUM(P3:P" + summaryLastRow + ")";

            All_Loaded_Summary_Sheet.Range["Q" + (summaryLastRow + 1)].Formula = "=SUM(Q3:Q" + summaryLastRow + ")";
            All_Loaded_Summary_Sheet.Range["R" + (summaryLastRow + 1)].Formula = "=SUM(R3:R" + summaryLastRow + ")";
            All_Loaded_Summary_Sheet.Range["S" + (summaryLastRow + 1)].Formula = "=SUM(S3:S" + summaryLastRow + ")";
            All_Loaded_Summary_Sheet.Range["T" + (summaryLastRow + 1)].Formula = "=SUM(T3:T" + summaryLastRow + ")";
            All_Loaded_Summary_Sheet.Range["U" + (summaryLastRow + 1)].Formula = "=SUM(U3:U" + summaryLastRow + ")";
            All_Loaded_Summary_Sheet.Range["V" + (summaryLastRow + 1)].Formula = "=SUM(V3:V" + summaryLastRow + ")";
            All_Loaded_Summary_Sheet.Range["W" + (summaryLastRow + 1)].Formula = "=SUM(W3:W" + summaryLastRow + ")";
            All_Loaded_Summary_Sheet.Range["X" + (summaryLastRow + 1)].Formula = "=SUM(X3:X" + summaryLastRow + ")";
            All_Loaded_Summary_Sheet.Range["Y" + (summaryLastRow + 1)].Formula = "=SUM(Y3:Y" + summaryLastRow + ")";
            All_Loaded_Summary_Sheet.Range["Z" + (summaryLastRow + 1)].Formula = "=SUM(Z3:Z" + summaryLastRow + ")";
            All_Loaded_Summary_Sheet.Range["AA" + (summaryLastRow + 1)].Formula = "=SUM(AA3:AA" + summaryLastRow + ")";
            All_Loaded_Summary_Sheet.Range["AB" + (summaryLastRow + 1)].Formula = "=SUM(AB3:AB" + summaryLastRow + ")";
            All_Loaded_Summary_Sheet.Range["AC" + (summaryLastRow + 1)].Formula = "=SUM(AC3:AC" + summaryLastRow + ")";
            All_Loaded_Summary_Sheet.Range["AD" + (summaryLastRow + 1)].Formula = "=SUM(AD3:AD" + summaryLastRow + ")";
            All_Loaded_Summary_Sheet.Range["AE" + (summaryLastRow + 1)].Formula = "=SUM(AE3:AE" + summaryLastRow + ")";
            //Console.WriteLine("Totals");
            All_Loaded_Summary_Sheet.Range["A1:AE" + (summaryLastRow + 1)].Copy();
            All_Loaded_Summary_Sheet.Range["A1:AE" + (summaryLastRow + 1)].PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
            All_Loaded_Summary_Sheet.Range["A" + (summaryLastRow + 1) + ":AE" + (summaryLastRow + 1)].Font.Bold = true;
            currentColor = System.Drawing.Color.FromArgb(35, 45, 55);
            All_Loaded_Summary_Sheet.Range["A" + summaryLastRow + 1 + ":AE" + (summaryLastRow + 1)].Interior.Color = currentColor;
            currentColor = System.Drawing.Color.FromArgb(100, 255, 255);
            All_Loaded_Summary_Sheet.Range["A" + (summaryLastRow + 1) + ":AE" + (summaryLastRow + 1)].Font.Color = currentColor;
            All_Loaded_Summary_Sheet.Range["A" + (summaryLastRow + 1) + ":AE" + (summaryLastRow + 1)].Font.Name = "Verdana";
            All_Loaded_Summary_Sheet.Range["A" + (summaryLastRow + 1) + ":AE" + (summaryLastRow + 1)].Font.Size = "8";

            DataFile_SHEET.Delete();

            //Make the Sum sheet of INITIAL and TRANSFERS
            getLastRowNumber(All_Loaded_Summary_Sheet, "");
            All_Loaded_Summary_Sheet.Activate();
            All_Loaded_Summary_Sheet.Range["A1:AE" + lastRow].Copy();
            Worksheet Total_Summary_Sheet = xlApp.Worksheets.Add();
            Total_Summary_Sheet.Activate();
            Total_Summary_Sheet.Name = "CPC - Combined Summary";
            currentColor = System.Drawing.Color.FromArgb(35, 45, 55);
            Total_Summary_Sheet.Range["A:XFD"].Interior.Color = currentColor;
            Total_Summary_Sheet.Range["A1:AE" + lastRow].PasteSpecial();
            Total_Summary_Sheet.Range["B1:P1"].Value = "[TOTAL: INITIAL + TRANSFERRED] - LOADED (DATE OF IMPORT BY QUOTAS)";
            Total_Summary_Sheet.Range["BB3"].Formula = "=$Q3+$B3";
            Total_Summary_Sheet.Range["BC3"].Formula = "=$R3+$C3";
            Total_Summary_Sheet.Range["BD3"].Formula = "=$S3+$D3";
            Total_Summary_Sheet.Range["BE3"].Formula = "=$T3+$E3";
            Total_Summary_Sheet.Range["BF3"].Formula = "=$U3+$F3";
            Total_Summary_Sheet.Range["BG3"].Formula = "=$V3+$G3";
            Total_Summary_Sheet.Range["BH3"].Formula = "=$W3+$H3";
            Total_Summary_Sheet.Range["BI3"].Formula = "=$X3+$I3";
            Total_Summary_Sheet.Range["BJ3"].Formula = "=$Y3+$J3";
            Total_Summary_Sheet.Range["BK3"].Formula = "=$Z3+$K3";
            Total_Summary_Sheet.Range["BL3"].Formula = "=$AA3+$L3";
            Total_Summary_Sheet.Range["BM3"].Formula = "=$AB3+$M3";
            Total_Summary_Sheet.Range["BN3"].Formula = "=$AC3+$N3";
            Total_Summary_Sheet.Range["BO3"].Formula = "=$AD3+$O3";
            Total_Summary_Sheet.Range["BP3"].Formula = "=$AE3+$P3";
            Total_Summary_Sheet.Range["BB3"].AutoFill(Total_Summary_Sheet.Range["BB3:BB" + (summaryLastRow + 1)]);
            Total_Summary_Sheet.Range["BC3"].AutoFill(Total_Summary_Sheet.Range["BC3:BC" + (summaryLastRow + 1)]);
            Total_Summary_Sheet.Range["BD3"].AutoFill(Total_Summary_Sheet.Range["BD3:BD" + (summaryLastRow + 1)]);
            Total_Summary_Sheet.Range["BE3"].AutoFill(Total_Summary_Sheet.Range["BE3:BE" + (summaryLastRow + 1)]);
            Total_Summary_Sheet.Range["BF3"].AutoFill(Total_Summary_Sheet.Range["BF3:BF" + (summaryLastRow + 1)]);
            Total_Summary_Sheet.Range["BG3"].AutoFill(Total_Summary_Sheet.Range["BG3:BG" + (summaryLastRow + 1)]);
            Total_Summary_Sheet.Range["BH3"].AutoFill(Total_Summary_Sheet.Range["BH3:BH" + (summaryLastRow + 1)]);
            Total_Summary_Sheet.Range["BI3"].AutoFill(Total_Summary_Sheet.Range["BI3:BI" + (summaryLastRow + 1)]);
            Total_Summary_Sheet.Range["BJ3"].AutoFill(Total_Summary_Sheet.Range["BJ3:BJ" + (summaryLastRow + 1)]);
            Total_Summary_Sheet.Range["BK3"].AutoFill(Total_Summary_Sheet.Range["BK3:BK" + (summaryLastRow + 1)]);
            Total_Summary_Sheet.Range["BL3"].AutoFill(Total_Summary_Sheet.Range["BL3:BL" + (summaryLastRow + 1)]);
            Total_Summary_Sheet.Range["BM3"].AutoFill(Total_Summary_Sheet.Range["BM3:BM" + (summaryLastRow + 1)]);
            Total_Summary_Sheet.Range["BN3"].AutoFill(Total_Summary_Sheet.Range["BN3:BN" + (summaryLastRow + 1)]);
            Total_Summary_Sheet.Range["BO3"].AutoFill(Total_Summary_Sheet.Range["BO3:BO" + (summaryLastRow + 1)]);
            Total_Summary_Sheet.Range["BP3"].AutoFill(Total_Summary_Sheet.Range["BP3:BP" + (summaryLastRow + 1)]);
            Total_Summary_Sheet.Range["BB3:BP" + (summaryLastRow + 1)].Copy();
            Total_Summary_Sheet.Range["B3"].PasteSpecial(XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
            //Console.WriteLine("tset");
            //string xxx = Console.ReadLine();
            Total_Summary_Sheet.Range["A:A"].Copy();
            Total_Summary_Sheet.Range["A:A"].PasteSpecial(XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
            Total_Summary_Sheet.Range["Q:BP"].Delete();
            currentColor = System.Drawing.Color.FromArgb(35, 45, 55);
            Total_Summary_Sheet.Range[(summaryLastRow + 1) + ":" + (summaryLastRow + 1)].Interior.Color = currentColor;
            Total_Summary_Sheet.Range["B3"].Select();
            Total_Summary_Sheet.Range["B3"].Application.ActiveWindow.FreezePanes = true;
            Total_Summary_Sheet.Range["A" + (summaryLastRow + 1)].Select();
            Total_Summary_Sheet.Move(After: (Worksheet)All_Loaded_Summary_Sheet);




            //HyperLink
            currentColor = System.Drawing.Color.FromArgb(0, 250, 150);
            Total_Summary_Sheet.Range["B" + (summaryLastRow + 3) + ":P" + (summaryLastRow + 3)].MergeCells = true;
            Total_Summary_Sheet.Range["B" + (summaryLastRow + 3)].Value = "NAVIGATE TO: INDIVIDUAL COUNTS FOR '[INITIAL / TRANSFERS]'";
            Total_Summary_Sheet.Hyperlinks.Add(Total_Summary_Sheet.Range["B" + (summaryLastRow + 3)], string.Empty, "'" + All_Loaded_Summary_Sheet.Name + "'!A" + (summaryLastRow + 1), "Breakdown of counts for: [Initial / Transferred calls]");
            Total_Summary_Sheet.Range["B" + (summaryLastRow + 3)].Font.Color = currentColor;


            All_Loaded_Summary_Sheet.Activate();
            All_Loaded_Summary_Sheet.Range["B3"].Select();
            All_Loaded_Summary_Sheet.Range["B3"].Application.ActiveWindow.FreezePanes = true;
            All_Loaded_Summary_Sheet.Range["A" + (summaryLastRow + 1)].Select();

            //HyperLink
            currentColor = System.Drawing.Color.FromArgb(0, 250, 150);
            All_Loaded_Summary_Sheet.Range["B" + (summaryLastRow + 3) + ":AE" + (summaryLastRow + 3)].MergeCells = true;
            All_Loaded_Summary_Sheet.Range["B" + (summaryLastRow + 3)].Value = "NAVIGATE TO: COUNTS FOR '[INITIAL + TRANSFERS]'";
            All_Loaded_Summary_Sheet.Hyperlinks.Add(All_Loaded_Summary_Sheet.Range["B" + (summaryLastRow + 3)], string.Empty, "'" + Total_Summary_Sheet.Name + "'!A" + (summaryLastRow + 1), "Sum of counts for: [Initial + Transferred calls]");
            All_Loaded_Summary_Sheet.Range["B" + (summaryLastRow + 3)].Font.Color = currentColor;
            DataFile_WB.SaveAs(SaveFilePathName, XlFileFormat.xlOpenXMLWorkbook, Missing.Value, Missing.Value, false, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlUserResolution, true, Missing.Value, Missing.Value, Missing.Value);
            DataFile_WB.SaveAs(AltSavePathName, XlFileFormat.xlOpenXMLWorkbook, Missing.Value, Missing.Value, false, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlUserResolution, true, Missing.Value, Missing.Value, Missing.Value);

            //Close all excel workbooks
            DataFile_WB.Close(false);


            //String[] FilterDates = { "1" };
            //All_Loaded_Summary_Sheet.ListObjects.AddEx(Microsoft.Office.Interop.Excel.XlListObjectSourceType.xlSrcRange, All_Loaded_Summary_Sheet.UsedRange, System.Type.Missing, Microsoft.Office.Interop.Excel.XlYesNoGuess.xlYes).Name = "SummaryList";
            //All_Loaded_Summary_Sheet.ListObjects["SummaryList"].Range.AutoFilter(30, FilterDates, Microsoft.Office.Interop.Excel.XlAutoFilterOperator.xlFilterValues);
            //This USED RANGE For SummaryList should exclude the 'Totals' Row, so filtering will not affect it's row placement

        }

        public static void AnalyzeRawSampleFiles(Microsoft.Office.Interop.Excel.Application xlApp, string mySampleFilePath, string fileInitialDateStamp)
        {
            int ESCA_Counts = 0, ESCC_Counts = 0, ESCH_Counts = 0, PHIL_Counts = 0, HPDK_Counts = 0, CMBC_Counts = 0, BUSN_Counts = 0, CONSA_Counts = 0, CONSH_Counts = 0, CONSC_Counts = 0, EPST_Counts = 0, CM12_Counts = 0, CML3_Counts = 0, CML4_Counts = 0, CML5_Counts = 0;
            int ESCA_Trans_Counts = 0, ESCC_Trans_Counts = 0, ESCH_Trans_Counts = 0, PHIL_Trans_Counts = 0, HPDK_Trans_Counts = 0, CMBC_Trans_Counts = 0, BUSN_Trans_Counts = 0, CONSA_Trans_Counts = 0, CONSH_Trans_Counts = 0, CONSC_Trans_Counts = 0, EPST_Trans_Counts = 0, CM12_Trans_Counts = 0, CML3_Trans_Counts = 0, CML4_Trans_Counts = 0, CML5_Trans_Counts = 0;

            DateTimeOffset dateToday;
            string dateTodayString = "";
            dateToday = DateTimeOffset.Now;
            dateTodayString = dateToday.ToString("yyyyMMdd");

            SaveFilePathName = @"\\CC3-MGMT2\EXPORT\" + projectCode + @"\Raw_Samplefile_" + projectCode + "_" + fileInitialDateStamp + ".xlsx";
            System.IO.File.Delete(SaveFilePathName);

            Console.WriteLine("Analyzing Raw Sample Counts for:\t" + mySampleFilePath);

            int integerValue = 0;
            string stringValue = "";
            int cCnt;
            int rCnt;
            int rw = 0;
            int cl = 0;
            float progress = 0.00f;

            //OPEN UP (NECESSARY FILES) AND OBTAIN LAST ROW NUMBER USED RESPECTIVELY
            SampleFilePath = mySampleFilePath;
            Microsoft.Office.Interop.Excel.Workbook DataFile_WB = xlApp.Workbooks.Open(SampleFilePath);
            Worksheet DataFile_SHEET = (Worksheet)DataFile_WB.Sheets[1];
            getLastRowNumber(DataFile_SHEET, SampleFilePath);

            //READ SAMPLEFILE DATA FOR GIVEN "RESPONDENT ROW"
            Microsoft.Office.Interop.Excel.Range range = (Microsoft.Office.Interop.Excel.Range)DataFile_SHEET.UsedRange;
            range = DataFile_SHEET.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;

            mSampleFileDetailEmail += "Raw Sample File Log Run On: " + dateTodayString + "<br>";
            mSampleFileDetailLog += "Raw Sample File Log Run On: " + dateTodayString + "\r\n";

            //Do things here
            DataFile_SHEET.Range["AA1"].Formula = "=COUNTIFS(L:L,\"EPST\",M:M,\"*\")";
            DataFile_SHEET.Range["AA1"].Copy();
            DataFile_SHEET.Range["AA1"].PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
            try { integerValue = (int)(DataFile_SHEET.Range["AA1"].Value2); stringValue = integerValue.ToString(); }
            catch { stringValue = (string)(DataFile_SHEET.Range["AA1"].Value2); }
            EPST_Counts = integerValue;
            DataFile_SHEET.Range["AA1"].Formula = "=COUNTIFS(L:L,\"EPST\",M:M,\">0\")";
            DataFile_SHEET.Range["AA1"].Copy();
            DataFile_SHEET.Range["AA1"].PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
            try { integerValue = (int)(DataFile_SHEET.Range["AA1"].Value2); stringValue = integerValue.ToString(); }
            catch { stringValue = (string)(DataFile_SHEET.Range["AA1"].Value2); }
            EPST_Trans_Counts = integerValue;

            DataFile_SHEET.Range["AA1"].Formula = "=COUNTIFS(L:L,\"PHIL\",M:M,\"*\")";
            DataFile_SHEET.Range["AA1"].Copy();
            DataFile_SHEET.Range["AA1"].PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
            try { integerValue = (int)(DataFile_SHEET.Range["AA1"].Value2); stringValue = integerValue.ToString(); }
            catch { stringValue = (string)(DataFile_SHEET.Range["AA1"].Value2); }
            PHIL_Counts = integerValue;
            DataFile_SHEET.Range["AA1"].Formula = "=COUNTIFS(T:T,\"PHIL\",M:M,\">0\")";
            DataFile_SHEET.Range["AA1"].Copy();
            DataFile_SHEET.Range["AA1"].PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
            try { integerValue = (int)(DataFile_SHEET.Range["AA1"].Value2); stringValue = integerValue.ToString(); }
            catch { stringValue = (string)(DataFile_SHEET.Range["AA1"].Value2); }
            PHIL_Trans_Counts = integerValue;

            DataFile_SHEET.Range["AA1"].Formula = "=COUNTIFS(L:L,\"BUSN\",M:M,\"*\")";
            DataFile_SHEET.Range["AA1"].Copy();
            DataFile_SHEET.Range["AA1"].PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
            try { integerValue = (int)(DataFile_SHEET.Range["AA1"].Value2); stringValue = integerValue.ToString(); }
            catch { stringValue = (string)(DataFile_SHEET.Range["AA1"].Value2); }
            BUSN_Counts = integerValue;
            DataFile_SHEET.Range["AA1"].Formula = "=COUNTIFS(T:T,\"BUSN\",M:M,\">0\")";
            DataFile_SHEET.Range["AA1"].Copy();
            DataFile_SHEET.Range["AA1"].PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
            try { integerValue = (int)(DataFile_SHEET.Range["AA1"].Value2); stringValue = integerValue.ToString(); }
            catch { stringValue = (string)(DataFile_SHEET.Range["AA1"].Value2); }
            BUSN_Trans_Counts = integerValue;

            DataFile_SHEET.Range["AA1"].Formula = "=COUNTIFS(L:L,\"CMBC\",M:M,\"*\")";
            DataFile_SHEET.Range["AA1"].Copy();
            DataFile_SHEET.Range["AA1"].PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
            try { integerValue = (int)(DataFile_SHEET.Range["AA1"].Value2); stringValue = integerValue.ToString(); }
            catch { stringValue = (string)(DataFile_SHEET.Range["AA1"].Value2); }
            CMBC_Counts = integerValue;
            DataFile_SHEET.Range["AA1"].Formula = "=COUNTIFS(T:T,\"CMBC\",M:M,\">0\")";
            DataFile_SHEET.Range["AA1"].Copy();
            DataFile_SHEET.Range["AA1"].PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
            try { integerValue = (int)(DataFile_SHEET.Range["AA1"].Value2); stringValue = integerValue.ToString(); }
            catch { stringValue = (string)(DataFile_SHEET.Range["AA1"].Value2); }
            CMBC_Trans_Counts = integerValue;

            DataFile_SHEET.Range["AA1"].Formula = "=COUNTIFS(L:L,\"HPDK\",M:M,\"*\")";
            DataFile_SHEET.Range["AA1"].Copy();
            DataFile_SHEET.Range["AA1"].PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
            try { integerValue = (int)(DataFile_SHEET.Range["AA1"].Value2); stringValue = integerValue.ToString(); }
            catch { stringValue = (string)(DataFile_SHEET.Range["AA1"].Value2); }
            HPDK_Counts = integerValue;
            DataFile_SHEET.Range["AA1"].Formula = "=COUNTIFS(T:T,\"HPDK\",M:M,\">0\")";
            DataFile_SHEET.Range["AA1"].Copy();
            DataFile_SHEET.Range["AA1"].PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
            try { integerValue = (int)(DataFile_SHEET.Range["AA1"].Value2); stringValue = integerValue.ToString(); }
            catch { stringValue = (string)(DataFile_SHEET.Range["AA1"].Value2); }
            HPDK_Trans_Counts = integerValue;

            DataFile_SHEET.Range["AA1"].Formula = "=COUNTIFS(L:L,\"CM12\",M:M,\"*\")";
            DataFile_SHEET.Range["AA1"].Copy();
            DataFile_SHEET.Range["AA1"].PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
            try { integerValue = (int)(DataFile_SHEET.Range["AA1"].Value2); stringValue = integerValue.ToString(); }
            catch { stringValue = (string)(DataFile_SHEET.Range["AA1"].Value2); }
            CM12_Counts = integerValue;
            DataFile_SHEET.Range["AA1"].Formula = "=COUNTIFS(T:T,\"CM12\",M:M,\">0\")";
            DataFile_SHEET.Range["AA1"].Copy();
            DataFile_SHEET.Range["AA1"].PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
            try { integerValue = (int)(DataFile_SHEET.Range["AA1"].Value2); stringValue = integerValue.ToString(); }
            catch { stringValue = (string)(DataFile_SHEET.Range["AA1"].Value2); }
            CM12_Trans_Counts = integerValue;

            DataFile_SHEET.Range["AA1"].Formula = "=COUNTIFS(L:L,\"CML3\",M:M,\"*\")";
            DataFile_SHEET.Range["AA1"].Copy();
            DataFile_SHEET.Range["AA1"].PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
            try { integerValue = (int)(DataFile_SHEET.Range["AA1"].Value2); stringValue = integerValue.ToString(); }
            catch { stringValue = (string)(DataFile_SHEET.Range["AA1"].Value2); }
            CML3_Counts = integerValue;
            DataFile_SHEET.Range["AA1"].Formula = "=COUNTIFS(T:T,\"CML3\",M:M,\">0\")";
            DataFile_SHEET.Range["AA1"].Copy();
            DataFile_SHEET.Range["AA1"].PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
            try { integerValue = (int)(DataFile_SHEET.Range["AA1"].Value2); stringValue = integerValue.ToString(); }
            catch { stringValue = (string)(DataFile_SHEET.Range["AA1"].Value2); }
            CML3_Trans_Counts = integerValue;

            DataFile_SHEET.Range["AA1"].Formula = "=COUNTIFS(L:L,\"CML4\",M:M,\"*\")";
            DataFile_SHEET.Range["AA1"].Copy();
            DataFile_SHEET.Range["AA1"].PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
            try { integerValue = (int)(DataFile_SHEET.Range["AA1"].Value2); stringValue = integerValue.ToString(); }
            catch { stringValue = (string)(DataFile_SHEET.Range["AA1"].Value2); }
            CML4_Counts = integerValue;
            DataFile_SHEET.Range["AA1"].Formula = "=COUNTIFS(T:T,\"CML4\",M:M,\">0\")";
            DataFile_SHEET.Range["AA1"].Copy();
            DataFile_SHEET.Range["AA1"].PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
            try { integerValue = (int)(DataFile_SHEET.Range["AA1"].Value2); stringValue = integerValue.ToString(); }
            catch { stringValue = (string)(DataFile_SHEET.Range["AA1"].Value2); }
            CML4_Trans_Counts = integerValue;

            DataFile_SHEET.Range["AA1"].Formula = "=COUNTIFS(L:L,\"CML5\",M:M,\"*\")";
            DataFile_SHEET.Range["AA1"].Copy();
            DataFile_SHEET.Range["AA1"].PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
            try { integerValue = (int)(DataFile_SHEET.Range["AA1"].Value2); stringValue = integerValue.ToString(); }
            catch { stringValue = (string)(DataFile_SHEET.Range["AA1"].Value2); }
            CML5_Counts = integerValue;
            DataFile_SHEET.Range["AA1"].Formula = "=COUNTIFS(T:T,\"CML5\",M:M,\">0\")";
            DataFile_SHEET.Range["AA1"].Copy();
            DataFile_SHEET.Range["AA1"].PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
            try { integerValue = (int)(DataFile_SHEET.Range["AA1"].Value2); stringValue = integerValue.ToString(); }
            catch { stringValue = (string)(DataFile_SHEET.Range["AA1"].Value2); }
            CML5_Trans_Counts = integerValue;

            DataFile_SHEET.Range["AA1"].Formula = "=COUNTIFS(L:L,\"CONS\",G:G,\"CPC\",M:M,\"*\")";
            DataFile_SHEET.Range["AA1"].Copy();
            DataFile_SHEET.Range["AA1"].PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
            try { integerValue = (int)(DataFile_SHEET.Range["AA1"].Value2); stringValue = integerValue.ToString(); }
            catch { stringValue = (string)(DataFile_SHEET.Range["AA1"].Value2); }
            CONSC_Counts = integerValue;
            DataFile_SHEET.Range["AA1"].Formula = "=COUNTIFS(T:T,\"CONS\",G:G,\"CPC\",M:M,\">0\")";
            DataFile_SHEET.Range["AA1"].Copy();
            DataFile_SHEET.Range["AA1"].PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
            try { integerValue = (int)(DataFile_SHEET.Range["AA1"].Value2); stringValue = integerValue.ToString(); }
            catch { stringValue = (string)(DataFile_SHEET.Range["AA1"].Value2); }
            CONSC_Trans_Counts = integerValue;

            DataFile_SHEET.Range["AA1"].Formula = "=COUNTIFS(L:L,\"CONS\",G:G,\"Atelka\",M:M,\"*\")";
            DataFile_SHEET.Range["AA1"].Copy();
            DataFile_SHEET.Range["AA1"].PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
            try { integerValue = (int)(DataFile_SHEET.Range["AA1"].Value2); stringValue = integerValue.ToString(); }
            catch { stringValue = (string)(DataFile_SHEET.Range["AA1"].Value2); }
            CONSA_Counts = integerValue;
            DataFile_SHEET.Range["AA1"].Formula = "=COUNTIFS(T:T,\"CONS\",G:G,\"Atelkla\",M:M,\">0\")";
            DataFile_SHEET.Range["AA1"].Copy();
            DataFile_SHEET.Range["AA1"].PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
            try { integerValue = (int)(DataFile_SHEET.Range["AA1"].Value2); stringValue = integerValue.ToString(); }
            catch { stringValue = (string)(DataFile_SHEET.Range["AA1"].Value2); }
            CONSA_Trans_Counts = integerValue;

            DataFile_SHEET.Range["AA1"].Formula = "=COUNTIFS(L:L,\"CONS\",G:G,\"HGS\",M:M,\"*\")";
            DataFile_SHEET.Range["AA1"].Copy();
            DataFile_SHEET.Range["AA1"].PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
            try { integerValue = (int)(DataFile_SHEET.Range["AA1"].Value2); stringValue = integerValue.ToString(); }
            catch { stringValue = (string)(DataFile_SHEET.Range["AA1"].Value2); }
            CONSH_Counts = integerValue;
            DataFile_SHEET.Range["AA1"].Formula = "=COUNTIFS(T:T,\"CONS\",G:G,\"HGS\",M:M,\">0\")";
            DataFile_SHEET.Range["AA1"].Copy();
            DataFile_SHEET.Range["AA1"].PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
            try { integerValue = (int)(DataFile_SHEET.Range["AA1"].Value2); stringValue = integerValue.ToString(); }
            catch { stringValue = (string)(DataFile_SHEET.Range["AA1"].Value2); }
            CONSH_Trans_Counts = integerValue;

            DataFile_SHEET.Range["AA1"].Formula = "=COUNTIFS(L:L,\"ESCA\",M:M,\"*\")";
            DataFile_SHEET.Range["AA1"].Copy();
            DataFile_SHEET.Range["AA1"].PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
            try { integerValue = (int)(DataFile_SHEET.Range["AA1"].Value2); stringValue = integerValue.ToString(); }
            catch { stringValue = (string)(DataFile_SHEET.Range["AA1"].Value2); }
            ESCA_Counts = integerValue;
            DataFile_SHEET.Range["AA1"].Formula = "=COUNTIFS(T:T,\"ESCA\",M:M,\">0\")";
            DataFile_SHEET.Range["AA1"].Copy();
            DataFile_SHEET.Range["AA1"].PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
            try { integerValue = (int)(DataFile_SHEET.Range["AA1"].Value2); stringValue = integerValue.ToString(); }
            catch { stringValue = (string)(DataFile_SHEET.Range["AA1"].Value2); }
            ESCA_Trans_Counts = integerValue;

            DataFile_SHEET.Range["AA1"].Formula = "=COUNTIFS(L:L,\"ESCC\",M:M,\"*\")";
            DataFile_SHEET.Range["AA1"].Copy();
            DataFile_SHEET.Range["AA1"].PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
            try { integerValue = (int)(DataFile_SHEET.Range["AA1"].Value2); stringValue = integerValue.ToString(); }
            catch { stringValue = (string)(DataFile_SHEET.Range["AA1"].Value2); }
            ESCC_Counts = integerValue;
            DataFile_SHEET.Range["AA1"].Formula = "=COUNTIFS(T:T,\"ESCC\",M:M,\">0\")";
            DataFile_SHEET.Range["AA1"].Copy();
            DataFile_SHEET.Range["AA1"].PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
            try { integerValue = (int)(DataFile_SHEET.Range["AA1"].Value2); stringValue = integerValue.ToString(); }
            catch { stringValue = (string)(DataFile_SHEET.Range["AA1"].Value2); }
            ESCC_Trans_Counts = integerValue;

            DataFile_SHEET.Range["AA1"].Formula = "=COUNTIFS(L:L,\"ESCH\",M:M,\"*\")";
            DataFile_SHEET.Range["AA1"].Copy();
            DataFile_SHEET.Range["AA1"].PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
            try { integerValue = (int)(DataFile_SHEET.Range["AA1"].Value2); stringValue = integerValue.ToString(); }
            catch { stringValue = (string)(DataFile_SHEET.Range["AA1"].Value2); }
            ESCH_Counts = integerValue;
            DataFile_SHEET.Range["AA1"].Formula = "=COUNTIFS(T:T,\"ESCH\",M:M,\">0\")";
            DataFile_SHEET.Range["AA1"].Copy();
            DataFile_SHEET.Range["AA1"].PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
            try { integerValue = (int)(DataFile_SHEET.Range["AA1"].Value2); stringValue = integerValue.ToString(); }
            catch { stringValue = (string)(DataFile_SHEET.Range["AA1"].Value2); }
            ESCH_Trans_Counts = integerValue;

            //End things here
            mSampleFileDetailLog += "ANALYZING FILE:\t" + mySampleFilePath + "\r\n";
            mSampleFileDetailLog += "[CONC - Initial Call]:\t" + CONSC_Counts + "\t[CONC - Transferred Call]:\t" + CONSC_Trans_Counts + "\r\n";
            mSampleFileDetailLog += "[CONA - Initial Call]:\t" + CONSA_Counts + "\t[CONA - Transferred Call]:\t" + CONSA_Trans_Counts + "\r\n";
            mSampleFileDetailLog += "[CONH - Initial Call]:\t" + CONSH_Counts + "\t[CONH - Transferred Call]:\t" + CONSH_Trans_Counts + "\r\n";
            mSampleFileDetailLog += "[EPST - Initial Call]:\t" + EPST_Counts + "\t[EPST - Transferred Call]:\t" + EPST_Trans_Counts + "\r\n";
            mSampleFileDetailLog += "[CM12 - Initial Call]:\t" + CM12_Counts + "\t[CM12 - Transferred Call]:\t" + CM12_Trans_Counts + "\r\n";
            mSampleFileDetailLog += "[CML3 - Initial Call]:\t" + CML3_Counts + "\t[CML3 - Transferred Call]:\t" + CML3_Trans_Counts + "\r\n";
            mSampleFileDetailLog += "[CML4 - Initial Call]:\t" + CML4_Counts + "\t[CML4 - Transferred Call]:\t" + CML4_Trans_Counts + "\r\n";
            mSampleFileDetailLog += "[CML5 - Initial Call]:\t" + CML5_Counts + "\t[CML5 - Transferred Call]:\t" + CML5_Trans_Counts + "\r\n";
            mSampleFileDetailLog += "[HPDK - Initial Call]:\t" + HPDK_Counts + "\t[HPDK - Transferred Call]:\t" + HPDK_Trans_Counts + "\r\n";
            mSampleFileDetailLog += "[CMBC - Initial Call]:\t" + CMBC_Counts + "\t[CMBC - Transferred Call]:\t" + CMBC_Trans_Counts + "\r\n";
            mSampleFileDetailLog += "[BUSN - Initial Call]:\t" + BUSN_Counts + "\t[BUSN - Transferred Call]:\t" + BUSN_Trans_Counts + "\r\n";
            mSampleFileDetailLog += "[PHIL - Initial Call]:\t" + PHIL_Counts + "\t[PHIL - Transferred Call]:\t" + PHIL_Trans_Counts + "\r\n";
            mSampleFileDetailLog += "[ESCA - Initial Call]:\t" + ESCA_Counts + "\t[ESCA - Transferred Call]:\t" + ESCA_Trans_Counts + "\r\n";
            mSampleFileDetailLog += "[ESCC - Initial Call]:\t" + ESCC_Counts + "\t[ESCC - Transferred Call]:\t" + ESCC_Trans_Counts + "\r\n";
            mSampleFileDetailLog += "[ESCH - Initial Call]:\t" + ESCH_Counts + "\t[ESCH - Transferred Call]:\t" + ESCH_Trans_Counts + "\r\n";

            //DataFile_WB.SaveAs(SaveFilePathName, XlFileFormat.xlOpenXMLWorkbook, Missing.Value, Missing.Value, false, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlUserResolution, true, Missing.Value, Missing.Value, Missing.Value);

            //Close all excel workbooks
            DataFile_WB.Close(false);


        }


        public static void CPC_Coding_Summary(Microsoft.Office.Interop.Excel.Application xlApp, string mySampleFilePath)
        {
            DateTimeOffset dateToday;
            string dateTodayString = "";
            dateToday = DateTimeOffset.Now;
            dateTodayString = dateToday.ToString("yyyyMMdd");

            int integerValue = 0;
            string stringValue = "";
            int cCnt;
            int rCnt;
            int rw = 0;
            int cl = 0;
            float progress = 0.00f;

            //OPEN UP (NECESSARY FILES) AND OBTAIN LAST ROW NUMBER USED RESPECTIVELY
            SampleFilePath = mySampleFilePath;
            Console.WriteLine("filePath is: " + mySampleFilePath);
            string s = Console.ReadLine();
            Microsoft.Office.Interop.Excel.Workbook DataFile_WB = xlApp.Workbooks.Open(SampleFilePath);
            Worksheet DataFile_SHEET = (Worksheet)DataFile_WB.Sheets[1];
            getLastRowNumber(DataFile_SHEET, SampleFilePath);

            //READ SAMPLEFILE DATA FOR GIVEN "RESPONDENT ROW"
            Microsoft.Office.Interop.Excel.Range range = (Microsoft.Office.Interop.Excel.Range)DataFile_SHEET.UsedRange;
            range = DataFile_SHEET.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;

            //DICTIONARY OF QUESTION CODES
            Dictionary<string, string> Q3CodesDictionary = new Dictionary<string, string>();
            Q3CodesDictionary.Add("1", "Transaction Mail");
            Q3CodesDictionary.Add("2", "A parcel (package)");
            Q3CodesDictionary.Add("3", "Direct Marketing");
            Q3CodesDictionary.Add("4", "Other Canada Post Products or Services");
            Q3CodesDictionary.Add("5", "No product associated");
            Q3CodesDictionary.Add("50", "Not specified [record verbatim]");
            Dictionary<string, string> Q3ACodesDictionary = new Dictionary<string, string>();
            Q3ACodesDictionary.Add("1", "Commercial Expedited (domestic)");
            Q3ACodesDictionary.Add("2", "Priority Courier (domestic)");
            Q3ACodesDictionary.Add("3", "Regular (domestic)");
            Q3ACodesDictionary.Add("4", "USA and International parcel");
            Q3ACodesDictionary.Add("5", "No product associated");
            Q3ACodesDictionary.Add("6", "Combined/ Not Specified/ Unknown");
            Q3ACodesDictionary.Add("50", "Combined/ Not Specified/ Unknown");
            Q3ACodesDictionary.Add("99", "Combined/ Not Specified/ Unknown");
            Dictionary<string, string> Q3BCodesDictionary = new Dictionary<string, string>();
            Q3BCodesDictionary.Add("1", "Registered Lettermail (domestic)");
            Q3BCodesDictionary.Add("2", "Incentive or Registered Lettermail (domestic)");
            Q3BCodesDictionary.Add("3", "Registered Lettermail (USA or International)");
            Q3BCodesDictionary.Add("4", "Incentive or Registered Lettermail (USA or International)");
            Q3BCodesDictionary.Add("50", "Combined/ Not Specified/ Unknown");
            Q3BCodesDictionary.Add("99", "Combined/ Not Specified/ Unknown");
            Dictionary<string, string> Q3CCodesDictionary = new Dictionary<string, string>();
            Q3CCodesDictionary.Add("1", "Unaddressed Admail");
            Q3CCodesDictionary.Add("2", "Addressed Admail");
            Q3CCodesDictionary.Add("3", "Business Reply Mail");
            Q3CCodesDictionary.Add("4", "Publications Mail");
            Q3CCodesDictionary.Add("50", "Combined/ Not Specified/ Unknown");
            Q3CCodesDictionary.Add("99", "Combined/ Not Specified/ Unknown");
            Dictionary<string, string> Q3DCodesDictionary = new Dictionary<string, string>();
            Q3DCodesDictionary.Add("1", "Change of Address/Hold Mail");
            Q3DCodesDictionary.Add("2", "Technical (ePost/Helpdesk)");
            Q3DCodesDictionary.Add("3", "Money Order");
            Q3DCodesDictionary.Add("50", "Other/Not specified/Unknown [RECORD VERBATIM]");
            Q3DCodesDictionary.Add("99", "Combined/ Not Specified/ Unknown");
            Dictionary<string, string> Q3ECodesDictionary = new Dictionary<string, string>();
            Q3ECodesDictionary.Add("1", "Community Mail Boxes");
            Q3ECodesDictionary.Add("2", "Venture One");
            Q3ECodesDictionary.Add("3", "Complaints");
            Q3ECodesDictionary.Add("50", "Combined/ Not Specified/ Unknown");
            Q3ECodesDictionary.Add("99", "Combined/ Not Specified/ Unknown");
            Dictionary<string, string> Q4CodesDictionary = new Dictionary<string, string>();
            Q4CodesDictionary.Add("1", "Delivery Notice Card/ Picking Up item at PO (inquiries or issues)");
            Q4CodesDictionary.Add("2", "Item not delivered");
            Q4CodesDictionary.Add("3", "Item already delivered (late or damaged)");
            Q4CodesDictionary.Add("4", "Other - Delivery Stat or Conf (please describe)");
            Q4CodesDictionary.Add("5", "All mailbox issues (excluding key related issues)");
            Q4CodesDictionary.Add("6", "Community Mail Box Key related request/issues (CMB)");
            Q4CodesDictionary.Add("7", "Method of delivery (location/type of mailbox) inquiries/complaints");
            Q4CodesDictionary.Add("8", "Mail delivery Issue");
            Q4CodesDictionary.Add("9", "Other - Mail Delivery (please describe)");
            Q4CodesDictionary.Add("10", "Household counts");
            Q4CodesDictionary.Add("11", "Mail Prep/Deposit - Help with a mailing - Commercial");
            Q4CodesDictionary.Add("12", "Mail Prep/Deposit - Online question");
            Q4CodesDictionary.Add("13", "Mail Prep/Deposit - general");
            Q4CodesDictionary.Add("14", "Commercial Mail Delivery Issue");
            Q4CodesDictionary.Add("15", "Rates");
            Q4CodesDictionary.Add("16", "COA or Hold Mail Delivery Issue");
            Q4CodesDictionary.Add("17", "General Questions on COA or Hold Mail");
            Q4CodesDictionary.Add("18", "Request to Cancel/Modify an existing COA");
            Q4CodesDictionary.Add("19", "Online COA issue Question");
            Q4CodesDictionary.Add("20", "Other - COA or Hold Mail");
            Q4CodesDictionary.Add("21", "Order/billing Adjustment or Cancelation");
            Q4CodesDictionary.Add("22", "General account enquiry/request");
            Q4CodesDictionary.Add("23", "New, modification or cancellation of customer account setup");
            Q4CodesDictionary.Add("24", "Venture One (Small Business Solution)");
            Q4CodesDictionary.Add("25", "Assistance with Electronic Shipping Tool Online");
            Q4CodesDictionary.Add("26", "Assistance with EST Desktop (EST 2.0)");
            Q4CodesDictionary.Add("27", "Manage my account/Online Payment");
            Q4CodesDictionary.Add("28", "Assistance with Precision Targeter");
            Q4CodesDictionary.Add("29", "Other - Helpdesk (please describe)");
            Q4CodesDictionary.Add("30", "ePost - Bill/Statement");
            Q4CodesDictionary.Add("31", "Epost/CanadaPost ID - Existing Account Access");
            Q4CodesDictionary.Add("32", "Sign up/Profile");
            Q4CodesDictionary.Add("33", "ePost Connect");
            Q4CodesDictionary.Add("34", "Other - ePost (please describe)");
            Q4CodesDictionary.Add("35", "Products or supplies (order and modification or issues)");
            Q4CodesDictionary.Add("36", "Pick-up request");
            Q4CodesDictionary.Add("37", "Pick up issues");
            Q4CodesDictionary.Add("38", "Product and Services (excluding COA) - Consumer");
            Q4CodesDictionary.Add("39", "Post Office Questions");
            Q4CodesDictionary.Add("40", "Other - General Information");
            Q4CodesDictionary.Add("42", "Complaint/Suggestion/Compliment (Canada Post in General or employee)");
            Q4CodesDictionary.Add("50", "Other (Please describe)");
            Q4CodesDictionary.Add("99", "DK/NA");

            mCodingDetailEmail += "Coding File Log For: " + dateTodayString + "<br>";
            mCodingDetailLog += "Coding File Log For: " + dateTodayString + "\r\n";
            for (rCnt = 1; rCnt <= rw; rCnt++)
            {
                if (programMode == "1")
                {
                    progress = (float)rCnt / (float)rw * 100;
                    Console.WriteLine("Progress: " + progress + "%");
                }

                for (cCnt = 1; cCnt <= cl; cCnt++)
                {
                    try
                    {
                        integerValue = (int)(range.Cells[rCnt, cCnt].Value2);
                        stringValue = integerValue.ToString();
                    }
                    catch
                    {
                        stringValue = (string)(range.Cells[rCnt, cCnt].Value2);
                    }
                    if (rCnt == 1)
                    {
                        headerList[cCnt] = stringValue;
                    }

                    //ACTUAL DATA (WITHOUT HEADER)
                    if (rCnt > 1)
                    {
                        //STORE QID OF CURRENT RECORD WE'RE ANALYZING IN CASE WE NEED TO REPORT ON THIS
                        if (cCnt == 1)
                        {
                            qid = stringValue;
                        }
                        questionList[cCnt] = stringValue;
                        //1) Q3 OTHERS COUNT
                        if (cCnt == 4)
                        {
                            //If there are '50's still remaining, this is an error. All Q3 responses must fall within, or be coded as codes 1-5.
                            if (questionList[cCnt] == "50")
                            {
                                OthersQ3++;
                                isCodingError = true;
                                mCodingDetailEmail += qid + " @ " + headerList[cCnt] + " cannot be: " + questionList[cCnt] + "<br>";
                                mCodingDetailLog += qid + " @ Q3 cannot be: " + questionList[cCnt] + "\r\n";
                            }
                            try
                            {
                                if (!Q3CodesDictionary.ContainsKey(questionList[cCnt]))
                                {
                                    isCodingError = true;
                                    mCodingDetailEmail += qid + " @ " + headerList[cCnt] + " value of: " + questionList[cCnt] + " does not exist<br>";
                                    mCodingDetailLog += qid + " @ " + headerList[cCnt] + " value of: " + questionList[cCnt] + " does not exist\r\n";
                                }
                            }
                            catch { }
                            try
                            {
                                if (questionList[cCnt] == null)
                                {
                                    isCodingError = true;
                                    mCodingDetailEmail += qid + " @ " + headerList[cCnt] + " should not be NULL<br>";
                                    mCodingDetailLog += qid + " @ " + headerList[cCnt] + " should not be NULL\r\n";
                                }
                            }
                            catch { }
                        }
                        //2) Q3A OTHERS COUNT
                        if (cCnt == 6)
                        {
                            if (questionList[cCnt] == "50")
                            {
                                OthersQ3A++;
                            }
                            try
                            {
                                if (!Q3ACodesDictionary.ContainsKey(questionList[cCnt]))
                                {
                                    isCodingError = true;
                                    mCodingDetailEmail += qid + " @ " + headerList[cCnt] + " value of: " + questionList[cCnt] + " does not exist<br>";
                                    mCodingDetailLog += qid + " @ " + headerList[cCnt] + " value of: " + questionList[cCnt] + " does not exist\r\n";
                                }
                            }
                            catch { }
                            try
                            {
                                if (questionList[4] == "2" && questionList[cCnt] == null)
                                {
                                    Console.WriteLine("qid: " + qid + ": @ " + headerList[cCnt] + " CODE is: " + questionList[cCnt]);
                                    isCodingError = true;
                                    mCodingDetailEmail += qid + " @ " + headerList[cCnt] + " Code is a NULL Value<br>";
                                    mCodingDetailLog += qid + " @ " + headerList[cCnt] + " Code is a NULL Value\r\n";
                                }
                            }
                            catch { }
                        }
                        //2.1) Q3A VERBATIM
                        if (cCnt == 7)
                        {
                            try
                            {
                                if (questionList[4] == "2" && questionList[cCnt] == null)
                                {
                                    Console.WriteLine("qid: " + qid + ": @ " + headerList[cCnt] + " verbatim is: " + questionList[cCnt]);
                                    isCodingError = true;
                                    mCodingDetailEmail += qid + " @ " + headerList[cCnt] + " Verbatim is Missing<br>";
                                    mCodingDetailLog += qid + " @ " + headerList[cCnt] + " verbatim is Missing\r\n";
                                }
                            }
                            catch { }
                        }
                        //3) Q3B OTHERS COUNT
                        if (cCnt == 8)
                        {
                            if (questionList[cCnt] == "50")
                            {
                                OthersQ3B++;
                            }
                            try
                            {
                                if (!Q3BCodesDictionary.ContainsKey(questionList[cCnt]))
                                {
                                    isCodingError = true;
                                    mCodingDetailEmail += qid + " @ " + headerList[cCnt] + " value of: " + questionList[cCnt] + " does not exist<br>";
                                    mCodingDetailLog += qid + " @ " + headerList[cCnt] + " value of: " + questionList[cCnt] + " does not exist\r\n";
                                }
                            }
                            catch { }
                            try
                            {
                                if (questionList[4] == "1" && questionList[cCnt] == null)
                                {
                                    Console.WriteLine("qid: " + qid + ": @ " + headerList[cCnt] + " CODE is: " + questionList[cCnt]);
                                    isCodingError = true;
                                    mCodingDetailEmail += qid + " @ " + headerList[cCnt] + " Code is a NULL Value<br>";
                                    mCodingDetailLog += qid + " @ " + headerList[cCnt] + " Code is a NULL Value\r\n";
                                }
                            }
                            catch { }
                        }
                        //3.1) Q3B VERBATIM
                        if (cCnt == 9)
                        {
                            try
                            {
                                if (questionList[4] == "1" && questionList[cCnt] == null)
                                {
                                    Console.WriteLine("qid: " + qid + ": @ " + headerList[cCnt] + " verbatim is: " + questionList[cCnt]);
                                    isCodingError = true;
                                    mCodingDetailEmail += qid + " @ " + headerList[cCnt] + " Verbatim is Missing<br>";
                                    mCodingDetailLog += qid + " @ " + headerList[cCnt] + " verbatim is Missing\r\n";
                                }
                            }
                            catch { }
                        }
                        //4) Q3C OTHERS COUNT
                        if (cCnt == 10)
                        {
                            if (questionList[cCnt] == "50")
                            {
                                OthersQ3C++;
                            }
                            try
                            {
                                if (!Q3CCodesDictionary.ContainsKey(questionList[cCnt]))
                                {
                                    isCodingError = true;
                                    mCodingDetailEmail += qid + " @ " + headerList[cCnt] + " value of: " + questionList[cCnt] + " does not exist<br>";
                                    mCodingDetailLog += qid + " @ " + headerList[cCnt] + " value of: " + questionList[cCnt] + " does not exist\r\n";
                                }
                            }
                            catch { }
                            try
                            {
                                if (questionList[4] == "3" && questionList[cCnt] == null)
                                {
                                    Console.WriteLine("qid: " + qid + ": @ " + headerList[cCnt] + " CODE is: " + questionList[cCnt]);
                                    isCodingError = true;
                                    mCodingDetailEmail += qid + " @ " + headerList[cCnt] + " Code is a NULL Value<br>";
                                    mCodingDetailLog += qid + " @ " + headerList[cCnt] + " Code is a NULL Value\r\n";
                                }
                            }
                            catch { }
                        }
                        //4.1) Q3C VERBATIM
                        if (cCnt == 11)
                        {
                            try
                            {
                                if (questionList[4] == "3" && questionList[cCnt] == null)
                                {
                                    Console.WriteLine("qid: " + qid + ": @ " + headerList[cCnt] + " verbatim is: " + questionList[cCnt]);
                                    isCodingError = true;
                                    mCodingDetailEmail += qid + " @ " + headerList[cCnt] + " Verbatim is Missing<br>";
                                    mCodingDetailLog += qid + " @ " + headerList[cCnt] + " verbatim is Missing\r\n";
                                }
                            }
                            catch { }
                        }
                        //5) Q3D OTHERS COUNT
                        if (cCnt == 12)
                        {
                            if (questionList[cCnt] == "50")
                            {
                                OthersQ3D++;
                            }
                            try
                            {
                                if (!Q3DCodesDictionary.ContainsKey(questionList[cCnt]))
                                {
                                    isCodingError = true;
                                    mCodingDetailEmail += qid + " @ " + headerList[cCnt] + " value of: " + questionList[cCnt] + " does not exist<br>";
                                    mCodingDetailLog += qid + " @ " + headerList[cCnt] + " value of: " + questionList[cCnt] + " does not exist\r\n";
                                }
                            }
                            catch { }
                            try
                            {
                                if (questionList[4] == "4" && questionList[cCnt] == null)
                                {
                                    Console.WriteLine("qid: " + qid + ": @ " + headerList[cCnt] + " CODE is: " + questionList[cCnt]);
                                    isCodingError = true;
                                    mCodingDetailEmail += qid + " @ " + headerList[cCnt] + " Code is a NULL Value<br>";
                                    mCodingDetailLog += qid + " @ " + headerList[cCnt] + " Code is a NULL Value\r\n";
                                }
                            }
                            catch { }
                        }
                        //5.1) Q3D VERBATIM
                        if (cCnt == 13)
                        {
                            try
                            {
                                if (questionList[4] == "4" && questionList[cCnt] == null)
                                {
                                    Console.WriteLine("qid: " + qid + ": @ " + headerList[cCnt] + " verbatim is: " + questionList[cCnt]);
                                    isCodingError = true;
                                    mCodingDetailEmail += qid + " @ " + headerList[cCnt] + " Verbatim is Missing<br>";
                                    mCodingDetailLog += qid + " @ " + headerList[cCnt] + " verbatim is Missing\r\n";
                                }
                            }
                            catch { }
                        }
                        //6) Q3E OTHERS COUNT
                        if (cCnt == 14)
                        {
                            if (questionList[cCnt] == "50")
                            {
                                OthersQ3E++;
                            }
                            try
                            {
                                if (!Q3ECodesDictionary.ContainsKey(questionList[cCnt]))
                                {
                                    isCodingError = true;
                                    mCodingDetailEmail += qid + " @ " + headerList[cCnt] + " value of: " + questionList[cCnt] + " does not exist<br>";
                                    mCodingDetailLog += qid + " @ " + headerList[cCnt] + " value of: " + questionList[cCnt] + " does not exist\r\n";
                                }
                            }
                            catch { }
                            try
                            {
                                if (questionList[4] == "5" && questionList[cCnt] == null)
                                {
                                    Console.WriteLine("qid: " + qid + ": @ " + headerList[cCnt] + " CODE is: " + questionList[cCnt]);
                                    isCodingError = true;
                                    mCodingDetailEmail += qid + " @ " + headerList[cCnt] + " Code is a NULL Value<br>";
                                    mCodingDetailLog += qid + " @ " + headerList[cCnt] + " Code is a NULL Value\r\n";
                                }
                            }
                            catch { }
                        }
                        //6.1) Q3E VERBATIM
                        if (cCnt == 15)
                        {
                            try
                            {
                                if (questionList[4] == "5" && questionList[cCnt] == null)
                                {
                                    Console.WriteLine("qid: " + qid + ": @ " + headerList[cCnt] + " verbatim is: " + questionList[cCnt]);
                                    isCodingError = true;
                                    mCodingDetailEmail += qid + " @ " + headerList[cCnt] + " Verbatim is Missing<br>";
                                    mCodingDetailLog += qid + " @ " + headerList[cCnt] + " verbatim is Missing\r\n";
                                }
                            }
                            catch { }
                        }
                        //7) Q4 OTHERS COUNT
                        if (cCnt == 16)
                        {
                            if (questionList[cCnt] == "50")
                            {
                                OthersQ4++;
                            }
                            try
                            {
                                if (!Q4CodesDictionary.ContainsKey(questionList[cCnt]))
                                {
                                    isCodingError = true;
                                    mCodingDetailEmail += qid + " @ " + headerList[cCnt] + " value of: " + questionList[cCnt] + " does not exist<br>";
                                    mCodingDetailLog += qid + " @ " + headerList[cCnt] + " value of: " + questionList[cCnt] + " does not exist\r\n";
                                }
                            }
                            catch { }
                            try
                            {
                                if (questionList[cCnt] == null)
                                {
                                    Console.WriteLine("qid: " + qid + ": @ " + headerList[cCnt] + " CODE is: " + questionList[cCnt]);
                                    isCodingError = true;
                                    mCodingDetailEmail += qid + " @ " + headerList[cCnt] + " Code is a NULL Value<br>";
                                    mCodingDetailLog += qid + " @ " + headerList[cCnt] + " Code is a NULL Value\r\n";
                                }
                            }
                            catch { }
                        }
                        //7.1) Q4 VERBATIM
                        if (cCnt == 17)
                        {
                            try
                            {
                                if (questionList[cCnt] == null)
                                {
                                    Console.WriteLine("qid: " + qid + ": @ " + headerList[cCnt] + " verbatim is: " + questionList[cCnt]);
                                    isCodingError = true;
                                    mCodingDetailEmail += qid + " @ " + headerList[cCnt] + " Verbatim is Missing<br>";
                                    mCodingDetailLog += qid + " @ " + headerList[cCnt] + " verbatim is Missing\r\n";
                                }
                            }
                            catch { }
                        }
                    }
                }
            }
            //Close all excel workbooks
            DataFile_WB.Close(false);
        }

        public static void SendReportEmail(string myEmail, string ccEmail1, string ccEmail2, string ccEmail3, string ccEmail4, string ccEmail5, string ccEmail6)
        {
            DateTimeOffset dateToday;
            string dateTodayString = "";
            dateToday = DateTimeOffset.Now;
            dateTodayString = dateToday.ToString("MMMM dd, yyyy");

            //CONDITIONAL OPERATOR USED TO POPULATE SUBJECT LINE IF IT'S BLANK
            subject = projectCode + " Week " + currentWeek + " of " + weekPeriod + " - " + (lastRow - 2) + " Completes";
            string[] emailContent = System.IO.File.ReadAllLines(@"\\QNAP\CATI\Peter_Tan\_____WORK_____\C#_Projects\CPC_Report_Summary\CPC_Report_Summary\CPC_Email.html");
            string modifiedEmailLine = "";
            string body = "";

            MailMessage message = new MailMessage();

            message.To.Add(myEmail);
            if (ccEmail1 != "")
            {
                message.CC.Add(ccEmail1);
                message.CC.Add(ccEmail2);
                message.CC.Add(ccEmail3);
                message.CC.Add(ccEmail4);
                message.CC.Add(ccEmail5);
                message.CC.Add(ccEmail6);
            }
            message.From = new MailAddress("ptan@forumresearch.com");
            message.Subject = subject;

            foreach (string emailLine in emailContent)
            {
                if (emailLine == "<font face=\"verdana\" size=\"2\">")
                {
                    modifiedEmailLine = "<font face=\"verdana\" size=\"2\">" + mReportDetailEmail + " <br>";
                    body += modifiedEmailLine;
                }
                else if (emailLine == @"\\CC3-MGMT2\EXPORT\")
                {
                    modifiedEmailLine = @"\\CC3-MGMT2\EXPORT\" + projectCode + @"\";
                    body += modifiedEmailLine;
                }
                else
                {
                    body += emailLine;
                }
            }

            message.Body = body;
            message.BodyEncoding = UTF8Encoding.UTF8;
            message.SubjectEncoding = UTF8Encoding.UTF8;
            message.IsBodyHtml = true;
            //USE OF DELEGATE CALLBACK TO DISPLAY MESSAGE INFO
            MessageDisplay(message, false);
            //ATTACHMENTS
            Attachment attachedFileData = null;
            for (int i = 0; i < fileID; i++)
            {
                attachedFileData = new Attachment(filePathToAttach[i]);
                message.Attachments.Add(attachedFileData);
            }
            SmtpClient smtpClient = new SmtpClient();
            smtpClient.Port = 587;
            smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
            smtpClient.UseDefaultCredentials = false;
            smtpClient.Host = "smtp.mandrillapp.com";
            smtpClient.Credentials = new System.Net.NetworkCredential("itaccounts@forumresearch.com", "G9PAqaU3hFS6Nv0UEnCN5w");
            smtpClient.EnableSsl = true;
            smtpClient.Send(message);
        }

        /* MAIN */
        [STAThreadAttribute]
        static void Main(string[] args)
        {
            try { RunAllFunctions(); } catch { Console.WriteLine("Main Function Failed"); }
        }

        static void RunAllFunctions()
        {
            //INITIALIZE THE (EXCEL APP) BY DECLARING IT ONCE AT THE START OF MAIN
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlApp.DisplayAlerts = false;

            DateTimeOffset dateToday;
            string dateTodayString = "";
            //TIMESTAMP PER RESPONDENT
            dateToday = DateTimeOffset.Now;
            dateTodayString = dateToday.ToString("yyyyMMdd");

            DateTimeOffset dateWeekend;
            string dateWeekendString = "";
            //TIMESTAMP PER RESPONDENT
            dateWeekend = DateTimeOffset.Now.AddDays(-2);
            dateWeekendString = dateWeekend.ToString("yyyyMMdd");

            DateTimeOffset dateYesterday;
            string dateYesterdayString = "";
            //TIMESTAMP PER RESPONDENT
            dateYesterday = DateTimeOffset.Now.AddDays(-1);
            dateYesterdayString = dateYesterday.ToString("yyyyMMdd");


            //MESSAGES
            string mFinishedSuccessfully = "\nProgram Successfully Completed. ENTER ANY VALUE TO EXIT\n";
            string mCheck = "\nChecking Data File For Date: " + dateTodayString + "\n";
            string mNoLogFile = "\nNo Log File Found... Creating Log File.\n";
            string mRunning = "";

            Console.Write("\r\nCANADA POST PROJECT MANAGER\r\n\r\n\r\n[1]\tCHECK CODING FILE\r\n[2]\tCHECK AND SUMMARIZE CCE WEEKLY REPORT GENERATED\r\n[3]\tBREAKDOWN OF RAW SAMPLE PROVIDED PER QUOTA\r\n[4]\tBREAKDOWN OF SAMPLE LOADED PER QUOTA\r\n[5]\tRUN CCE WEEKLY REPORT (EXECUTE PYTHON SCRIPT)\r\n[6]\tRUN (TEST VERSION) CCE SAMPLE IMPORTATION (EXECUTE PYTHON SCRIPT)\r\n[0]\tTERMINATE\r\n\r\nChoose an option: ");
            programMode = Console.ReadLine();
            mRunning = programMode.Equals("0") || programMode.Equals("") ? "" : "\r\n\r\nRunning option: " + programMode + ". Please wait...\r\n\r\n";
            Console.WriteLine(mRunning);

            Stream myStream = null;
            Application2.OpenFileDialog theDialog = new Application2.OpenFileDialog();

            string[] logSampleFileContent;
            try
            {
                logSampleFileContent = System.IO.File.ReadAllLines(@"G:\Peter_Tan\_____WORK_____\C#_Projects\CPC_Report_Summary\CPC_Report_Summary\DialogLog.log");
            }
            catch
            {
                System.IO.StreamWriter LogFileDialog = new System.IO.StreamWriter(@"G:\Peter_Tan\_____WORK_____\C#_Projects\CPC_Report_Summary\CPC_Report_Summary\DialogLog.log");
                LogFileDialog.Close();
                logSampleFileContent = System.IO.File.ReadAllLines(@"G:\Peter_Tan\_____WORK_____\C#_Projects\CPC_Report_Summary\CPC_Report_Summary\DialogLog.log");
            }

            foreach (string line in logSampleFileContent)
            {
                DialogLogDetails += line;
            }
            if (DialogLogDetails.Equals(""))
            {
                if (!programMode.Equals("3"))
                {
                    theDialog.InitialDirectory = @"Y:\CD28";
                }
                else
                {
                    theDialog.InitialDirectory = @"W:\CPC";
                }
            }
            else
            {
                theDialog.InitialDirectory = DialogLogDetails;
            }

            System.IO.StreamWriter LogFileDialog2 = new System.IO.StreamWriter(@"G:\Peter_Tan\_____WORK_____\C#_Projects\CPC_Report_Summary\CPC_Report_Summary\DialogLog.log");
            LogFileDialog2.WriteLine((DialogLogDetails));
            LogFileDialog2.Close();


            if (programMode.Equals("0") || programMode.Equals(""))
            {
                xlApp.Quit();
                Console.WriteLine("\r\n\r\nTerminating CPC Project Manager...");
                System.Threading.Thread.Sleep(1000);
                System.Environment.Exit(1);
            }

            //Executing python scripts through this program
            if (programMode.Equals("5"))
            {
                try { Execute_Python_Script(@"G:\CPC\cpc_scripts_update_2016\cpc_weekly_report_2016_Oct_19_New.py"); xlApp.Quit(); Console.WriteLine("\r\n\r\nDaily CCE Sample Preparation Script was successful."); System.Threading.Thread.Sleep(1000); System.Environment.Exit(1); } catch { Console.WriteLine("Failed To Exectue Python Script for Generating Weekly Report"); xlApp.Quit(); System.Threading.Thread.Sleep(1000); System.Environment.Exit(1); }
            }
            if (programMode.Equals("6"))
            {
                try { Execute_Python_Script(@"G:\CPC\cpc_scripts_update_2016\test_cce.py"); xlApp.Quit(); Console.WriteLine("\r\n\r\nDaily CCE Sample Preparation Script was successful."); System.Threading.Thread.Sleep(1000); System.Environment.Exit(1); } catch { Console.WriteLine("Failed To Exectue Python Script for Generating Daily Sample"); xlApp.Quit(); System.Threading.Thread.Sleep(1000); System.Environment.Exit(1); }
            }



            if (programMode.Equals("4"))
            {
                theDialog.Title = "Open CSV File";
                theDialog.Filter = "Comma Separated Values|*.csv|Excel files|*.xlsx|All files|*";
            }
            else
            {
                theDialog.Title = "Open Excel File";
                theDialog.Filter = "Excel files|*.xlsx|Comma Separated Values|*.csv|All files|*";
            }
            if (theDialog.ShowDialog() == Application2.DialogResult.OK)
            {
                try
                {
                    if ((myStream = theDialog.OpenFile()) != null)
                    {
                        FilePath = theDialog.FileName;
                        FileName = theDialog.SafeFileName;
                        FilePath = FilePath.Replace(FileName, "");
                    }
                }
                catch { Console.WriteLine("Error: Could not read file from disk. Original error: "); }
            }
            Console.WriteLine("File Path Selected:\t" + FilePath);
            Console.WriteLine("File Name Selected:\t" + FileName);

            LogFileDialog2 = new System.IO.StreamWriter(@"G:\Peter_Tan\_____WORK_____\C#_Projects\CPC_Report_Summary\CPC_Report_Summary\DialogLog.log");
            LogFileDialog2.WriteLine(FilePath);
            LogFileDialog2.Close();


            //PROGRAM PROMPTS BEGINS
            Console.WriteLine("Enter the CPC Report PROJECT CODE:");
            projectCode = Console.ReadLine();
            Console.Write("# WEEKS IN " + projectCode + ": ");
            weekPeriod = Console.ReadLine();
            Console.Write("# CURRENT WEEK IN " + projectCode + ": ");
            currentWeek = Console.ReadLine();


            string ReportFilePathName = @"\\CC3-MGMT2\EXPORT\" + projectCode + @"\Weekly Data Report " + projectCode + "_" + dateTodayString + ".xlsx";
            string CodingFilePathName = @"\\CC3-MGMT2\EXPORT\" + projectCode + @"\" + projectCode + " " + dateTodayString + " Coding week " + currentWeek + " of " + weekPeriod + ".xlsx";

            //Count Loaded Sample per quota analysis for current date
            if (programMode == "4")
            {
                //try { AnalyzeDailyLoadedPerQuota(xlApp, @"\\CC3-MGMT2\EXPORT\" + projectCode + @"\CCE-Weekly_" + dateTodayString + "-" + projectCode + " All Loaded.CSV"); } catch { Console.WriteLine("Something went wrong."); }
                try { AnalyzeDailyLoadedPerQuota(xlApp, FilePath + FileName); } catch { Console.WriteLine("Something went wrong in retrieving the selected file from: " + FilePath + FileName); }
            }

            //Count Raw Sample provided per quota analysis for current date
            if (programMode == "3")
            {
                //try { AnalyzeRawSampleFiles(xlApp, @"W:\CPC\dailycontactfile" + dateYesterdayString + ".xlsx", dateTodayString); } catch { AnalyzeRawSampleFiles(xlApp, @"W:\CPC\Friday and Saturday file" + dateWeekendString + ".xlsx", dateTodayString); }
                try { AnalyzeRawSampleFiles(xlApp, FilePath + FileName, dateTodayString); } catch { Console.WriteLine("Something went wrong in retrieving the selected file from: " + FilePath + FileName); }
                //try { AnalyzeRawSampleFiles(xlApp, @"W:\CPC\Friday and Saturday file" + dateWeekendString + ".xlsx", dateTodayString); } catch { AnalyzeRawSampleFiles(xlApp, @"W:\CPC\Friday and Saturday file" + dateWeekendString + ".xlsx", dateTodayString); }

                string currentSampleFileLogDetails = "";
                string[] logRawSampleFileContent;
                try
                {
                    logRawSampleFileContent = System.IO.File.ReadAllLines(@"\\CC3-MGMT2\EXPORT\" + projectCode + @"\" + projectCode + @"_RawSampleBreakdownByQuota_Log.txt");
                }
                catch
                {
                    Console.WriteLine(mNoLogFile);
                    System.Diagnostics.Debug.WriteLine(mNoLogFile);
                    System.IO.StreamWriter newLogFile = new System.IO.StreamWriter(@"\\CC3-MGMT2\EXPORT\" + projectCode + @"\" + projectCode + @"_RawSampleBreakdownByQuota_Log.txt");
                    newLogFile.Close();
                    logRawSampleFileContent = System.IO.File.ReadAllLines(@"\\CC3-MGMT2\EXPORT\" + projectCode + @"\" + projectCode + @"_RawSampleBreakdownByQuota_Log.txt");
                }
                foreach (string logCodingFileLine in logRawSampleFileContent)
                {
                    currentSampleFileLogDetails += (logCodingFileLine + "\r\n");
                }
                System.IO.StreamWriter logFileRawSample = new System.IO.StreamWriter(@"\\CC3-MGMT2\EXPORT\" + projectCode + @"\" + projectCode + @"_RawSampleBreakdownByQuota_Log.txt");
                logFileRawSample.WriteLine((currentSampleFileLogDetails + mSampleFileDetailLog));
                logFileRawSample.Close();

            }

            if (programMode == "1")
            {
                CPC_Coding_Summary(xlApp, FilePath + FileName);

                string currentCodingLogDetails = "";
                string[] logCodingFileContent;
                try
                {
                    logCodingFileContent = System.IO.File.ReadAllLines(@"\\CC3-MGMT2\EXPORT\" + projectCode + @"\" + projectCode + @"_Coding_Log.txt");
                }
                catch
                {
                    Console.WriteLine(mNoLogFile);
                    System.Diagnostics.Debug.WriteLine(mNoLogFile);
                    System.IO.StreamWriter newLogFile = new System.IO.StreamWriter(@"\\CC3-MGMT2\EXPORT\" + projectCode + @"\" + projectCode + @"_Coding_Log.txt");
                    newLogFile.Close();
                    logCodingFileContent = System.IO.File.ReadAllLines(@"\\CC3-MGMT2\EXPORT\" + projectCode + @"\" + projectCode + @"_Coding_Log.txt");
                }
                foreach (string logCodingFileLine in logCodingFileContent)
                {
                    currentCodingLogDetails += (logCodingFileLine + "\r\n");
                }
                System.IO.StreamWriter logFileCodingCheck = new System.IO.StreamWriter(@"\\CC3-MGMT2\EXPORT\" + projectCode + @"\" + projectCode + @"_Coding_Log.txt");
                logFileCodingCheck.WriteLine((currentCodingLogDetails + mCodingDetailLog));
                logFileCodingCheck.Close();


                if (isCodingError)
                {
                    xlApp.Quit();
                    Console.WriteLine("Coding File Has Errors! " + mCodingDetailLog);
                    string aa = Console.ReadLine();
                    System.Environment.Exit(1);
                }
                else
                {
                    xlApp.Quit();
                    Console.WriteLine("Coding File Is Good!");
                    string aa = Console.ReadLine();
                    System.Environment.Exit(1);
                }
            }

            if (programMode == "2")
            {
                try
                {
                    string emailInput = "";

                    Console.WriteLine(mCheck);
                    System.Diagnostics.Debug.WriteLine(mCheck);

                    //MAIN FUNCTION
                    CPC_Report_Summary(xlApp, FilePath + FileName);

                    xlApp.Quit();

                    //MAIL OUT
                    mReportDetailEmail += "To whom it may concern,<br><br>Here's the Week " + currentWeek + " of " + weekPeriod + " " + projectCode + " Report, attached.<br><br>";
                    mReportDetailLog += "\tDate: " + dateTodayString + "\r\n\tCounts of Code 50\r\n\t----------------------------------------\r\n";
                    mReportDetailEmail += "&emsp;<font color=\"purple\">Date: " + dateTodayString + "<br>&emsp;Counts of Code 50<br>&emsp;----------------------------------------<br>";
                    mReportDetailLog += "\tQ3:\t" + OthersQ3 + "\r\n";
                    mReportDetailEmail += "&emsp;Q3:&emsp;" + OthersQ3 + "<br>";
                    mReportDetailLog += "\tQ3A:\t" + OthersQ3 + "\r\n";
                    mReportDetailEmail += "&emsp;Q3A:&emsp;" + OthersQ3A + "<br>";
                    mReportDetailLog += "\tQ3B:\t" + OthersQ3 + "\r\n";
                    mReportDetailEmail += "&emsp;Q3B:&emsp;" + OthersQ3B + "<br>";
                    mReportDetailLog += "\tQ3C:\t" + OthersQ3 + "\r\n";
                    mReportDetailEmail += "&emsp;Q3C:&emsp;" + OthersQ3C + "<br>";
                    mReportDetailLog += "\tQ3D:\t" + OthersQ3 + "\r\n";
                    mReportDetailEmail += "&emsp;Q3D:&emsp;" + OthersQ3D + "<br>";
                    mReportDetailLog += "\tQ3E:\t" + OthersQ3 + "\r\n";
                    mReportDetailEmail += "&emsp;Q3E:&emsp;" + OthersQ3E + "<br>";
                    mReportDetailLog += "\tQ4:\t" + OthersQ3 + "\r\n";
                    mReportDetailEmail += "&emsp;Q4:&emsp;" + OthersQ4 + "</font><br>";

                    if (InitialSurveyId_Index > 0 || TransferredSurveyId_Index > 0)
                    {
                        mReportDetailLog += ("\r\n" + missingAgentRow_Initial_Index + " record(s) missing AGNTUSER (<i>Initial</i>, Cell V");
                        mReportDetailEmail += ("<br><b><font color=\"indianred\">" + missingAgentRow_Initial_Index + "</font></b> record(s) missing AGNTUSER (<i>Initial</i>, Cell V");
                        if (missingAgentRow_Initial_Index > 0)
                        {
                            for (int i = 0; i < missingAgentRow_Initial_Index; i++)
                            {
                                mReportDetailLog += missingAgentRow_Initial_List[i];
                                mReportDetailEmail += missingAgentRow_Initial_List[i];
                                if ((i + 1) == missingAgentRow_Initial_Index)
                                {
                                    mReportDetailLog += ")";
                                    mReportDetailEmail += ")";
                                }
                                else
                                {
                                    mReportDetailLog += ", ";
                                    mReportDetailEmail += ", ";
                                }
                            }
                            mReportDetailLog += "\r\n";
                            mReportDetailEmail += "<br>";
                        }
                        else
                        {
                            mReportDetailLog += ("0 ROWS IDENTIFIED WHERE AGENT USER IS MISSING (INITIAL AND TRANSFERRED) IN " + projectCode + "\r\n");
                            mReportDetailEmail += ("<b>0</b> ROWS IDENTIFIED WHERE AGENT USER IS MISSING (INITIAL AND TRANSFERRED) IN " + projectCode + "<br>");
                        }
                        for (int i = 0; i < InitialSurveyId_Index; i++)
                        {
                            mReportDetailLog += ("Initial Call Survey ID:\t\t" + InitialSurveyId_List[i] + "\r\n");
                            mReportDetailEmail += ("<font color=\"teal\">Initial Call Survey ID:</font>\t\t" + InitialSurveyId_List[i] + "<br>");
                        }

                        mReportDetailLog += ("\r\n" + missingAgentRow_Transferred_Index + " record(s) missing AGNTUSER (<i>Transferred</i>, Cell AE");
                        mReportDetailEmail += ("<br><b><font color=\"indianred\">" + missingAgentRow_Transferred_Index + "</font></b> record(s) missing AGNTUSER (<i>Transferred</i>, Cell AE");
                        if (missingAgentRow_Transferred_Index > 0)
                        {
                            for (int i = 0; i < missingAgentRow_Transferred_Index; i++)
                            {
                                mReportDetailLog += missingAgentRow_Transferred_List[i];
                                mReportDetailEmail += missingAgentRow_Transferred_List[i];
                                if ((i + 1) == missingAgentRow_Transferred_Index)
                                {
                                    mReportDetailLog += ")";
                                    mReportDetailEmail += ")";
                                }
                                else
                                {
                                    mReportDetailLog += ", ";
                                    mReportDetailEmail += ", ";
                                }
                            }
                            mReportDetailLog += "\r\n";
                            mReportDetailEmail += "<br>";
                        }
                        for (int i = 0; i < TransferredSurveyId_Index; i++)
                        {
                            mReportDetailLog += ("Transferred Call Survey ID:\t" + TransferredSurveyId_List[i] + "\r\n");
                            mReportDetailEmail += ("<font color=\"teal\">Transferred Call Survey ID:</font>\t" + TransferredSurveyId_List[i] + "<br>");
                        }
                    }
                    else
                    {
                        mReportDetailLog += ("0 MISSING AGENTS (INITIAL AND TRANSFERRED) IN " + projectCode + "\r\n");
                        mReportDetailEmail += ("<b>0</b> MISSING AGENTS (INITIAL AND TRANSFERRED) IN " + projectCode + "<br>");
                    }
                    mReportDetailLog += ("\r\nCompleted at:\t" + dateTodayString + "\t----------------------------------------\r\n");

                    Console.WriteLine("Choose Mailing Lists [1 = Mail to yourself] [2 = Mail to General Mailing List]");
                    string emailOption = Console.ReadLine();

                    if (emailOption == "2")
                    {
                        SendReportEmail("AArgyropoulos@forumresearch.com", "ssinnott@forumresearch.com", "CHollyer@access-research.com", "dcadieux@forumresearch.com", "CVanHerpt@forumresearch.com", "ptan@forumresearch.com", "MChand@forumresearch.com");
                    }
                    else if (emailOption == "1")
                    {
                        SendReportEmail("ptan@forumresearch.com", "", "", "", "", "", "");
                    }

                    string currentLogDetails = "";
                    string[] logFileContent;
                    try
                    {
                        logFileContent = System.IO.File.ReadAllLines(@"\\CC3-MGMT2\EXPORT\" + projectCode + @"\" + projectCode + @"_Report_Log.txt");
                    }
                    catch
                    {
                        Console.WriteLine(mNoLogFile);
                        System.Diagnostics.Debug.WriteLine(mNoLogFile);
                        System.IO.StreamWriter newLogFile = new System.IO.StreamWriter(@"\\CC3-MGMT2\EXPORT\" + projectCode + @"\" + projectCode + @"_Report_Log.txt");
                        newLogFile.Close();
                        logFileContent = System.IO.File.ReadAllLines(@"\\CC3-MGMT2\EXPORT\" + projectCode + @"\" + projectCode + @"_Report_Log.txt");
                    }
                    foreach (string logFileLine in logFileContent)
                    {
                        currentLogDetails += (logFileLine + "\r\n");
                    }
                    System.IO.StreamWriter logFileWeeklyReport = new System.IO.StreamWriter(@"\\CC3-MGMT2\EXPORT\" + projectCode + @"\" + projectCode + @"_Report_Log.txt");
                    logFileWeeklyReport.WriteLine((currentLogDetails + mReportDetailLog));
                    logFileWeeklyReport.Close();

                    Console.WriteLine(mFinishedSuccessfully);
                    System.Diagnostics.Debug.WriteLine(mFinishedSuccessfully);
                    string a = Console.ReadLine();
                    System.Environment.Exit(1);

                }
                catch
                {
                    Console.WriteLine("ERROR CAUGHT. QUITTING APPLICATION");
                    string a = Console.ReadLine();
                    System.Environment.Exit(1);
                }
            }
        }

        public static void Execute_Python_Script(string fileName)
        {

            //string fileName = @"G:\CPC\cpc_scripts_update_2016\cpc_weekly_report_2016_Oct_19_New.py";
            
            Process p = new Process();
            p.StartInfo = new ProcessStartInfo(@"C:\Python34\python.exe", fileName)
            {
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            p.Start();

            string output = p.StandardOutput.ReadToEnd();
            p.WaitForExit();

            Console.WriteLine(output);
            Console.WriteLine("\r\n\r\nPRESS ENTER OR ANY KEY TO TERMINATE");
            Console.ReadLine();

        }

    }
}