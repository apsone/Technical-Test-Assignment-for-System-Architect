/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * Jan Källman                      Added                       2012-03-04  
 *******************************************************************************/
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Exceptions;
namespace OfficeOpenXml
{

    public static class FundingConstants_V2
    {
        public struct PROGRAM_TYPE
        {
            public static string FTS = "FTS";
            public static string IBF_STS = "IBF-STS";
        }

        public struct CLAIMS_WORKSHEET_FORMULAS
        {
            public static string ROW_NO = "A{0}+1";

            public struct FI_WORKSHEET
            {
                public static string TOTAL_INTERNAL_COST = "SUM(L{0}:M{1})";
                public static string TOTAL_PROGRAMME_COST = "IF(N{0}=0,O{1},N{2})";
                public static string GRANT_AMOUNT_FTS = "IF(EXACT(P{0},\"FTS\"),IF(AND(EXACT(E{1},\"Yes\"),EXACT(D{2},$X$7),IF(ISERR(DATEVALUE(LEFT(I{3},4)&\"-\"&MID(I{4},5,2)&\"-\"&RIGHT(I{5},2))),FALSE,DATEVALUE(LEFT(I{6},4)&\"-\"&MID(I{7},5,2)&\"-\"&RIGHT(I{8},2))>=$Z$7)),ROUND(MIN(Q{9}*0.9,2000),2),ROUND(MIN(Q{10}*0.5,2000),2)),\"\")";
                public static string GRANT_AMOUNT_IBFSTS = "IF(EXACT(P{0},\"IBF-STS\"),IF(AND(EXACT(E{1},\"Yes\"),EXACT(D{2},$X$7),IF(ISERR(DATEVALUE(LEFT(I{3},4)&\"-\"&MID(I{4},5,2)&\"-\"&RIGHT(I{5},2))),FALSE,DATEVALUE(LEFT(I{6},4)&\"-\"&MID(I{7},5,2)&\"-\"&RIGHT(I{8},2))>=$Z$7)),ROUND(MIN(Q{9}*0.9,7000),2),ROUND(MIN(Q{10}*0.7,7000),2)),\"\")";

            }

            public struct FTP_WORKSHEET
            {
                public static string PROGRAMME_FEE = "";
                public static string GRANT_AMOUNT_IBFSTS = "IF(AND(EXACT(E{0},\"Yes\"),EXACT(D{1},$X$10),IF(ISERR(DATEVALUE(LEFT(J{2},4)&\"-\"&MID(J{3},5,2)&\"-\"&RIGHT(J{4},2))),FALSE,DATEVALUE(LEFT(J{5},4)&\"-\"&MID(J{6},5,2)&\"-\"&RIGHT(J{7},2))>=$Z$10)),ROUND(MIN(M{8}*0.9,7000),2),ROUND(MIN(M{9}*0.7,7000),2))";
            }
        }

        public struct CLAIMS_WORKSHEET_COLUMN_COUNTS
        {
            public static int FTPColumnNumbers = 15;
            public static int FIColumnNumbers = 19;
        }

        public struct CLAIMS_WORKSHEET_START_ROWS
        {
            public static int FI_START_ROW = 22;
            public static int FTP_START_ROW = 16;
        }

        public struct CLAIMS_WORKSHEET_VALIDATOR_COLUMNS
        {
            public static string FI_COLUMN_HEADER = "To be completed for In-house Developed Programmes only";
            public static int FI_COLUMN = 12;
            public static int FI_ROW = 19;

            public static string FTP_COLUMN_HEADER = "programme period";
            public static int FTP_COLUMN = 10;
            public static int FTP_ROW = 13;
        }

        public struct CLAIMS_WORKSHEET_CONSTANTS
        {
            public static string NRIC_SC = "Singapore Citizen";
            public static DateTime SCAGEFILTER_DATE_THRESHOLD = new DateTime(2016, 7, 1);
        }

        public struct CLAIMS_WORKSHEET_COLUMN_INDEXES
        {
            public struct FI_WORKSHEET
            {

                public static int NRIC_TYPE = 4;
                public static int SC_AGE_FILTER = 5;
                public static int PROG_START_DATE = 9;
                public static int PROG_COMPL_DATE = 11;

                public static int INHOUSE_COST = 12;
                public static int INHOUSE_FEECHARGED = 13;

                public static int EXTERNAL_TOTAL = 15;
                public static int TOTAL_INTERNAL_COST = 14;
                public static int TYPE_OF_SCHEME = 16;
                public static int TOTAL_PROGRAMME_COST = 17;
                public static int GRANT_AMOUNT_FTS = 18;
                public static int GRANT_AMOUNT_IBFSTS = 19;
            }

            public struct FTP_WORKSHEET
            {

                public static int NRIC_TYPE = 4;
                public static int SC_AGE_FILTER = 5;
                public static int PROG_START_DATE = 10;
                public static int PROG_COMPL_DATE = 12;

                public static int PROGRAMME_FEE = 13;
                public static int GRANT_AMOUNT_IBFSTS = 15;
            }
        }
    }

    public static class CalculationExtensions_V2
    {
        public static void Calculate_V2(this ExcelWorkbook workbook)
        {
            Calculate_V2(workbook, new ExcelCalculationOption() { AllowCirculareReferences = false });
        }
        public static void Calculate_V2(this ExcelWorkbook workbook, ExcelCalculationOption options)
        {
            Init(workbook);

            var dc = DependencyChainFactory.Create(workbook, options);
            workbook._formulaParser = null;
            var parser = workbook.FormulaParser;

            //TODO: Remove when tests are done. Outputs the dc to a text file. 
            //var fileDc = new System.IO.StreamWriter("c:\\temp\\dc.txt");

            //for (int i = 0; i < dc.list.Count; i++)
            //{
            //    fileDc.WriteLine(i.ToString() + "," + dc.list[i].Column.ToString() + "," + dc.list[i].Row.ToString() + "," + (dc.list[i].ws==null ? "" : dc.list[i].ws.Name) + "," + dc.list[i].Formula);
            //}
            //fileDc.Close();
            //fileDc = new System.IO.StreamWriter("c:\\temp\\dcorder.txt");
            //for (int i = 0; i < dc.CalcOrder.Count; i++)
            //{
            //    fileDc.WriteLine(dc.CalcOrder[i].ToString());
            //}
            //fileDc.Close();
            //fileDc = null;

            //TODO: Add calculation here

            CalcChain_V2(workbook, parser, dc);

            workbook._isCalculated = true;
        }
        public static void Calculate_V2(this ExcelWorksheet worksheet)
        {
            Calculate_V2(worksheet, new ExcelCalculationOption());
        }
        public static void Calculate_V2(this ExcelWorksheet worksheet, ExcelCalculationOption options)
        {
            Init(worksheet.Workbook);
            worksheet.Workbook._formulaParser = null;
            var parser = worksheet.Workbook.FormulaParser;
            var dc = DependencyChainFactory.Create(worksheet, options);
            CalcChain_V2(worksheet.Workbook, parser, dc);
        }
        public static void Calculate_V2(this ExcelRangeBase range)
        {
            Calculate_V2(range, new ExcelCalculationOption());
        }
        public static void Calculate_V2(this ExcelRangeBase range, ExcelCalculationOption options)
        {
            Init(range._workbook);
            var parser = range._workbook.FormulaParser;
            var dc = DependencyChainFactory.Create(range, options);
            CalcChain_V2(range._workbook, parser, dc, range);
        }
        public static object Calculate_V2(this ExcelWorksheet worksheet, string Formula)
        {
            return Calculate_V2(worksheet, Formula, new ExcelCalculationOption());
        }
        public static object Calculate_V2(this ExcelWorksheet worksheet, string Formula, ExcelCalculationOption options)
        {
            try
            {
                worksheet.CheckSheetType();
                if (string.IsNullOrEmpty(Formula.Trim())) return null;
                Init(worksheet.Workbook);
                var parser = worksheet.Workbook.FormulaParser;
                if (Formula[0] == '=') Formula = Formula.Substring(1); //Remove any starting equal sign
                var dc = DependencyChainFactory.Create(worksheet, Formula, options);
                var f = dc.list[0];
                dc.CalcOrder.RemoveAt(dc.CalcOrder.Count - 1);

                CalcChain_V2(worksheet.Workbook, parser, dc);

                return parser.ParseCell(f.Tokens, worksheet.Name, -1, -1);
            }
            catch (Exception ex)
            {
                return new ExcelErrorValueException(ex.Message, ExcelErrorValue.Create(eErrorType.Value));
            }
        }

        private static HashSet<int> rows = new HashSet<int>();
        private static int sheetId = 0;
        private static void CalcChain_V2(ExcelWorkbook wb, FormulaParser parser, DependencyChain dc, ExcelRangeBase range = null)
        {
            if (dc.CalcOrder.Count == 0)
            {
                if (range != null && range.Address.ToString().StartsWith("P"))
                {

                }
            }

            int fiColValidator = FundingConstants_V2.CLAIMS_WORKSHEET_VALIDATOR_COLUMNS.FI_COLUMN;
            int fiRowValidator = FundingConstants_V2.CLAIMS_WORKSHEET_VALIDATOR_COLUMNS.FI_ROW;

            int ftpColValidator = FundingConstants_V2.CLAIMS_WORKSHEET_VALIDATOR_COLUMNS.FTP_COLUMN;
            int ftpRowValidator = FundingConstants_V2.CLAIMS_WORKSHEET_VALIDATOR_COLUMNS.FTP_ROW;

            foreach (var ix in dc.CalcOrder)
            {
                var item = dc.list[ix];
                rows.Add(item.Row);
                sheetId = item.SheetID;

                try
                {
                    var ws = wb.Worksheets.GetBySheetID(item.SheetID);

                    bool goDefault = true;
                    if (ws.Cells[fiRowValidator, fiColValidator] != null &&
                        ws.Cells[fiRowValidator, fiColValidator].Text != "" &&
                        ws.Cells[fiRowValidator, fiColValidator].Text.ToLower() == FundingConstants_V2.CLAIMS_WORKSHEET_VALIDATOR_COLUMNS.FI_COLUMN_HEADER.ToLower())
                    { // FI Worksheet
                        #region FI WORKSHEETS
                        if (range != null &&
                            item.Column == FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.TOTAL_PROGRAMME_COST &&
                            range.Address.ToString().StartsWith("P") ||
                            item.Column == FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.GRANT_AMOUNT_FTS)
                        {
                            if (ws.Cells[item.Row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.TYPE_OF_SCHEME].Text != null &&
                                ws.Cells[item.Row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.TYPE_OF_SCHEME].Text != "" &&
                                ws.Cells[item.Row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.TYPE_OF_SCHEME].Text.ToLower() == FundingConstants_V2.PROGRAM_TYPE.FTS.ToLower())
                            {
                                decimal x1 = 0;
                                bool isValidX1 = Decimal.TryParse(ws.Cells[item.Row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.TOTAL_PROGRAMME_COST].Text.Replace("$", ""), out x1);

                                if (isValidX1)
                                {
                                    decimal grantPercentage = .5m;
                                    string nrictype = ws.Cells[item.Row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.NRIC_TYPE].Text.ToString();
                                    string progStartDateStr = ws.Cells[item.Row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.PROG_START_DATE].Text.ToString();
                                    string scAgeFilterStr = ws.Cells[item.Row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.SC_AGE_FILTER].Text.ToString();
                                    bool isSCAgeFilter = !string.IsNullOrEmpty(scAgeFilterStr) && scAgeFilterStr.ToLower() == "yes";
                                    DateTime progStartDate = parseDate(progStartDateStr);

                                    if (nrictype == FundingConstants_V2.CLAIMS_WORKSHEET_CONSTANTS.NRIC_SC
                                        && progStartDate.CompareTo(FundingConstants_V2.CLAIMS_WORKSHEET_CONSTANTS.SCAGEFILTER_DATE_THRESHOLD) >= 0
                                        && isSCAgeFilter
                                       )
                                    {
                                        grantPercentage = .9m;
                                    }

                                    x1 = x1 * grantPercentage;
                                    x1 = Math.Round(x1, 2, MidpointRounding.AwayFromZero);

                                    string formula = string.Format(FundingConstants_V2.CLAIMS_WORKSHEET_FORMULAS.FI_WORKSHEET.GRANT_AMOUNT_FTS, item.Row, item.Row, item.Row, item.Row, item.Row, item.Row, item.Row, item.Row, item.Row, item.Row, item.Row);

                                    if (item.Formula == null)
                                    {
                                        x1 = 0;
                                    }
                                    else if (item.Formula != null && formula != item.Formula.Trim().Replace(" ", ""))
                                    {
                                        x1 = 0;
                                    }
                                    else if (x1 > 2000.00m)
                                    {
                                        x1 = 2000.00m;
                                    }

                                    SetValue(wb, item, x1);
                                    goDefault = false;
                                }
                                else
                                {
                                    SetValue(wb, item, "");
                                    goDefault = false;
                                }
                            }
                            else
                            {
                                SetValue(wb, item, "");
                                goDefault = false;
                            }

                        }
                        else if (range != null &&
                            item.Column == FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.TOTAL_PROGRAMME_COST &&
                            range.Address.ToString().StartsWith("P") ||
                            item.Column == FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.GRANT_AMOUNT_IBFSTS)
                        {
                            if (ws.Cells[item.Row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.TYPE_OF_SCHEME].Text != null &&
                                ws.Cells[item.Row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.TYPE_OF_SCHEME].Text != "" &&
                                ws.Cells[item.Row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.TYPE_OF_SCHEME].Text.ToLower() == FundingConstants_V2.PROGRAM_TYPE.IBF_STS.ToLower())
                            {
                                decimal x1 = 0;
                                bool isValidX1 = Decimal.TryParse(ws.Cells[item.Row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.TOTAL_PROGRAMME_COST].Text.Replace("$", ""), out x1);

                                //if(isValidX1)
                                //{


                                decimal grantPercentage = .7m;
                                string nrictype = ws.Cells[item.Row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.NRIC_TYPE].Text.ToString();
                                string progStartDateStr = ws.Cells[item.Row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.PROG_START_DATE].Text.ToString();
                                string scAgeFilterStr = ws.Cells[item.Row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.SC_AGE_FILTER].Text.ToString();
                                bool isSCAgeFilter = !string.IsNullOrEmpty(scAgeFilterStr) && scAgeFilterStr.ToLower() == "yes";
                                DateTime progStartDate = parseDate(progStartDateStr);

                                if (nrictype == FundingConstants_V2.CLAIMS_WORKSHEET_CONSTANTS.NRIC_SC
                                    && progStartDate.CompareTo(FundingConstants_V2.CLAIMS_WORKSHEET_CONSTANTS.SCAGEFILTER_DATE_THRESHOLD) >= 0
                                    && isSCAgeFilter
                                   )
                                {
                                    grantPercentage = .9m;
                                }

                                x1 = x1 * grantPercentage;
                                x1 = Math.Round(x1, 2, MidpointRounding.AwayFromZero);

                                string formula = string.Format(FundingConstants_V2.CLAIMS_WORKSHEET_FORMULAS.FI_WORKSHEET.GRANT_AMOUNT_IBFSTS, item.Row, item.Row, item.Row, item.Row, item.Row, item.Row, item.Row, item.Row, item.Row, item.Row, item.Row);

                                if (item.Formula == null)
                                {
                                    x1 = 0;
                                }
                                else if (item.Formula != null && formula != item.Formula.Trim().Replace(" ", ""))
                                {
                                    x1 = 0;
                                }
                                else if (x1 > 7000.00m)
                                {
                                    x1 = 7000.00m;
                                }

                                SetValue(wb, item, x1);
                                goDefault = false;
                                //}
                                //else
                                //{
                                //    SetValue(wb, item, "");
                                //    goDefault = false;
                                //}
                            }
                            else
                            {
                                SetValue(wb, item, "");
                                goDefault = false;
                            }

                        }
                        else if (item.Column == FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.TOTAL_INTERNAL_COST)
                        {
                            //if (ws.Cells[item.Row, FundingConstants.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.INHOUSE_COST].Text != "" ||
                            //    ws.Cells[item.Row, FundingConstants.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.INHOUSE_FEECHARGED].Text != "")
                            //{
                            decimal x1 = 0;
                            decimal x2 = 0;

                            bool isValidX1 = Decimal.TryParse(ws.Cells[item.Row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.INHOUSE_COST].Text.Replace("$", ""), out x1);
                            bool isValidX2 = decimal.TryParse(ws.Cells[item.Row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.INHOUSE_FEECHARGED].Text.Replace("$", ""), out x2);

                            string formula = string.Format(FundingConstants_V2.CLAIMS_WORKSHEET_FORMULAS.FI_WORKSHEET.TOTAL_INTERNAL_COST, item.Row, item.Row);

                            decimal sum = x1 + x2;

                            if (item.Formula == null)
                            {
                                sum = 0;
                            }
                            else if (item.Formula != null && formula != item.Formula.Trim().Replace(" ", ""))
                            {
                                sum = 0;
                            }
                            else
                            {
                                SetValue(wb, item, sum);
                                goDefault = false;
                            }
                        }
                        else if (item.Column == FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.TOTAL_PROGRAMME_COST)
                        {
                            //if (ws.Cells[item.Row, FundingConstants.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.INHOUSE_COST].Text != "" && 
                            //    ws.Cells[item.Row, FundingConstants.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.INHOUSE_FEECHARGED].Text != "")
                            decimal x1a = 0;
                            decimal x2a = 0;

                            bool isValidX1a = Decimal.TryParse(ws.Cells[item.Row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.INHOUSE_COST].Text.Replace("$", ""), out x1a);
                            bool isValidX2a = decimal.TryParse(ws.Cells[item.Row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.INHOUSE_FEECHARGED].Text.Replace("$", ""), out x2a);

                            if ((x1a + x2a) != 0m)
                            {
                                decimal x1 = 0;
                                decimal x2 = 0;

                                bool isValidX1 = Decimal.TryParse(ws.Cells[item.Row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.INHOUSE_COST].Text.Replace("$", ""), out x1);
                                bool isValidX2 = decimal.TryParse(ws.Cells[item.Row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.INHOUSE_FEECHARGED].Text.Replace("$", ""), out x2);

                                string formula = string.Format(FundingConstants_V2.CLAIMS_WORKSHEET_FORMULAS.FI_WORKSHEET.TOTAL_PROGRAMME_COST, item.Row, item.Row, item.Row);

                                decimal sum = x1 + x2;

                                if (item.Formula == null)
                                {
                                    sum = 0;
                                }
                                else if (item.Formula != null && formula != item.Formula.Trim().Replace(" ", ""))
                                {
                                    sum = 0;
                                }
                                else
                                {
                                    SetValue(wb, item, sum);
                                    goDefault = false;
                                }
                            }
                            else if (ws.Cells[item.Row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.EXTERNAL_TOTAL].Text != "")
                            {
                                decimal x1 = 0;

                                bool isValidX1 = Decimal.TryParse(ws.Cells[item.Row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.EXTERNAL_TOTAL].Value + "", out x1);

                                string formula = string.Format(FundingConstants_V2.CLAIMS_WORKSHEET_FORMULAS.FI_WORKSHEET.TOTAL_PROGRAMME_COST, item.Row, item.Row, item.Row);


                                if (item.Formula == null)
                                {
                                    x1 = 0;
                                }
                                else if (item.Formula != null && formula != item.Formula.Trim().Replace(" ", ""))
                                {
                                    x1 = 0;
                                }
                                else
                                {
                                    SetValue(wb, item, x1);
                                    goDefault = false;
                                }
                            }
                        }
                        #endregion
                    }
                    else if (ws.Cells[ftpRowValidator, ftpColValidator] != null &&
                        ws.Cells[ftpRowValidator, ftpColValidator].Text != "" &&
                        ws.Cells[ftpRowValidator, ftpColValidator].Text.ToLower() == FundingConstants_V2.CLAIMS_WORKSHEET_VALIDATOR_COLUMNS.FTP_COLUMN_HEADER.ToLower())
                    { // FTP Worksheet
                        #region FTP Worksheet
                        if (range != null &&
                            item.Column == FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FTP_WORKSHEET.GRANT_AMOUNT_IBFSTS &&
                            range.Address.ToString().StartsWith("M") ||
                            item.Column == FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FTP_WORKSHEET.GRANT_AMOUNT_IBFSTS)
                        {
                            decimal x1 = 0;
                            bool isValidX1 = Decimal.TryParse(ws.Cells[item.Row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FTP_WORKSHEET.PROGRAMME_FEE].Value.ToString(), out x1);

                            //if(isValidX1)
                            //{

                            decimal grantPercentage = .7m;
                            string nrictype = ws.Cells[item.Row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FTP_WORKSHEET.NRIC_TYPE].Text.ToString();
                            string progStartDateStr = ws.Cells[item.Row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FTP_WORKSHEET.PROG_START_DATE].Text.ToString();
                            string scAgeFilterStr = ws.Cells[item.Row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FTP_WORKSHEET.SC_AGE_FILTER].Text.ToString();
                            bool isSCAgeFilter = !string.IsNullOrEmpty(scAgeFilterStr) && scAgeFilterStr.ToLower() == "yes";
                            DateTime progStartDate = parseDate(progStartDateStr);

                            if (nrictype == FundingConstants_V2.CLAIMS_WORKSHEET_CONSTANTS.NRIC_SC
                                && progStartDate.CompareTo(FundingConstants_V2.CLAIMS_WORKSHEET_CONSTANTS.SCAGEFILTER_DATE_THRESHOLD) >= 0
                                && isSCAgeFilter
                               )
                            {
                                grantPercentage = .9m;


                            }
                            x1 = x1 * grantPercentage;
                            x1 = Math.Round(x1, 2, MidpointRounding.AwayFromZero);


                            //string formula = "IF(EXACT(N"+item.Row+",\"IBF-STS\"),MIN(O"+item.Row+"*0.5,2000),\"\")";
                            // =ROUND(MIN(K16*0.7,7000),2)
                            // string formula = "MIN(K"+item.Row+"*0.7,7000)";
                            string formula = string.Format(FundingConstants_V2.CLAIMS_WORKSHEET_FORMULAS.FTP_WORKSHEET.GRANT_AMOUNT_IBFSTS, item.Row, item.Row, item.Row, item.Row, item.Row, item.Row, item.Row, item.Row, item.Row, item.Row);

                            if (item.Formula == null)
                            {
                                x1 = 0;
                            }
                            else if (item.Formula != null && formula != item.Formula.Trim().Replace(" ", ""))
                            {
                                x1 = 0;
                            }
                            else if (x1 > 7000.00m)
                            {
                                x1 = 7000.00m;
                            }

                            SetValue(wb, item, x1);
                            goDefault = false;
                            //}
                            //else
                            //{
                            //    SetValue(wb, item, "");
                            //    goDefault = false;
                            //}

                        }
                        #endregion
                    }

                    if (goDefault)
                    {
                        var v = parser.ParseCell(item.Tokens, ws == null ? "" : ws.Name, item.Row, item.Column);
                        SetValue(wb, item, v);
                    }
                }
                catch (FormatException fe)
                {
                    throw (fe);
                }
                catch (Exception e)
                {
                    var error = ExcelErrorValue.Parse(ExcelErrorValue.Values.Value);
                    SetValue(wb, item, error);
                }
            }

            var worksheet = wb.Worksheets.GetBySheetID(sheetId);
            foreach (int row in rows)
            {
                if (worksheet.Cells[fiRowValidator, fiColValidator] != null &&
                    worksheet.Cells[fiRowValidator, fiColValidator].Text != "" &&
                    worksheet.Cells[fiRowValidator, fiColValidator].Text.ToLower() == FundingConstants_V2.CLAIMS_WORKSHEET_VALIDATOR_COLUMNS.FI_COLUMN_HEADER.ToLower())
                {
                    #region FI WORKSHEETS
                    if (row >= 22)
                    {
                        if (String.IsNullOrEmpty(worksheet.Cells[row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.GRANT_AMOUNT_FTS].Formula) ||
                        (!String.IsNullOrEmpty(worksheet.Cells[row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.GRANT_AMOUNT_FTS].Formula) &&
                        worksheet.Cells[row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.GRANT_AMOUNT_FTS].Formula != string.Format(FundingConstants_V2.CLAIMS_WORKSHEET_FORMULAS.FI_WORKSHEET.GRANT_AMOUNT_FTS, row, row, row, row, row, row, row, row, row, row, row)))
                        {
                            worksheet._values.SetValue(row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.GRANT_AMOUNT_FTS, 0);
                        }

                        if (String.IsNullOrEmpty(worksheet.Cells[row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.GRANT_AMOUNT_IBFSTS].Formula) ||
                            (!String.IsNullOrEmpty(worksheet.Cells[row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.GRANT_AMOUNT_IBFSTS].Formula) &&
                            worksheet.Cells[row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.GRANT_AMOUNT_IBFSTS].Formula != string.Format(FundingConstants_V2.CLAIMS_WORKSHEET_FORMULAS.FI_WORKSHEET.GRANT_AMOUNT_IBFSTS, row, row, row, row, row, row, row, row, row, row, row)))
                        {
                            worksheet._values.SetValue(row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.GRANT_AMOUNT_IBFSTS, 0);
                        }

                        if (String.IsNullOrEmpty(worksheet.Cells[row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.TOTAL_INTERNAL_COST].Formula) ||
                            (!String.IsNullOrEmpty(worksheet.Cells[row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.TOTAL_INTERNAL_COST].Formula) &&
                            worksheet.Cells[row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.TOTAL_INTERNAL_COST].Formula != string.Format(FundingConstants_V2.CLAIMS_WORKSHEET_FORMULAS.FI_WORKSHEET.TOTAL_INTERNAL_COST, row, row)))
                        {
                            worksheet._values.SetValue(row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.TOTAL_INTERNAL_COST, 0);
                        }


                        if (String.IsNullOrEmpty(worksheet.Cells[row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.TOTAL_PROGRAMME_COST].Formula) ||
                            (!String.IsNullOrEmpty(worksheet.Cells[row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.TOTAL_PROGRAMME_COST].Formula) &&
                            worksheet.Cells[row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.TOTAL_PROGRAMME_COST].Formula != string.Format(FundingConstants_V2.CLAIMS_WORKSHEET_FORMULAS.FI_WORKSHEET.TOTAL_PROGRAMME_COST, row, row, row)))
                        {
                            worksheet._values.SetValue(row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FI_WORKSHEET.TOTAL_PROGRAMME_COST, 0);
                        }
                    }
                    #endregion
                }
                else if (worksheet.Cells[ftpRowValidator, ftpColValidator] != null &&
                        worksheet.Cells[ftpRowValidator, ftpColValidator].Text != "" &&
                        worksheet.Cells[ftpRowValidator, ftpColValidator].Text.ToLower() == FundingConstants_V2.CLAIMS_WORKSHEET_VALIDATOR_COLUMNS.FTP_COLUMN_HEADER.ToLower())
                {
                    #region FTP Worksheets
                    if (row >= 16)
                    {
                        if (String.IsNullOrEmpty(worksheet.Cells[row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FTP_WORKSHEET.GRANT_AMOUNT_IBFSTS].Formula) ||
                        (!String.IsNullOrEmpty(worksheet.Cells[row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FTP_WORKSHEET.GRANT_AMOUNT_IBFSTS].Formula) &&
                        worksheet.Cells[row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FTP_WORKSHEET.GRANT_AMOUNT_IBFSTS].Formula != string.Format(FundingConstants_V2.CLAIMS_WORKSHEET_FORMULAS.FTP_WORKSHEET.GRANT_AMOUNT_IBFSTS, row, row, row, row, row, row, row, row, row, row)))
                        {
                            worksheet.Cells[row, FundingConstants_V2.CLAIMS_WORKSHEET_COLUMN_INDEXES.FTP_WORKSHEET.GRANT_AMOUNT_IBFSTS].Value = 0;
                        }
                    }
                    #endregion
                }
            }
        }
        private static void Init(ExcelWorkbook workbook)
        {
            workbook._formulaTokens = new CellStore<List<Token>>(); ;
            foreach (var ws in workbook.Worksheets)
            {
                if (!(ws is ExcelChartsheet))
                {
                    if (ws._formulaTokens != null)
                    {
                        ws._formulaTokens.Dispose();
                    }
                    ws._formulaTokens = new CellStore<List<Token>>();
                }
            }
        }

        private static void SetValue(ExcelWorkbook workbook, FormulaCell item, object v)
        {
            if (item.Column == 0)
            {
                if (item.SheetID <= 0)
                {
                    workbook.Names[item.Row].NameValue = v;
                }
                else
                {
                    var sh = workbook.Worksheets.GetBySheetID(item.SheetID);
                    sh.Names[item.Row].NameValue = v;
                }
            }
            else
            {
                var sheet = workbook.Worksheets.GetBySheetID(item.SheetID);
                sheet._values.SetValue(item.Row, item.Column, v);
            }
        }

        private static DateTime parseDate(string str)
        {
            // yyyymmdd
            if (str.Length == 8)
            {
                int year = -1;
                int month = -1;
                int date = -1;
                int.TryParse(str.Substring(0, 4), out year);
                if (year == -1) return DateTime.MinValue;
                int.TryParse(str.Substring(4, 2), out month);
                if (month == -1) return DateTime.MinValue;
                int.TryParse(str.Substring(6, 2), out date);
                if (date == -1) return DateTime.MinValue;
                return new DateTime(year, month, date);
            }
            return DateTime.MinValue;
        }
    }
}
