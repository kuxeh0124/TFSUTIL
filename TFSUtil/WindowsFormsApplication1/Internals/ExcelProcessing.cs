using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using excelInterop = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using TFSUtil.Internals;
using System.IO;
using HtmlAgilityPack;
using System.Text.RegularExpressions;
using System.Drawing;

namespace TFSUtil.Internals
{
    class ExcelProcessing
    {
        static string outputPath = (System.IO.Path.GetDirectoryName(
                System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase).ToString()).Substring(6);
        public string destPath = outputPath + @"\" + "Upload";
        public string templateSource = outputPath + @"\Templates";
        public Dictionary<string, string[]> dicExData = new Dictionary<string, string[]>();
        public List<int> getAllID = new List<int>();
        public int stepStart = 1;

        public static void ReleaseExcel(excelInterop.Workbook thisWb = null,
                                                         excelInterop.Worksheet thisWs = null,
                                                         excelInterop.Application thisApp = null,
                                                         excelInterop.Sheets thisSheet = null)
        {
            //release all memory - stop EXCEL.exe from hanging around.
            try
            {
                thisWb.Close(false, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                while (System.Runtime.InteropServices.Marshal.ReleaseComObject(thisWb) != 0)
                { }
                thisApp.Quit();
                while (System.Runtime.InteropServices.Marshal.ReleaseComObject(thisApp) != 0)
                { }
                if (thisWs != null) { Marshal.ReleaseComObject(thisWs); } //release each worksheet like this     
                if (thisSheet != null) { Marshal.ReleaseComObject(thisSheet); } //release each worksheet like this     
                if (thisWb != null) { Marshal.ReleaseComObject(thisWb); } //release each workbook like this                
                if (thisApp != null) { Marshal.ReleaseComObject(thisApp); } //release the Excel application
                thisApp = null;
                thisSheet = null;
                thisWs = null;
                thisWb = null; //set each memory reference to null.
            }
            catch (Exception)
            {
                if (thisWs != null) { Marshal.ReleaseComObject(thisWs); } //release each worksheet like this     
                if (thisSheet != null) { Marshal.ReleaseComObject(thisSheet); } //release each worksheet like this     
                if (thisWb != null) { Marshal.ReleaseComObject(thisWb); } //release each workbook like this                
                if (thisApp != null) { Marshal.ReleaseComObject(thisApp); } //release the Excel application
                thisApp = null;
                thisSheet = null;
                thisWs = null;
                thisWb = null; //set each memory reference to null.
            }
            finally
            {
                GC.Collect();
                //GC.WaitForPendingFinalizers();
            }

        }

        public void openShowWorkbook(string fileName)
        {
            excelInterop.Application xlApp = new excelInterop.Application();
            excelInterop.Workbook wb = xlApp.Workbooks.Open(fileName);
            xlApp.Visible = true;

            while(xlApp.Visible)
            {
                string foo = "bar";
            }
            ReleaseExcel(wb, null, xlApp);
        }

        public void openShowWorkbook()
        {
            excelInterop.Application xlApp = new excelInterop.Application();
            excelInterop.Workbook wb = xlApp.Workbooks.Add();
            xlApp.Visible = true;
            ReleaseExcel(wb, null, xlApp);
        }
        public void createNewExcelTemplate(string typeTemplate)
        {
            excelInterop.Application xlApp = new excelInterop.Application();
            excelInterop.Workbook wb = xlApp.Workbooks.Add();
            excelInterop.Worksheet ws = wb.Worksheets.Add();
            try
            {
                saveExcelFile(wb, outputPath + @"\Templates\" + typeTemplate + ".xlsx");                
            }
            finally
            {
                ReleaseExcel(wb, ws, xlApp);
            }
        }

        public void createNewReport()
        {
            excelInterop.Application xlApp = new excelInterop.Application();
            excelInterop.Workbook wb = xlApp.Workbooks.Add();
            excelInterop.Worksheet ws = (excelInterop.Worksheet)wb.Worksheets[1];
            Globals.getReportPath = outputPath + @"\Reports\Report" + DateTime.Now.ToString("ddMMyyyyhhmmss") + ".xlsx";
            try
            {
                ws.Cells[1, 1].value2 = "WorkItemID";
                ws.Cells[1, 2].value2 = "FieldName";
                ws.Cells[1, 3].value2 = "OriginalValue";
                ws.Cells[1, 4].value2 = "NewValue";
                ws.Cells[1, 5].value2 = "Status";
                saveExcelFile(wb, Globals.getReportPath);
            }
            finally
            {
                ReleaseExcel(wb, null, xlApp);
            }
        }

        public void updateReport(string getFile, Dictionary<string, string[]>getDic)
        {
            excelInterop.Application xlApp = new excelInterop.Application();
            excelInterop.Workbook wb = xlApp.Workbooks.Open(getFile);
            excelInterop.Worksheet ws = (excelInterop.Worksheet)wb.Worksheets[1];
            Globals.getReportPath = outputPath + @"\Reports\Report" + DateTime.Now.ToString("ddMMyyyyhhmmss") + ".xlsx";
            int currentRowCount = 1;
            string[] getKeys = getDic.Keys.ToArray();
            try
            {
                for (int x = 0; x<=getDic.Keys.Count-1; x++)
                {
                    for(int y = 0; y<= getDic[getKeys[x]].Count()-1; y++)
                    {
                        ws.Cells[currentRowCount + 1, y+1].value2 = Convert.ToString(getDic[getKeys[x]][y]);
                    }
                    currentRowCount++;
                }
                saveExcelFile(wb, Globals.getReportPath);
            }
            finally
            {
                ReleaseExcel(wb, null, xlApp);
            }
        }

        private static void saveExcelFile(excelInterop.Workbook thisWb, string fullFilePath)
        {
            DateTime thisDate = DateTime.Now;
            if (File.Exists(fullFilePath))
            {
                File.Move(fullFilePath, fullFilePath.Replace(@"\Templates\", @"\Templates\Archive\")
                    .Replace(".xlsx", thisDate.ToString("ddMMyyyyhhmmss") + ".xlsx"));
            }
            thisWb.SaveAs(fullFilePath);
        }

        public void generateDefectTemplate(string outputFile=null)
        {
            defectsTFS deftfs = new defectsTFS();
            string defTemp = "Defect";
            string getExcelLetter = "";
            string valueListRange = "";
            string getCurrentLetter = "";
            string getTotalRows = "9999"; //TODO: This should come from a config file
            int getTotalItems = 0;
            deftfs.getTFSDefectFields();
            excelInterop.Application xlApp = new excelInterop.Application();
            excelInterop.Workbook wb = xlApp.Workbooks.Open(outputPath + @"\Templates\" + defTemp + ".xlsx");
            excelInterop.Worksheet ws = (excelInterop.Worksheet)wb.Worksheets[1];
            excelInterop.Worksheet ws2 = (excelInterop.Worksheet)wb.Worksheets[2];
            ws.Name = "Defects";
            ws2.Name = "LOV";
            Dictionary<string, double> dicColumnWidth = new Dictionary<string, double>();
            //TODO: this should be changed when gui comes in and select the fields included

            //This is for LOV sheet========================================
            try
            {
                int m = 1;
                foreach (KeyValuePair<string, string[]> entry in deftfs.fieldAllowedValues)
                {
                    int n = 2;
                    int valLen = 0;
                    int topValLen = entry.Key.ToString().Length;
                    double getWidth = 0;
                    ws2.Cells[1, m].Value2 = entry.Key;
                    ws2.Cells[1, m].Columns.AutoFit();
                    foreach (string vals in entry.Value)
                    {
                        ws2.Cells[n, m].Value2 = vals;
                        valLen = vals.Length;
                        if(valLen > topValLen)
                        {
                            topValLen = valLen;
                            ws2.Cells[n, m].Columns.AutoFit();
                        }                                              
                        n++;
                    }
                    getWidth = ws2.Columns[m].ColumnWidth;
                    dicColumnWidth[entry.Key] = getWidth;
                    m++;                    
                }
                //excelInterop.Range cell = ws2.Cells[1,1].
                //This is for LOV sheet========================================

                //This is for Defects Sheet====================================
                ws.Cells[1, 1] = "SNo";
                ws.Cells[1, 2] = "ID";
                for (int i = 0; i <= deftfs.xmlDefectFields.Count - 1; i++)
                {
                    ws.Cells[1, i + 3].Value2 = deftfs.xmlDefectFields[i];
                    ws.Cells[1, i + 3].Columns.AutoFit();
                    if (deftfs.xmlDefectFields[i].Contains("Date"))
                    {
                        ws.Cells[2, i + 3].NumberFormat = "M/d/yyyy hh:mm";
                        ws.Cells[2, i + 3].Select();
                        ws.Cells[2, i + 3].Copy();
                        ws.Range[getCurrentLetter + "3", getCurrentLetter + getTotalRows].
                            PasteSpecial(excelInterop.XlPasteType.xlPasteValuesAndNumberFormats,
                            excelInterop.XlPasteSpecialOperation.xlPasteSpecialOperationNone, Type.Missing, Type.Missing);
                    }
                    try
                    {
                        if (deftfs.fieldAllowedValues[deftfs.xmlDefectFields[i]] != null)
                        {
                            getExcelLetter = GetMatchColumnHeader(deftfs.xmlDefectFields[i], ws2);
                            getCurrentLetter = GetExcelColumnName(i + 3);
                            getTotalItems = deftfs.fieldAllowedValues[deftfs.xmlDefectFields[i]].Count() + 1;
                            valueListRange = "=" + ws2.Name.ToString() + "!$" + getExcelLetter + "$2:$" +
                                getExcelLetter + "$" + getTotalItems.ToString();
                            ws.Cells[2, i + 3].Validation.Delete();
                            ws.Cells[2, i + 3].Validation.Add(excelInterop.XlDVType.xlValidateList,
                                excelInterop.XlDVAlertStyle.xlValidAlertStop, excelInterop.XlFormatConditionOperator.xlBetween, valueListRange, Type.Missing);
                            ws.Cells[2, i + 3].Validation.IgnoreBlank = true;
                            ws.Cells[2, i + 3].Select();
                            ws.Cells[2, i + 3].Copy();
                            ws.Range[getCurrentLetter + "3", getCurrentLetter + getTotalRows].
                                PasteSpecial(excelInterop.XlPasteType.xlPasteValidation,
                                excelInterop.XlPasteSpecialOperation.xlPasteSpecialOperationNone, Type.Missing, Type.Missing);
                            ws.Columns[i + 3].ColumnWidth = dicColumnWidth[deftfs.xmlDefectFields[i]];
                        }
                    }
                    catch
                    {

                    }
                }
                //This is for Defects Sheet====================================
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                ReleaseExcel(wb, ws, xlApp);
                ReleaseExcel(null, ws2, null);
            }            
            wb.Save();
            ReleaseExcel(wb, ws, xlApp);
            ReleaseExcel(null, ws2, null);
            if (!String.IsNullOrEmpty(outputFile))
            {
                updateAndCopy(outputFile + @"\DefectTemplate.xlsx", "Defect", true);
            }            
        }

        public void generateTestCaseTemplate(string outputFile = null)
        {
            testmanTFS tmtfs = new testmanTFS();
            string defTemp = "TestCase";
            string getExcelLetter = "";
            string valueListRange = "";
            string getCurrentLetter = "";
            string getTotalRows = "9999"; //TODO: This should come from a config file
            int getTotalItems = 0;
            tmtfs.getTFSTestCaseFields();
            excelInterop.Application xlApp = new excelInterop.Application();
            excelInterop.Workbook wb = xlApp.Workbooks.Open(outputPath + @"\Templates\" + defTemp + ".xlsx");
            excelInterop.Worksheet ws = (excelInterop.Worksheet)wb.Worksheets[1];
            excelInterop.Worksheet ws2 = (excelInterop.Worksheet)wb.Worksheets[2];
            ws.Name = "TestCase";
            ws2.Name = "LOV";
            Dictionary<string, double> dicColumnWidth = new Dictionary<string, double>();
            //TODO: this should be changed when gui comes in and select the fields included

            //This is for LOV sheet========================================
            try
            {
                int m = 1;
                foreach (KeyValuePair<string, string[]> entry in tmtfs.tcFieldAllowedValues)
                {
                    int n = 2;
                    int valLen = 0;
                    int topValLen = entry.Key.ToString().Length;
                    double getWidth = 0;
                    ws2.Cells[1, m].Value2 = entry.Key;
                    ws2.Cells[1, m].Columns.AutoFit();
                    foreach (string vals in entry.Value)
                    {
                        ws2.Cells[n, m].Value2 = vals;
                        valLen = vals.Length;
                        if (valLen > topValLen)
                        {
                            topValLen = valLen;
                            ws2.Cells[n, m].Columns.AutoFit();
                        }
                        n++;
                    }
                    getWidth = ws2.Columns[m].ColumnWidth;
                    dicColumnWidth[entry.Key] = getWidth;
                    m++;
                }
                //excelInterop.Range cell = ws2.Cells[1,1].
                //This is for LOV sheet========================================

                //This is for Defects Sheet====================================
                ws.Cells[1, 1] = "SNo";
                ws.Cells[1, 2] = "ID";
                for (int i = 0; i <= tmtfs.xmlTestCaseFields.Count - 1; i++)
                {
                    ws.Cells[1, i + 3].Value2 = tmtfs.xmlTestCaseFields[i];
                    ws.Cells[1, i + 3].Columns.AutoFit();
                    ws.Cells[2, i + 3].WrapText=true;
                    ws.Cells[2, i + 3].VerticalAlignment = excelInterop.XlVAlign.xlVAlignTop;
                    if (tmtfs.xmlTestCaseFields[i].Contains("Date"))
                    {
                        ws.Cells[2, i + 3].NumberFormat = "M/d/yyyy hh:mm";
                        ws.Cells[2, i + 3].Select();
                        ws.Cells[2, i + 3].Copy();                        
                        ws.Range[getCurrentLetter + "3", getCurrentLetter + getTotalRows].
                            PasteSpecial(excelInterop.XlPasteType.xlPasteValuesAndNumberFormats,
                            excelInterop.XlPasteSpecialOperation.xlPasteSpecialOperationNone, Type.Missing, Type.Missing);
                        ws.Range[getCurrentLetter + "3", getCurrentLetter + getTotalRows].
                            PasteSpecial(excelInterop.XlPasteType.xlPasteFormats,
                            excelInterop.XlPasteSpecialOperation.xlPasteSpecialOperationNone, Type.Missing, Type.Missing);
                    }
                    try
                    {
                        if (tmtfs.tcFieldAllowedValues[tmtfs.xmlTestCaseFields[i]] != null)
                        {
                            getExcelLetter = GetMatchColumnHeader(tmtfs.xmlTestCaseFields[i], ws2);
                            getCurrentLetter = GetExcelColumnName(i + 3);
                            getTotalItems = tmtfs.tcFieldAllowedValues[tmtfs.xmlTestCaseFields[i]].Count() + 1;
                            valueListRange = "=" + ws2.Name.ToString() + "!$" + getExcelLetter + "$2:$" +
                                getExcelLetter + "$" + getTotalItems.ToString();
                            ws.Cells[2, i + 3].Validation.Delete();
                            ws.Cells[2, i + 3].Validation.Add(excelInterop.XlDVType.xlValidateList,
                                excelInterop.XlDVAlertStyle.xlValidAlertStop, excelInterop.XlFormatConditionOperator.xlBetween, valueListRange, Type.Missing);
                            ws.Cells[2, i + 3].Validation.IgnoreBlank = true;
                            ws.Cells[2, i + 3].Select();
                            ws.Cells[2, i + 3].Copy();
                            ws.Range[getCurrentLetter + "3", getCurrentLetter + getTotalRows].
                                PasteSpecial(excelInterop.XlPasteType.xlPasteValidation,
                                excelInterop.XlPasteSpecialOperation.xlPasteSpecialOperationNone, Type.Missing, Type.Missing);
                            ws.Columns[i + 3].ColumnWidth = dicColumnWidth[tmtfs.xmlTestCaseFields[i]];
                        }
                    }
                    catch
                    {

                    }
                }
                //This is for Defects Sheet====================================
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                ReleaseExcel(wb, ws, xlApp);
                ReleaseExcel(null, ws2, null);
            }
            wb.Save();
            ReleaseExcel(wb, ws, xlApp);
            ReleaseExcel(null, ws2, null);
            if (!String.IsNullOrEmpty(outputFile))
            {
                updateAndCopy(outputFile + @"\TestCaseTemplate.xlsx", "TestCase", true);
            }
        }

        private static string GetMatchColumnHeader(string columnName, excelInterop.Worksheet ws)
        {
            string retVal = "";
            for(int x=1; x<=3000; x++)
            {
                if (ws.Cells[1, x].value2 == columnName)
                {
                    retVal = GetExcelColumnName(x);
                    break;
                }
            }
            return retVal;
        }

        private static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        public void archiveAndCopy(string filePath, string wiType)
        {
            DateTime getTime = DateTime.Now;
            string fileName = filePath.Substring(filePath.LastIndexOf("\\"));
            string compDestPath = destPath + wiType + fileName;

            string archiveFile = destPath + wiType + @"\Archive" + fileName.Replace(".xlsx", getTime.ToString("ddMMyyyyhhmmss") + ".xlsx");
            if (File.Exists(compDestPath))
            {
                File.Copy(filePath, archiveFile);
                File.Copy(filePath, compDestPath, true);
            }
            else
            {
                File.Copy(filePath, compDestPath);
            }
        }

        public void updateAndCopy(string filePath, string wiType)
        {            
            string fileName = filePath.Substring(filePath.LastIndexOf("\\"));
            string compDestPath = destPath + wiType + fileName;            
            File.Copy(compDestPath, filePath, true);
        }

        public void updateAndCopy(string filePath, string wiType, bool fromTemp, string fileName=null)
        {
            try
            {
                if (fromTemp)
                {
                    if (fileName == null)
                    {
                        fileName = filePath.Substring(filePath.LastIndexOf("\\"));
                    }
                    string compDestPath = templateSource + @"\" + wiType + ".xlsx";
                    File.Copy(compDestPath, filePath, true);
                }
                else
                {
                    if (fileName == null)
                    {
                        fileName = filePath.Substring(filePath.LastIndexOf("\\"));
                    }
                    else
                    {
                        fileName = @"\" + fileName;
                    }
                    string compDestPath = destPath + wiType + fileName;
                    File.Copy(compDestPath, filePath, true);
                }
            }
            catch(FileNotFoundException)
            {
                Console.WriteLine("Source not Found");
            }

        }
        public void getDefectTFSIDs(string filePath, string wiType)
        {
            string fileName = filePath.Substring(filePath.LastIndexOf("\\"));
            string compDestPath = destPath + wiType + fileName;
            List<string> getValList = new List<string>();
            excelInterop.Application xlApp = new excelInterop.Application();
            excelInterop.Workbook wb = xlApp.Workbooks.Open(compDestPath);
            excelInterop.Worksheet ws = (excelInterop.Worksheet)wb.Worksheets[1];
            try
            {
                getAllID.Clear();                
                int totalRows = ws.UsedRange.Rows.Count;
                for(int x = 2; x<=totalRows; x++)
                {
                    getAllID.Add(Convert.ToInt32(ws.Cells[x, 2].value2));
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            finally
            {
                ReleaseExcel(wb, ws, xlApp);
            }
        }
        public void readExcelData(string filePath, string wiType)
        {
            string fileName = filePath.Substring(filePath.LastIndexOf("\\"));
            string compDestPath = destPath + wiType + fileName;            
            string getKey = "";
            string getValue = "";
            List<string> getValList = new List<string>();            
            excelInterop.Application xlApp = new excelInterop.Application();
            excelInterop.Workbook wb = xlApp.Workbooks.Open(compDestPath);
            excelInterop.Worksheet ws = (excelInterop.Worksheet)wb.Worksheets[1];
            try
            {
                dicExData.Clear();
                int totalColumns = ws.UsedRange.Columns.Count;
                int totalRows = 0;
                for(int u=1; u<=9999; u++)
                {
                    if(String.IsNullOrEmpty(Convert.ToString(ws.Cells[u, 1].value2)))
                    {
                        totalRows = u - 1;
                        break;
                    }
                }
                for(int x = 1; x<= totalColumns; x++)
                {
                    getValList.Clear();
                    for(int y = 2; y<=totalRows; y++)
                    {
                        getKey = ws.Cells[1, x].value2;
                        getValue = Convert.ToString(ws.Cells[y, x].value2);                        
                        getValList.Add(getValue);                        
                    }
                    dicExData[getKey] = getValList.ToArray();
                }
            }
            catch(Exception e)
            {
                Console.WriteLine(e.ToString());                
            }
            finally
            {
                ReleaseExcel(wb, ws, xlApp);
            }
        }

        public void readExcelDataForTC(string filePath, string wiType)
        {
            string fileName = filePath.Substring(filePath.LastIndexOf("\\"));
            string compDestPath = destPath + wiType + fileName;
            string getKey = "";
            string getValue = "";
            List<string> getValList = new List<string>();
            excelInterop.Application xlApp = new excelInterop.Application();
            excelInterop.Workbook wb = xlApp.Workbooks.Open(compDestPath);
            excelInterop.Worksheet ws = (excelInterop.Worksheet)wb.Worksheets[1];
            try
            {
                dicExData.Clear();
                int totalColumns = ws.UsedRange.Columns.Count;
                int totalRows = 0;
                int getCols = 0;
                int getRows = 0;
                for(int cols=totalColumns; cols<=1; cols--)
                {
                    if(ws.Cells[1, cols].value2=="Step No")
                    {
                        getCols = cols;
                        break;
                    }
                }
                for(int rows=1; rows<=9999; rows++)
                {
                    if(String.IsNullOrEmpty(Convert.ToString(ws.Cells[rows, getCols])))
                    {
                        getRows = rows;
                    }
                }
                for(int rows = 1; rows <= getRows; rows++)
                {
                    for (int cols = 1; cols <= totalColumns; cols++)
                    {
                        if (String.IsNullOrEmpty(Convert.ToString(ws.Cells[rows, cols].value2)) && cols==getCols)
                        {
                            break;
                        }
                        else if(String.IsNullOrEmpty(Convert.ToString(ws.Cells[rows, cols].value2)) && cols != getCols)
                        {
                            if (Convert.ToString(ws.Cells[1, cols].value2)=="ID")
                            {
                                //Put into dictionary
                            }
                        }
                        else if (!String.IsNullOrEmpty(Convert.ToString(ws.Cells[rows, cols].value2)) && cols != getCols)
                        {
                        
                        }                                   
                    }
                }

            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            finally
            {
                ReleaseExcel(wb, ws, xlApp);
            }
        }

        public void updateExcelData(string filePath, string wiType, Dictionary<string, string[]> dicUpdData, int valCtr = 0)
        {
            string fileName = filePath.Substring(filePath.LastIndexOf("\\"));
            string compDestPath = destPath + wiType + fileName;
            int getTotalData = dicUpdData["ID"].Count()-1;
            List<string> getValList = new List<string>();
            excelInterop.Application xlApp = new excelInterop.Application();
            excelInterop.Workbook wb = xlApp.Workbooks.Open(compDestPath);
            excelInterop.Worksheet ws = (excelInterop.Worksheet)wb.Worksheets[1];

            try
            {
                for (int i = valCtr; i <= getTotalData; i++)
                {
                    int q = 1;
                    foreach (KeyValuePair<string, string[]> updData in dicUpdData)
                    {
                        for(int d = q; d<=dicUpdData.Count()+4; d++)
                        {
                            if(ws.Cells[1, d].value2 == updData.Key)
                            {                                
                                ws.Cells[i + 2, d].value2 = HtmlToText.ConvertHtml(checkNull(Convert.ToString(updData.Value[i])));
                                //ws.Cells[i + 2, d].value2 = updData.Value[i];
                                q = d+1;
                                if (q > dicUpdData.Count())
                                {
                                    q = 1;
                                }
                                break;
                            }
                            else
                            {
                                q = 1;
                            }
                        }                    
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                ReleaseExcel(wb, ws, xlApp);
            }
            finally
            {
                wb.Save();
                ReleaseExcel(wb, ws, xlApp);
            }
        }

        public void updateTestCaseExcelData(string filePath, string wiType, Dictionary<string, Dictionary<string,string>> dicUpdData, int valCtr = 0)
        {
            string fileName = filePath.Substring(filePath.LastIndexOf("\\"));
            string compDestPath = destPath + wiType + fileName;
            int getTotalData = dicUpdData.Count() - 1;
            List<string> getValList = new List<string>();
            excelInterop.Application xlApp = new excelInterop.Application();
            excelInterop.Workbook wb = xlApp.Workbooks.Open(compDestPath);
            excelInterop.Worksheet ws = (excelInterop.Worksheet)wb.Worksheets[1];

            try
            {
                int getInitialRow = 2;
                foreach (KeyValuePair<string, Dictionary<string,string>> entry in dicUpdData)
                {                    
                    int ctrEntry = 1;
                    int ctrColumn = 1;
                    ws.Cells[getInitialRow, ctrColumn].value2 = ctrEntry;
                    ctrColumn++;
                    ws.Cells[getInitialRow, ctrColumn].value2 = entry.Key;
                    ctrColumn++;
                    ws.Cells[getInitialRow, ctrColumn].value2 = Globals.getTestPlan;
                    ctrColumn++;
                    foreach (KeyValuePair<string,string> getInfo in entry.Value)
                    {
                        if(getInfo.Key!= "getSteps")
                        {
                            ws.Cells[getInitialRow, ctrColumn].value2 = HtmlToText.ConvertHtml(checkNull(Convert.ToString(getInfo.Value)));
                            ctrColumn++;
                        }
                        else
                        {
                            string[] splitSteps = getInfo.Value.Split(new string[] { "<...>" }, StringSplitOptions.None);
                            foreach(string allSteps in splitSteps)
                            {
                                int stepColumn = ctrColumn;
                                string[] splitCtr = allSteps.Split(new string[] { "<+++>" }, StringSplitOptions.None);
                                string[] splitstepExp = splitCtr[1].Split(new string[] { "<--->" }, StringSplitOptions.None);
                                string stepNo = splitCtr[0];                                                                
                                string stepTitle = HtmlToText.ConvertHtml(checkNull(Convert.ToString(splitstepExp[0]))); 
                                string stepExp = HtmlToText.ConvertHtml(checkNull(Convert.ToString(splitstepExp[1])));
                                ws.Cells[getInitialRow, stepColumn].value2 = stepNo;
                                stepColumn++;
                                ws.Cells[getInitialRow, stepColumn].value2 = stepTitle;
                                stepColumn++;
                                ws.Cells[getInitialRow, stepColumn].value2 = stepExp;
                                stepColumn++;
                                getInitialRow++;
                            }                            
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                ReleaseExcel(wb, ws, xlApp);
            }
            finally
            {
                wb.Save();
                ReleaseExcel(wb, ws, xlApp);
            }
        }

        public void updateTestCaseExcelDataSpecFormat(string filePath, string wiType, Dictionary<string, Dictionary<string, string>> dicUpdData, int valCtr = 0)
        {
            string fileName = filePath.Substring(filePath.LastIndexOf("\\"));
            string compDestPath = destPath + wiType + fileName;
            int getTotalData = dicUpdData.Count() - 1;
            List<string> getValList = new List<string>();
            excelInterop.Application xlApp = new excelInterop.Application();
            excelInterop.Workbook wb = xlApp.Workbooks.Open(compDestPath);
            excelInterop.Worksheet ws = (excelInterop.Worksheet)wb.Worksheets[1];

            try
            {
                foreach (KeyValuePair<string, Dictionary<string, string>> entry in dicUpdData)
                {
                    createSpecialFormatting(xlApp, wb, ws, stepStart, entry.Value);
                    CreateStepsSpecial(xlApp, wb, ws, entry.Value, stepStart + 5);                                           
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                ReleaseExcel(wb, ws, xlApp);
            }
            finally
            {
                wb.Save();
                ReleaseExcel(wb, ws, xlApp);
            }
        }
    
        private void createSpecialFormatting(excelInterop.Application xlApp, excelInterop.Workbook wb, 
            excelInterop.Worksheet ws, int sRow, Dictionary<string, string> getDic)
        {
            //System Name
            xlApp.DisplayAlerts = false;
            excelInterop.Range lblsysNameRange = ws.get_Range("A" + sRow, "B" + sRow);
            excelInterop.Range valsysNameRange = ws.get_Range("C" + sRow, "G" + sRow);
            lblsysNameRange.Merge();
            lblsysNameRange.Font.Name = "Times New Roman";
            lblsysNameRange.Font.Size = "12";
            valsysNameRange.Merge();
            lblsysNameRange.Value = "System Name:";
            valsysNameRange.Value = "IPOS-DASH";
            valsysNameRange.Font.Name = "Arial";
            valsysNameRange.Font.Size = "10";
            lblsysNameRange.Font.Bold = true;

            //Module Name
            excelInterop.Range lblFuncModuleRange = ws.get_Range("A" + (sRow + 1), "B" + (sRow + 1));
            excelInterop.Range valFuncModuleRange = ws.get_Range("C" + (sRow + 1), "E" + (sRow + 1));
            lblFuncModuleRange.Merge();
            lblFuncModuleRange.Font.Name = "Times New Roman";
            lblFuncModuleRange.Font.Size = "12";
            valFuncModuleRange.Merge();
            lblFuncModuleRange.Value = "Function / Module Name:";
            valFuncModuleRange.Value = getDic["Area Path"];
            //valFuncModuleRange.Value = "IPOS-DASH End To End Testing";
            valFuncModuleRange.Font.Name = "Arial";
            valFuncModuleRange.Font.Size = "10";
            lblFuncModuleRange.Font.Bold = true;

            //Test No
            excelInterop.Range lbltestNoRange = ws.get_Range("A" + (sRow + 2), "B" + (sRow + 2));
            excelInterop.Range valtestNoRange = ws.get_Range("C" + (sRow + 2), "E" + (sRow + 2));
            lbltestNoRange.Merge();
            valtestNoRange.Merge();
            lbltestNoRange.Value = "Test No:";
            lbltestNoRange.Font.Name = "Times New Roman";
            lbltestNoRange.Font.Size = "12";
            valtestNoRange.Value = getDic["Test Case ID"];
            valtestNoRange.Font.Name = "Arial";
            valtestNoRange.Font.Size = "10";
            lbltestNoRange.Font.Bold = true;

            //Test Case (Label)
            excelInterop.Range lblTestCase = ws.get_Range("B" + (sRow + 3), "C" + (sRow + 3));
            lblTestCase.Merge();
            lblTestCase.Value = "Test Case";
            lblTestCase.Font.Name = "Arial";
            lblTestCase.Font.Size = "10";

            //Description Merge
            string description = getDic["Description"];
            excelInterop.Range lblDescription = ws.get_Range("A" + (sRow + 4), "G" + (sRow + 4));
            lblDescription.Merge();
            if (description.Length < 1)
            {
                description = getDic["Title"];
            }
            lblDescription.Value = "Scenario: " + description;
            lblDescription.Font.Name = "Arial";
            lblDescription.Font.Size = "10";
            lblDescription.EntireRow.RowHeight = ((description.Length / 2) * 1.20).ToString("##.##");
            lblDescription.VerticalAlignment = excelInterop.XlVAlign.xlVAlignTop;
            lblDescription.WrapText = true;

            //Prepared By
            excelInterop.Range lblPreparedBy = ws.get_Range("F" + (sRow + 1));
            excelInterop.Range valPreparedBy = ws.get_Range("G" + (sRow + 1));
            lblPreparedBy.Value = "Prepared By / Date:";
            lblPreparedBy.Font.Name = "Times New Roman";
            lblPreparedBy.Font.Size = "12";
            lblPreparedBy.Font.Bold = true;
            valPreparedBy.Value = "CrimsonLogic / "; //+ getValues[8];
            valPreparedBy.Font.Name = "Arial";
            valPreparedBy.Font.Size = "10";
            valPreparedBy.WrapText = true;

            //Tested By
            excelInterop.Range lblTestedBy = ws.get_Range("F" + (sRow + 2));
            excelInterop.Range valTestedBy = ws.get_Range("G" + (sRow + 2));
            lblTestedBy.Value = "Tested By / Date:";
            lblTestedBy.Font.Name = "Times New Roman";
            lblTestedBy.Font.Size = "12";
            lblTestedBy.Font.Bold = true;
            valTestedBy.Value = "CrimsonLogic / ";// + getValues[6];
            valTestedBy.Font.Name = "Arial";
            valTestedBy.Font.Size = "10";
            valTestedBy.WrapText = true;

            //Headers
            excelInterop.Range lblSno = ws.get_Range("A" + (sRow + 3));
            excelInterop.Range lblInputData = ws.get_Range("D" + (sRow + 3));
            excelInterop.Range lblExpRes = ws.get_Range("E" + (sRow + 3));
            excelInterop.Range lblActRes = ws.get_Range("F" + (sRow + 3));
            excelInterop.Range lblRemarks = ws.get_Range("G" + (sRow + 3));
            lblSno.Value = "S/No";
            lblInputData.Value = "Input Data";
            lblExpRes.Value = "Expected Results";
            lblActRes.Value = "Actual Results";
            lblRemarks.Value = "Remarks";
            Color headerRow = Color.FromArgb(217, 225, 242);
            ws.get_Range("A" + (sRow + 3), "G" + (sRow + 3)).Font.Name = "Arial";
            ws.get_Range("A" + (sRow + 3), "G" + (sRow + 3)).Font.Size = "10";
            ws.get_Range("A" + (sRow + 3), "G" + (sRow + 3)).Interior.Color = headerRow;
            ws.get_Range("A" + (sRow + 3), "G" + (sRow + 3)).HorizontalAlignment = excelInterop.XlHAlign.xlHAlignCenter;
            ws.get_Range("A" + (sRow + 3), "G" + (sRow + 3)).VerticalAlignment = excelInterop.XlVAlign.xlVAlignTop;
            ws.get_Range("A" + (sRow + 3), "G" + (sRow + 3)).Font.Bold = true;
            ws.get_Range("A1").ColumnWidth = "7.43";
            ws.get_Range("B1").ColumnWidth = "18.71";
            ws.get_Range("C1").ColumnWidth = "17.86";
            ws.get_Range("D1").ColumnWidth = "12.57";
            ws.get_Range("E1").ColumnWidth = "26.14";
            ws.get_Range("F1").ColumnWidth = "21.00";
            ws.get_Range("G1").ColumnWidth = "22.14";
        }

        private void CreateStepsSpecial(excelInterop.Application xlApp, excelInterop.Workbook wb,
            excelInterop.Worksheet ws, Dictionary<string, string> getDic, int rowNum)
        {
            int i = 0; 
            double getLen = 0;
            double rowRatio = 2.2;
            double multiPlier = 0.00;
            double getRowHeight = 0;
            string[] splitSteps = getDic["getSteps"].Split(new string[] { "<...>" }, StringSplitOptions.None);

            for (int x = rowNum; x < splitSteps.Length + rowNum; x++)
            {
                string[] splitCtr = splitSteps[i].Split(new string[] { "<+++>" }, StringSplitOptions.None);
                string[] splitstepExp = splitCtr[1].Split(new string[] { "<--->" }, StringSplitOptions.None);
                string stepNo = splitCtr[0];
                string stepTitle = HtmlToText.ConvertHtml(checkNull(Convert.ToString(splitstepExp[0])));
                string stepExp = HtmlToText.ConvertHtml(checkNull(Convert.ToString(splitstepExp[1])));

                if (stepExp.Length > stepTitle.Length)
                {
                    getLen = stepExp.Length;
                }
                else
                {
                    getLen = stepTitle.Length;
                }
                if (getLen < 100)
                {
                    multiPlier = 1.5;
                }
                else
                {
                    multiPlier = 1.2;
                }
                ws.get_Range("A" + x).Value = stepNo;
                ws.get_Range("B" + x, "C" + x).Merge();
                ws.get_Range("B" + x).Value = stepTitle;
                ws.get_Range("E" + x).Value = stepExp;
                //if (actRes[i] == "Passed")
                //{
                //    ws.get_Range("F" + x).Value = "Actual result is as per expected";
                //}
                //else
                //{
                //    ws.get_Range("F" + x).Value = actRes[i];
                //}
                //ws.get_Range("G" + x).Value = actRes[i];
                ws.get_Range("A" + x, "G" + x).Font.Name = "Arial";
                ws.get_Range("A" + x, "G" + x).Font.Size = "10";
                getRowHeight = ((getLen / rowRatio) * multiPlier);
                if (getRowHeight > 409)
                {
                    getRowHeight = 409;
                }
                else if (getRowHeight < 15)
                {
                    getRowHeight = 15;
                }
                ws.get_Range("A" + x, "G" + x).EntireRow.RowHeight = getRowHeight.ToString("##.##");
                ws.get_Range("A" + x, "G" + x).VerticalAlignment = excelInterop.XlVAlign.xlVAlignTop;
                ws.get_Range("A" + x, "G" + x).WrapText = true;
                i++;
                stepStart++;
            }
            stepStart = stepStart + 5;
        }
        private string checkNull(string input)
        {
            if (String.IsNullOrEmpty(input))
            {
                return "";
            }
            else
            {
                return input;
            }
            
        }
        public bool validateFileFormat(List<string> fieldsToUpload, string fileToUpload)
        {
            bool retval = false;
            List<string> getCols = new List<string>();
            try
            {
                excelInterop.Application xlApp = new excelInterop.Application();
                excelInterop.Workbook wb = xlApp.Workbooks.Open(fileToUpload);
                excelInterop.Sheets sheets = wb.Worksheets;
                excelInterop.Worksheet ws = sheets[1];
                for(int i=1; i<=9999; i++)
                {
                    if (!String.IsNullOrEmpty(Convert.ToString(ws.Cells[1, i].value2)))
                    {
                        getCols.Add(ws.Cells[1, i].value2);
                    }
                    else
                    {
                        //xlApp.Visible = true;
                        ReleaseExcel(wb, ws, xlApp, sheets);
                        break;
                    }
                }
                foreach(string item in getCols)
                {
                    if(fieldsToUpload.Contains(item))
                    {
                        retval = true;
                    }
                    else
                    {
                        retval = false;
                        break;
                    }
                }
            }
            catch (ArgumentOutOfRangeException)
            {
                
            }
            finally
            {
                
            }
            return retval;
        }
    }
    public static class HtmlToText
    {

        public static string Convert(string path)
        {
            HtmlDocument doc = new HtmlDocument();
            doc.Load(path);
            return ConvertDoc(doc);
        }

        public static string ConvertHtml(string html)
        {
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(html);
            return ConvertDoc(doc);
        }

        public static string ConvertDoc(HtmlDocument doc)
        {
            using (StringWriter sw = new StringWriter())
            {
                ConvertTo(doc.DocumentNode, sw);
                sw.Flush();
                return sw.ToString();
            }
        }

        internal static void ConvertContentTo(HtmlNode node, TextWriter outText, PreceedingDomTextInfo textInfo)
        {
            foreach (HtmlNode subnode in node.ChildNodes)
            {
                ConvertTo(subnode, outText, textInfo);
            }
        }
        public static void ConvertTo(HtmlNode node, TextWriter outText)
        {
            ConvertTo(node, outText, new PreceedingDomTextInfo(false));
        }
        internal static void ConvertTo(HtmlNode node, TextWriter outText, PreceedingDomTextInfo textInfo)
        {
            string html;
            switch (node.NodeType)
            {
                case HtmlNodeType.Comment:
                    // don't output comments
                    break;
                case HtmlNodeType.Document:
                    ConvertContentTo(node, outText, textInfo);
                    break;
                case HtmlNodeType.Text:
                    // script and style must not be output
                    string parentName = node.ParentNode.Name;
                    if ((parentName == "script") || (parentName == "style"))
                    {
                        break;
                    }
                    // get text
                    html = ((HtmlTextNode)node).Text;
                    // is it in fact a special closing node output as text?
                    if (HtmlNode.IsOverlappedClosingElement(html))
                    {
                        break;
                    }
                    // check the text is meaningful and not a bunch of whitespaces
                    if (html.Length == 0)
                    {
                        break;
                    }
                    if (!textInfo.WritePrecedingWhiteSpace || textInfo.LastCharWasSpace)
                    {
                        html = html.TrimStart();
                        if (html.Length == 0) { break; }
                        textInfo.IsFirstTextOfDocWritten.Value = textInfo.WritePrecedingWhiteSpace = true;
                    }
                    outText.Write(HtmlEntity.DeEntitize(Regex.Replace(html.TrimEnd(), @"\s{2,}", " ")));
                    if (textInfo.LastCharWasSpace = char.IsWhiteSpace(html[html.Length - 1]))
                    {
                        outText.Write(' ');
                    }
                    break;
                case HtmlNodeType.Element:
                    string endElementString = null;
                    bool isInline;
                    bool skip = false;
                    int listIndex = 0;
                    switch (node.Name)
                    {
                        case "nav":
                            skip = true;
                            isInline = false;
                            break;
                        case "body":
                        case "section":
                        case "article":
                        case "aside":
                        case "h1":
                        case "h2":
                        case "header":
                        case "footer":
                        case "address":
                        case "main":
                        case "div":
                        case "p": // stylistic - adjust as you tend to use
                            if (textInfo.IsFirstTextOfDocWritten)
                            {
                                outText.Write("\n");
                            }
                            endElementString = "";
                            isInline = false;
                            break;
                        case "br":
                            outText.Write("\n");
                            skip = true;
                            textInfo.WritePrecedingWhiteSpace = false;
                            isInline = true;
                            break;
                        case "a":
                            if (node.Attributes.Contains("href"))
                            {
                                string href = node.Attributes["href"].Value.Trim();
                                if (node.InnerText.IndexOf(href, StringComparison.InvariantCultureIgnoreCase) == -1)
                                {
                                    endElementString = "<" + href + ">";
                                }
                            }
                            isInline = true;
                            break;
                        case "li":
                            if (textInfo.ListIndex > 0)
                            {
                                outText.Write("\r\n{0}.\t", textInfo.ListIndex++);
                            }
                            else
                            {
                                outText.Write("\r\n*\t"); //using '*' as bullet char, with tab after, but whatever you want eg "\t->", if utf-8 0x2022
                            }
                            isInline = false;
                            break;
                        case "ol":
                            listIndex = 1;
                            goto case "ul";
                        case "ul": //not handling nested lists any differently at this stage - that is getting close to rendering problems
                            endElementString = "\r\n";
                            isInline = false;
                            break;
                        case "img": //inline-block in reality
                            if (node.Attributes.Contains("alt"))
                            {
                                outText.Write('[' + node.Attributes["alt"].Value);
                                endElementString = "]";
                            }
                            if (node.Attributes.Contains("src"))
                            {
                                outText.Write('<' + node.Attributes["src"].Value + '>');
                            }
                            isInline = true;
                            break;
                        default:
                            isInline = true;
                            break;
                    }
                    if (!skip && node.HasChildNodes)
                    {
                        ConvertContentTo(node, outText, isInline ? textInfo : new PreceedingDomTextInfo(textInfo.IsFirstTextOfDocWritten) { ListIndex = listIndex });
                    }
                    if (endElementString != null)
                    {
                        outText.Write(endElementString);
                    }
                    break;
            }
        }
    }
    internal class PreceedingDomTextInfo
    {
        public PreceedingDomTextInfo(BoolWrapper isFirstTextOfDocWritten)
        {
            IsFirstTextOfDocWritten = isFirstTextOfDocWritten;
        }
        public bool WritePrecedingWhiteSpace { get; set; }
        public bool LastCharWasSpace { get; set; }
        public readonly BoolWrapper IsFirstTextOfDocWritten;
        public int ListIndex { get; set; }
    }
    internal class BoolWrapper
    {
        public BoolWrapper() { }
        public bool Value { get; set; }
        public static implicit operator bool(BoolWrapper boolWrapper)
        {
            return boolWrapper.Value;
        }
        public static implicit operator BoolWrapper(bool boolWrapper)
        {
            return new BoolWrapper { Value = boolWrapper };
        }
    }
}
