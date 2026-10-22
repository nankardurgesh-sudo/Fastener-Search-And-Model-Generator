using Microsoft.Office.Interop.Excel;
using NXOpen;
using NXOpen.BlockStyler;
using NXOpen.Features;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static NXOpen.Motion.HydrodynamicBearingBuilder;
using static NXOpen.UF.UFUgmgr;

namespace SmartFastener
{
    public class Excel
    {
        public static System.Data.DataTable FetchData(string category)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            Range range = null;

            try
            {
                excelApp = new Microsoft.Office.Interop.Excel.Application();

                workbook = excelApp.Workbooks.Open(
                    @"C:\Users\PC\Desktop\SmartFastener\FastenersData.xlsx");

                worksheet = workbook.Worksheets[category];

                System.Data.DataTable data = new System.Data.DataTable();

                range = worksheet.UsedRange;

                int rowCount = range.Rows.Count;
                int columnCount = range.Columns.Count;

                // Add columns
                for (int i = 1; i <= columnCount; i++)
                {
                    string colName = Convert.ToString(
                        worksheet.Cells[1, i].Value2);

                    data.Columns.Add(colName);
                }

                // Add rows
                for (int i = 2; i <= rowCount; i++)
                {
                    DataRow dataRow = data.NewRow();

                    for (int j = 1; j <= columnCount; j++)
                    {
                        dataRow[j - 1] = Convert.ToString(
                            worksheet.Cells[i, j].Value2);
                    }

                    data.Rows.Add(dataRow);
                }

                return data;
            }
            catch (Exception ex)
            {
                UI.GetUI().NXMessageBox.Show(
                    "Error in Excel Data Fetch",
                    NXMessageBox.DialogType.Error,
                    ex.Message);

                return null;
            }
            finally
            {
                // Close workbook first
                if (workbook != null)
                {
                    try
                    {
                        workbook.Close(false);
                    }
                    catch
                    {
                        // Workbook may already be disconnected
                    }
                }

                // Quit Excel
                if (excelApp != null)
                {
                    try
                    {
                        excelApp.Quit();
                    }
                    catch
                    {
                        // Excel may already be disconnected
                    }
                }

                // Release COM objects
                if (range != null)
                    Marshal.ReleaseComObject(range);

                if (worksheet != null)
                    Marshal.ReleaseComObject(worksheet);

                if (workbook != null)
                    Marshal.ReleaseComObject(workbook);

                if (excelApp != null)
                    Marshal.ReleaseComObject(excelApp);

                range = null;
                worksheet = null;
                workbook = null;
                excelApp = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        public static void UpdateSHCSData(string msize, string length)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;

            try
            {
                excelApp = new Microsoft.Office.Interop.Excel.Application();
                workbook = excelApp.Workbooks.Open(@"C:\Users\PC\Desktop\SmartFastener\FastenersData.xlsx");

                worksheet = workbook.Worksheets["Socket Head Cap Screw"];

                Range range = worksheet.UsedRange;

                int rowCount = range.Rows.Count;
                int columnCount = range.Columns.Count;

                if (rowCount > 1)
                {
                    string lastPN = Convert.ToString(worksheet.Cells[rowCount, 1].Value2);
                    string last3 = lastPN.Substring(4);
                    double PN = Convert.ToDouble(last3) + 1;
                    string newPN = PN.ToString();
                    worksheet.Cells[rowCount + 1, 1].Value = "SHCS" + newPN;
                    worksheet.Cells[rowCount + 1, 2].Value = msize;
                    worksheet.Cells[rowCount + 1, 3].Value = length;
                }
                else
                {
                    worksheet.Cells[rowCount + 1, 1].Value = "SHCS" + "1";
                    worksheet.Cells[rowCount + 1, 2].Value = msize;
                    worksheet.Cells[rowCount + 1, 3].Value = length;
                }
                workbook.Save();
            }
            catch (Exception ex)
            {
                UI.GetUI().NXMessageBox.Show(
                    "Error in Excel Update",
                    NXMessageBox.DialogType.Error,
                    ex.Message);
            }
            finally
            {
                if (workbook != null)
                {
                    try
                    {
                        workbook.Close(false);
                    }
                    catch
                    {
                    }
                }

                if (excelApp != null)
                {
                    try
                    {
                        excelApp.Quit();
                    }
                    catch
                    {
                    }
                }

                if (worksheet != null)
                    Marshal.ReleaseComObject(worksheet);

                if (workbook != null)
                    Marshal.ReleaseComObject(workbook);

                if (excelApp != null)
                    Marshal.ReleaseComObject(excelApp);

                worksheet = null;
                workbook = null;
                excelApp = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

        }

        public static void UpdateHSCSData(string msize, string length)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;

            try
            {
                excelApp = new Microsoft.Office.Interop.Excel.Application();
                workbook = excelApp.Workbooks.Open(@"C:\Users\PC\Desktop\SmartFastener\FastenersData.xlsx");
                worksheet = workbook.Worksheets["Hex Socket Countersunk Screw"];

                Range range = worksheet.UsedRange;

                int rowCount = range.Rows.Count;
                int columnCount = range.Columns.Count;

                if (rowCount > 1)
                {
                    string lastPN = Convert.ToString(worksheet.Cells[rowCount, 1].Value2);
                    string last3 = lastPN.Substring(4);
                    double PN = Convert.ToDouble(last3) + 1;
                    string newPN = PN.ToString();
                    worksheet.Cells[rowCount + 1, 1].Value = "HSCS" + newPN;
                    worksheet.Cells[rowCount + 1, 2].Value = msize;
                    worksheet.Cells[rowCount + 1, 3].Value = length;
                }
                else
                {
                    worksheet.Cells[rowCount + 1, 1].Value = "HSCS" + "1";
                    worksheet.Cells[rowCount + 1, 2].Value = msize;
                    worksheet.Cells[rowCount + 1, 3].Value = length;
                }
                workbook.Save();
            }
            catch (Exception ex)
            {
                UI.GetUI().NXMessageBox.Show(
                    "Error in Excel Update",
                    NXMessageBox.DialogType.Error,
                    ex.Message);
            }
            finally
            {
                if (workbook != null)
                {
                    try
                    {
                        workbook.Close(false);
                    }
                    catch
                    {
                    }
                }

                if (excelApp != null)
                {
                    try
                    {
                        excelApp.Quit();
                    }
                    catch
                    {
                    }
                }

                if (worksheet != null)
                    Marshal.ReleaseComObject(worksheet);

                if (workbook != null)
                    Marshal.ReleaseComObject(workbook);

                if (excelApp != null)
                    Marshal.ReleaseComObject(excelApp);

                worksheet = null;
                workbook = null;
                excelApp = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

        }

        public static void UpdateHRNData(string msize)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;

            try
            {
                excelApp = new Microsoft.Office.Interop.Excel.Application();
                workbook = excelApp.Workbooks.Open(@"C:\Users\PC\Desktop\SmartFastener\FastenersData.xlsx");

                worksheet = workbook.Worksheets["Hexagon Regular Nut"];

                Range range = worksheet.UsedRange;

                int rowCount = range.Rows.Count;
                int columnCount = range.Columns.Count;

                if (rowCount > 1)
                {
                    string lastPN = Convert.ToString(worksheet.Cells[rowCount, 1].Value2);
                    string last3 = lastPN.Substring(3);
                    double PN = Convert.ToDouble(last3) + 1;
                    string newPN = PN.ToString();
                    worksheet.Cells[rowCount + 1, 1].Value = "HRN" + newPN;
                    worksheet.Cells[rowCount + 1, 2].Value = msize;
                }
                else
                {
                    worksheet.Cells[rowCount + 1, 1].Value = "HRN" + "1";
                    worksheet.Cells[rowCount + 1, 2].Value = msize;
                }
                workbook.Save();
            }
            catch (Exception ex)
            {
                UI.GetUI().NXMessageBox.Show(
                    "Error in Excel Update",
                    NXMessageBox.DialogType.Error,
                    ex.Message);
            }
            finally
            {
                if (workbook != null)
                {
                    try
                    {
                        workbook.Close(false);
                    }
                    catch
                    {
                    }
                }

                if (excelApp != null)
                {
                    try
                    {
                        excelApp.Quit();
                    }
                    catch
                    {
                    }
                }

                if (worksheet != null)
                    Marshal.ReleaseComObject(worksheet);

                if (workbook != null)
                    Marshal.ReleaseComObject(workbook);

                if (excelApp != null)
                    Marshal.ReleaseComObject(excelApp);

                worksheet = null;
                workbook = null;
                excelApp = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        public static void UpdatePWData(string id, string od, string thickness)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;

            try
            {
                excelApp = new Microsoft.Office.Interop.Excel.Application();
                workbook = excelApp.Workbooks.Open(@"C:\Users\PC\Desktop\SmartFastener\FastenersData.xlsx");

                worksheet = workbook.Worksheets["Plain Washer"];

                Range range = worksheet.UsedRange;

                int rowCount = range.Rows.Count;
                int columnCount = range.Columns.Count;

                if (rowCount > 1)
                {
                    string lastPN = Convert.ToString(worksheet.Cells[rowCount, 1].Value2);
                    string last3 = lastPN.Substring(2);
                    double PN = Convert.ToDouble(last3) + 1;
                    string newPN = PN.ToString();
                    worksheet.Cells[rowCount + 1, 1].Value = "PW" + newPN;
                    worksheet.Cells[rowCount + 1, 2].Value = id;
                    worksheet.Cells[rowCount + 1, 3].Value = od;
                    worksheet.Cells[rowCount + 1, 4].Value = thickness;
                }
                else
                {
                    worksheet.Cells[rowCount + 1, 1].Value = "PW" + "1";
                    worksheet.Cells[rowCount + 1, 2].Value = id;
                    worksheet.Cells[rowCount + 1, 3].Value = od;
                    worksheet.Cells[rowCount + 1, 4].Value = thickness;
                }
                workbook.Save();
            }
            catch (Exception ex)
            {
                UI.GetUI().NXMessageBox.Show(
                    "Error in Excel Update",
                    NXMessageBox.DialogType.Error,
                    ex.Message);
            }
            finally
            {
                if (workbook != null)
                {
                    try
                    {
                        workbook.Close(false);
                    }
                    catch
                    {
                    }
                }

                if (excelApp != null)
                {
                    try
                    {
                        excelApp.Quit();
                    }
                    catch
                    {
                    }
                }

                if (worksheet != null)
                    Marshal.ReleaseComObject(worksheet);

                if (workbook != null)
                    Marshal.ReleaseComObject(workbook);

                if (excelApp != null)
                    Marshal.ReleaseComObject(excelApp);

                worksheet = null;
                workbook = null;
                excelApp = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }


        }

        public static string GetNewSHCSPN()
        {
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            Workbook workbook = null;

            try
            {
                excelApp = new Microsoft.Office.Interop.Excel.Application();

                workbook = excelApp.Workbooks.Open(@"C:\Users\PC\Desktop\SmartFastener\FastenersData.xlsx");

                Worksheet worksheet = workbook.Worksheets["Socket Head Cap Screw"];

                Range range = worksheet.UsedRange;

                int rowCount = range.Rows.Count;

                if (rowCount > 1)
                {
                    string lastPN = Convert.ToString(worksheet.Cells[rowCount, 1].Value2);
                    string last3 = lastPN.Substring(4);
                    double PN = Convert.ToDouble(last3) + 1;
                    return PN.ToString();
                }

                return "1";
            }
            finally
            {
                if(workbook != null)
                {
                    workbook.Close();
                }
                if(excelApp != null)
                {  
                    excelApp.Quit(); 
                }
            }
        }

        public static string GetNewHSCSPN()
        {
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            Workbook workbook = null;

            try
            {
                excelApp = new Microsoft.Office.Interop.Excel.Application();
                workbook = excelApp.Workbooks.Open(@"C:\Users\PC\Desktop\SmartFastener\FastenersData.xlsx");

                Worksheet worksheet = workbook.Worksheets["Hex Socket Countersunk Screw"];

                Range range = worksheet.UsedRange;

                int rowCount = range.Rows.Count;
                int columnCount = range.Columns.Count;

                if (rowCount > 1)
                {
                    string lastPN = Convert.ToString(worksheet.Cells[rowCount, 1].Value2);
                    string last3 = lastPN.Substring(4);
                    double PN = Convert.ToDouble(last3) + 1;
                    string newPN = PN.ToString();
                    return newPN;
                }
                return "1";
            }
            finally
            {
                if (workbook != null)
                {
                    workbook.Close();
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                }
            }

        }

        public static string GetNewHRNPN()
        {
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            Workbook workbook = null;
            try
            {
                excelApp = new Microsoft.Office.Interop.Excel.Application();
                workbook = excelApp.Workbooks.Open(@"C:\Users\PC\Desktop\SmartFastener\FastenersData.xlsx");

                Worksheet worksheet = workbook.Worksheets["Hexagon Regular Nut"];

                Range range = worksheet.UsedRange;

                int rowCount = range.Rows.Count;
                int columnCount = range.Columns.Count;

                if (rowCount > 1)
                {
                    string lastPN = Convert.ToString(worksheet.Cells[rowCount, 1].Value2);
                    string last3 = lastPN.Substring(3);
                    double PN = Convert.ToDouble(last3) + 1;
                    string newPN = PN.ToString();
                    return newPN;
                }             
                return "1";   
            }
            finally
            {
                if (workbook != null)
                {
                    workbook.Close();
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                }
            }
        }

        public static string GetNewPWPN()
        {
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            Workbook workbook = null;
            try
            {
                excelApp = new Microsoft.Office.Interop.Excel.Application();
                workbook = excelApp.Workbooks.Open(@"C:\Users\PC\Desktop\SmartFastener\FastenersData.xlsx");

                Worksheet worksheet = workbook.Worksheets["Plain Washer"];

                Range range = worksheet.UsedRange;

                int rowCount = range.Rows.Count;
                int columnCount = range.Columns.Count;

                if (rowCount > 1)
                {
                    string lastPN = Convert.ToString(worksheet.Cells[rowCount, 1].Value2);
                    string last3 = lastPN.Substring(2);
                    double PN = Convert.ToDouble(last3) + 1;
                    string newPN = PN.ToString();
                    return newPN;
                }
                return "1";
            }
            finally
            {
                if (workbook != null)
                {
                    workbook.Close();
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                }
            }
        }

        public static void CheckSHCSAvailibilty(string mSize, string length)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            Workbook workbook = null;

            Session theSession = Session.GetSession();

            try
            {
                excelApp = new Microsoft.Office.Interop.Excel.Application();

                workbook = excelApp.Workbooks.Open(@"C:\Users\PC\Desktop\SmartFastener\FastenersData.xlsx");

                Worksheet worksheet = workbook.Worksheets["Socket Head Cap Screw"];

                Range range = worksheet.UsedRange;

                int rowCount = range.Rows.Count;
                int columnCount = range.Columns.Count;

                bool matchFound = false;

                if(rowCount>1)
                {
                    for(int i = 2; i <= rowCount; i++)
                    {
                        if (Convert.ToString(worksheet.Cells[i, 2].Value2) == mSize)
                        {
                            if(Convert.ToString(worksheet.Cells[i, 3].Value2) == length)
                            {
                                theSession.ListingWindow.Open();
                                theSession.ListingWindow.WriteLine("Please use PN : " + Convert.ToString(worksheet.Cells[i, 1].Value2));
                                matchFound = true;
                                break;
                            }
                        }       
                    }
                }
                if(!matchFound)
                {
                    theSession.ListingWindow.Open();
                    theSession.ListingWindow.WriteLine("There is no PN available for this configuration. Please generate a new model");
                }
            }
            finally
            {
                if (workbook != null)
                {
                    workbook.Close();
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                }
            }
        }

        public static void CheckHSCSAvailibilty(string mSize, string length)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            Workbook workbook = null;

            Session theSession = Session.GetSession();

            try
            {
                excelApp = new Microsoft.Office.Interop.Excel.Application();

                workbook = excelApp.Workbooks.Open(@"C:\Users\PC\Desktop\SmartFastener\FastenersData.xlsx");

                Worksheet worksheet = workbook.Worksheets["Hex Socket Countersunk Screw"];

                Range range = worksheet.UsedRange;

                int rowCount = range.Rows.Count;
                int columnCount = range.Columns.Count;

                bool matchFound = false;

                if (rowCount > 1)
                {
                    for (int i = 2; i <= rowCount; i++)
                    {
                        if (Convert.ToString(worksheet.Cells[i, 2].Value2) == mSize)
                        {
                            if (Convert.ToString(worksheet.Cells[i, 3].Value2) == length)
                            {
                                theSession.ListingWindow.Open();
                                theSession.ListingWindow.WriteLine("Please use PN : " + Convert.ToString(worksheet.Cells[i, 1].Value2));
                                matchFound = true;
                                break;
                            }
                        }
                    }
                }
                if (!matchFound)
                {
                    theSession.ListingWindow.Open();
                    theSession.ListingWindow.WriteLine("There is no PN available for this configuration. Please generate a new model");
                }
            }
            finally
            {
                if (workbook != null)
                {
                    workbook.Close();
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                }
            }
        }

        public static void CheckHRNAvailibilty(string mSize)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            Workbook workbook = null;

            Session theSession = Session.GetSession();

            try
            {
                excelApp = new Microsoft.Office.Interop.Excel.Application();

                workbook = excelApp.Workbooks.Open(@"C:\Users\PC\Desktop\SmartFastener\FastenersData.xlsx");

                Worksheet worksheet = workbook.Worksheets["Hexagon Regular Nut"];

                Range range = worksheet.UsedRange;

                int rowCount = range.Rows.Count;
                int columnCount = range.Columns.Count;

                bool matchFound = false;

                if (rowCount > 1)
                {
                    for (int i = 2; i <= rowCount; i++)
                    {
                        if (Convert.ToString(worksheet.Cells[i, 2].Value2) == mSize)
                        {
                            theSession.ListingWindow.Open();
                            theSession.ListingWindow.WriteLine("Please use PN : " + Convert.ToString(worksheet.Cells[i, 1].Value2));
                            matchFound = true;
                            break;
                        }
                    }
                }
                if (!matchFound)
                {
                    theSession.ListingWindow.Open();
                    theSession.ListingWindow.WriteLine("There is no PN available for this configuration. Please generate a new model");
                }
            }
            finally
            {
                if (workbook != null)
                {
                    workbook.Close();
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                }
            }
        }

        public static void CheckPWAvailibilty(string id, string od, string thickness)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            Workbook workbook = null;

            Session theSession = Session.GetSession();

            try
            {
                excelApp = new Microsoft.Office.Interop.Excel.Application();

                workbook = excelApp.Workbooks.Open(@"C:\Users\PC\Desktop\SmartFastener\FastenersData.xlsx");

                Worksheet worksheet = workbook.Worksheets["Hexagon Regular Nut"];

                Range range = worksheet.UsedRange;

                int rowCount = range.Rows.Count;
                int columnCount = range.Columns.Count;

                bool matchFound = false;

                if (rowCount > 1)
                {
                    for (int i = 2; i <= rowCount; i++)
                    {
                        if (Convert.ToString(worksheet.Cells[i, 2].Value2) == id)
                        {
                            if (Convert.ToString(worksheet.Cells[i, 3].Value2) == od)
                            {
                                if(Convert.ToString(worksheet.Cells[i, 4].Value2) == thickness)
                                {   
                                    theSession.ListingWindow.Open();
                                    theSession.ListingWindow.WriteLine("Please use PN : " + Convert.ToString(worksheet.Cells[i, 1].Value2));
                                    matchFound = true;
                                    break;
                                }
                            }
                        }
                    }
                }
                if (!matchFound)
                {
                    theSession.ListingWindow.Open();
                    theSession.ListingWindow.WriteLine("There is no PN available for this configuration. Please generate a new model");
                }
            }
            finally
            {
                if (workbook != null)
                {
                    workbook.Close();
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                }
            }
        }

    }
}
