using System;
//using System.Windows.Forms;
using System.Collections.Generic;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using NextGenCodeMetricsAnalyzer.Model;
using System.Diagnostics;
using System.IO;

namespace NextGenCodeMetricsAnalyzer.Engine
{
    public class ExcelEngine
    {
        public void LoadExcelModel(string fileName, ExcelRowDataModelList listExcelRowDataModel)
        {
            try
            {
                string cell1, cell2, excelmodelcell3;
                if (File.Exists(fileName))
                {
                    Excel.Application xlApp = new Excel.Application();
                    Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fileName);
                    Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1]; // assume it is the first sheet  
                    Excel.Range xlRange = xlWorksheet.UsedRange;
                    int rowCount = xlRange.Rows.Count;

                    for (int i = 2; i <= rowCount; i++)
                    {
                        ExcelRowDataModel excelmodel = new ExcelRowDataModel();

                        if (xlRange.Cells[i, 1] != null && xlRange.Cells[i, 1].Value2 != null)
                            excelmodel.Scope = xlRange.Cells[i, 1].Value2.ToString();
                        if (xlRange.Cells[i, 2] != null && xlRange.Cells[i, 2].Value2 != null)
                            excelmodel.Project = xlRange.Cells[i, 2].Value2.ToString();
                        if (xlRange.Cells[i, 3] != null && xlRange.Cells[i, 3].Value2 != null)
                            excelmodel.Namespace = xlRange.Cells[i, 3].Value2.ToString();
                        if (xlRange.Cells[i, 4] != null && xlRange.Cells[i, 4].Value2 != null)
                            excelmodel.Type = xlRange.Cells[i, 4].Value2.ToString();
                        if (xlRange.Cells[i, 5] != null && xlRange.Cells[i, 5].Value2 != null)
                            excelmodel.Member = xlRange.Cells[i, 5].Value2.ToString();
                        if (xlRange.Cells[i, 6] != null && xlRange.Cells[i, 6].Value2 != null)
                            excelmodel.MaintainabilityIndex = xlRange.Cells[i, 6].Value2.ToString();
                        if (xlRange.Cells[i, 7] != null && xlRange.Cells[i, 7].Value2 != null)
                            excelmodel.CyclomaticComplexity = xlRange.Cells[i, 7].Value2.ToString();
                        if (xlRange.Cells[i, 8] != null && xlRange.Cells[i, 8].Value2 != null)
                            excelmodel.DepthofInheritance = xlRange.Cells[i, 8].Value2.ToString();
                        if (xlRange.Cells[i, 9] != null && xlRange.Cells[i, 9].Value2 != null)
                            excelmodel.ClassCoupling = xlRange.Cells[i, 9].Value2.ToString();
                        if (xlRange.Cells[i, 10] != null && xlRange.Cells[i, 10].Value2 != null)
                            excelmodel.LinesofCode = xlRange.Cells[i, 10].Value2.ToString();

                        listExcelRowDataModel.Add(excelmodel);
                    }


                }

            }
            catch (AccessViolationException)
            {

            }
            catch (Exception ex)
            {
                //System.Windows.Forms.MessageBox.Show("Unknown error", "Unknown error");
            }
        }
    }
}
