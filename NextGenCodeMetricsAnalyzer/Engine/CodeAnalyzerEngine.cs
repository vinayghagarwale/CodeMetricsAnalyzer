using NextGenCodeMetricsAnalyzer.Contract;
using System.Linq;
using NextGenCodeMetricsAnalyzer.Model;
using System;
namespace NextGenCodeMetricsAnalyzer.Engine
{
    public  class CodeAnalyzerEngine
    {
        private readonly ExcelRowDataModelList _lstexcelRowdataModel;
        private readonly CodeAnalyzerPreferences _codeAnalyzerPreferences;
        private readonly ExcelEngine _ExcelEngine;
        public CodeAnalyzerEngine(ExcelRowDataModelList lstexcelRowdataModel, CodeAnalyzerPreferences codeAnalyzerPreferences)
        {
            _lstexcelRowdataModel = lstexcelRowdataModel;
            _codeAnalyzerPreferences = codeAnalyzerPreferences;
            _ExcelEngine = new ExcelEngine();
        }

        public void ProcessExcel()
        {

            CodeAnalysisProcessedExcelData excel = new CodeAnalysisProcessedExcelData();


            foreach( ExcelRowDataModel excelrow in _lstexcelRowdataModel)
            {
                if(excelrow.Scope == "Type")
                {
                    

                }
                else if(excelrow.Scope == "Member")
                {
                    CyclometicComplexity cc = new CyclometicComplexity();
                    cc.Range1 = (Convert.ToInt16(excelrow.CyclomaticComplexity ) < 10) ? 1 : 1;

                    


                    excel.CyclomatiComplexityList.Add();
                }
            }



        }

        public void LoadExcelModel()
        {
            if(!string.IsNullOrEmpty(_codeAnalyzerPreferences.ExcelFileName))
            {
                _ExcelEngine.LoadExcelModel(_codeAnalyzerPreferences.ExcelFileName, _lstexcelRowdataModel);
            }
            
        }

    }
}
