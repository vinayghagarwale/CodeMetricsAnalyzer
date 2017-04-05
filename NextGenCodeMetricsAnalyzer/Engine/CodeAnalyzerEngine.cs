using NextGenCodeMetricsAnalyzer.Contract;
using NextGenCodeMetricsAnalyzer.Model;
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
