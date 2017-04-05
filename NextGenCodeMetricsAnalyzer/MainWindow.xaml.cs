using NextGenCodeMetricsAnalyzer.Model;
using NextGenCodeMetricsAnalyzer.Contract;
using NextGenCodeMetricsAnalyzer.Engine;

namespace NextGenCodeMetricsAnalyzer
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        //TODO : Implement IOC
        ExcelRowDataModelList _excelRowDataModel;
        CodeAnalyzerPreferences _codeAnalyzerPreferences;
        CodeAnalyzerEngine codeAnalyzer;

        public MainWindow()
        {
            InitializeComponent();

            _excelRowDataModel = new ExcelRowDataModelList();
            _codeAnalyzerPreferences = new CodeAnalyzerPreferences();

            _codeAnalyzerPreferences.ExcelFileName = strFilePath.Text;

            codeAnalyzer = new CodeAnalyzerEngine(_excelRowDataModel, _codeAnalyzerPreferences);

        }

        private void btnDownload_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            codeAnalyzer.LoadExcelModel();
            codeAnalyzer.ProcessExcel();
        }
    }
}
