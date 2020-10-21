using System.Windows.Controls;

namespace ExcelTools.Pages
{
    public partial class MultiCounter : Page
    {
        private Logger Logger { get; }
        private ExcelAnalysis ExcelAnalysis { get; }

        public MultiCounter(Logger logger)
        {
            this.InitializeComponent();

            this.Logger = logger;

            this.ExcelAnalysis = new ExcelAnalysis(this.Logger);
        }
    }
}