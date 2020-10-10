namespace ExcelTools
{
    public class ExcelAnalysis
    {
        private Logger Logger { get; }
        public ExcelAnalysis(Logger logger)
        {
            this.Logger = logger;
        }

        public void FindDuplicates()
        {
            this.Logger.Log("Test");
        }
    }
}