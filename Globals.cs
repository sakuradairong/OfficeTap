namespace OfficeTap
{
    internal sealed partial class Globals
    {
        private static ThisAddIn _ThisAddIn;
        private static Microsoft.Office.Tools.Excel.ApplicationFactory _factory;

        internal static ThisAddIn ThisAddIn
        {
            get { return _ThisAddIn; }
            set
            {
                if (_ThisAddIn == null)
                {
                    _ThisAddIn = value;
                    return;
                }

                throw new System.NotSupportedException();
            }
        }

        internal static Microsoft.Office.Tools.Excel.ApplicationFactory Factory
        {
            get { return _factory; }
            set
            {
                if (_factory == null)
                {
                    _factory = value;
                    return;
                }

                throw new System.NotSupportedException();
            }
        }
    }
}
