namespace OfficeTap
{
    internal sealed partial class Globals
    {
        private static ThisAddIn _ThisAddIn;

        internal static ThisAddIn ThisAddIn
        {
            get { return _ThisAddIn; }
            set { _ThisAddIn = value; }
        }
    }
}
