namespace OfficeTap
{
    public partial class ThisAddIn : Microsoft.Office.Tools.Excel.AddInBase
    {
        internal Microsoft.Office.Interop.Excel.Application Application
        {
            get { return (Microsoft.Office.Interop.Excel.Application)HostApplication; }
        }

        public ThisAddIn(
            global::Microsoft.Office.Tools.Excel.ApplicationFactory factory,
            global::Microsoft.Office.Interop.Excel.Application application)
            : base(factory, application)
        {
        }

        protected override void Initialize()
        {
            base.Initialize();
            Globals.ThisAddIn = this;
        }

        protected override void FinishInitialize()
        {
            base.FinishInitialize();
            this.InternalStartup();
        }

        protected override void OnShutdown()
        {
            this.InternalShutdown();
            base.OnShutdown();
        }

        private void InternalShutdown()
        {
            Shutdown -= ThisAddIn_Shutdown;
        }
    }
}
