namespace TimeTracker.Helper
{
    public class InfoToExcel
    {
        public string UserName { get; set; }
        public string ProjectName { get; set; }
        public string Description { get; set; }
        public string Data = DateTime.Now.ToShortDateString();
        public string Time = DateTime.Now.ToShortTimeString();
    }
}
