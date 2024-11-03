using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using Nice3point.Revit.Toolkit.External;
using System.IO;
using TimeTracker.Commands;
using TimeTracker.Helper;

namespace TimeTracker
{
    /// <summary>
    ///     Application entry point
    /// </summary>
    [UsedImplicitly]
    public class Application : ExternalApplication
    {
        private string _folderPath = "C:\\Users\\davydov\\source\\repos";
        private string _fileName = "TimeTracker.xlsx";
        private string filePath { get; set; }

        public override void OnStartup()
        {
            filePath = Path.Combine(_folderPath, _fileName);
            Host.Start();
            CreateRibbon();
            // Подписка на событие DocumentOpened
            Application.ControlledApplication.DocumentOpened += OnDocumentOpened;
            Application.ControlledApplication.DocumentClosing += OnDocumentClosing;
            Application.ControlledApplication.DocumentCreated += OnDocumentCreated;
            Application.ViewActivated += OnViewActivated;
        }
        public override void OnShutdown()
        {
            // Отписка от события DocumentOpened
            Application.ControlledApplication.DocumentOpened -= OnDocumentOpened;
            Application.ControlledApplication.DocumentClosing -= OnDocumentClosing;
            Application.ControlledApplication.DocumentCreated -= OnDocumentCreated;
            Application.ViewActivated -= OnViewActivated;
        }
        private void OnDocumentCreated(object sender, DocumentCreatedEventArgs e)
        {
            var test = e.Document;
            // Выполняем действия при открытии документа
            ToExcel(e.Document, "is created");
        }
        private void OnDocumentClosing(object sender, DocumentClosingEventArgs e)
        {
            var test = e.Document;
            // Выполняем действия при открытии документа
            ToExcel(e.Document, "is closed");
        }
        private void OnDocumentOpened(object sender, DocumentOpenedEventArgs e)
        {
            // Выполняем действия при открытии документа
            ToExcel(e.Document, "is opened");
        }
        // Метод-обработчик события ViewActivated
        private void OnViewActivated(object sender, ViewActivatedEventArgs e)
        {
            var currentView = e.CurrentActiveView;
            // Получаем текущий документ
            var doc = e.Document;

            // Выполняем действия при открытии документа
            InfoToExcel infoToExcel = new InfoToExcel();
            infoToExcel.ProjectName = doc.Title;
            infoToExcel.UserName = doc.Application.Username;
            infoToExcel.Description = $"View {currentView.Title} is activated";
            ExcelWriter.WriteToExcel(filePath, infoToExcel);
        }
        private void CreateRibbon()
        {
            var panel = Application.CreatePanel("Commands", "TimeTracker");

            panel.AddPushButton<StartupCommand>("Execute")
                .SetImage("/TimeTracker;component/Resources/Icons/RibbonIcon16.png")
                .SetLargeImage("/TimeTracker;component/Resources/Icons/RibbonIcon32.png");

            AddStackedButtons(panel);
        }
        private void ToExcel(Document doc, string text)
        {
            InfoToExcel infoToExcel = new InfoToExcel();
            infoToExcel.ProjectName = doc.Title;
            infoToExcel.UserName = doc.Application.Username;
            infoToExcel.Description = $"Document {doc.Title} + {text}";
            ExcelWriter.WriteToExcel(filePath, infoToExcel);
        }
        private void ProcessText(object sender, Autodesk.Revit.UI.Events.TextBoxEnterPressedEventArgs args)
        {
            // cast sender as TextBox to retrieve text value
            TextBox textBox = sender as TextBox;
            string strText = textBox.Value as string;
        }
        private void AddStackedButtons(RibbonPanel panel)
        {
            ComboBoxData cbData = new ComboBoxData("comboBox");

            TextBoxData textData = new TextBoxData("Text Box");
            textData.Name = "Text Box";
            textData.ToolTip = "Enter some text here";
            textData.LongDescription = "This is text that will appear next to the image"
                    + "when the user hovers the mouse over the control";

            IList<RibbonItem> stackedItems = panel.AddStackedItems(textData, cbData);
            if (stackedItems.Count > 1)
            {
                TextBox tBox = stackedItems[0] as TextBox;
                if (tBox != null)
                {
                    tBox.PromptText = "Enter a comment";
                    tBox.ShowImageAsButton = true;
                    tBox.ToolTip = "Enter some text";
                    // Register event handler ProcessText
                    tBox.EnterPressed += new EventHandler<Autodesk.Revit.UI.Events.TextBoxEnterPressedEventArgs>(ProcessText);
                }
            }
        }
    }
}