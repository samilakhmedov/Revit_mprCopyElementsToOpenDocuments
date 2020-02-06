namespace mprCopyElementsToOpenDocuments.Views
{
    /// <summary>
    /// Окно журнала приложения
    /// </summary>
    public partial class LoggerView
    {
        /// <summary>
        /// Создает экземпляр класса <see cref="LoggerView"/>
        /// </summary>
        public LoggerView()
        {
            InitializeComponent();
            Title = ModPlusAPI.Language.GetItem("mprCopyElementsToOpenDocuments", "h5");
        }
    }
}
