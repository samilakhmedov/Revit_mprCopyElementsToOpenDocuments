namespace mprCopyElementsToOpenDocuments.Views
{
    using System.Windows;
    using System.Windows.Controls;
    using mprCopySheetsToOpenDocuments;

    /// <summary>
    /// Главное окно плагина
    /// </summary>
    public partial class MainView
    {
        /// <summary>
        /// Создает экземпляр класса <see cref="MainView"/>
        /// </summary>
        public MainView()
        {
            InitializeComponent();
            Title = ModPlusAPI.Language.GetFunctionLocalName(ModPlusConnector.Instance.Name, ModPlusConnector.Instance.LName);
        }

        // https://stackoverflow.com/a/9494484/4944499
        private void TreeViewSelectedItem_OnHandler(object sender, RoutedEventArgs e)
        {
            if (sender is TreeViewItem item)
            {
                item.BringIntoView();
                e.Handled = true;
            }
        }
    }
}
