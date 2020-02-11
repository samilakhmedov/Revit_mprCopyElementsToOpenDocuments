namespace mprCopyElementsToOpenDocuments.Views
{
    /// <summary>
    /// Главное окно плагина
    /// </summary>
    public partial class MainView
    {
        private bool _isChangeableFieldsEnabled;

        /// <summary>
        /// Создает экземпляр класса <see cref="MainView"/>
        /// </summary>
        public MainView()
        {
            InitializeComponent();
            Title = ModPlusAPI.Language.GetFunctionLocalName(ModPlusConnector.Instance.Name, ModPlusConnector.Instance.LName);
        }

        /// <summary>
        /// Указывает, включены ли изменяемые поля приложения
        /// </summary>
        public bool IsChangeableFieldsEnabled
        {
            get => _isChangeableFieldsEnabled;
            set
            {
                _isChangeableFieldsEnabled = value;

                ExpandAll.IsEnabled = value;
                CollapseAll.IsEnabled = value;
                CheckAll.IsEnabled = value;
                UncheckAll.IsEnabled = value;
                ElementsTreeView.IsEnabled = value;
                AllowDuplicate.IsEnabled = value;
                RefuseDuplicate.IsEnabled = value;
                AskUser.IsEnabled = value;
                FromDoc.IsEnabled = value;
                ToDoc.IsEnabled = value;
                LogButton.IsEnabled = value;
                TransferButton.IsEnabled = value;
            }
        }
    }
}
