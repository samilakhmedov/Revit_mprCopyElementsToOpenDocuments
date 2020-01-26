namespace mprCopyElementsToOpenDocuments.ViewModels
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Linq;
    using System.Windows.Input;
    using Autodesk.Revit.UI;
    using Helpers;
    using Models;
    using ModPlusAPI.Mvvm;
    using Views;

    /// <summary>
    /// Модель представления главного окна плагина
    /// </summary>
    public class MainViewModel : VmBase
    {
        private int _selectedItemsGroup;
        private readonly RevitOperationService _revitOperationService;
        private MainView _mainView;
        private RevitDocument _fromDocument;
        private List<RevitDocument> _toDocuments;
        private ObservableCollection<BrowserGeneralGroup> _generalGroups = new ObservableCollection<BrowserGeneralGroup>();
        private ObservableCollection<RevitDocument> _openedDocuments = new ObservableCollection<RevitDocument>();

        /// <summary>
        /// Создает экземпляр класса <see cref="MainViewModel"/>
        /// </summary>
        /// <param name="uiApplication">Активная сессия пользовательского интерфейса Revit</param>
        /// <param name="mainView">Главное окно плагина</param>
        public MainViewModel(UIApplication uiApplication, MainView mainView)
        {
            _revitOperationService = new RevitOperationService(uiApplication);
            _mainView = mainView;

            var docs = _revitOperationService.GetAllDocuments();
            foreach (var doc in docs)
            {
                Documents.Add(doc);
            }
        }

        /// <summary>
        /// Команда выбора текущего документа Revit
        /// </summary>
        public ICommand LoadCurrentDocumentElementsCommand =>
            new RelayCommandWithoutParameter(LoadCurrentDocumentElements);

        /// <summary>
        /// Обобщенные группы элементов для отображения в дереве
        /// </summary>
        public ObservableCollection<BrowserGeneralGroup> GeneralGroups
        {
            get => _generalGroups;
            set
            {
                _generalGroups = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Открытые документы
        /// </summary>
        public ObservableCollection<RevitDocument> Documents
        {
            get => _openedDocuments;
            set
            {
                _openedDocuments = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Документ Revit из которого производится копирование элементов
        /// </summary>
        public RevitDocument FromDocument
        {
            get => _fromDocument;
            set
            {
                _fromDocument = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Документ Revit из которого производится копирование элементов
        /// </summary>
        public List<RevitDocument> ToDocuments
        {
            get => _toDocuments;
            set
            {
                _toDocuments = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Количество выбранных элементов
        /// </summary>
        public int SelectedItemsCount
        {
            get => _selectedItemsGroup;
            set
            {
                _selectedItemsGroup = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Метод обработки выделения
        /// </summary>
        private void OnCheckedElementsCountChanged(object sender, EventArgs e)
        {
            SelectedItemsCount = _generalGroups
                .SelectMany(generalGroup => generalGroup.Items)
                .SelectMany(itemsGroup => itemsGroup.Items)
                .Count(item => item.Checked == true);
        }

        /// <summary>
        /// Загружает все элементы выбранного документа Revit
        /// </summary>
        private void LoadCurrentDocumentElements()
        {
            var generalGroup = _revitOperationService.GetAllRevitElements(FromDocument);
            generalGroup.SelectionChanged += OnCheckedElementsCountChanged;
            _generalGroups.Add(generalGroup);
        }
    }
}
