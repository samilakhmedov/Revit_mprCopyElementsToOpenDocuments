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
        private readonly RevitOperationService _revitOperationService;
        private MainView _mainView;
        private RevitDocument _fromDocument;
        private CopyingOptions _copyingOptions = CopyingOptions.AllowDuplicates;
        private List<BrowserItem> _selectedItems = new List<BrowserItem>();
        private ObservableCollection<BrowserGeneralGroup> _generalGroups = new ObservableCollection<BrowserGeneralGroup>();
        private ObservableCollection<RevitDocument> _openedDocuments = new ObservableCollection<RevitDocument>();
        private ObservableCollection<RevitDocument> _toDocuments = new ObservableCollection<RevitDocument>();

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

            foreach (var document in Documents)
            {
                if (FromDocument != null)
                {
                    if (document.Title != FromDocument.Title)
                        ToDocuments.Add(document);
                }
                else
                {
                    ToDocuments.Add(document);
                }
            }
        }

        /// <summary>
        /// Команда выбора текущего документа Revit
        /// </summary>
        public ICommand ProcessSelectedDocumentCommand =>
            new RelayCommandWithoutParameter(ProcessSelectedDocument);

        /// <summary>
        /// Команда выбора настроек копирования
        /// </summary>
        public ICommand ChangeCopyingOptionsCommand =>
            new RelayCommand<string>(ChangeCopyingOptions);

        /// <summary>
        /// Команда раскрытия всех элементов групп браузера
        /// </summary>
        public ICommand ExpandAllCommand =>
            new RelayCommandWithoutParameter(ExpandAll);

        /// <summary>
        /// Команда скрытия всех элементов групп браузера
        /// </summary>
        public ICommand CollapseAllCommand =>
            new RelayCommandWithoutParameter(CollapseAll);

        /// <summary>
        /// Команда выделения всех элементов групп браузера
        /// </summary>
        public ICommand CheckAllCommand =>
            new RelayCommandWithoutParameter(CheckAll);

        /// <summary>
        /// Команда снятия выделения всех элементов групп браузера
        /// </summary>
        public ICommand UncheckAllCommand =>
            new RelayCommandWithoutParameter(UncheckAll);

        /// <summary>
        /// Команда начала копирования
        /// </summary>
        public ICommand StartCopyingCommand =>
            new RelayCommandWithoutParameter(StartCopying, CanStartCopying);

        /// <summary>
        /// Настройки копирования элементов
        /// </summary>
        public CopyingOptions CopyingOptions
        {
            get => _copyingOptions;
            set
            {
                _copyingOptions = value;
                OnPropertyChanged();
            }
        }

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
        public ObservableCollection<RevitDocument> ToDocuments
        {
            get => _toDocuments;
            set
            {
                _toDocuments = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Элементы текущего выбора
        /// </summary>
        public List<BrowserItem> SelectedItems
        {
            get => _selectedItems;
            set
            {
                _selectedItems = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Разворачивает все элементы групп в браузере
        /// </summary>
        private void ExpandAll()
        {
            foreach (var generalGroup in GeneralGroups)
            {
                generalGroup.IsExpanded = true;

                foreach (var itemsGroup in generalGroup.Items)
                {
                    itemsGroup.IsExpanded = true;
                }
            }
        }

        /// <summary>
        /// Сворачивает все элементы групп в браузере
        /// </summary>
        private void CollapseAll()
        {
            foreach (var generalGroup in GeneralGroups)
            {
                generalGroup.IsExpanded = false;

                foreach (var itemsGroup in generalGroup.Items)
                {
                    itemsGroup.IsExpanded = false;
                }
            }
        }

        /// <summary>
        /// Отмечает все элементы групп в браузере
        /// </summary>
        private void CheckAll()
        {
            foreach (var generalGroup in GeneralGroups)
            {
                generalGroup.Checked = true;
            }
        }

        /// <summary>
        /// Снимает выделение со всех элементов групп в браузере
        /// </summary>
        private void UncheckAll()
        {
            foreach (var generalGroup in GeneralGroups)
            {
                generalGroup.Checked = false;
            }
        }

        /// <summary>
        /// Метод обработки выделения
        /// </summary>
        private void OnCheckedElementsCountChanged(object sender, EventArgs e)
        {
            SelectedItems = _generalGroups
                .SelectMany(generalGroup => generalGroup.Items)
                .SelectMany(itemsGroup => itemsGroup.Items)
                .Where(item => item.Checked == true)
                .ToList();
        }

        /// <summary>
        /// Проверка возможности начала копирования
        /// </summary>
        /// <returns></returns>
        private bool CanStartCopying(object obj)
        {
            return SelectedItems.Count > 0
                   && FromDocument != null
                   && ToDocuments.Any(doc => doc.Selected);
        }

        /// <summary>
        /// Загружает все элементы выбранного документа Revit
        /// </summary>
        private void ProcessSelectedDocument()
        {
            var generalGroup = _revitOperationService.GetAllRevitElements(FromDocument);
            generalGroup.SelectionChanged += OnCheckedElementsCountChanged;

            GeneralGroups.Clear();
            GeneralGroups.Add(generalGroup);
            ToDocuments.Clear();
            SelectedItems.Clear();
            OnPropertyChanged(nameof(SelectedItems));

            foreach (var doc in Documents)
            {
                doc.Selected = false;
            }

            foreach (var document in Documents)
            {
                if (FromDocument != null)
                {
                    if (document.Title != FromDocument.Title)
                        ToDocuments.Add(document);
                }
                else
                {
                    ToDocuments.Add(document);
                }
            }
        }

        /// <summary>
        /// Изменяет текущие настройки копирования
        /// </summary>
        /// <param name="name">Имя выбранного условия</param>
        private void ChangeCopyingOptions(string name)
        {
            switch (name)
            {
                case "AllowDuplicate":
                    CopyingOptions = CopyingOptions.AllowDuplicates;
                    break;
                case "DisallowDuplicate":
                    CopyingOptions = CopyingOptions.DisallowDuplicates;
                    break;
                case "AskUser":
                    CopyingOptions = CopyingOptions.AskUser;
                    break;
            }
        }

        /// <summary>
        /// Выполняет копирование элементов
        /// </summary>
        private void StartCopying()
        {
        }
    }
}
