namespace mprCopyElementsToOpenDocuments.ViewModels
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Linq;
    using System.Windows;
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
        private readonly string _langItem = ModPlusConnector.Instance.Name;
        private readonly RevitOperationService _revitOperationService;
        private int _passedElements;
        private int _brokenElements;
        private int _totalElements = 1;
        private string _searchString = string.Empty;
        private Visibility _isShowing = Visibility.Hidden;
        private readonly MainView _mainView;
        private RevitDocument _fromDocument;
        private CopyingOptions _copyingOptions = CopyingOptions.AllowDuplicates;
        private List<BrowserItem> _selectedItems = new List<BrowserItem>();
        private ObservableCollection<RevitDocument> _openedDocuments = new ObservableCollection<RevitDocument>();
        private ObservableCollection<RevitDocument> _toDocuments = new ObservableCollection<RevitDocument>();

        /// <summary>
        /// Создает экземпляр класса <see cref="MainViewModel"/>
        /// </summary>
        /// <param name="uiApplication">Активная сессия пользовательского интерфейса Revit</param>
        /// <param name="mainView">Главное окно плагина</param>
        public MainViewModel(UIApplication uiApplication, MainView mainView)
        {
            _mainView = mainView;
            _revitOperationService = new RevitOperationService(uiApplication);
            _revitOperationService.PassedElementsCountChanged +=
                OnPassedElementsCountChanged;
            _revitOperationService.BrokenElementsCountChanged += OnBrokenElementsCountChanged;

            GeneralGroups = new ObservableCollection<GeneralItemsGroup>();

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
        /// Команда обработки выбранного документа Revit
        /// </summary>
        public ICommand ProcessSelectedDocumentCommand => new RelayCommandWithoutParameter(ProcessSelectedDocument);

        /// <summary>
        /// Команда выбора настроек копирования
        /// </summary>
        public ICommand ChangeCopyingOptionsCommand => new RelayCommand<string>(ChangeCopyingOptions);

        /// <summary>
        /// Команда раскрытия всех элементов групп браузера
        /// </summary>
        public ICommand ExpandAllCommand => new RelayCommandWithoutParameter(ExpandAll);

        /// <summary>
        /// Команда скрытия всех элементов групп браузера
        /// </summary>
        public ICommand CollapseAllCommand => new RelayCommandWithoutParameter(CollapseAll);

        /// <summary>
        /// Команда выделения всех элементов групп браузера
        /// </summary>
        public ICommand CheckAllCommand => new RelayCommandWithoutParameter(CheckAll);

        /// <summary>
        /// Команда снятия выделения всех элементов групп браузера
        /// </summary>
        public ICommand UncheckAllCommand => new RelayCommandWithoutParameter(UncheckAll);

        /// <summary>
        /// Команда начала копирования
        /// </summary>
        public ICommand StartCopyingCommand => new RelayCommandWithoutParameter(StartCopying, CanStartCopying);

        /// <summary>
        /// Команда открытия журнала работы приложения
        /// </summary>
        public ICommand OpenLogCommand => new RelayCommandWithoutParameter(OpenLog);

        /// <summary>
        /// Команда остановки процесса копирования
        /// </summary>
        public ICommand StopCopyingCommand => new RelayCommandWithoutParameter(StopCopying);

        /// <summary>
        /// Указывает, выполняет ли приложение копирование
        /// </summary>
        public Visibility IsVisible
        {
            get => _isShowing;
            set
            {
                _isShowing = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Количество скопированных элементов
        /// </summary>
        public int PassedElements
        {
            get => _passedElements;
            set
            {
                _passedElements = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Количество ошибок
        /// </summary>
        public int BrokenElements
        {
            get => _brokenElements;
            set
            {
                _brokenElements = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Общее количество элементов для копирования
        /// </summary>
        public int TotalElements
        {
            get => _totalElements;
            set
            {
                _totalElements = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Текст строки поиска
        /// </summary>
        public string SearchString
        {
            get => _searchString;
            set
            {
                _searchString = value;
                UpdateItemsVisibility();
            }
        }

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
        public ObservableCollection<GeneralItemsGroup> GeneralGroups { get; }

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
        /// Документы Revit в которые осуществляется копирование элементов
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
        /// Обновляет видимость элементов дерева
        /// </summary>
        public void UpdateItemsVisibility()
        {
            if (string.IsNullOrEmpty(SearchString))
            {
                foreach (var generalGroup in GeneralGroups)
                {
                    foreach (var categoryGroup in generalGroup.Items)
                    {
                        categoryGroup.ShowItem();
                        foreach (var typeGroup in categoryGroup.Items)
                        {
                            categoryGroup.IsExpanded = false;
                            typeGroup.Visibility = Visibility.Visible;
                            typeGroup.ShowAllItems();
                        }
                    }
                }
            }

            foreach (var generalGroup in GeneralGroups)
            {
                foreach (var categoryGroup in generalGroup.Items)
                {
                    var typeFound = false;
                    if (categoryGroup.Name.ToUpperInvariant().Contains(SearchString.ToUpperInvariant()))
                    {
                        categoryGroup.ShowItem();
                        foreach (var typeGroup in categoryGroup.Items)
                        {
                            typeGroup.ShowItem();
                            typeGroup.ShowAllItems();
                        }
                    }
                    else
                    {
                        foreach (var typeGroup in categoryGroup.Items)
                        {
                            if (typeGroup.Name.ToUpperInvariant().Contains(SearchString.ToUpperInvariant()))
                            {
                                typeFound = true;
                                categoryGroup.ShowItem(true);
                                typeGroup.ShowItem();
                                typeGroup.ShowAllItems();
                            }
                            else
                            {
                                foreach (var item in typeGroup.Items)
                                {
                                    item.Visibility = item.Name.ToUpperInvariant()
                                        .Contains(SearchString.ToUpperInvariant())
                                        ? Visibility.Visible
                                        : Visibility.Collapsed;
                                }

                                if (typeGroup.Items.Any(item => item.Visibility == Visibility.Visible))
                                {
                                    generalGroup.IsExpanded = true;
                                    categoryGroup.ShowItem(true);
                                    typeGroup.ShowItem(true);
                                }
                                else
                                {
                                    if (!typeFound)
                                    {
                                        categoryGroup.HideItem();
                                    }

                                    typeGroup.HideItem();
                                }
                            }
                        }
                    }
                }
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
                foreach (var categoryGroup in generalGroup.Items)
                {
                    categoryGroup.IsExpanded = true;
                    foreach (var typeGroup in categoryGroup.Items)
                    {
                        typeGroup.IsExpanded = true;
                    }
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
                foreach (var categoryGroup in generalGroup.Items)
                {
                    categoryGroup.IsExpanded = false;
                    foreach (var typeGroup in categoryGroup.Items)
                    {
                        typeGroup.IsExpanded = false;
                    }
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
                foreach (var categoryGroup in generalGroup.Items)
                {
                    categoryGroup.Checked = false;
                }

                generalGroup.Checked = false;
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
                case "RefuseDuplicate":
                    CopyingOptions = CopyingOptions.RefuseDuplicate;
                    break;
                case "AskUser":
                    CopyingOptions = CopyingOptions.AskUser;
                    break;
            }
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
            TotalElements = 1;
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

            UpdateItemsVisibility();
        }

        /// <summary>
        /// Выполняет копирование элементов
        /// </summary>
        private void StartCopying()
        {
            IsVisible = Visibility.Visible;
            _mainView.IsChangeableFieldsEnabled = false;
            TotalElements = SelectedItems.Any()
                ? SelectedItems.Count * ToDocuments.Count(doc => doc.Selected)
                : 1;

            _revitOperationService.CopyElements(
                            FromDocument,
                            ToDocuments.Where(doc => doc.Selected),
                            SelectedItems,
                            CopyingOptions);
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
        /// Открывает окно журнала работы приложения
        /// </summary>
        private void OpenLog()
        {
            var loggerViewModel = new LoggerViewModel();
            var loggerView = new LoggerView { DataContext = loggerViewModel };
            loggerView.ShowDialog();
        }

        /// <summary>
        /// Остановка операции копирования
        /// </summary>
        private void StopCopying()
        {
            _revitOperationService.StopCopyingOperation();
        }

        /// <summary>
        /// Метод обработки выделения
        /// </summary>
        private void OnCheckedElementsCountChanged(object sender, EventArgs e)
        {
            var allCheckedElements = new List<BrowserItem>();

            foreach (var generalGroup in GeneralGroups)
            {
                foreach (var categoryGroup in generalGroup.Items)
                {
                    foreach (var typeGroup in categoryGroup.Items)
                    {
                        if (typeGroup.Items.Any())
                        {
                            allCheckedElements.AddRange(
                                typeGroup.Items
                                    .Where(instance => instance.Checked == true));
                        }
                        else
                        {
                            if (typeGroup.Checked == true)
                                allCheckedElements.Add(typeGroup);
                        }
                    }
                }
            }

            SelectedItems = allCheckedElements;
        }

        /// <summary>
        /// Метод обработки события изменения
        /// количества элементов, прошедших проверку
        /// </summary>
        private void OnPassedElementsCountChanged(object sender, bool e)
        {
            if (e)
            {
                PassedElements++;
                IsVisible = Visibility.Hidden;
                _mainView.IsChangeableFieldsEnabled = true;
                var resultMessage = string.Format(
                    ModPlusAPI.Language.GetItem(_langItem, "m31"),
                    PassedElements - BrokenElements,
                    Environment.NewLine,
                    BrokenElements,
                    Environment.NewLine);
                TaskDialog.Show(ModPlusAPI.Language.GetItem(_langItem, "m30"), resultMessage);
                _mainView.Activate();
                PassedElements = 0;
                BrokenElements = 0;
                TotalElements = 1;
            }
            else
            {
                PassedElements++;
            }
        }

        private void OnBrokenElementsCountChanged(object sender, EventArgs e)
        {
            BrokenElements++;
        }
    }
}
