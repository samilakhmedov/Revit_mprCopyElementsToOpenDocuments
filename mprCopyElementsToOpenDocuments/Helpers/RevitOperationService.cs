using mprCopyElementsToOpenDocuments.Models.Interfaces;

namespace mprCopyElementsToOpenDocuments.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using Autodesk.Revit.DB;
    using Autodesk.Revit.UI;
    using Models;

    /// <summary>
    /// Сервис работы с Revit
    /// </summary>
    public class RevitOperationService
    {
        private const string ParamError = "Ошибка при получении элементов категории {0}.";
        private const string ParamStart = "{0} - Начало операции сбора элементов модели.";
        private const string ParamFinish = "{0} - Завершение операции сбора элементов модели.";
        private const string CopyStart = "{0} - Начало копирования элементов из документа \"{1}\" в документы \"{2}\".";
        private const string CopyFinish = "{0} - Завершение копирования элементов из документа \"{1}\" в документы \"{2}\".";
        private const string CopyElementError = "{0} - В процессе копирования элемента \"{1}\" категории \"{2}\" произошла ошибка: \"{3}\".";
        private const string CopyingOptions = "Настройки копирования элементов: \"{0}\".";
        private readonly UIApplication _uiApplication;
        private CopyPasteOptions _copyPasteOption = new CopyPasteOptions();

        /// <summary>
        /// Создает экземпляр класса <see cref="UIApplication"/>
        /// </summary>
        /// <param name="uiApplication">Активная сессия пользовательского интерфейса Revit</param>
        public RevitOperationService(UIApplication uiApplication)
        {
            _uiApplication = uiApplication;
        }

        /// <summary>
        /// Получает все открытые документы Revit
        /// </summary>
        public IEnumerable<RevitDocument> GetAllDocuments()
        {
            foreach (Document document in _uiApplication.Application.Documents)
            {
                yield return new RevitDocument(document);
            }
        }

        /// <summary>
        /// Получает все элементы выбранного документа Revit
        /// </summary>
        /// <param name="revitDocument">Документ Revit для извлечения данных</param>
        /// <returns></returns>
        public BrowserGeneralGroup GetAllRevitElements(RevitDocument revitDocument)
        {
            Logger.Instance.Add(string.Format(ParamStart, DateTime.Now.ToLocalTime()));

            var groups = new List<BrowserItemGroup>();

            var allElements = new FilteredElementCollector(revitDocument.Document)
                .WherePasses(new LogicalOrFilter(
                    new ElementIsElementTypeFilter(false),
                    new ElementIsElementTypeFilter(true)))
                .Where(e => e.Category != null && e.IsValidObject);

            var groupedElements = allElements.GroupBy(e => e.Category.Name);
            foreach (var group in groupedElements)
            {
                try
                {
                    var elements = group
                        .Select(element => new BrowserItem(element.Name, element.Id.IntegerValue))
                        .ToList();
                    groups.Add(new BrowserItemGroup(group.Key, elements));
                }
                catch
                {
                    Logger.Instance.Add(string.Format(ParamError, group.Key));
                }
            }

            Logger.Instance.Add(string.Format(ParamFinish, DateTime.Now.ToLocalTime()));
            Logger.Instance.Add("---------");

            return new BrowserGeneralGroup("Все группы", groups);
        }

        /// <summary>
        /// Копирует все выбранные элементы в выбранные документы
        /// </summary>
        /// <param name="documentFrom">Документ из которого производится копирование</param>
        /// <param name="documentsTo">Список документов в которые осуществляется копирование</param>
        /// <param name="elements">Список элементов Revit</param>
        /// <param name="copyingOptions">Настройки копирования элементов</param>
        public void CopyAllElements(
            RevitDocument documentFrom,
            List<RevitDocument> documentsTo,
            List<IRevitEntity> elements,
            CopyingOptions copyingOptions)
        {
            Logger.Instance.Add(string.Format(CopyStart, DateTime.Now.ToLocalTime(), documentFrom.Title, string.Join(", ", documentsTo)));
            Logger.Instance.Add(string.Format(CopyingOptions, GetCopyingOptionsName(copyingOptions)));

            try
            {
                switch (copyingOptions)
                {
                    case Helpers.CopyingOptions.AllowDuplicates:
                        _copyPasteOption.SetDuplicateTypeNamesHandler(new CustomCopyHandlerAllow());
                        break;
                    case Helpers.CopyingOptions.DisallowDuplicates:
                        _copyPasteOption.SetDuplicateTypeNamesHandler(new CustomCopyHandlerAbort());
                        break;
                }

                foreach (var element in elements)
                {
                    var elementId = new ElementId(element.Id);
                    var revitElement = documentFrom.Document.GetElement(elementId);

                    foreach (var documentTo in documentsTo)
                    {
                        if (revitElement.GetType() == typeof(Workset) && documentTo.Document.IsWorkshared)
                        {
                            Workset.Create(documentTo.Document, revitElement.Name);
                        }
                        else
                        {
                            // TODO: Logger event
                        }
                    }

                    // TODO: General copying case
                }
            }
            catch (Exception e)
            {
                Logger.Instance.Add(string.Format(CopyElementError, DateTime.Now.ToLocalTime(), "", "", e.Message));
            }

            Logger.Instance.Add(string.Format(CopyFinish, DateTime.Now.ToLocalTime(), documentFrom.Title, string.Join(", ", documentsTo)));
            Logger.Instance.Add("---------");
        }

        /// <summary>
        /// Возвращает текстовое представление настроек копирования элементов
        /// </summary>
        /// <param name="copyingOptions">Настройки копирования элементов</param>
        /// <returns>Текстовое представление настроек копирования элементов</returns>
        private string GetCopyingOptionsName(CopyingOptions copyingOptions)
        {
            switch (copyingOptions)
            {
                case Helpers.CopyingOptions.AllowDuplicates:
                    return "Разрешить дублирование";
                case Helpers.CopyingOptions.DisallowDuplicates:
                    return "Запретить дублирование";
                case Helpers.CopyingOptions.AskUser:
                    return "Запросить разрешение";
                default:
                    return string.Empty;
            }
        }
    }
}
