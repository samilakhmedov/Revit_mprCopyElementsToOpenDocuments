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
        private const string ParamError = "Ошибка при получении элементов категории {0}";
        private const string ParamStart = "{0} - Начало операции сбора элементов модели";
        private const string ParamFinish = "{0} - Завершение операции сбора элементов модели";
        private readonly UIApplication _uiApplication;

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
                    var elements = group.Select(element => new BrowserItem(element.Name, element.Id.IntegerValue))
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
    }
}
