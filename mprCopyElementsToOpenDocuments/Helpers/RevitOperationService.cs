using System.Collections.Generic;

namespace mprCopyElementsToOpenDocuments.Helpers
{
    using Autodesk.Revit.DB;
    using Autodesk.Revit.UI;
    using Models;

    /// <summary>
    /// Сервис работы с Revit
    /// </summary>
    public class RevitOperationService
    {
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
            // test data
            var groups = new List<BrowserItemGroup>();
            for (var i = 0; i < 50; i++)
            {
                var elements = new List<BrowserItem>();
                for (var j = 0; j < 20; j++)
                {
                    elements.Add(new BrowserItem($"Группа {(j + 1).ToString()}"));
                }

                groups.Add(new BrowserItemGroup($"Группа {(i + 1).ToString()}", elements));
            }

            return new BrowserGeneralGroup("Все группы", groups);
        }
    }
}
