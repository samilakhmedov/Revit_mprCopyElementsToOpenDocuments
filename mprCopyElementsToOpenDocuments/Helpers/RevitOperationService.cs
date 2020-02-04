namespace mprCopyElementsToOpenDocuments.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using Autodesk.Revit.DB;
    using Autodesk.Revit.UI;
    using Models;
    using Models.Interfaces;

    /// <summary>
    /// Сервис работы с Revit
    /// </summary>
    public class RevitOperationService
    {
        private const string ParamError = "{0} - ОШИБКА! В процессе получения элемента произошла ошибка. Id: \"{1}\". Ошибка: \"{2}\"";
        private const string ParamCollectingStart = "{0} - Начало операции сбора элементов модели ({1}).";
        private const string ParamCollectingFinish = "{0} - Завершение операции сбора элементов модели ({1}).";
        private const string ParamCollectingError = "{0} - ОШИБКА! Критическая ошибка в процессе операции сбора элементов модели ({1}).";
        private const string CopyStart = "{0} - Начало копирования элементов из документа \"{1}\" в документы \"{2}\".";
        private const string CopyFinish = "{0} - Завершение копирования элементов из документа \"{1}\" в документы \"{2}\".";
        private const string CopyElementError = "{0} - ОШИБКА! В процессе копирования элемента \"{1}\" категории \"{2}\" произошла ошибка: \"{3}\".";
        private const string CopyingOptions = "Настройки копирования элементов: \"{0}\".";
        private const string CopyingNotShared = "{0} - ПРЕДУПРЕЖДЕНИЕ! В данном документе ({1}) отключен общий доступ. Копирование рабочих наборов не производится.";
        private const string WorksetsStr = "Рабочие наборы";
        private const string TypeSuffix = " (Тип)";
        private const string GeneralGroupTitle = "Все группы";
        private readonly UIApplication _uiApplication;
        private readonly List<Type> _elementTypes = new List<Type>
        {
            typeof(ExportDWGSettings), typeof(Material), typeof(ProjectInfo),
            typeof(ProjectLocation), typeof(SiteLocation), typeof(ParameterElement),
            typeof(SharedParameterElement), typeof(SunAndShadowSettings), typeof(SpatialElement),
            typeof(BrowserOrganization), typeof(DimensionType), typeof(FillPatternElement),
            typeof(ParameterFilterElement), typeof(LinePatternElement), typeof(Family),
            typeof(PhaseFilter), typeof(PrintSetting), typeof(Revision),
            typeof(RevisionSettings), typeof(TextNoteType), typeof(ViewFamilyType)
        };

        private readonly Dictionary<string, string> _specialTypeCategoryNames = new Dictionary<string, string>
        {
            { nameof(BrowserOrganization), "Обозреватель" },
            { nameof(DimensionType), "Размеры" },
            { nameof(SpotDimensionType), "Размеры" },
            { nameof(FillPatternElement), "Штриховки" },
            { nameof(ParameterFilterElement), "Фильтры параметров" },
            { nameof(LinePatternElement), "Типы линий" },
            { nameof(Family), "Загружаемые семейства" },
            { nameof(PhaseFilter), "Фильтры стадии" },
            { nameof(PrintSetting), "Настройки печати" },
            { nameof(Revision), "Изменения" },
            { nameof(RevisionSettings), "Настройки изменений" },
            { nameof(TextNoteType), "Стили текста" },
            { nameof(ViewFamilyType), "Типы семейств видов" },
            { nameof(View), "Виды" },
            { nameof(ParameterElement), "Параметры" },
            { nameof(SharedParameterElement), "Общие параметры" }
        };

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
        /// <returns>Общая группа элементов в браузере</returns>
        public GeneralItemsGroup GetAllRevitElements(RevitDocument revitDocument)
        {
            Logger.Instance.Add(string.Format(ParamCollectingStart, DateTime.Now.ToLocalTime(), revitDocument.Title));

            try
            {
                var allElements = new List<BrowserItem>();

                var elementTypes = new FilteredElementCollector(revitDocument.Document)
                    .WhereElementIsElementType()
                    .Where(e => e.Category != null && e.GetType() != typeof(ViewSheet))
                    .Select(e =>
                    {
                        try
                        {
                            return new BrowserItem(
                                e.Id.IntegerValue,
                                e.Category.Name + TypeSuffix,
                                ((ElementType)e).FamilyName,
                                e.Name);
                        }
                        catch (Exception ex)
                        {
                            Logger.Instance.Add(string.Format(ParamError, DateTime.Now.ToLocalTime(), e.Id.IntegerValue, ex.Message));
                            return null;
                        }
                    });

                var elements = new FilteredElementCollector(revitDocument.Document)
                    .WherePasses(new ElementMulticlassFilter(_elementTypes))
                    .Select(e =>
                    {
                        try
                        {
                            var elementType = (ElementType)revitDocument.Document.GetElement(e.GetTypeId());
                            return new BrowserItem(
                                e.Id.IntegerValue,
                                e.Category?.Name ?? _specialTypeCategoryNames[e.GetType().Name],
                                elementType != null ? elementType.FamilyName : string.Empty,
                                e.Name);
                        }
                        catch (Exception ex)
                        {
                            Logger.Instance.Add(string.Format(ParamError, DateTime.Now.ToLocalTime(), e.Id.IntegerValue, ex.Message));
                            return null;
                        }
                    });

                var specialCategoryElements = new FilteredElementCollector(revitDocument.Document)
                    .WherePasses(new LogicalOrFilter(new List<ElementFilter>
                    {
                        new ElementCategoryFilter(BuiltInCategory.OST_ColorFillSchema),
                        new ElementCategoryFilter(BuiltInCategory.OST_AreaSchemes),
                        new ElementCategoryFilter(BuiltInCategory.OST_Phases),
                        new ElementCategoryFilter(BuiltInCategory.OST_VolumeOfInterest)
                    }))
                    .Select(e =>
                    {
                        try
                        {
                            var elementType = (ElementType)revitDocument.Document.GetElement(e.GetTypeId());
                            return new BrowserItem(
                                e.Id.IntegerValue,
                                e.Category.Name,
                                elementType != null ? elementType.FamilyName : string.Empty,
                                e.Name);
                        }
                        catch (Exception ex)
                        {
                            Logger.Instance.Add(string.Format(ParamError, DateTime.Now.ToLocalTime(), e.Id.IntegerValue, ex.Message));
                            return null;
                        }
                    });

                var viewTemplates = new FilteredElementCollector(revitDocument.Document)
                    .OfClass(typeof(View))
                    .Where(e => ((View)e).IsTemplate)
                    .Select(e =>
                    {
                        try
                        {
                            var elementType = (ElementType)revitDocument.Document.GetElement(e.GetTypeId());
                            return new BrowserItem(
                                e.Id.IntegerValue,
                                "Виды (шаблоны)",
                                elementType != null ? elementType.FamilyName : string.Empty,
                                e.Name);
                        }
                        catch (Exception ex)
                        {
                            Logger.Instance.Add(string.Format(ParamError, DateTime.Now.ToLocalTime(), e.Id.IntegerValue, ex.Message));
                            return null;
                        }
                    });

                var views = new FilteredElementCollector(revitDocument.Document)
                    .OfClass(typeof(View))
                    .WhereElementIsNotElementType()
                    .Where(e => !((View)e).IsTemplate)
                    .Select(e =>
                    {
                        try
                        {
                            var elementType = (ElementType)revitDocument.Document.GetElement(e.GetTypeId());
                            return new BrowserItem(
                                e.Id.IntegerValue,
                                e.Category?.Name ?? _specialTypeCategoryNames[e.GetType().Name],
                                elementType != null ? elementType.FamilyName : string.Empty,
                                e.Name);
                        }
                        catch (Exception ex)
                        {
                            Logger.Instance.Add(string.Format(ParamError, DateTime.Now.ToLocalTime(), e.Id.IntegerValue, ex.Message));
                            return null;
                        }
                    });

                var elevationMarkers = new FilteredElementCollector(revitDocument.Document)
                    .OfClass(typeof(ElevationMarker))
                    .WhereElementIsNotElementType()
                    .Where(e => ((ElevationMarker)e).CurrentViewCount > 0)
                    .Select(e =>
                    {
                        try
                        {
                            var elementType = (ElementType)revitDocument.Document.GetElement(e.GetTypeId());
                            return new BrowserItem(
                                e.Id.IntegerValue,
                                e.Category.Name,
                                elementType != null ? elementType.FamilyName : string.Empty,
                                e.Name);
                        }
                        catch (Exception ex)
                        {
                            Logger.Instance.Add(string.Format(ParamError, DateTime.Now.ToLocalTime(), e.Id.IntegerValue, ex.Message));
                            return null;
                        }
                    });

                var viewports = new FilteredElementCollector(revitDocument.Document)
                    .OfClass(typeof(ElementType))
                    .Where(e => ((ElementType)e).FamilyName == "Viewport")
                    .Select(e =>
                    {
                        try
                        {
                            var elementType = (ElementType)revitDocument.Document.GetElement(e.GetTypeId());
                            return new BrowserItem(
                                e.Id.IntegerValue,
                                e.Category.Name,
                                elementType != null ? elementType.FamilyName : string.Empty,
                                e.Name);
                        }
                        catch (Exception ex)
                        {
                            Logger.Instance.Add(string.Format(ParamError, DateTime.Now.ToLocalTime(), e.Id.IntegerValue, ex.Message));
                            return null;
                        }
                    });

                IEnumerable<BrowserItem> worksets = new List<BrowserItem>();
                if (revitDocument.Document.IsWorkshared)
                {
                    worksets = new FilteredWorksetCollector(revitDocument.Document)
                        .OfKind(WorksetKind.UserWorkset)
                        .Select(e =>
                        {
                            try
                            {
                                return new BrowserItem(
                                    e.Id.IntegerValue,
                                    WorksetsStr,
                                    "-",
                                    e.Name);
                            }
                            catch (Exception ex)
                            {
                                Logger.Instance.Add(string.Format(ParamError, DateTime.Now.ToLocalTime(), e.Id.IntegerValue, ex.Message));
                                return null;
                            }
                        });
                }

                var gridAndLevelFilter = new ElementMulticlassFilter(
                    new List<Type> { typeof(Grid), typeof(Level) });
                var gridsAndLevels = new FilteredElementCollector(revitDocument.Document)
                    .WherePasses(gridAndLevelFilter)
                    .WhereElementIsNotElementType()
                    .Select(e =>
                    {
                        try
                        {
                            var elementType = (ElementType)revitDocument.Document.GetElement(e.GetTypeId());
                            return new BrowserItem(
                                e.Id.IntegerValue,
                                e.Category.Name,
                                elementType != null ? elementType.FamilyName : string.Empty,
                                e.Name);
                        }
                        catch (Exception ex)
                        {
                            Logger.Instance.Add(string.Format(ParamError, DateTime.Now.ToLocalTime(), e.Id.IntegerValue, ex.Message));
                            return null;
                        }
                    });

                var parameters = new List<BrowserItem>();
                if (!revitDocument.Document.IsFamilyDocument)
                {
                    var definitionBindingMapIterator = revitDocument.Document.ParameterBindings.ForwardIterator();
                    definitionBindingMapIterator.Reset();
                    while (definitionBindingMapIterator.MoveNext())
                    {
                        Element element = null;
                        try
                        {
                            var key = (InternalDefinition)definitionBindingMapIterator.Key;
                            element = revitDocument.Document.GetElement(key.Id);
                            var elementType = (ElementType)revitDocument.Document.GetElement(element.GetTypeId());
                            parameters.Add(new BrowserItem(
                                    element.Id.IntegerValue,
                                    element.Category?.Name ?? _specialTypeCategoryNames[element.GetType().Name],
                                    elementType != null ? elementType.FamilyName : string.Empty,
                                    element.Name));
                        }
                        catch (Exception ex)
                        {
                            Logger.Instance.Add(string.Format(
                                ParamError,
                                DateTime.Now.ToLocalTime(),
                                element != null ? element.Id.IntegerValue.ToString() : "null",
                                ex.Message));
                            return null;
                        }
                    }
                }

                var categories = new List<BrowserItem>();
                var categoriesList = revitDocument.Document.Settings.Categories;
                foreach (Category category in categoriesList)
                {
                    if (category.Id.IntegerValue > 0)
                        continue;

                    var subCategories = category.SubCategories;
                    if (subCategories == null || subCategories.Size == 0)
                        continue;

                    foreach (Category subCategory in subCategories)
                    {
                        var element = revitDocument.Document.GetElement(subCategory.Id);
                        if (element != null)
                        {
                            categories.Add(new BrowserItem(element.Id.IntegerValue, "Категории", string.Empty, element.Name));
                        }
                    }
                }

                allElements.AddRange(elementTypes);
                allElements.AddRange(elements);
                allElements.AddRange(specialCategoryElements);
                allElements.AddRange(viewTemplates);
                allElements.AddRange(views);
                allElements.AddRange(elevationMarkers);
                allElements.AddRange(viewports);
                allElements.AddRange(worksets);
                allElements.AddRange(gridsAndLevels);
                allElements.AddRange(parameters);
                allElements.AddRange(categories);

                var elementsGroupedByCategory = allElements
                    .GroupBy(e => e.CategoryName)
                    .ToList();
                var categoryGroups = new List<BrowserItem>();
                foreach (var categoryGroup in elementsGroupedByCategory)
                {
                    var elementsGroupedByType = categoryGroup
                        .GroupBy(e => e.FamilyName)
                        .ToList();
                    var typeGroups = new List<BrowserItem>();
                    foreach (var typeGroup in elementsGroupedByType)
                    {
                        var instances = typeGroup.ToList();
                        instances = instances.OrderBy(instance => instance.Name).ToList();

                        if (string.IsNullOrEmpty(typeGroup.Key))
                        {
                            instances.ForEach(instance => typeGroups.Add(instance));
                            continue;
                        }

                        typeGroups.Add(new BrowserItem(typeGroup.Key, instances));
                    }

                    typeGroups = typeGroups.OrderBy(type => type.Name).ToList();
                    categoryGroups.Add(new BrowserItem(categoryGroup.Key, typeGroups));
                }

                categoryGroups = categoryGroups.OrderBy(category => category.Name).ToList();

                Logger.Instance.Add(string.Format(ParamCollectingFinish, DateTime.Now.ToLocalTime(), revitDocument.Title));
                Logger.Instance.Add("---------");

                return new GeneralItemsGroup(GeneralGroupTitle, categoryGroups);
            }
            catch
            {
                Logger.Instance.Add(string.Format(ParamCollectingError, DateTime.Now.ToLocalTime(), revitDocument.Title));

                return new GeneralItemsGroup(GeneralGroupTitle, new List<BrowserItem>());
            }
        }

        /// <summary>
        /// Копирует все выбранные элементы в выбранные документы
        /// </summary>
        /// <param name="documentFrom">Документ из которого производится копирование</param>
        /// <param name="documentsTo">Список документов в которые осуществляется копирование</param>
        /// <param name="elements">Список элементов Revit</param>
        /// <param name="copyingOptions">Настройки копирования элементов</param>
        public void CopyElements(
            RevitDocument documentFrom,
            IEnumerable<RevitDocument> documentsTo,
            IEnumerable<BrowserItem> elements,
            CopyingOptions copyingOptions)
        {
            var revitDocuments = documentsTo.ToList();

            Logger.Instance.Add(string.Format(
                CopyStart,
                DateTime.Now.ToLocalTime(),
                documentFrom.Title,
                string.Join(", ", revitDocuments.Select(doc => doc.Title))));
            Logger.Instance.Add(string.Format(
                CopyingOptions,
                GetCopyingOptionsName(copyingOptions)));

            foreach (var element in elements)
            {
                try
                {
                    var elementId = new ElementId(element.Id);
                    var revitElement = documentFrom.Document.GetElement(elementId);
                    ICollection<ElementId> elementIds = new List<ElementId> { elementId };

                    foreach (var documentTo in revitDocuments)
                    {
                        using (var transaction = new Transaction(documentTo.Document, "Копирование элементов"))
                        {
                            var copyPasteOption = new CopyPasteOptions();
                            switch (copyingOptions)
                            {
                                case Helpers.CopyingOptions.AllowDuplicates:
                                    copyPasteOption.SetDuplicateTypeNamesHandler(new CustomCopyHandlerAllow());
                                    break;
                                case Helpers.CopyingOptions.DisallowDuplicates:
                                    copyPasteOption.SetDuplicateTypeNamesHandler(new CustomCopyHandlerAbort());
                                    break;
                            }

                            transaction.Start();

                            try
                            {
                                if (revitElement.GetType() == typeof(Workset))
                                {
                                    if (documentTo.Document.IsWorkshared)
                                    {
                                        Workset.Create(documentTo.Document, revitElement.Name);
                                    }
                                    else
                                    {
                                        Logger.Instance.Add(string.Format(
                                            CopyingNotShared,
                                            DateTime.Now.ToLocalTime(),
                                            documentTo.Title));
                                    }
                                }

                                ElementTransformUtils.CopyElements(
                                documentFrom.Document,
                                elementIds,
                                documentTo.Document,
                                null,
                                copyPasteOption);
                            }
                            catch (Exception e)
                            {
                                Logger.Instance.Add(string.Format(
                                    CopyElementError,
                                    DateTime.Now.ToLocalTime(),
                                    element.Name,
                                    element.CategoryName,
                                    e.Message));
                            }

                            transaction.Commit();
                        }
                    }
                }
                catch (Exception e)
                {
                    Logger.Instance.Add(string.Format(
                        CopyElementError,
                        DateTime.Now.ToLocalTime(),
                        element.Name,
                        element.CategoryName,
                        e.Message));
                }
            }

            Logger.Instance.Add(string.Format(
                CopyFinish,
                DateTime.Now.ToLocalTime(),
                documentFrom.Title,
                string.Join(", ", revitDocuments.Select(doc => doc.Title))));
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
