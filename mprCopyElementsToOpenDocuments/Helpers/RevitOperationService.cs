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
        private const string ParamError = "Ошибка при получении элементов категории {0}.";
        private const string ParamStart = "{0} - Начало операции сбора элементов модели.";
        private const string ParamFinish = "{0} - Завершение операции сбора элементов модели.";
        private const string CopyStart = "{0} - Начало копирования элементов из документа \"{1}\" в документы \"{2}\".";
        private const string CopyFinish = "{0} - Завершение копирования элементов из документа \"{1}\" в документы \"{2}\".";
        private const string CopyElementError = "{0} - В процессе копирования элемента \"{1}\" категории \"{2}\" произошла ошибка: \"{3}\".";
        private const string CopyingOptions = "Настройки копирования элементов: \"{0}\".";
        private const string WorksetsStr = "Рабочие наборы";
        private readonly UIApplication _uiApplication;
        private readonly CopyPasteOptions _copyPasteOption = new CopyPasteOptions();
        private readonly List<Type> _elementTypes = new List<Type>
        {
            typeof(ExportDWGSettings), typeof(Material), typeof(ProjectInfo),
            typeof(ProjectLocation), typeof(SiteLocation), typeof(Revision),
            typeof(PhaseFilter), typeof(LinePatternElement), typeof(FillPatternElement),
            typeof(ParameterElement), typeof(SharedParameterElement), typeof(SunAndShadowSettings),
            typeof(SpatialElement)
        };

        private readonly List<Type> _elementTypesWithoutCategories = new List<Type>
        {
            typeof(BrowserOrganization), typeof(DimensionType), typeof(FillPatternElement),
            typeof(ParameterFilterElement), typeof(LinePatternElement), typeof(Family),
            typeof(PhaseFilter), typeof(PrintSetting), typeof(Revision),
            typeof(RevisionSettings), typeof(TextNoteType), typeof(ViewFamilyType)
        };

        private readonly Dictionary<string, string> _specialTypeNames = new Dictionary<string, string>
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
        /// <returns></returns>
        public BrowserGeneralGroup GetAllRevitElements(RevitDocument revitDocument)
        {
            Logger.Instance.Add(string.Format(ParamStart, DateTime.Now.ToLocalTime()));

            try
            {
                var groups = new List<BrowserItemsGroup>();
                var allElements = new List<BrowserItem>();

                var elementTypes = new FilteredElementCollector(revitDocument.Document)
                    .WhereElementIsElementType()
                    .Where(e => e.Category != null && e.GetType() != typeof(ViewSheet))
                    .Select(e =>
                    {
                        var elementType = (ElementType)revitDocument.Document.GetElement(e.GetTypeId());
                        return new BrowserItem(
                            e.Id.IntegerValue,
                            e.Category.Name + " (Тип)",
                            elementType != null ? elementType.FamilyName : "-",
                            e.Name,
                            true);
                    });

                var elementsWithoutCategories = new FilteredElementCollector(revitDocument.Document)
                    .WherePasses(new ElementMulticlassFilter(_elementTypesWithoutCategories))
                    .Select(e =>
                    {
                        var elementType = (ElementType)revitDocument.Document.GetElement(e.GetTypeId());
                        return new BrowserItem(
                            e.Id.IntegerValue,
                            e.Category?.Name ?? _specialTypeNames[e.GetType().Name],
                            elementType != null ? elementType.FamilyName : "-",
                            e.Name);
                    });

                var searchedElementsWithCategory = new FilteredElementCollector(revitDocument.Document)
                    .WherePasses(new ElementMulticlassFilter(_elementTypes))
                    .Where(e => e.Category != null)
                    .Select(e =>
                    {
                        var elementType = (ElementType)revitDocument.Document.GetElement(e.GetTypeId());
                        return new BrowserItem(
                            e.Id.IntegerValue,
                            e.Category.Name,
                            elementType != null ? elementType.FamilyName : "-",
                            e.Name);
                    });

                var searchedElementsWithoutCategory = new FilteredElementCollector(revitDocument.Document)
                    .WherePasses(new ElementMulticlassFilter(_elementTypes))
                    .Where(e => e.Category == null)
                    .Select(e =>
                    {
                        var elementType = (ElementType)revitDocument.Document.GetElement(e.GetTypeId());
                        return new BrowserItem(
                            e.Id.IntegerValue,
                            e.Category?.Name ?? _specialTypeNames[e.GetType().Name],
                            elementType != null ? elementType.FamilyName : "-",
                            e.Name);
                    });

                var specialCategories = new FilteredElementCollector(revitDocument.Document)
                    .WherePasses(new LogicalOrFilter(new List<ElementFilter>
                    {
                        new ElementCategoryFilter(BuiltInCategory.OST_ColorFillSchema),
                        new ElementCategoryFilter(BuiltInCategory.OST_AreaSchemes),
                        new ElementCategoryFilter(BuiltInCategory.OST_Phases),
                        new ElementCategoryFilter(BuiltInCategory.OST_VolumeOfInterest)
                    }))
                    .Select(e =>
                    {
                        var elementType = (ElementType)revitDocument.Document.GetElement(e.GetTypeId());
                        return new BrowserItem(
                            e.Id.IntegerValue,
                            e.Category.Name,
                            elementType != null ? elementType.FamilyName : "-",
                            e.Name);
                    });

                var viewTemplates = new FilteredElementCollector(revitDocument.Document)
                    .OfClass(typeof(View))
                    .Where(e => ((View)e).IsTemplate)
                    .Select(e =>
                    {
                        var elementType = (ElementType)revitDocument.Document.GetElement(e.GetTypeId());
                        return new BrowserItem(
                            e.Id.IntegerValue,
                            "Виды (шаблоны)",
                            elementType != null ? elementType.FamilyName : "-",
                            e.Name);
                    });

                var views = new FilteredElementCollector(revitDocument.Document)
                    .OfClass(typeof(View))
                    .WhereElementIsNotElementType()
                    .Where(e => !((View)e).IsTemplate)
                    .Select(e =>
                    {
                        var elementType = (ElementType)revitDocument.Document.GetElement(e.GetTypeId());

                        return new BrowserItem(
                            e.Id.IntegerValue,
                            e.Category?.Name ?? _specialTypeNames[e.GetType().Name],
                            elementType != null ? elementType.FamilyName : "-",
                            e.Name);
                    });

                var elevationMarkers = new FilteredElementCollector(revitDocument.Document)
                    .OfClass(typeof(ElevationMarker))
                    .WhereElementIsNotElementType()
                    .Where(e => ((ElevationMarker)e).CurrentViewCount > 0)
                    .Select(e =>
                    {
                        var elementType = (ElementType)revitDocument.Document.GetElement(e.GetTypeId());
                        return new BrowserItem(
                            e.Id.IntegerValue,
                            e.Category.Name,
                            elementType != null ? elementType.FamilyName : "-",
                            e.Name);
                    });

                var viewports = new FilteredElementCollector(revitDocument.Document)
                    .OfClass(typeof(ElementType))
                    .Where(e => ((ElementType)e).FamilyName == "Viewport")
                    .Select(e =>
                    {
                        var elementType = (ElementType)revitDocument.Document.GetElement(e.GetTypeId());
                        return new BrowserItem(
                            e.Id.IntegerValue,
                            e.Category.Name,
                            elementType != null ? elementType.FamilyName : "-",
                            e.Name);
                    });

                IEnumerable<BrowserItem> worksets = new List<BrowserItem>();
                if (revitDocument.Document.IsWorkshared)
                {
                    worksets = new FilteredWorksetCollector(revitDocument.Document)
                        .OfKind(WorksetKind.UserWorkset)
                        .Select(e =>
                            new BrowserItem(
                                e.Id.IntegerValue,
                                WorksetsStr,
                                "-",
                                e.Name));
                }

                var grids = new FilteredElementCollector(revitDocument.Document)
                    .OfClass(typeof(Grid))
                    .WhereElementIsNotElementType()
                    .Select(e =>
                    {
                        var elementType = (ElementType)revitDocument.Document.GetElement(e.GetTypeId());
                        return new BrowserItem(
                            e.Id.IntegerValue,
                            e.Category.Name,
                            elementType != null ? elementType.FamilyName : "-",
                            e.Name);
                    });

                var levels = new FilteredElementCollector(revitDocument.Document)
                    .OfClass(typeof(Level))
                    .WhereElementIsNotElementType()
                    .Select(e =>
                    {
                        var elementType = (ElementType)revitDocument.Document.GetElement(e.GetTypeId());
                        return new BrowserItem(
                            e.Id.IntegerValue,
                            e.Category.Name,
                            elementType != null ? elementType.FamilyName : "-",
                            e.Name);
                    });

                var parameters = new List<BrowserItem>();
                if (!revitDocument.Document.IsFamilyDocument)
                {
                    var definitionBindingMapIterator = revitDocument.Document.ParameterBindings.ForwardIterator();
                    definitionBindingMapIterator.Reset();
                    while (definitionBindingMapIterator.MoveNext())
                    {
                        var key = (InternalDefinition)definitionBindingMapIterator.Key;
                        var element = revitDocument.Document.GetElement(key.Id);
                        var elementType = (ElementType)revitDocument.Document.GetElement(element.GetTypeId());
                        parameters.Add(new BrowserItem(
                                element.Id.IntegerValue,
                                element.Category?.Name ?? _specialTypeNames[element.GetType().Name],
                                elementType != null ? elementType.FamilyName : "-",
                                element.Name));
                    }
                }

                var categoriesList = new List<BrowserItem>();
                var categories = revitDocument.Document.Settings.Categories;
                foreach (Category category in categories)
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
                            categoriesList.Add(new BrowserItem(element.Id.IntegerValue, "Категории", "-", element.Name));
                        }
                    }
                }

                allElements.AddRange(elementTypes);
                allElements.AddRange(elementsWithoutCategories);
                allElements.AddRange(searchedElementsWithCategory);
                allElements.AddRange(searchedElementsWithoutCategory);
                allElements.AddRange(specialCategories);
                allElements.AddRange(viewTemplates);
                allElements.AddRange(views);
                allElements.AddRange(elevationMarkers);
                allElements.AddRange(viewports);
                allElements.AddRange(worksets);
                allElements.AddRange(grids);
                allElements.AddRange(levels);
                allElements.AddRange(parameters);
                allElements.AddRange(categoriesList);

                var groupedElements = allElements
                    .GroupBy(e => e.CategoryName);

                foreach (var group in groupedElements)
                {
                    try
                    {
                        var elementsList = group.Distinct(new BrowserItemEqualityComparer()).ToList();

                        groups.Add(new BrowserItemsGroup(group.Key, elementsList));
                    }
                    catch
                    {
                        Logger.Instance.Add(string.Format(ParamError, group.Key));
                    }
                }

                groups = groups.OrderBy(group => group.Name).ToList();

                Logger.Instance.Add(string.Format(ParamFinish, DateTime.Now.ToLocalTime()));
                Logger.Instance.Add("---------");

                return new BrowserGeneralGroup("Все группы", groups);
            }
            catch (Exception e)
            {
            }

            return new BrowserGeneralGroup("Все группы", new List<BrowserItemsGroup>());
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
            List<IRevitElement> elements,
            CopyingOptions copyingOptions)
        {
            Logger.Instance.Add(string.Format(CopyStart, DateTime.Now.ToLocalTime(), documentFrom.Title, string.Join(", ", documentsTo)));
            Logger.Instance.Add(string.Format(CopyingOptions, GetCopyingOptionsName(copyingOptions)));

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
                try
                {
                    var elementId = new ElementId(element.Id);
                    var revitElement = documentFrom.Document.GetElement(elementId);

                    foreach (var documentTo in documentsTo)
                    {
                        using (var transaction = new Transaction(documentTo.Document, "Копирование элементов"))
                        {
                            transaction.Start();

                            try
                            {
                                if (revitElement.GetType() == typeof(Workset) && documentTo.Document.IsWorkshared)
                                {
                                    Workset.Create(documentTo.Document, revitElement.Name);
                                }
                                else if (revitElement.GetType() == typeof(View))
                                {
                                    // TODO: Logger event
                                }

                                // TODO: General copying case
                                ElementTransformUtils.CopyElement(documentTo.Document, elementId, null);
                            }
                            catch (Exception e)
                            {
                                Logger.Instance.Add(string.Format(CopyElementError, DateTime.Now.ToLocalTime(), "", "",
                                    e.Message));
                            }

                            transaction.Commit();
                        }
                    }
                }
                catch (Exception e)
                {
                    Logger.Instance.Add(string.Format(CopyElementError, DateTime.Now.ToLocalTime(), "", "",
                        e.Message));
                }
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
