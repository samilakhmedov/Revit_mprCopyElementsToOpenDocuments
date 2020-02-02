namespace mprCopyElementsToOpenDocuments.Models.Interfaces
{
    /// <summary>
    /// Интерфейс объекта Revit
    /// </summary>
    public interface IRevitElement
    {
        /// <summary>
        /// Идентификатор элемента или категории элементов Revit
        /// </summary>
        int Id { get; }

        /// <summary>
        /// Имя категории Revit
        /// </summary>
        string CategoryName { get; }

        /// <summary>
        /// Имя семейства Revit
        /// </summary>
        string FamilyName { get; }

        /// <summary>
        /// Является ли элемент типом
        /// </summary>
        bool IsType { get; }
    }
}
