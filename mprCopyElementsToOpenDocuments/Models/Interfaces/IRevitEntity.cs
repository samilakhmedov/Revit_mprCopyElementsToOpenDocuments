namespace mprCopyElementsToOpenDocuments.Models.Interfaces
{
    /// <summary>
    /// Интерфейс объекта Revit
    /// </summary>
    public interface IRevitEntity
    {
        /// <summary>
        /// Идентификатор элемента или категории элементов Revit
        /// </summary>
        int Id { get; }
    }
}
