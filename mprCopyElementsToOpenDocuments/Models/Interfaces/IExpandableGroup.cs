namespace mprCopyElementsToOpenDocuments.Models.Interfaces
{
    /// <summary>
    /// Интерфейс группы элементов, поддерживающей развертывание списка
    /// </summary>
    public interface IExpandableGroup
    {
        /// <summary>
        /// Показывает, развернута ли группа
        /// </summary>
        bool IsExpanded { get; set; }
    }
}
