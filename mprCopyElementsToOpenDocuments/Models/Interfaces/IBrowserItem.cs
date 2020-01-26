namespace mprCopyElementsToOpenDocuments.Models.Interfaces
{
    /// <summary>
    /// Интерфейс элемента дерева
    /// </summary>
    public interface IBrowserItem
    {
        /// <summary>
        /// Указывает, отмечен ли элемент в браузере
        /// </summary>
        bool? Checked { get; set; }
    }
}