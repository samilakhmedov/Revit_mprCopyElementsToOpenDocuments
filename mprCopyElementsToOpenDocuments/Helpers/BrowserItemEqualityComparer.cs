namespace mprCopyElementsToOpenDocuments.Helpers
{
    using System.Collections.Generic;
    using Models;

    /// <summary>
    /// Производит сравнение элементов Revit
    /// </summary>
    public class BrowserItemEqualityComparer : IEqualityComparer<BrowserItem>
    {
        /// <summary>
        /// Сравнивает элементы Revit
        /// </summary>
        /// <param name="x">Первый элемент сравнения</param>
        /// <param name="y">Второй элемент сравнения</param>
        public bool Equals(BrowserItem x, BrowserItem y)
        {
            if (x is null)
                return false;

            if (y is null)
                return false;

            return ReferenceEquals(x, y) || (string.Equals(x.Name, y.Name) && string.Equals(x.Name, y.Name));
        }

        /// <summary>
        /// Возвращает хэш код элемента
        /// </summary>
        /// <param name="obj">Элемент Revit</param>
        /// <returns>Хэш код элемента</returns>
        public int GetHashCode(BrowserItem obj)
        {
            return obj.Name.GetHashCode() ^ obj.Id.GetHashCode();
        }
    }
}