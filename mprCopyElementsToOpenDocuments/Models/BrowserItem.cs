namespace mprCopyElementsToOpenDocuments.Models
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using Interfaces;
    using ModPlusAPI.Mvvm;

    /// <summary>
    /// Элемент в браузере
    /// </summary>
    public class BrowserItem : VmBase, IBrowserItem, IRevitElement, IEquatable<BrowserItem>
    {
        private bool? _checked = false;

        /// <summary>
        /// Создает экземпляр класса <see cref="BrowserItem"/>
        /// </summary>
        /// <param name="id">Идентификатор элемента</param>
        /// <param name="categoryName">Имя категории</param>
        /// <param name="familyName">Имя семейства</param>
        /// <param name="name">Имя элемента</param>
        /// <param name="isType">Является ли элемент типом</param>
        public BrowserItem(int id, string categoryName, string familyName, string name, bool isType = false)
        {
            Id = id;
            CategoryName = categoryName;
            FamilyName = familyName;
            Name = name;
            IsType = isType;
        }

        /// <summary>
        /// Событие выделения элемента
        /// </summary>
        public event EventHandler SelectionChanged;

        /// <inheritdoc/>
        public bool? Checked
        {
            get => _checked;
            set
            {
                _checked = value;
                OnPropertyChanged();
                OnSelectionChanged();
            }
        }

        /// <inheritdoc/>
        public int Id { get; }

        /// <inheritdoc/>
        public string CategoryName { get; }

        /// <inheritdoc/>
        public string FamilyName { get; }

        /// <inheritdoc/>
        public string Name { get; }

        /// <inheritdoc/>
        public bool IsType { get; }

        /// <summary>
        /// Метод вызова события
        /// </summary>
        protected virtual void OnSelectionChanged()
        {
            SelectionChanged?.Invoke(this, EventArgs.Empty);
        }

        /// <summary>
        /// Сравнивает элементы браузера
        /// </summary>
        /// <param name="other">Элемент сравнения</param>
        public bool Equals(BrowserItem other)
        {
            if (other is null)
                return false;

            if (ReferenceEquals(this, other))
                return true;

            return string.Equals(Name, other.Name);
        }

        public int SaveHashCode<T>(T x)
        {
            if (EqualityComparer<T>.Default.Equals(x, default(T)))
                return 0;

            if (x is IEnumerable<object> xEnumerable)
                return xEnumerable.Aggregate(17, (acc, item) => (acc * 19) + SaveHashCode(item));

            return x.GetHashCode();
        }

        /// <summary>
        /// Возвращает хэш код элемента
        /// </summary>
        /// <param name="obj">Элемент браузера</param>
        /// <returns>Хэш код элемента</returns>
        public int GetHashCode(BrowserItem obj)
        {
            return obj.Name.GetHashCode();
        }
    }
}
