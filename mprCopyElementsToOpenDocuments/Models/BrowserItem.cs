namespace mprCopyElementsToOpenDocuments.Models
{
    using System;
    using Interfaces;
    using ModPlusAPI.Mvvm;

    /// <summary>
    /// Элемент в браузере
    /// </summary>
    public class BrowserItem : VmBase, IBrowserItem, IRevitEntity
    {
        private bool? _checked = false;

        /// <summary>
        /// Создает экземпляр класса <see cref="BrowserItem"/>
        /// </summary>
        /// <param name="name">Имя элемента</param>
        /// <param name="id">Идентификатор элемента</param>
        public BrowserItem(string name, int id)
        {
            Name = name;
            Id = id;
        }

        /// <summary>
        /// Событие выделения элемента
        /// </summary>
        public event EventHandler SelectionChanged;

        /// <summary>
        /// Имя элемента
        /// </summary>
        public string Name { get; }

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

        /// <summary>
        /// Метод вызова события
        /// </summary>
        protected virtual void OnSelectionChanged()
        {
            SelectionChanged?.Invoke(this, EventArgs.Empty);
        }
    }
}
