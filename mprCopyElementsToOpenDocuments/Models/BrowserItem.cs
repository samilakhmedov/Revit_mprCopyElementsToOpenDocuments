namespace mprCopyElementsToOpenDocuments.Models
{
    using System;
    using Interfaces;
    using ModPlusAPI.Mvvm;

    /// <summary>
    /// Элемент в браузере
    /// </summary>
    public class BrowserItem : VmBase, IBrowserItem
    {
        private bool? _checked = false;

        /// <summary>
        /// Создает экземпляр класса <see cref="BrowserItem"/>
        /// </summary>
        /// <param name="name">Имя элемента</param>
        public BrowserItem(string name)
        {
            Name = name;
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

        /// <summary>
        /// Имя элемента
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Метод вызова события
        /// </summary>
        protected virtual void OnSelectionChanged()
        {
            SelectionChanged?.Invoke(this, EventArgs.Empty);
        }
    }
}
