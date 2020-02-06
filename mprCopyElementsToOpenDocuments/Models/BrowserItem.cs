namespace mprCopyElementsToOpenDocuments.Models
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Linq;
    using Interfaces;
    using ModPlusAPI.Mvvm;

    /// <summary>
    /// Элемент в браузере
    /// </summary>
    public class BrowserItem : VmBase, IBrowserItem, IExpandableGroup, IRevitElement
    {
        private bool? _checked = false;
        private bool _isExpanded;
        private ObservableCollection<BrowserItem> _items = new ObservableCollection<BrowserItem>();

        /// <summary>
        /// Создает экземпляр класса <see cref="BrowserItem"/>
        /// </summary>
        /// <param name="id">Идентификатор элемента</param>
        /// <param name="categoryName">Имя категории</param>
        /// <param name="familyName">Имя семейства</param>
        /// <param name="name">Имя элемента</param>
        public BrowserItem(int id, string categoryName, string familyName, string name)
        {
            Id = id;
            CategoryName = categoryName;
            FamilyName = familyName;
            Name = name;
        }

        /// <summary>
        /// Создает экземпляр класса <see cref="BrowserItem"/>
        /// </summary>
        /// <param name="name">Имя группы</param>
        /// <param name="items">Список элементов группы</param>
        public BrowserItem(string name, List<BrowserItem> items)
        {
            Name = name;

            items.ForEach(item =>
            {
                item.SelectionChanged += OnItemSelectionChanged;

                if (item.Items.Any())
                {
                    foreach (var browserItem in item.Items)
                    {
                        browserItem.SelectionChanged += OnItemSelectionChanged;
                    }
                }

                _items.Add(item);
            });
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

                if (Items.Any())
                {
                    foreach (var item in Items)
                    {
                        item.Checked = value;
                    }
                }

                OnPropertyChanged();
                OnSelectionChanged();
            }
        }

        /// <inheritdoc />
        public bool IsExpanded
        {
            get => _isExpanded;
            set
            {
                _isExpanded = value;
                OnPropertyChanged();
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

        /// <summary>
        /// Список элементов группы
        /// </summary>
        public ObservableCollection<BrowserItem> Items
        {
            get => _items;
            set
            {
                _items = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Метод обработки выделения элементов в браузере
        /// </summary>
        private void OnItemSelectionChanged(object sender, EventArgs e)
        {
            OnSelectionChanged();
        }

        /// <summary>
        /// Метод вызова события
        /// </summary>
        protected virtual void OnSelectionChanged()
        {
            SelectionChanged?.Invoke(this, EventArgs.Empty);
        }
    }
}
