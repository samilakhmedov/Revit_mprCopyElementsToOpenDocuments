namespace mprCopyElementsToOpenDocuments.Models
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using Interfaces;
    using ModPlusAPI.Mvvm;

    /// <summary>
    /// Общая группа элементов в браузере
    /// </summary>
    public class BrowserGeneralGroup : VmBase, IBrowserItem, IExpandableGroup
    {
        private bool? _checked = false;
        private bool _isExpanded = true;
        private ObservableCollection<BrowserItemsGroup> _groups = new ObservableCollection<BrowserItemsGroup>();

        /// <summary>
        /// Создает экземпляр класса <see cref="BrowserGeneralGroup"/>
        /// </summary>
        /// <param name="name">Имя группы</param>
        /// <param name="groups">Список групп элементов</param>
        public BrowserGeneralGroup(string name, List<BrowserItemsGroup> groups)
        {
            Name = name;

            groups.ForEach(group =>
            {
                group.SelectionChanged += OnGroupSelectionChanged;
                _groups.Add(group);
            });
        }

        /// <summary>
        /// Событие изменения количества выделенных элементов
        /// </summary>
        public event EventHandler SelectionChanged;

        /// <inheritdoc />
        public string Name { get; }

        /// <inheritdoc />
        public bool? Checked
        {
            get => _checked;
            set
            {
                _checked = value;

                foreach (var group in _groups)
                {
                    group.Checked = value;
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

        /// <summary>
        /// Список элементов группы
        /// </summary>
        public ObservableCollection<BrowserItemsGroup> Items
        {
            get => _groups;
            set
            {
                _groups = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Метод обработки выделения элементов в браузере
        /// </summary>
        private void OnGroupSelectionChanged(object sender, EventArgs e)
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
