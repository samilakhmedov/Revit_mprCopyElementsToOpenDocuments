namespace mprCopyElementsToOpenDocuments.ViewModels
{
    using System;
    using System.Windows.Input;
    using Helpers;
    using ModPlusAPI.Mvvm;

    /// <summary>
    /// Модель представления главного окна плагина
    /// </summary>
    public class LoggerViewModel : VmBase
    {
        private readonly string _langItem = ModPlusConnector.Instance.Name;

        /// <summary>
        /// Команда открытия лога в блокноте
        /// </summary>
        public ICommand OpenInNotepadCommand => new RelayCommandWithoutParameter(OpenInNotepad);

        /// <summary>
        /// Текущее состояние журнала событий
        /// </summary>
        public string CurrentLogState => string.Join(Environment.NewLine, Logger.Instance);

        /// <summary>
        /// Сохраняет данные журнала в файл
        /// </summary>
        private void OpenInNotepad()
        {
            ModPlusAPI.IO.String.ShowTextWithNotepad(CurrentLogState, ModPlusAPI.Language.GetItem(_langItem, "h5"));
        }
    }
}
