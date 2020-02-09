namespace mprCopyElementsToOpenDocuments.ViewModels
{
    using System;
    using System.IO;
    using System.Windows.Input;
    using Helpers;
    using Microsoft.Win32;
    using ModPlusAPI.Mvvm;

    /// <summary>
    /// Модель представления главного окна плагина
    /// </summary>
    public class LoggerViewModel : VmBase
    {
        private const string LangItem = "mprCopyElementsToOpenDocuments";

        /// <summary>
        /// Команда обработки выбранного документа Revit
        /// </summary>
        public ICommand SaveResultsCommand =>
            new RelayCommandWithoutParameter(SaveResults);

        /// <summary>
        /// Текущее состояние журнала событий
        /// </summary>
        public string CurrentLogState => string.Join(Environment.NewLine, Logger.Instance);

        /// <summary>
        /// Сохраняет данные журнала в файл
        /// </summary>
        private void SaveResults()
        {
            var saveFileDialog = new SaveFileDialog
            {
                Filter = "Text file (*.txt)|*.txt",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                FileName = ModPlusAPI.Language.GetItem(LangItem, "h5")
            };

            if (saveFileDialog.ShowDialog() == true)
                File.WriteAllText(saveFileDialog.FileName, CurrentLogState);
        }
    }
}
