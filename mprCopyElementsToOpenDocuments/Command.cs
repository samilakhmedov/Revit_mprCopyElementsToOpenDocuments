namespace mprCopyElementsToOpenDocuments
{
    using System;
    using Autodesk.Revit.Attributes;
    using Autodesk.Revit.DB;
    using Autodesk.Revit.UI;
    using ModPlusAPI;
    using ModPlusAPI.Windows;
    using mprCopySheetsToOpenDocuments;
    using ViewModels;
    using Views;

    /// <inheritdoc />
    [Regeneration(RegenerationOption.Manual)]
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        private MainView _mainView;

        /// <inheritdoc />
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                ////Statistic.SendCommandStarting(ModPlusConnector.Instance);

                if (_mainView == null)
                {
                    ////if (commandData.Application.Application.Documents.Size < 2)
                    ////{
                    ////    // Необходимо открыть не менее двух документов
                    ////    MessageBox.Show(Language.GetItem(ModPlusConnector.Instance.Name, "m1"), MessageBoxIcon.Close);
                    ////    return Result.Cancelled;
                    ////}

                    _mainView = new MainView();
                    var viewModel = new MainViewModel(commandData.Application, _mainView);
                    _mainView.DataContext = viewModel;
                    _mainView.Closed += (sender, args) => _mainView = null;
                    _mainView.Show();

                    return Result.Succeeded;
                }

                _mainView.Activate();
                _mainView.Focus();
                return Result.Succeeded;
            }
            catch (Exception exception)
            {
                ExceptionBox.Show(exception);
                return Result.Failed;
            }
        }
    }
}
