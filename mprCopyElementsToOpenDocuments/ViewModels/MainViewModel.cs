namespace mprCopyElementsToOpenDocuments.ViewModels
{
    using Autodesk.Revit.UI;
    using ModPlusAPI.Mvvm;
    using Views;

    /// <summary>
    /// Модель представления главного окна плагина
    /// </summary>
    public class MainViewModel : VmBase
    {
        /// <summary>
        /// Создает экземпляр класса <see cref="MainViewModel"/>
        /// <param name="uiApplication">Активная сессия пользовательского интерфейса Revit</param>
        /// <param name="mainView">Главное окно плагина</param>
        /// </summary>
        public MainViewModel(UIApplication uiApplication, MainView mainView)
        {
        }
    }
}
