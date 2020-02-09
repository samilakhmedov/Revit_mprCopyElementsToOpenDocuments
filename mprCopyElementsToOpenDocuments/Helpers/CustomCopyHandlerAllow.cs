namespace mprCopyElementsToOpenDocuments.Helpers
{
    using Autodesk.Revit.DB;

    /// <summary>
    /// Обработчик копирования элементов
    /// </summary>
    public class CustomCopyHandlerAllow : IDuplicateTypeNamesHandler
    {
        /// <inheritdoc/>
        public DuplicateTypeAction OnDuplicateTypeNamesFound(DuplicateTypeNamesHandlerArgs args)
        {
            return DuplicateTypeAction.UseDestinationTypes;
        }
    }
}
