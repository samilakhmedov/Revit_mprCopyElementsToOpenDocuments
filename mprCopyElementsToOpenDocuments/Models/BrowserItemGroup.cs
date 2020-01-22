namespace mprCopyElementsToOpenDocuments.Models
{
    using Interfaces;
    using ModPlusAPI.Mvvm;

    /// <inheritdoc cref="IBrowserItem"/>
    public class BrowserItemGroup : VmBase, IBrowserItem
    {
        /// <inheritdoc/>
        public bool Checked { get; set; }
    }
}
