namespace mprCopyElementsToOpenDocuments.Models
{
    using Interfaces;
    using ModPlusAPI.Mvvm;

    /// <inheritdoc cref="IBrowserItem"/>
    public class BrowserItem : VmBase, IBrowserItem
    {
        /// <inheritdoc/>
        public bool Checked { get; set; }
    }
}
