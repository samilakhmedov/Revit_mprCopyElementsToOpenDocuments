namespace mprCopyElementsToOpenDocuments.Helpers
{
    using System.Collections.Generic;

    /// <summary>
    /// Журнал работы приложения
    /// </summary>
    public class Logger
    {
        private static List<string> _logger = null;
        private static readonly object Mutex = new object();

        /// <summary>
        /// Экземпляр списка для ведения журнала работы приложения
        /// </summary>
        public static List<string> Instance
        {
            get
            {
                lock (Mutex)
                {
                    return _logger ?? (_logger = new List<string>());
                }
            }
        }
    }
}
