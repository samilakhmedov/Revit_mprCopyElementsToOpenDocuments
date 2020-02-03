namespace mprCopyElementsToOpenDocuments.Helpers
{
    using System;
    using System.Linq;
    using Autodesk.Revit.DB;
    using Autodesk.Revit.UI;

    /// <inheritdoc />
    public class RevitExternalEventHandler : IExternalEventHandler
    {
        private readonly ExternalEvent _exEvent;
        private Action _doAction;
        private Document _doc;
        private bool _skipFailures;

        /// <summary>
        /// Конструктор
        /// </summary>
        public RevitExternalEventHandler()
        {
            _exEvent = ExternalEvent.Create(this);
        }

        /// <summary>
        /// Экземпляр обработчика
        /// </summary>
        public static RevitExternalEventHandler Instance { get; private set; }

        /// <summary>
        /// Инициализация обработчика
        /// </summary>
        public static void Init()
        {
            Instance = new RevitExternalEventHandler();
        }

        /// <inheritdoc/>
        public void Execute(UIApplication app)
        {
            try
            {
                if (_doAction == null)
                    return;

                if (_doc == null)
                    _doc = app.ActiveUIDocument.Document;

                if (_skipFailures)
                    app.Application.FailuresProcessing += Application_FailuresProcessing;

                _doAction();

                if (_skipFailures)
                    app.Application.FailuresProcessing -= Application_FailuresProcessing;
            }
            catch (Exception)
            {
                if (_skipFailures)
                    app.Application.FailuresProcessing -= Application_FailuresProcessing;
                throw;
            }
        }

        /// <inheritdoc/>
        public string GetName()
        {
            return "PluginEvent";
        }

        /// <summary>
        /// Запускает выполнение Revit-ом переданного Action
        /// </summary>
        /// <param name="doAction">Action, который будет выполнен</param>
        /// <param name="skipFailures">Пропуск предупреждений</param>
        /// <param name="doc">Текущий документ</param>
        public void Run(Action doAction, bool skipFailures, Document doc = null)
        {
            _doAction = doAction;
            _skipFailures = skipFailures;
            _doc = doc;
            _exEvent.Raise();
        }

        private static void Application_FailuresProcessing(object sender,
            Autodesk.Revit.DB.Events.FailuresProcessingEventArgs e)
        {
            // Получаем все предупреждения
            var failList = e.GetFailuresAccessor().GetFailureMessages();
            if (!failList.Any())
                return;

            // Пропускаем все ошибки
            e.GetFailuresAccessor().DeleteAllWarnings();
            e.SetProcessingResult(FailureProcessingResult.Continue);
        }
    }
}
