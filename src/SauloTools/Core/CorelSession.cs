using System;
using Corel.Interop.VGCore;

namespace SauloTools.Core
{
    public enum StatusLevel
    {
        Info,
        Success,
        Warning,
        Error
    }

    public class ActionResult
    {
        public bool Ok { get; set; }
        public StatusLevel Level { get; set; }
        public string Message { get; set; }

        public static ActionResult Success(string msg) => new ActionResult { Ok = true, Level = StatusLevel.Success, Message = msg };
        public static ActionResult Info(string msg) => new ActionResult { Ok = true, Level = StatusLevel.Info, Message = msg };
        public static ActionResult Warn(string msg) => new ActionResult { Ok = false, Level = StatusLevel.Warning, Message = msg };
        public static ActionResult Fail(string msg) => new ActionResult { Ok = false, Level = StatusLevel.Error, Message = msg };
    }

    public static class CorelSession
    {
        /// <summary>
        /// Runs a CorelDRAW document operation inside a single Undo command group,
        /// safely temporarily switching document units and enabling Optimization.
        /// </summary>
        public static ActionResult RunUndoable(
            Application app,
            string commandGroupName,
            cdrUnit? workingUnit,
            bool useOptimization,
            Func<Document, ActionResult> work)
        {
            if (app == null)
                return ActionResult.Fail("CorelDRAW não conectado.");

            Document doc = app.ActiveDocument;
            if (doc == null)
                return ActionResult.Warn("Abra um documento no CorelDRAW primeiro.");

            cdrUnit originalUnit = doc.Unit;
            cdrReferencePoint originalRef = doc.ReferencePoint;
            bool originalOpt = app.Optimization;
            bool groupStarted = false;

            try
            {
                if (workingUnit.HasValue)
                    doc.Unit = workingUnit.Value;

                if (useOptimization)
                    app.Optimization = true;

                doc.BeginCommandGroup(commandGroupName);
                groupStarted = true;

                return work(doc);
            }
            catch (Exception ex)
            {
                return ActionResult.Fail($"Erro em '{commandGroupName}': {ex.Message}");
            }
            finally
            {
                if (groupStarted)
                {
                    try { doc.EndCommandGroup(); } catch { }
                }

                try { doc.Unit = originalUnit; } catch { }
                try { doc.ReferencePoint = originalRef; } catch { }
                try { app.Optimization = originalOpt; } catch { }
                try { app.ActiveWindow?.Refresh(); } catch { }
                try { app.Refresh(); } catch { }
            }
        }
    }
}
