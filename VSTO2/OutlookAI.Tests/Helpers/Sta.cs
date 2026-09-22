using System;
using System.Runtime.ExceptionServices;
using System.Threading;

namespace OutlookAI.Tests.Helpers
{
    /// <summary>Runs WinForms code on a dedicated STA thread and rethrows its failure.</summary>
    internal static class Sta
    {
        public static void Run(Action action)
        {
            Exception failure = null;
            var thread = new Thread(() =>
            {
                try { action(); }
                catch (Exception ex) { failure = ex; }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();
            if (failure != null) ExceptionDispatchInfo.Capture(failure).Throw();
        }
    }
}
