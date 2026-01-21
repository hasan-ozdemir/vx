using System;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace VxRemoteControl
{
    internal sealed class VsOutputLogger : IDisposable
    {
        private static readonly Guid PaneGuid = new Guid("a8ce5d88-2df7-4879-93c3-0bfa9b7c6424");
        private readonly IVsOutputWindowPane _pane;

        private VsOutputLogger(IVsOutputWindowPane pane)
        {
            _pane = pane;
        }

        public static async Task<VsOutputLogger> CreateAsync(AsyncPackage package, CancellationToken token)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(token);

            var outputWindow = await package.GetServiceAsync(typeof(SVsOutputWindow)) as IVsOutputWindow;
            if (outputWindow == null)
            {
                throw new InvalidOperationException("SVsOutputWindow not available.");
            }

            outputWindow.CreatePane(ref PaneGuid, "VX Remote Control", 1, 0);
            outputWindow.GetPane(ref PaneGuid, out var pane);
            return new VsOutputLogger(pane);
        }

        public void Log(string message)
        {
            var timestamp = DateTime.Now.ToString("HH:mm:ss");
            _pane?.OutputStringThreadSafe($"[{timestamp}] {message}{Environment.NewLine}");
        }

        public void Dispose()
        {
        }
    }
}
