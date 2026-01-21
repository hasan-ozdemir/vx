using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Shell;

namespace VxRemoteControl
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [Guid(PackageGuidString)]
    [InstalledProductRegistration("VX Remote Control", "Remote control server for Visual Studio 2022", "0.1")]
    [ProvideAutoLoad(UIContextGuids.ShellInitialized, PackageAutoLoadFlags.BackgroundLoad)]
    public sealed class VxRemoteControlPackage : AsyncPackage
    {
        public const string PackageGuidString = "d4728ef3-d6a6-480a-a4e2-aa5d975c2575";

        private RemoteControlServer _server;
        private VsOutputLogger _logger;

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await base.InitializeAsync(cancellationToken, progress);

            _logger = await VsOutputLogger.CreateAsync(this, cancellationToken);
            _server = new RemoteControlServer(this, _logger);
            await _server.StartAsync(cancellationToken);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _server?.Dispose();
                _logger?.Dispose();
            }

            base.Dispose(disposing);
        }
    }
}
