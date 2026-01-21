using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace VxRemoteControl
{
    internal sealed class RemoteControlServer : IDisposable
    {
        private const int DefaultPort = 53100;
        private static readonly string[] DefaultPrefixes =
        {
            $"http://127.0.0.1:{DefaultPort}/",
            $"http://localhost:{DefaultPort}/"
        };

        private readonly AsyncPackage _package;
        private readonly VsOutputLogger _logger;
        private readonly HttpListener _listener;
        private readonly CancellationTokenSource _cts;
        private readonly DateTime _startedUtc;
        private readonly string[] _prefixes;

        public RemoteControlServer(AsyncPackage package, VsOutputLogger logger)
        {
            _package = package ?? throw new ArgumentNullException(nameof(package));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _listener = new HttpListener();
            _cts = new CancellationTokenSource();
            _startedUtc = DateTime.UtcNow;
            _prefixes = GetPrefixes();
        }

        public async Task StartAsync(CancellationToken token)
        {
            foreach (var prefix in _prefixes)
            {
                _listener.Prefixes.Add(prefix);
            }

            try
            {
                _listener.Start();
            }
            catch (HttpListenerException ex)
            {
                _logger.Log($"HttpListener start failed: {ex.Message}");
                _logger.Log("If access is denied, reserve the URL: netsh http add urlacl url=http://127.0.0.1:53100/ user=USERNAME");
                throw;
            }

            Task.Run(() => ListenLoopAsync(_cts.Token), _cts.Token);
            _logger.Log($"Listening on {string.Join(", ", _prefixes)}");
            await Task.CompletedTask;
        }

        public void Dispose()
        {
            try
            {
                _cts.Cancel();
                _listener.Stop();
                _listener.Close();
            }
            catch
            {
                // Ignore shutdown errors.
            }
        }

        private async Task ListenLoopAsync(CancellationToken token)
        {
            while (!token.IsCancellationRequested)
            {
                HttpListenerContext context;
                try
                {
                    context = await _listener.GetContextAsync();
                }
                catch (HttpListenerException)
                {
                    break;
                }
                catch (ObjectDisposedException)
                {
                    break;
                }

                _ = Task.Run(() => HandleContextAsync(context, token), token);
            }
        }

        private async Task HandleContextAsync(HttpListenerContext context, CancellationToken token)
        {
            try
            {
                var request = context.Request;
                var method = request.HttpMethod.ToUpperInvariant();
                var path = request.Url.AbsolutePath.TrimEnd('/');
                if (string.IsNullOrEmpty(path))
                {
                    path = "/";
                }

                object data;

                if (method == "GET" && path == "/api/status")
                {
                    data = BuildStatus();
                }
                else if (method == "GET" && path == "/api/help")
                {
                    data = BuildHelp();
                }
                else if (method == "POST" && path == "/api/dte/execute")
                {
                    data = await HandleExecuteCommandAsync(request, token);
                }
                else if (method == "GET" && path == "/api/commands")
                {
                    data = await HandleListCommandsAsync(token);
                }
                else if (method == "GET" && path == "/api/solution")
                {
                    data = await HandleSolutionInfoAsync(token);
                }
                else if (method == "POST" && path == "/api/solution/open")
                {
                    data = await HandleSolutionOpenAsync(request, token);
                }
                else if (method == "POST" && path == "/api/solution/close")
                {
                    data = await HandleSolutionCloseAsync(request, token);
                }
                else if (method == "POST" && path == "/api/solution/build")
                {
                    data = await HandleSolutionBuildAsync(request, token);
                }
                else if (method == "POST" && path == "/api/solution/clean")
                {
                    data = await HandleSolutionCleanAsync(request, token);
                }
                else if (method == "GET" && path == "/api/solution/projects")
                {
                    data = await HandleSolutionProjectsAsync(token);
                }
                else if (method == "GET" && path == "/api/documents/active")
                {
                    data = await HandleActiveDocumentAsync(token);
                }
                else if (method == "POST" && path == "/api/documents/open")
                {
                    data = await HandleOpenDocumentAsync(request, token);
                }
                else if (method == "POST" && path == "/api/documents/save")
                {
                    data = await HandleSaveDocumentAsync(request, token);
                }
                else if (method == "POST" && path == "/api/documents/get-text")
                {
                    data = await HandleGetDocumentTextAsync(request, token);
                }
                else if (method == "POST" && path == "/api/documents/set-text")
                {
                    data = await HandleSetDocumentTextAsync(request, token);
                }
                else if (method == "POST" && path == "/api/debug/start")
                {
                    data = await HandleDebugCommandAsync("Debug.Start", token);
                }
                else if (method == "POST" && path == "/api/debug/stop")
                {
                    data = await HandleDebugCommandAsync("Debug.StopDebugging", token);
                }
                else if (method == "POST" && path == "/api/debug/break")
                {
                    data = await HandleDebugCommandAsync("Debug.Break", token);
                }
                else if (method == "POST" && path == "/api/server/stop")
                {
                    data = HandleStopServer();
                }
                else
                {
                    await WriteJsonAsync(context.Response, ApiResponse.Fail($"No route for {method} {path}"), 404);
                    return;
                }

                await WriteJsonAsync(context.Response, ApiResponse.Success(data), 200);
            }
            catch (Exception ex)
            {
                _logger.Log($"Request error: {ex}");
                try
                {
                    await WriteJsonAsync(context.Response, ApiResponse.Fail(ex.Message), 500);
                }
                catch
                {
                    // Ignore response failures.
                }
            }
        }

        private object BuildStatus()
        {
            return new
            {
                name = "VX Remote Control",
                version = "0.1",
                pid = Process.GetCurrentProcess().Id,
                uptimeSeconds = (int)(DateTime.UtcNow - _startedUtc).TotalSeconds,
                serverTime = DateTime.Now,
                prefixes = _prefixes
            };
        }

        private object BuildHelp()
        {
            return new
            {
                endpoints = new[]
                {
                    "GET /api/status",
                    "GET /api/help",
                    "POST /api/dte/execute",
                    "GET /api/commands",
                    "GET /api/solution",
                    "POST /api/solution/open",
                    "POST /api/solution/close",
                    "POST /api/solution/build",
                    "POST /api/solution/clean",
                    "GET /api/solution/projects",
                    "GET /api/documents/active",
                    "POST /api/documents/open",
                    "POST /api/documents/save",
                    "POST /api/documents/get-text",
                    "POST /api/documents/set-text",
                    "POST /api/debug/start",
                    "POST /api/debug/stop",
                    "POST /api/debug/break",
                    "POST /api/server/stop"
                }
            };
        }

        private async Task<object> HandleExecuteCommandAsync(HttpListenerRequest request, CancellationToken token)
        {
            var body = await ReadBodyAsync(request);
            var command = body.Value<string>("command");
            if (string.IsNullOrWhiteSpace(command))
            {
                throw new InvalidOperationException("Missing 'command' field.");
            }

            var args = body.Value<string>("args") ?? string.Empty;

            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(token);
            var dte = await GetDteAsync();
            dte.ExecuteCommand(command, args);

            return new { command, args };
        }

        private async Task<object> HandleListCommandsAsync(CancellationToken token)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(token);
            var dte = await GetDteAsync();
            var commands = new List<string>();
            foreach (Command command in dte.Commands)
            {
                commands.Add(command.Name);
            }

            return new { count = commands.Count, commands };
        }

        private async Task<object> HandleSolutionInfoAsync(CancellationToken token)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(token);
            var dte = await GetDteAsync();
            var solution = dte.Solution;
            var build = solution?.SolutionBuild;
            var activeConfig = build?.ActiveConfiguration;

            return new
            {
                isOpen = solution != null && solution.IsOpen,
                fullName = solution?.FullName,
                activeConfiguration = activeConfig?.Name,
                activePlatform = activeConfig?.PlatformName,
                buildState = build?.BuildState.ToString()
            };
        }

        private async Task<object> HandleSolutionOpenAsync(HttpListenerRequest request, CancellationToken token)
        {
            var body = await ReadBodyAsync(request);
            var path = body.Value<string>("path");
            if (string.IsNullOrWhiteSpace(path))
            {
                throw new InvalidOperationException("Missing 'path' field.");
            }

            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(token);
            var dte = await GetDteAsync();
            dte.Solution.Open(path);

            return new { opened = path };
        }

        private async Task<object> HandleSolutionCloseAsync(HttpListenerRequest request, CancellationToken token)
        {
            var body = await ReadBodyAsync(request);
            var saveAll = body.Value<bool?>("saveAll") ?? false;

            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(token);
            var dte = await GetDteAsync();
            if (dte.Solution != null && dte.Solution.IsOpen)
            {
                dte.Solution.Close(saveAll);
            }

            return new { closed = true, saveAll };
        }

        private async Task<object> HandleSolutionBuildAsync(HttpListenerRequest request, CancellationToken token)
        {
            var body = await ReadBodyAsync(request);
            var configuration = body.Value<string>("configuration");

            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(token);
            var dte = await GetDteAsync();
            ActivateSolutionConfiguration(dte, configuration);

            dte.Solution.SolutionBuild.Build(true);
            var activeConfig = dte.Solution.SolutionBuild.ActiveConfiguration;

            return new
            {
                buildRequested = true,
                activeConfiguration = activeConfig?.Name,
                activePlatform = activeConfig?.PlatformName
            };
        }

        private async Task<object> HandleSolutionCleanAsync(HttpListenerRequest request, CancellationToken token)
        {
            var body = await ReadBodyAsync(request);
            var configuration = body.Value<string>("configuration");

            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(token);
            var dte = await GetDteAsync();
            ActivateSolutionConfiguration(dte, configuration);

            dte.Solution.SolutionBuild.Clean(true);
            var activeConfig = dte.Solution.SolutionBuild.ActiveConfiguration;

            return new
            {
                cleanRequested = true,
                activeConfiguration = activeConfig?.Name,
                activePlatform = activeConfig?.PlatformName
            };
        }

        private async Task<object> HandleSolutionProjectsAsync(CancellationToken token)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(token);
            var dte = await GetDteAsync();
            var list = new List<object>();
            foreach (var project in DteHelpers.EnumerateProjects(dte.Solution))
            {
                list.Add(new
                {
                    name = project.Name,
                    fullName = project.FullName,
                    kind = project.Kind
                });
            }

            return new { count = list.Count, projects = list };
        }

        private async Task<object> HandleActiveDocumentAsync(CancellationToken token)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(token);
            var dte = await GetDteAsync();
            var doc = dte.ActiveDocument;
            var text = DteHelpers.GetDocumentText(doc);
            return new
            {
                name = doc?.Name,
                fullName = doc?.FullName,
                kind = doc?.Kind,
                text
            };
        }

        private async Task<object> HandleOpenDocumentAsync(HttpListenerRequest request, CancellationToken token)
        {
            var body = await ReadBodyAsync(request);
            var path = body.Value<string>("path");
            if (string.IsNullOrWhiteSpace(path))
            {
                throw new InvalidOperationException("Missing 'path' field.");
            }

            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(token);
            var dte = await GetDteAsync();
            var window = dte.ItemOperations.OpenFile(path);
            window?.Activate();

            return new { opened = path };
        }

        private async Task<object> HandleSaveDocumentAsync(HttpListenerRequest request, CancellationToken token)
        {
            var body = await ReadBodyAsync(request);
            var path = body.Value<string>("path");

            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(token);
            var dte = await GetDteAsync();
            var document = FindDocument(dte, path) ?? dte.ActiveDocument;
            document?.Save();

            return new { saved = document?.FullName };
        }

        private async Task<object> HandleGetDocumentTextAsync(HttpListenerRequest request, CancellationToken token)
        {
            var body = await ReadBodyAsync(request);
            var path = body.Value<string>("path");

            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(token);
            var dte = await GetDteAsync();
            var document = FindDocument(dte, path) ?? dte.ActiveDocument;
            var text = DteHelpers.GetDocumentText(document);

            return new
            {
                name = document?.Name,
                fullName = document?.FullName,
                text
            };
        }

        private async Task<object> HandleSetDocumentTextAsync(HttpListenerRequest request, CancellationToken token)
        {
            var body = await ReadBodyAsync(request);
            var path = body.Value<string>("path");
            var text = body.Value<string>("text") ?? string.Empty;

            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(token);
            var dte = await GetDteAsync();
            var document = FindDocument(dte, path);
            if (document == null && !string.IsNullOrWhiteSpace(path))
            {
                document = dte.ItemOperations.OpenFile(path)?.Document;
            }

            document ??= dte.ActiveDocument;

            var updated = DteHelpers.SetDocumentText(document, text);
            return new
            {
                updated,
                fullName = document?.FullName
            };
        }

        private async Task<object> HandleDebugCommandAsync(string command, CancellationToken token)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(token);
            var dte = await GetDteAsync();
            dte.ExecuteCommand(command);

            return new { command };
        }

        private object HandleStopServer()
        {
            Dispose();
            return new { stopped = true };
        }

        private static async Task WriteJsonAsync(HttpListenerResponse response, ApiResponse payload, int statusCode)
        {
            var json = JsonConvert.SerializeObject(payload, Formatting.None);
            var buffer = Encoding.UTF8.GetBytes(json);

            response.StatusCode = statusCode;
            response.ContentType = "application/json";
            response.ContentEncoding = Encoding.UTF8;
            response.ContentLength64 = buffer.Length;

            await response.OutputStream.WriteAsync(buffer, 0, buffer.Length);
            response.OutputStream.Close();
        }

        private static async Task<JObject> ReadBodyAsync(HttpListenerRequest request)
        {
            if (request == null || !request.HasEntityBody)
            {
                return new JObject();
            }

            using (var reader = new StreamReader(request.InputStream, request.ContentEncoding ?? Encoding.UTF8))
            {
                var body = await reader.ReadToEndAsync();
                if (string.IsNullOrWhiteSpace(body))
                {
                    return new JObject();
                }

                return JObject.Parse(body);
            }
        }

        private async Task<DTE2> GetDteAsync()
        {
            var service = await _package.GetServiceAsync(typeof(DTE));
            var dte = service as DTE2;
            if (dte == null)
            {
                throw new InvalidOperationException("DTE service not available.");
            }

            return dte;
        }

        private static string[] GetPrefixes()
        {
            var portValue = Environment.GetEnvironmentVariable("VX_REMOTE_PORT");
            if (int.TryParse(portValue, out var port) && port > 0 && port < 65536)
            {
                return new[]
                {
                    $"http://127.0.0.1:{port}/",
                    $"http://localhost:{port}/"
                };
            }

            return DefaultPrefixes;
        }

        private static void ActivateSolutionConfiguration(DTE2 dte, string configuration)
        {
            if (dte?.Solution?.SolutionBuild == null)
            {
                return;
            }

            if (string.IsNullOrWhiteSpace(configuration))
            {
                return;
            }

            foreach (SolutionConfiguration config in dte.Solution.SolutionBuild.SolutionConfigurations)
            {
                var fullName = $"{config.Name}|{config.PlatformName}";
                if (string.Equals(config.Name, configuration, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(fullName, configuration, StringComparison.OrdinalIgnoreCase))
                {
                    config.Activate();
                    break;
                }
            }
        }

        private static Document FindDocument(DTE2 dte, string path)
        {
            if (dte == null || string.IsNullOrWhiteSpace(path))
            {
                return null;
            }

            foreach (Document doc in dte.Documents)
            {
                if (string.Equals(doc.FullName, path, StringComparison.OrdinalIgnoreCase))
                {
                    return doc;
                }
            }

            return null;
        }
    }
}
