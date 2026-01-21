using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        if (args.Length == 0 || IsHelp(args[0]))
        {
            PrintHelp();
            return 0;
        }

        if (args[0].Trim().StartsWith("!", StringComparison.Ordinal))
        {
            return RunProjectBuildCommand(args);
        }

        var command = args[0].Trim().ToLowerInvariant();
        switch (command)
        {
            case "info":
                return RunInfo(args.Skip(1).ToArray());
            case "open":
                return RunOpen(args.Skip(1).ToArray());
            case "view":
                return RunView(args.Skip(1).ToArray());
            case "ls":
                return RunList(args.Skip(1).ToArray());
            case "build":
                return RunSolutionBuild("build", args.Skip(1).ToArray());
            case "rebuild":
                return RunSolutionBuild("rebuild", args.Skip(1).ToArray());
            case "clean":
                return RunSolutionBuild("clean", args.Skip(1).ToArray());
            case "deploy":
                return RunDeploy(args.Skip(1).ToArray());
            case "startup":
                return RunStartup(args.Skip(1).ToArray());
            case "startups":
                return RunStartups(args.Skip(1).ToArray());
            default:
                Console.Error.WriteLine($"Unknown command: {args[0]}");
                PrintHelp();
                return 1;
        }
    }

    private static bool IsHelp(string arg)
    {
        if (string.IsNullOrWhiteSpace(arg))
        {
            return true;
        }

        return string.Equals(arg, "-h", StringComparison.OrdinalIgnoreCase)
            || string.Equals(arg, "--help", StringComparison.OrdinalIgnoreCase)
            || string.Equals(arg, "help", StringComparison.OrdinalIgnoreCase)
            || string.Equals(arg, "/?", StringComparison.OrdinalIgnoreCase);
    }

    private static void PrintHelp()
    {
        Console.WriteLine("vx - Visual Studio 2022 local CLI");
        Console.WriteLine();
        Console.WriteLine("Usage:");
        Console.WriteLine("  vx info [--all] [--instance N] [--json]");
        Console.WriteLine("  vx open solution <path>");
        Console.WriteLine("  vx view <filename>");
        Console.WriteLine("  vx view !ProjectName:<filename>");
        Console.WriteLine("  vx ls us | cl | ns | mt:ClassName");
        Console.WriteLine("  vx build [projectPattern]");
        Console.WriteLine("  vx rebuild [projectPattern]");
        Console.WriteLine("  vx clean [projectPattern]");
        Console.WriteLine("  vx deploy !projectPattern");
        Console.WriteLine("  vx startup !projectPattern");
        Console.WriteLine("  vx startups !pattern1;!pattern2");
        Console.WriteLine("  vx !ProjectName:build|rebuild|clean|deploy");
        Console.WriteLine();
        Console.WriteLine("Commands:");
        Console.WriteLine("  info   Show details about running VS2022 instances.");
        Console.WriteLine("  open   Open a solution in Visual Studio 2022.");
        Console.WriteLine("  view   Open a file in the active solution.");
        Console.WriteLine("  ls     List items from the active document.");
        Console.WriteLine("  build  Build the active solution or matching project.");
        Console.WriteLine("  rebuild Rebuild the active solution or matching project.");
        Console.WriteLine("  clean  Clean the active solution or matching project.");
        Console.WriteLine("  deploy Deploy a matching project.");
        Console.WriteLine("  startup Set the startup project.");
        Console.WriteLine("  startups Set multiple startup projects.");
    }

    private static int RunInfo(string[] args)
    {
        var options = InfoOptions.Parse(args);
        var instances = VsRot.GetRunningDteInstances();

        if (instances.Count == 0)
        {
            if (options.Json)
            {
                var empty = new VxInfoSnapshot { InstanceCount = 0 };
                WriteJson(empty);
            }
            else
            {
                Console.WriteLine("VS2022 running: no");
            }

            return 0;
        }

        var activeIndex = VsSelector.GetActiveIndex(instances);
        if (activeIndex < 0)
        {
            activeIndex = 0;
        }

        if (options.InstanceIndex.HasValue && !options.All)
        {
            if (options.InstanceIndex.Value < 0 || options.InstanceIndex.Value >= instances.Count)
            {
                Console.Error.WriteLine($"Instance index {options.InstanceIndex.Value} is out of range.");
                VsRot.ReleaseInstances(instances);
                return 1;
            }
        }

        var snapshots = new List<VsInstanceSnapshot>();
        for (var i = 0; i < instances.Count; i++)
        {
            if (options.All)
            {
                snapshots.Add(SnapshotBuilder.Build(instances[i], i, i == activeIndex));
                continue;
            }

            if (options.InstanceIndex.HasValue)
            {
                if (options.InstanceIndex.Value == i)
                {
                    snapshots.Add(SnapshotBuilder.Build(instances[i], i, i == activeIndex));
                }

                continue;
            }

            if (i == activeIndex)
            {
                snapshots.Add(SnapshotBuilder.Build(instances[i], i, true));
            }
        }

        var info = new VxInfoSnapshot
        {
            InstanceCount = instances.Count,
            Instances = snapshots
        };

        if (options.Json)
        {
            WriteJson(info);
        }
        else
        {
            PrintInfoText(info, activeIndex);
        }

        VsRot.ReleaseInstances(instances);
        return 0;
    }

    private static int RunOpen(string[] args)
    {
        if (args.Length < 2 || !string.Equals(args[0], "solution", StringComparison.OrdinalIgnoreCase))
        {
            Console.Error.WriteLine("Usage: vx open solution <path>");
            return 1;
        }

        var rawPath = string.Join(" ", args.Skip(1)).Trim();
        if (string.IsNullOrWhiteSpace(rawPath))
        {
            Console.Error.WriteLine("Missing solution path.");
            return 1;
        }

        var expandedPath = Environment.ExpandEnvironmentVariables(rawPath);
        var fullPath = Path.GetFullPath(expandedPath);
        if (!File.Exists(fullPath))
        {
            Console.Error.WriteLine($"Solution not found: {fullPath}");
            return 1;
        }

        var instances = VsRot.GetRunningDteInstances();
        try
        {
            OleMessageFilter.Register();
            if (instances.Count > 0)
            {
                var target = VsSelector.GetActiveInstance(instances) ?? instances[0];
                if (target.Dte == null)
                {
                    Console.Error.WriteLine("Failed to access running Visual Studio instance.");
                    return 1;
                }

                OpenInExistingInstance(target.Dte, fullPath);
                return 0;
            }

            OpenInNewInstance(fullPath);
            return 0;
        }
        catch (COMException ex)
        {
            Console.Error.WriteLine($"Visual Studio automation error: {ex.Message}");
            return 1;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unexpected error: {ex.Message}");
            return 1;
        }
        finally
        {
            OleMessageFilter.Revoke();
            VsRot.ReleaseInstances(instances);
        }
    }

    private static int RunView(string[] args)
    {
        if (args.Length == 0)
        {
            Console.Error.WriteLine("Usage: vx view <filename> | vx view !ProjectName:<filename>");
            return 1;
        }

        var rawTarget = string.Join(" ", args).Trim();
        if (string.IsNullOrWhiteSpace(rawTarget))
        {
            Console.Error.WriteLine("Missing file name.");
            return 1;
        }

        string? projectName = null;
        var fileSpec = rawTarget;
        if (rawTarget.StartsWith("!", StringComparison.Ordinal))
        {
            var separator = rawTarget.IndexOf(':');
            if (separator <= 1 || separator == rawTarget.Length - 1)
            {
                Console.Error.WriteLine("Usage: vx view !ProjectName:<filename>");
                return 1;
            }

            projectName = rawTarget.Substring(1, separator - 1).Trim();
            fileSpec = rawTarget.Substring(separator + 1).Trim();
            if (string.IsNullOrWhiteSpace(projectName) || string.IsNullOrWhiteSpace(fileSpec))
            {
                Console.Error.WriteLine("Usage: vx view !ProjectName:<filename>");
                return 1;
            }
        }

        var instances = VsRot.GetRunningDteInstances();
        try
        {
            OleMessageFilter.Register();
            if (instances.Count == 0)
            {
                Console.Error.WriteLine("No running Visual Studio 2022 instance found.");
                return 1;
            }

            var target = VsSelector.GetActiveInstance(instances) ?? instances[0];
            if (target.Dte == null)
            {
                Console.Error.WriteLine("Failed to access running Visual Studio instance.");
                return 1;
            }

            dynamic dte = target.Dte;
            var solution = TryGetValue(() => (dynamic)dte.Solution);
            if (solution == null)
            {
                Console.Error.WriteLine("No solution is open in the active Visual Studio instance.");
                return 1;
            }

            var isOpen = TryGetValue(() => (bool)solution.IsOpen, false);
            if (!isOpen)
            {
                Console.Error.WriteLine("No solution is open in the active Visual Studio instance.");
                return 1;
            }

            dynamic? project = null;
            if (!string.IsNullOrWhiteSpace(projectName))
            {
                project = FindProjectBySelector(solution, projectName);
                if (project == null)
                {
                    Console.Error.WriteLine($"Project not found: {projectName}");
                    return 1;
                }
            }

            var path = project != null ? FindFileInProject(project, fileSpec) : FindFileInSolution(solution, fileSpec);
            if (string.IsNullOrWhiteSpace(path))
            {
                var scope = projectName != null ? $"project '{projectName}'" : "solution";
                Console.Error.WriteLine($"File not found in {scope}: {fileSpec}");
                return 1;
            }

            dte.ItemOperations.OpenFile(path);
            TryActivateMainWindow(dte);
            return 0;
        }
        catch (COMException ex)
        {
            Console.Error.WriteLine($"Visual Studio automation error: {ex.Message}");
            return 1;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unexpected error: {ex.Message}");
            return 1;
        }
        finally
        {
            OleMessageFilter.Revoke();
            VsRot.ReleaseInstances(instances);
        }
    }

    private static int RunList(string[] args)
    {
        if (args.Length == 0)
        {
            Console.Error.WriteLine("Usage: vx ls us | cl | ns | mt:ClassName");
            return 1;
        }

        var rawTarget = string.Join(" ", args).Trim();
        if (string.IsNullOrWhiteSpace(rawTarget))
        {
            Console.Error.WriteLine("Usage: vx ls us | cl | ns | mt:ClassName");
            return 1;
        }

        var query = rawTarget.Trim();
        var mode = query;
        string? className = null;
        if (query.StartsWith("mt", StringComparison.OrdinalIgnoreCase))
        {
            var separator = query.IndexOf(':');
            if (separator <= 1 || separator == query.Length - 1)
            {
                Console.Error.WriteLine("Usage: vx ls mt:ClassName");
                return 1;
            }

            className = query.Substring(separator + 1).Trim().TrimEnd(':');
            if (string.IsNullOrWhiteSpace(className))
            {
                Console.Error.WriteLine("Usage: vx ls mt:ClassName");
                return 1;
            }

            mode = "mt";
        }
        else
        {
            mode = query.ToLowerInvariant();
        }

        if (mode != "us" && mode != "cl" && mode != "ns" && mode != "mt")
        {
            Console.Error.WriteLine("Usage: vx ls us | cl | ns | mt:ClassName");
            return 1;
        }

        var instances = VsRot.GetRunningDteInstances();
        try
        {
            OleMessageFilter.Register();
            if (instances.Count == 0)
            {
                Console.Error.WriteLine("No running Visual Studio 2022 instance found.");
                return 1;
            }

            var target = VsSelector.GetActiveInstance(instances) ?? instances[0];
            if (target.Dte == null)
            {
                Console.Error.WriteLine("Failed to access running Visual Studio instance.");
                return 1;
            }

            dynamic dte = target.Dte;
            var activeDoc = TryGetValue(() => (dynamic)dte.ActiveDocument);
            if (activeDoc == null)
            {
                Console.Error.WriteLine("No active document is selected.");
                return 1;
            }

            var text = GetActiveDocumentText(activeDoc);
            if (string.IsNullOrWhiteSpace(text))
            {
                Console.Error.WriteLine("Failed to read the active document.");
                return 1;
            }

            SyntaxTree tree = CSharpSyntaxTree.ParseText(text);
            var diagnostics = tree.GetDiagnostics().Where(d => d.Severity == DiagnosticSeverity.Error).ToList();
            if (diagnostics.Count > 0)
            {
                Console.Error.WriteLine("Syntax errors found in the active document:");
                foreach (var diagnostic in diagnostics)
                {
                    var span = diagnostic.Location.GetLineSpan();
                    var line = span.StartLinePosition.Line + 1;
                    var column = span.StartLinePosition.Character + 1;
                    Console.Error.WriteLine($"  L{line}:C{column} {diagnostic.GetMessage()}");
                }

                return 1;
            }

            var root = tree.GetRoot();
            switch (mode)
            {
                case "us":
                    return PrintUsingList(root);
                case "ns":
                    return PrintNamespaceList(root);
                case "cl":
                    return PrintClassList(root);
                case "mt":
                    return PrintMethodList(root, className!);
                default:
                    Console.Error.WriteLine("Usage: vx ls us | cl | ns | mt:ClassName");
                    return 1;
            }
        }
        catch (COMException ex)
        {
            Console.Error.WriteLine($"Visual Studio automation error: {ex.Message}");
            return 1;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unexpected error: {ex.Message}");
            return 1;
        }
        finally
        {
            OleMessageFilter.Revoke();
            VsRot.ReleaseInstances(instances);
        }
    }

    private static int RunSolutionBuild(string action, string[] args)
    {
        if (args.Length == 0)
        {
            return ExecuteBuildAction(action, null);
        }

        if (args.Length == 1)
        {
            var selector = args[0].Trim();
            if (selector.StartsWith("!", StringComparison.Ordinal))
            {
                selector = selector.Substring(1).Trim();
            }

            if (string.IsNullOrWhiteSpace(selector))
            {
                Console.Error.WriteLine("Usage: vx build|rebuild|clean [projectPattern]");
                return 1;
            }

            return ExecuteBuildAction(action, selector);
        }

        {
            Console.Error.WriteLine("Usage: vx build|rebuild|clean [projectPattern]");
            return 1;
        }
    }

    private static int RunDeploy(string[] args)
    {
        if (args.Length != 1)
        {
            Console.Error.WriteLine("Usage: vx deploy !projectPattern");
            return 1;
        }

        var selector = args[0].Trim();
        if (selector.StartsWith("!", StringComparison.Ordinal))
        {
            selector = selector.Substring(1).Trim();
        }

        if (string.IsNullOrWhiteSpace(selector))
        {
            Console.Error.WriteLine("Usage: vx deploy !projectPattern");
            return 1;
        }

        return ExecuteBuildAction("deploy", selector);
    }

    private static int RunStartup(string[] args)
    {
        if (args.Length != 1)
        {
            Console.Error.WriteLine("Usage: vx startup !projectPattern");
            return 1;
        }

        var selector = NormalizeProjectSelector(args[0]);
        if (string.IsNullOrWhiteSpace(selector))
        {
            Console.Error.WriteLine("Usage: vx startup !projectPattern");
            return 1;
        }

        return SetStartupProjects(new[] { selector });
    }

    private static int RunStartups(string[] args)
    {
        if (args.Length == 0)
        {
            Console.Error.WriteLine("Usage: vx startups !pattern1;!pattern2");
            return 1;
        }

        var tokens = args
            .SelectMany(arg => arg.Split(';', StringSplitOptions.RemoveEmptyEntries))
            .Select(token => token.Trim())
            .Where(token => !string.IsNullOrWhiteSpace(token))
            .ToList();

        if (tokens.Count == 0)
        {
            Console.Error.WriteLine("Usage: vx startups !pattern1;!pattern2");
            return 1;
        }

        var selectors = tokens
            .Select(s => NormalizeProjectSelector(s))
            .Where(s => !string.IsNullOrWhiteSpace(s))
            .ToList();

        if (selectors.Count == 0)
        {
            Console.Error.WriteLine("Usage: vx startups !pattern1;!pattern2");
            return 1;
        }

        return SetStartupProjects(selectors);
    }

    private static string NormalizeProjectSelector(string raw)
    {
        if (string.IsNullOrWhiteSpace(raw))
        {
            return string.Empty;
        }

        var selector = raw.Trim();
        if (selector.StartsWith("!", StringComparison.Ordinal))
        {
            selector = selector.Substring(1).Trim();
        }

        return selector;
    }

    private static int RunProjectBuildCommand(string[] args)
    {
        if (args.Length != 1)
        {
            Console.Error.WriteLine("Usage: vx !ProjectName:build|rebuild|clean|deploy");
            return 1;
        }

        var token = args[0];
        if (!token.StartsWith("!", StringComparison.Ordinal))
        {
            Console.Error.WriteLine("Usage: vx !ProjectName:build|rebuild|clean|deploy");
            return 1;
        }

        var separator = token.IndexOf(':');
        if (separator <= 1 || separator == token.Length - 1)
        {
            Console.Error.WriteLine("Usage: vx !ProjectName:build|rebuild|clean|deploy");
            return 1;
        }

        var projectName = token.Substring(1, separator - 1).Trim();
        var action = token.Substring(separator + 1).Trim().ToLowerInvariant();
        if (string.IsNullOrWhiteSpace(projectName) || string.IsNullOrWhiteSpace(action))
        {
            Console.Error.WriteLine("Usage: vx !ProjectName:build|rebuild|clean|deploy");
            return 1;
        }

        if (action is not ("build" or "rebuild" or "clean" or "deploy"))
        {
            Console.Error.WriteLine("Usage: vx !ProjectName:build|rebuild|clean|deploy");
            return 1;
        }

        return ExecuteBuildAction(action, projectName);
    }

    private static int ExecuteBuildAction(string action, string? projectName)
    {
        var instances = VsRot.GetRunningDteInstances();
        try
        {
            OleMessageFilter.Register();
            if (instances.Count == 0)
            {
                Console.Error.WriteLine("No running Visual Studio 2022 instance found.");
                return 1;
            }

            var target = VsSelector.GetActiveInstance(instances) ?? instances[0];
            if (target.Dte == null)
            {
                Console.Error.WriteLine("Failed to access running Visual Studio instance.");
                return 1;
            }

            dynamic dte = target.Dte;
            var solution = TryGetValue(() => (dynamic)dte.Solution);
            if (solution == null)
            {
                Console.Error.WriteLine("No solution is open in the active Visual Studio instance.");
                return 1;
            }

            var isOpen = TryGetValue(() => (bool)solution.IsOpen, false);
            if (!isOpen)
            {
                Console.Error.WriteLine("No solution is open in the active Visual Studio instance.");
                return 1;
            }

            var build = TryGetValue(() => (dynamic)solution.SolutionBuild);
            if (build == null)
            {
                Console.Error.WriteLine("Solution build service is unavailable.");
                return 1;
            }

            if (string.IsNullOrWhiteSpace(projectName))
            {
                if (string.Equals(action, "deploy", StringComparison.OrdinalIgnoreCase))
                {
                    Console.Error.WriteLine("Usage: vx deploy !projectPattern");
                    return 1;
                }

                var result = ExecuteSolutionBuildAction(dte, build, action);
                if (result == 0)
                {
                    ReportBuildResult(action, build, GetActiveSolutionConfiguration(build), null);
                }

                return result;
            }

            var project = FindProjectBySelector(solution, projectName);
            if (project == null)
            {
                Console.Error.WriteLine($"Project not found: {projectName}");
                return 1;
            }

            var projectDisplayName = TryGetValue(() => (string?)project.Name) ?? projectName;
            var projectUniqueName = TryGetValue(() => (string?)project.UniqueName)
                ?? TryGetValue(() => (string?)project.FullName)
                ?? projectDisplayName;

            if (string.IsNullOrWhiteSpace(projectUniqueName))
            {
                Console.Error.WriteLine("Project unique name is unavailable.");
                return 1;
            }

            var configurationName = GetActiveSolutionConfiguration(build);
            if (string.IsNullOrWhiteSpace(configurationName))
            {
                Console.Error.WriteLine("Active solution configuration is unavailable.");
                return 1;
            }

            var projectResult = ExecuteProjectBuildAction(build, action, configurationName, projectUniqueName, projectDisplayName);
            if (projectResult == 0)
            {
                ReportBuildResult(action, build, configurationName, projectDisplayName);
            }

            return projectResult;
        }
        catch (COMException ex)
        {
            Console.Error.WriteLine($"Visual Studio automation error: {ex.Message}");
            return 1;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unexpected error: {ex.Message}");
            return 1;
        }
        finally
        {
            OleMessageFilter.Revoke();
            VsRot.ReleaseInstances(instances);
        }
    }

    private static int SetStartupProjects(IReadOnlyList<string> selectors)
    {
        var instances = VsRot.GetRunningDteInstances();
        try
        {
            OleMessageFilter.Register();
            if (instances.Count == 0)
            {
                Console.Error.WriteLine("No running Visual Studio 2022 instance found.");
                return 1;
            }

            var target = VsSelector.GetActiveInstance(instances) ?? instances[0];
            if (target.Dte == null)
            {
                Console.Error.WriteLine("Failed to access running Visual Studio instance.");
                return 1;
            }

            dynamic dte = target.Dte;
            var solution = TryGetValue(() => (dynamic)dte.Solution);
            if (solution == null)
            {
                Console.Error.WriteLine("No solution is open in the active Visual Studio instance.");
                return 1;
            }

            var isOpen = TryGetValue(() => (bool)solution.IsOpen, false);
            if (!isOpen)
            {
                Console.Error.WriteLine("No solution is open in the active Visual Studio instance.");
                return 1;
            }

            var startupProjects = new List<string>();
            var displayNames = new List<string>();

            foreach (var selector in selectors)
            {
                var project = FindProjectBySelector(solution, selector);
                if (project == null)
                {
                    Console.Error.WriteLine($"Project not found: {selector}");
                    return 1;
                }

                var displayName = TryGetValue(() => (string?)project.Name) ?? selector;
                var uniqueName = TryGetValue(() => (string?)project.UniqueName)
                    ?? TryGetValue(() => (string?)project.FullName)
                    ?? displayName;

                if (string.IsNullOrWhiteSpace(uniqueName))
                {
                    Console.Error.WriteLine($"Project unique name is unavailable: {displayName}");
                    return 1;
                }

                if (!startupProjects.Contains(uniqueName, StringComparer.OrdinalIgnoreCase))
                {
                    startupProjects.Add(uniqueName);
                    displayNames.Add(displayName);
                }
            }

            if (startupProjects.Count == 0)
            {
                Console.Error.WriteLine("No matching projects found.");
                return 1;
            }

            var build = TryGetValue(() => (dynamic)solution.SolutionBuild);
            if (build == null)
            {
                Console.Error.WriteLine("Solution build service is unavailable.");
                return 1;
            }

            build.StartupProjects = startupProjects.ToArray();

            Console.WriteLine("Startup projects set to:");
            foreach (var name in displayNames)
            {
                Console.WriteLine($"  - {name}");
            }

            return 0;
        }
        catch (COMException ex)
        {
            Console.Error.WriteLine($"Visual Studio automation error: {ex.Message}");
            return 1;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unexpected error: {ex.Message}");
            return 1;
        }
        finally
        {
            OleMessageFilter.Revoke();
            VsRot.ReleaseInstances(instances);
        }
    }

    private static int ExecuteSolutionBuildAction(dynamic dte, dynamic build, string action)
    {
        string? error;
        switch (action)
        {
            case "build":
                if (!TryInvoke(() => build.Build(true), out error) &&
                    !TryInvoke(() => dte.ExecuteCommand("Build.BuildSolution"), out error))
                {
                    Console.Error.WriteLine($"Build failed: {error}");
                    return 1;
                }

                return 0;
            case "rebuild":
                if (!TryInvoke(() => build.Rebuild(true), out error) &&
                    !TryInvoke(() => dte.ExecuteCommand("Build.RebuildSolution"), out error))
                {
                    Console.Error.WriteLine($"Rebuild failed: {error}");
                    return 1;
                }

                return 0;
            case "clean":
                if (!TryInvoke(() => build.Clean(true), out error) &&
                    !TryInvoke(() => dte.ExecuteCommand("Build.CleanSolution"), out error))
                {
                    Console.Error.WriteLine($"Clean failed: {error}");
                    return 1;
                }

                return 0;
            default:
                Console.Error.WriteLine("Usage: vx build | rebuild | clean");
                return 1;
        }
    }

    private static int ExecuteProjectBuildAction(dynamic build, string action, string configurationName, string projectUniqueName, string projectName)
    {
        string? error;
        switch (action)
        {
            case "build":
                if (!TryInvoke(() => build.BuildProject(configurationName, projectUniqueName, true), out error))
                {
                    Console.Error.WriteLine($"Build failed for {projectName}: {error}");
                    return 1;
                }

                return 0;
            case "rebuild":
                if (!TryInvoke(() => build.CleanProject(configurationName, projectUniqueName, true), out error))
                {
                    Console.Error.WriteLine($"Rebuild failed (clean) for {projectName}: {error}");
                    return 1;
                }

                if (!TryInvoke(() => build.BuildProject(configurationName, projectUniqueName, true), out error))
                {
                    Console.Error.WriteLine($"Rebuild failed (build) for {projectName}: {error}");
                    return 1;
                }

                return 0;
            case "clean":
                if (!TryInvoke(() => build.CleanProject(configurationName, projectUniqueName, true), out error))
                {
                    Console.Error.WriteLine($"Clean failed for {projectName}: {error}");
                    return 1;
                }

                return 0;
            case "deploy":
                if (!TryInvoke(() => build.DeployProject(configurationName, projectUniqueName, true), out error))
                {
                    Console.Error.WriteLine($"Deploy failed for {projectName}: {error}");
                    return 1;
                }

                return 0;
            default:
                Console.Error.WriteLine("Usage: vx !ProjectName:build|rebuild|clean|deploy");
                return 1;
        }
    }

    private static string? GetActiveSolutionConfiguration(dynamic build)
    {
        var activeConfig = TryGetValue(() => (dynamic)build.ActiveConfiguration);
        if (activeConfig == null)
        {
            return null;
        }

        var name = TryGetValue(() => (string?)activeConfig.Name);
        var platform = TryGetValue(() => (string?)activeConfig.PlatformName);
        if (string.IsNullOrWhiteSpace(name))
        {
            return null;
        }

        if (!string.IsNullOrWhiteSpace(platform))
        {
            return $"{name}|{platform}";
        }

        return name;
    }

    private static void ReportBuildResult(string action, dynamic build, string? configuration, string? projectName)
    {
        var lastBuildInfo = TryGetValue(() => (int)build.LastBuildInfo, -1);
        var buildState = TryGetValue(() => build.BuildState)?.ToString();

        Console.WriteLine($"{action} completed.");
        if (!string.IsNullOrWhiteSpace(configuration))
        {
            Console.WriteLine($"  Configuration: {configuration}");
        }

        if (!string.IsNullOrWhiteSpace(projectName))
        {
            Console.WriteLine($"  Project: {projectName}");
        }

        if (lastBuildInfo >= 0)
        {
            Console.WriteLine($"  Errors: {lastBuildInfo}");
        }

        if (!string.IsNullOrWhiteSpace(buildState))
        {
            Console.WriteLine($"  Build state: {buildState}");
        }
    }

    private static bool TryInvoke(Action action, out string? error)
    {
        try
        {
            action();
            error = null;
            return true;
        }
        catch (COMException ex)
        {
            error = ex.Message;
            return false;
        }
        catch (Exception ex)
        {
            error = ex.Message;
            return false;
        }
    }

    private static void OpenInExistingInstance(object dteObject, string solutionPath)
    {
        dynamic dte = dteObject;
        dte.Solution.Open(solutionPath);
        TryActivateMainWindow(dte);
    }

    private static void OpenInNewInstance(string solutionPath)
    {
        object? dteObject = null;
        try
        {
            var type = Type.GetTypeFromProgID("VisualStudio.DTE.17.0", true);
            dteObject = Activator.CreateInstance(type!, true);
            if (dteObject == null)
            {
                throw new InvalidOperationException("Failed to create Visual Studio instance.");
            }

            dynamic dte = dteObject;
            dte.MainWindow.Visible = true;
            dte.UserControl = true;
            dte.Solution.Open(solutionPath);
            TryActivateMainWindow(dte);
        }
        finally
        {
            if (dteObject != null && Marshal.IsComObject(dteObject))
            {
                Marshal.FinalReleaseComObject(dteObject);
            }
        }
    }

    private static void TryActivateMainWindow(dynamic dte)
    {
        try
        {
            dte.MainWindow.Activate();
        }
        catch
        {
            // Ignore activation failures.
        }
    }

    private static string? GetActiveDocumentText(dynamic activeDocument)
    {
        if (activeDocument == null)
        {
            return null;
        }

        var textDoc = TryGetValue(() => (dynamic)activeDocument.Object("TextDocument"));
        if (textDoc == null)
        {
            return null;
        }

        var startPoint = TryGetValue(() => (dynamic)textDoc.StartPoint);
        var endPoint = TryGetValue(() => (dynamic)textDoc.EndPoint);
        var startObj = (object?)startPoint;
        var endObj = (object?)endPoint;
        if (startObj == null || endObj == null)
        {
            return null;
        }

        var editPoint = TryGetValue(() => ((dynamic)startObj).CreateEditPoint());
        if (editPoint == null)
        {
            return null;
        }

        return TryGetValue(() => (string)((dynamic)editPoint).GetText(endObj));
    }

    private static int PrintUsingList(SyntaxNode root)
    {
        var usings = root.DescendantNodes().OfType<UsingDirectiveSyntax>()
            .Select(u => u.ToString().Trim())
            .Where(u => !string.IsNullOrWhiteSpace(u))
            .ToList();

        if (usings.Count == 0)
        {
            Console.Error.WriteLine("No using directives found.");
            return 1;
        }

        foreach (var entry in usings)
        {
            Console.WriteLine(entry);
        }

        return 0;
    }

    private static int PrintNamespaceList(SyntaxNode root)
    {
        var namespaces = root.DescendantNodes().OfType<BaseNamespaceDeclarationSyntax>()
            .Select(n => n.Name.ToString())
            .Where(n => !string.IsNullOrWhiteSpace(n))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (namespaces.Count == 0)
        {
            Console.Error.WriteLine("No namespace declarations found.");
            return 1;
        }

        foreach (var entry in namespaces)
        {
            Console.WriteLine(entry);
        }

        return 0;
    }

    private static int PrintClassList(SyntaxNode root)
    {
        var classes = root.DescendantNodes().OfType<ClassDeclarationSyntax>()
            .Select(GetFullTypeName)
            .Where(n => !string.IsNullOrWhiteSpace(n))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (classes.Count == 0)
        {
            Console.Error.WriteLine("No class declarations found.");
            return 1;
        }

        foreach (var entry in classes)
        {
            Console.WriteLine(entry);
        }

        return 0;
    }

    private static int PrintMethodList(SyntaxNode root, string className)
    {
        var classDecl = root.DescendantNodes().OfType<ClassDeclarationSyntax>()
            .FirstOrDefault(c => string.Equals(c.Identifier.Text, className, StringComparison.OrdinalIgnoreCase));

        if (classDecl == null)
        {
            Console.Error.WriteLine($"Class not found: {className}");
            return 1;
        }

        var methods = classDecl.Members.OfType<MethodDeclarationSyntax>().ToList();
        if (methods.Count == 0)
        {
            Console.Error.WriteLine($"No methods found in class {className}.");
            return 1;
        }

        foreach (var method in methods)
        {
            var signature = method.Identifier.Text;
            if (method.TypeParameterList != null)
            {
                signature += method.TypeParameterList.ToString();
            }

            signature += method.ParameterList.ToString();
            var returnType = method.ReturnType.ToString();
            Console.WriteLine($"{returnType} {signature}");
        }

        return 0;
    }

    private static string GetFullTypeName(ClassDeclarationSyntax classDecl)
    {
        var parts = new List<string>
        {
            classDecl.Identifier.Text
        };

        for (var parent = classDecl.Parent; parent != null; parent = parent.Parent)
        {
            switch (parent)
            {
                case ClassDeclarationSyntax parentClass:
                    parts.Add(parentClass.Identifier.Text);
                    break;
                case BaseNamespaceDeclarationSyntax ns:
                    parts.Add(ns.Name.ToString());
                    break;
            }
        }

        parts.Reverse();
        return string.Join(".", parts);
    }

    private const string SolutionFolderKind = "{66A26720-8FB5-11D2-AA7E-00C04F688DDE}";

    private static dynamic? FindProjectBySelector(dynamic solution, string projectSelector)
    {
        var projects = TryGetValue(() => (dynamic)solution.Projects);
        foreach (var project in EnumerateProjects(projects))
        {
            var name = TryGetValue(() => (string?)project.Name);
            var uniqueName = TryGetValue(() => (string?)project.UniqueName);
            var fullName = TryGetValue(() => (string?)project.FullName);
            if (IsProjectMatch(projectSelector, name, uniqueName, fullName))
            {
                return project;
            }
        }

        return null;
    }

    private static string? FindFileInSolution(dynamic solution, string fileSpec)
    {
        var projects = TryGetValue(() => (dynamic)solution.Projects);
        foreach (var project in EnumerateProjects(projects))
        {
            var path = FindFileInProject(project, fileSpec);
            if (!string.IsNullOrWhiteSpace(path))
            {
                return path;
            }
        }

        return null;
    }

    private static string? FindFileInProject(dynamic project, string fileSpec)
    {
        var items = TryGetValue(() => (dynamic)project.ProjectItems);
        var path = FindFileInProjectItems(items, fileSpec);
        if (!string.IsNullOrWhiteSpace(path))
        {
            return path;
        }

        var projectPath = TryGetValue(() => (string?)project.FullName);
        if (string.IsNullOrWhiteSpace(projectPath))
        {
            return null;
        }

        var projectDir = Path.GetDirectoryName(projectPath);
        if (string.IsNullOrWhiteSpace(projectDir))
        {
            return null;
        }

        return FindFileOnDisk(projectDir, fileSpec);
    }

    private static string? FindFileInProjectItems(dynamic items, string fileSpec)
    {
        if (items == null)
        {
            return null;
        }

        foreach (var item in EnumerateProjectItems(items))
        {
            var fileCount = TryGetValue(() => (short)item.FileCount, (short)0);
            for (short i = 1; i <= fileCount; i++)
            {
                var path = TryGetValue(() => (string?)item.FileNames(i));
                if (IsPathMatch(path, fileSpec))
                {
                    return path;
                }
            }
        }

        return null;
    }

    private static IEnumerable<dynamic> EnumerateProjects(dynamic projects)
    {
        foreach (var project in EnumerateComCollection(projects))
        {
            foreach (var nested in EnumerateProjectNode(project))
            {
                yield return nested;
            }
        }
    }

    private static IEnumerable<dynamic> EnumerateProjectNode(dynamic project)
    {
        if (project == null)
        {
            yield break;
        }

        var kind = TryGetValue(() => (string?)project.Kind);
        if (string.Equals(kind, SolutionFolderKind, StringComparison.OrdinalIgnoreCase))
        {
            var items = TryGetValue(() => (dynamic)project.ProjectItems);
            foreach (var item in EnumerateComCollection(items))
            {
                var subProject = TryGetValue(() => (dynamic)item.SubProject);
                foreach (var nested in EnumerateProjectNode(subProject))
                {
                    yield return nested;
                }
            }
        }
        else
        {
            yield return project;
        }
    }

    private static IEnumerable<dynamic> EnumerateProjectItems(dynamic items)
    {
        foreach (var item in EnumerateComCollection(items))
        {
            if (item == null)
            {
                continue;
            }

            yield return item;

            var nestedItems = TryGetValue(() => (dynamic)item.ProjectItems);
            if (nestedItems == null)
            {
                continue;
            }

            foreach (var nested in EnumerateProjectItems(nestedItems))
            {
                yield return nested;
            }
        }
    }

    private static string? FindFileOnDisk(string rootDirectory, string fileSpec)
    {
        if (string.IsNullOrWhiteSpace(rootDirectory) || string.IsNullOrWhiteSpace(fileSpec))
        {
            return null;
        }

        if (Path.IsPathRooted(fileSpec))
        {
            return File.Exists(fileSpec) ? Path.GetFullPath(fileSpec) : null;
        }

        var combined = Path.Combine(rootDirectory, fileSpec);
        if (File.Exists(combined))
        {
            return Path.GetFullPath(combined);
        }

        var normalizedSpec = fileSpec.Replace('/', '\\');
        var hasSubPath = normalizedSpec.Contains('\\');
        var ignore = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "bin",
            "obj",
            ".vs",
            ".git",
            "node_modules",
            "packages"
        };

        var stack = new Stack<string>();
        stack.Push(rootDirectory);

        while (stack.Count > 0)
        {
            var current = stack.Pop();
            string currentName;
            try
            {
                currentName = new DirectoryInfo(current).Name;
            }
            catch
            {
                continue;
            }

            if (ignore.Contains(currentName))
            {
                continue;
            }

            IEnumerable<string> directories;
            try
            {
                directories = Directory.EnumerateDirectories(current);
            }
            catch
            {
                directories = Array.Empty<string>();
            }

            foreach (var dir in directories)
            {
                stack.Push(dir);
            }

            IEnumerable<string> files;
            try
            {
                files = hasSubPath
                    ? Directory.EnumerateFiles(current)
                    : Directory.EnumerateFiles(current, normalizedSpec);
            }
            catch
            {
                continue;
            }

            foreach (var file in files)
            {
                if (hasSubPath)
                {
                    var normalizedFile = file.Replace('/', '\\');
                    if (!normalizedFile.EndsWith(normalizedSpec, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }
                }

                return file;
            }
        }

        return null;
    }

    private static bool IsProjectMatch(string selector, string? name, string? uniqueName, string? fullName)
    {
        if (string.IsNullOrWhiteSpace(selector))
        {
            return false;
        }

        var pattern = selector.Trim();
        var hasWildcard = pattern.Contains('*');

        bool Match(string? candidate)
        {
            if (string.IsNullOrWhiteSpace(candidate))
            {
                return false;
            }

            return hasWildcard
                ? IsWildcardMatch(candidate, pattern)
                : string.Equals(candidate, pattern, StringComparison.OrdinalIgnoreCase);
        }

        if (Match(name) || Match(uniqueName) || Match(fullName))
        {
            return true;
        }

        if (!string.IsNullOrWhiteSpace(fullName))
        {
            var fileName = Path.GetFileNameWithoutExtension(fullName);
            if (Match(fileName))
            {
                return true;
            }
        }

        return false;
    }

    private static bool IsWildcardMatch(string input, string pattern)
    {
        if (string.IsNullOrWhiteSpace(pattern))
        {
            return false;
        }

        if (pattern == "*")
        {
            return true;
        }

        var regexPattern = "^" + Regex.Escape(pattern).Replace("\\*", ".*") + "$";
        return Regex.IsMatch(input ?? string.Empty, regexPattern, RegexOptions.IgnoreCase);
    }

    private static IEnumerable<dynamic> EnumerateComCollection(dynamic collection)
    {
        if (collection == null)
        {
            yield break;
        }

        if (collection is IEnumerable enumerable)
        {
            foreach (var item in enumerable)
            {
                yield return item;
            }

            yield break;
        }

        int count;
        try
        {
            count = collection.Count;
        }
        catch
        {
            yield break;
        }

        for (var i = 1; i <= count; i++)
        {
            object? item = null;
            try
            {
                item = collection.Item(i);
            }
            catch
            {
                // Skip invalid items.
            }

            if (item != null)
            {
                yield return item;
            }
        }
    }

    private static bool IsPathMatch(string? candidatePath, string fileSpec)
    {
        if (string.IsNullOrWhiteSpace(candidatePath) || string.IsNullOrWhiteSpace(fileSpec))
        {
            return false;
        }

        var candidate = candidatePath.Replace('/', '\\');
        var spec = fileSpec.Replace('/', '\\');

        if (Path.IsPathRooted(spec))
        {
            try
            {
                var fullCandidate = Path.GetFullPath(candidate);
                var fullSpec = Path.GetFullPath(spec);
                return string.Equals(fullCandidate, fullSpec, StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return string.Equals(candidate, spec, StringComparison.OrdinalIgnoreCase);
            }
        }

        if (spec.Contains('\\'))
        {
            return candidate.EndsWith(spec, StringComparison.OrdinalIgnoreCase);
        }

        return string.Equals(Path.GetFileName(candidate), spec, StringComparison.OrdinalIgnoreCase);
    }

    private static T? TryGetValue<T>(Func<T> action, T? fallback = default)
    {
        try
        {
            return action();
        }
        catch (COMException)
        {
            return fallback;
        }
        catch (InvalidCastException)
        {
            return fallback;
        }
        catch
        {
            return fallback;
        }
    }

    private static void WriteJson(VxInfoSnapshot info)
    {
        var options = new JsonSerializerOptions { WriteIndented = true };
        Console.WriteLine(JsonSerializer.Serialize(info, options));
    }

    private static void PrintInfoText(VxInfoSnapshot info, int activeIndex)
    {
        Console.WriteLine($"VS2022 running: yes ({info.InstanceCount} instance{(info.InstanceCount == 1 ? "" : "s")})");
        if (info.InstanceCount > 1)
        {
            Console.WriteLine($"Active instance index: {activeIndex}");
        }

        foreach (var instance in info.Instances)
        {
            Console.WriteLine($"Instance {instance.Index}{(instance.IsActive ? " (active)" : string.Empty)}");
            Console.WriteLine($"  PID: {instance.ProcessId?.ToString() ?? "unknown"}");
            Console.WriteLine($"  Version: {instance.Version ?? "unknown"}");
            Console.WriteLine($"  Edition: {instance.Edition ?? "unknown"}");
            Console.WriteLine($"  Moniker: {instance.Moniker ?? "unknown"}");

            if (instance.Solution != null && instance.Solution.IsOpen)
            {
                var solution = instance.Solution;
                var name = solution.Name;
                if (string.IsNullOrWhiteSpace(name))
                {
                    name = solution.FullName;
                }

                Console.WriteLine($"  Solution: {name ?? "none"}");
                if (!string.IsNullOrWhiteSpace(solution.FullName))
                {
                    Console.WriteLine($"  Solution path: {solution.FullName}");
                }

                if (!string.IsNullOrWhiteSpace(solution.ActiveConfiguration))
                {
                    Console.WriteLine($"  Active configuration: {solution.ActiveConfiguration}");
                }

                if (!string.IsNullOrWhiteSpace(solution.ActivePlatform))
                {
                    Console.WriteLine($"  Active platform: {solution.ActivePlatform}");
                }

                Console.WriteLine($"  Projects ({instance.Projects.Count}):");
                foreach (var project in instance.Projects)
                {
                    var kindLabel = project.KindName ?? project.Kind ?? "unknown";
                    var line = $"    - {project.Name ?? "unknown"} [{kindLabel}]";
                    if (!string.IsNullOrWhiteSpace(project.FullName))
                    {
                        line += $" - {project.FullName}";
                    }

                    Console.WriteLine(line);
                }
            }

            Console.WriteLine($"  Open documents ({instance.Documents.Count}):");
            foreach (var document in instance.Documents)
            {
                var name = document.Name ?? "unknown";
                var line = $"    - {name}";
                if (!string.IsNullOrWhiteSpace(document.FullName))
                {
                    line += $" - {document.FullName}";
                }

                Console.WriteLine(line);
            }

            Console.WriteLine($"  Active document: {instance.ActiveDocument ?? "none"}");
        }
    }
}

internal sealed class InfoOptions
{
    public bool All { get; set; }
    public bool Json { get; set; }
    public int? InstanceIndex { get; set; }

    public static InfoOptions Parse(string[] args)
    {
        var options = new InfoOptions();

        for (var i = 0; i < args.Length; i++)
        {
            var arg = args[i];
            if (string.Equals(arg, "--all", StringComparison.OrdinalIgnoreCase))
            {
                options.All = true;
                continue;
            }

            if (string.Equals(arg, "--json", StringComparison.OrdinalIgnoreCase))
            {
                options.Json = true;
                continue;
            }

            if (string.Equals(arg, "--instance", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(arg, "-i", StringComparison.OrdinalIgnoreCase))
            {
                if (i + 1 < args.Length && int.TryParse(args[i + 1], out var index))
                {
                    options.InstanceIndex = index;
                    i++;
                }
            }
        }

        return options;
    }
}

internal sealed class VxInfoSnapshot
{
    public int InstanceCount { get; set; }
    public List<VsInstanceSnapshot> Instances { get; set; } = new();
}

internal sealed class VsInstanceSnapshot
{
    public int Index { get; set; }
    public bool IsActive { get; set; }
    public string? Moniker { get; set; }
    public int? ProcessId { get; set; }
    public string? Version { get; set; }
    public string? Edition { get; set; }
    public SolutionSnapshot? Solution { get; set; }
    public List<ProjectSnapshot> Projects { get; set; } = new();
    public List<DocumentSnapshot> Documents { get; set; } = new();
    public string? ActiveDocument { get; set; }
}

internal sealed class SolutionSnapshot
{
    public bool IsOpen { get; set; }
    public string? Name { get; set; }
    public string? FullName { get; set; }
    public string? ActiveConfiguration { get; set; }
    public string? ActivePlatform { get; set; }
}

internal sealed class ProjectSnapshot
{
    public string? Name { get; set; }
    public string? FullName { get; set; }
    public string? Kind { get; set; }
    public string? KindName { get; set; }
}

internal sealed class DocumentSnapshot
{
    public string? Name { get; set; }
    public string? FullName { get; set; }
    public string? Kind { get; set; }
}

internal static class SnapshotBuilder
{
    private const string SolutionFolderKind = "{66A26720-8FB5-11D2-AA7E-00C04F688DDE}";

    private static readonly Dictionary<string, string> ProjectKindMap = new(StringComparer.OrdinalIgnoreCase)
    {
        ["{FAE04EC0-301F-11D3-BF4B-00C04F79EFBC}"] = "C#",
        ["{9A19103F-16F7-4668-BE54-9A1E7A4F7556}"] = "C# (SDK)",
        ["{F184B08F-C81C-45F6-A57F-5ABD9991F28F}"] = "VB.NET",
        ["{F2A71F9B-5D33-465A-A702-920D77279786}"] = "F#",
        ["{BC8A1FFA-BEE3-4634-8014-F334798102B3}"] = "C++",
        ["{E24C65DC-7377-472B-9ABA-BC803B73C61A}"] = "Web Site",
        ["{349C5851-65DF-11DA-9384-00065B846F21}"] = "Web Application",
        ["{D954291E-2A0B-460D-934E-DC6B0785DB48}"] = "Shared Project"
    };

    public static VsInstanceSnapshot Build(VsInstance instance, int index, bool isActive)
    {
        var snapshot = new VsInstanceSnapshot
        {
            Index = index,
            IsActive = isActive,
            Moniker = instance.Moniker,
            ProcessId = instance.ProcessId
        };

        if (instance.Dte == null)
        {
            return snapshot;
        }

        dynamic dte = instance.Dte;
        snapshot.Version = TryGet(() => (string?)dte.Version);
        snapshot.Edition = TryGet(() => (string?)dte.Edition);

        var solution = TryGet(() => (dynamic)dte.Solution);
        snapshot.Solution = BuildSolutionSnapshot(solution);
        snapshot.Projects = BuildProjectSnapshots(solution);
        snapshot.Documents = BuildDocumentSnapshots(dte);

        var activeDocument = TryGet(() => (dynamic)dte.ActiveDocument);
        if (activeDocument != null)
        {
            snapshot.ActiveDocument = TryGet(() => (string?)activeDocument.FullName) ?? TryGet(() => (string?)activeDocument.Name);
        }

        return snapshot;
    }

    private static SolutionSnapshot? BuildSolutionSnapshot(dynamic solution)
    {
        if (solution == null)
        {
            return null;
        }

        var snapshot = new SolutionSnapshot
        {
            IsOpen = TryGet(() => (bool)solution.IsOpen, false),
            Name = TryGet(() => (string?)solution.Name),
            FullName = TryGet(() => (string?)solution.FullName)
        };

        var build = TryGet(() => (dynamic)solution.SolutionBuild);
        if (build != null)
        {
            var activeConfig = TryGet(() => (dynamic)build.ActiveConfiguration);
            if (activeConfig != null)
            {
                snapshot.ActiveConfiguration = TryGet(() => (string?)activeConfig.Name);
                snapshot.ActivePlatform = TryGet(() => (string?)activeConfig.PlatformName);
            }
        }

        if (string.IsNullOrWhiteSpace(snapshot.Name) && !string.IsNullOrWhiteSpace(snapshot.FullName))
        {
            snapshot.Name = snapshot.FullName;
        }

        return snapshot;
    }

    private static List<ProjectSnapshot> BuildProjectSnapshots(dynamic solution)
    {
        var list = new List<ProjectSnapshot>();
        if (solution == null)
        {
            return list;
        }

        var projects = TryGet(() => (dynamic)solution.Projects);
        foreach (var project in EnumerateProjects(projects))
        {
            if (project == null)
            {
                continue;
            }

            var kind = TryGet(() => (string?)project.Kind);
            var snapshot = new ProjectSnapshot
            {
                Name = TryGet(() => (string?)project.Name),
                FullName = TryGet(() => (string?)project.FullName),
                Kind = kind,
                KindName = ResolveProjectKind(kind)
            };

            list.Add(snapshot);
        }

        return list;
    }

    private static List<DocumentSnapshot> BuildDocumentSnapshots(dynamic dte)
    {
        var list = new List<DocumentSnapshot>();
        if (dte == null)
        {
            return list;
        }

        var documents = TryGet(() => (dynamic)dte.Documents);
        foreach (var document in EnumerateComCollection(documents))
        {
            if (document == null)
            {
                continue;
            }

            list.Add(new DocumentSnapshot
            {
                Name = TryGet(() => (string?)document.Name),
                FullName = TryGet(() => (string?)document.FullName),
                Kind = TryGet(() => (string?)document.Kind)
            });
        }

        return list;
    }

    private static IEnumerable<dynamic> EnumerateProjects(dynamic projects)
    {
        foreach (var project in EnumerateComCollection(projects))
        {
            foreach (var nested in EnumerateProjectNode(project))
            {
                yield return nested;
            }
        }
    }

    private static IEnumerable<dynamic> EnumerateProjectNode(dynamic project)
    {
        if (project == null)
        {
            yield break;
        }

        var kind = TryGet(() => (string?)project.Kind);
        if (string.Equals(kind, SolutionFolderKind, StringComparison.OrdinalIgnoreCase))
        {
            var items = TryGet(() => (dynamic)project.ProjectItems);
            foreach (var item in EnumerateComCollection(items))
            {
                var subProject = TryGet(() => (dynamic)item.SubProject);
                foreach (var nested in EnumerateProjectNode(subProject))
                {
                    yield return nested;
                }
            }
        }
        else
        {
            yield return project;
        }
    }

    private static string? ResolveProjectKind(string? kind)
    {
        if (string.IsNullOrWhiteSpace(kind))
        {
            return null;
        }

        if (ProjectKindMap.TryGetValue(kind, out var name))
        {
            return name;
        }

        return null;
    }

    private static IEnumerable<dynamic> EnumerateComCollection(dynamic collection)
    {
        if (collection == null)
        {
            yield break;
        }

        if (collection is IEnumerable enumerable)
        {
            foreach (var item in enumerable)
            {
                yield return item;
            }

            yield break;
        }

        int count;
        try
        {
            count = collection.Count;
        }
        catch
        {
            yield break;
        }

        for (var i = 1; i <= count; i++)
        {
            object? item = null;
            try
            {
                item = collection.Item(i);
            }
            catch
            {
                // Skip invalid items.
            }

            if (item != null)
            {
                yield return item;
            }
        }
    }

    private static T? TryGet<T>(Func<T> action, T? fallback = default)
    {
        try
        {
            return action();
        }
        catch (COMException)
        {
            return fallback;
        }
        catch (InvalidCastException)
        {
            return fallback;
        }
        catch
        {
            return fallback;
        }
    }
}

internal sealed class VsInstance
{
    public string? Moniker { get; set; }
    public object? Dte { get; set; }
    public int? ProcessId { get; set; }
}

internal static class VsSelector
{
    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    public static int GetActiveIndex(List<VsInstance> instances)
    {
        var foreground = GetForegroundWindow();
        if (foreground == IntPtr.Zero)
        {
            return -1;
        }

        GetWindowThreadProcessId(foreground, out var foregroundPid);
        if (foregroundPid == 0)
        {
            return -1;
        }

        for (var i = 0; i < instances.Count; i++)
        {
            var instance = instances[i];
            if (!instance.ProcessId.HasValue)
            {
                continue;
            }

            if (instance.ProcessId.Value == (int)foregroundPid)
            {
                return i;
            }
        }

        return -1;
    }

    public static VsInstance? GetActiveInstance(List<VsInstance> instances)
    {
        if (instances == null || instances.Count == 0)
        {
            return null;
        }

        var activeIndex = GetActiveIndex(instances);
        if (activeIndex >= 0 && activeIndex < instances.Count)
        {
            return instances[activeIndex];
        }

        return instances[0];
    }
}

internal static class VsRot
{
    [DllImport("ole32.dll")]
    private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable prot);

    [DllImport("ole32.dll")]
    private static extern int CreateBindCtx(int reserved, out IBindCtx ppbc);

    public static List<VsInstance> GetRunningDteInstances()
    {
        var list = new List<VsInstance>();
        if (GetRunningObjectTable(0, out var rot) != 0 || rot == null)
        {
            return list;
        }

        rot.EnumRunning(out var enumMoniker);
        enumMoniker.Reset();
        var fetched = new IMoniker[1];

        while (enumMoniker.Next(1, fetched, IntPtr.Zero) == 0)
        {
            var moniker = fetched[0];
            if (moniker == null)
            {
                continue;
            }

            if (CreateBindCtx(0, out var bindCtx) != 0 || bindCtx == null)
            {
                continue;
            }

            moniker.GetDisplayName(bindCtx, null, out var displayName);
            if (string.IsNullOrWhiteSpace(displayName) || !displayName.StartsWith("!VisualStudio.DTE.17.0", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            try
            {
                rot.GetObject(moniker, out var comObject);
                if (comObject == null)
                {
                    continue;
                }

                list.Add(new VsInstance
                {
                    Moniker = displayName,
                    Dte = comObject,
                    ProcessId = ParsePid(displayName)
                });
            }
            catch
            {
                // Skip ROT entries we cannot access.
            }
        }

        return list;
    }

    public static void ReleaseInstances(IEnumerable<VsInstance> instances)
    {
        foreach (var instance in instances)
        {
            if (instance.Dte != null && Marshal.IsComObject(instance.Dte))
            {
                Marshal.FinalReleaseComObject(instance.Dte);
            }
        }
    }

    private static int? ParsePid(string? moniker)
    {
        if (string.IsNullOrWhiteSpace(moniker))
        {
            return null;
        }

        var index = moniker.LastIndexOf(':');
        if (index < 0 || index + 1 >= moniker.Length)
        {
            return null;
        }

        var tail = moniker.Substring(index + 1);
        if (int.TryParse(tail, out var pid))
        {
            return pid;
        }

        return null;
    }
}

[ComImport]
[Guid("00000016-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IOleMessageFilter
{
    int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);
    int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);
    int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
}

internal sealed class OleMessageFilter : IOleMessageFilter
{
    private const int ServerCallIshandled = 0;
    private const int RetryLater = 2;
    private const int PendingMsgWaitDefProcess = 2;

    [DllImport("ole32.dll")]
    private static extern int CoRegisterMessageFilter(IOleMessageFilter? newFilter, out IOleMessageFilter? oldFilter);

    private OleMessageFilter()
    {
    }

    public static void Register()
    {
        CoRegisterMessageFilter(new OleMessageFilter(), out _);
    }

    public static void Revoke()
    {
        CoRegisterMessageFilter(null, out _);
    }

    int IOleMessageFilter.HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo)
    {
        return ServerCallIshandled;
    }

    int IOleMessageFilter.RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
    {
        if (dwRejectType == RetryLater)
        {
            return 100;
        }

        return -1;
    }

    int IOleMessageFilter.MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType)
    {
        return PendingMsgWaitDefProcess;
    }
}
