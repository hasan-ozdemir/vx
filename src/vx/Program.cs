using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.Loader;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using System.Windows.Automation;
using System.Xml.Linq;
using Microsoft.Build.Construction;
using Microsoft.Build.Evaluation;
using Microsoft.Build.Locator;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        RegisterAssemblyResolvers();

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
            case "props":
                return RunProps(args.Skip(1).ToArray());
            case "ios:mf":
                return RunManifest("ios", args.Skip(1).ToArray());
            case "droid:mf":
            case "android:mf":
                return RunManifest("android", args.Skip(1).ToArray());
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
        Console.WriteLine("  vx props list !projectPattern");
        Console.WriteLine("  vx props set [!projectPattern]");
        Console.WriteLine("  vx ios:mf list [!projectPattern]");
        Console.WriteLine("  vx ios:mf set [!projectPattern]");
        Console.WriteLine("  vx droid:mf list [!projectPattern]");
        Console.WriteLine("  vx droid:mf set [!projectPattern]");
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
        Console.WriteLine("  props  List or edit project properties.");
        Console.WriteLine("  ios:mf List or edit iOS Info.plist manifest.");
        Console.WriteLine("  droid:mf List or edit AndroidManifest.xml.");
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
                if (IsVisualStudioProcessRunning())
                {
                    Console.WriteLine("VS2022 running: yes (process detected, DTE unavailable)");
                }
                else
                {
                    Console.WriteLine("VS2022 running: no");
                }
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

            if (IsVisualStudioProcessRunning())
            {
                Console.Error.WriteLine("Visual Studio process detected but DTE is not accessible. Run vx as Administrator or restart Visual Studio without elevation.");
                return 1;
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
            Console.Error.WriteLine($"Unexpected error: {FormatUnexpectedError(ex)}");
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
            if (!EnsureVsInstanceAvailable(instances))
            {
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
            Console.Error.WriteLine($"Unexpected error: {FormatUnexpectedError(ex)}");
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
            if (!EnsureVsInstanceAvailable(instances))
            {
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
            Console.Error.WriteLine($"Unexpected error: {FormatUnexpectedError(ex)}");
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

    private static int RunProps(string[] args)
    {
        if (args.Length == 0)
        {
            Console.Error.WriteLine("Usage: vx props list !projectPattern | vx props set [!projectPattern]");
            return 1;
        }

        var subCommand = args[0].Trim();
        if (string.Equals(subCommand, "list", StringComparison.OrdinalIgnoreCase))
        {
            return RunPropsList(args.Skip(1).ToArray());
        }

        if (string.Equals(subCommand, "set", StringComparison.OrdinalIgnoreCase))
        {
            return RunPropsSet(args.Skip(1).ToArray());
        }

        Console.Error.WriteLine("Usage: vx props list !projectPattern | vx props set [!projectPattern]");
        return 1;
    }

    private static int RunPropsList(string[] args)
    {
        if (args.Length == 0)
        {
            Console.Error.WriteLine("Usage: vx props list !projectPattern");
            return 1;
        }

        var selector = NormalizeProjectSelector(string.Join(" ", args));
        if (string.IsNullOrWhiteSpace(selector))
        {
            Console.Error.WriteLine("Usage: vx props list !projectPattern");
            return 1;
        }

        return ListProjectProperties(selector);
    }

    private static int RunPropsSet(string[] args)
    {
        string? selector = null;
        if (args.Length > 0)
        {
            selector = NormalizeProjectSelector(string.Join(" ", args));
        }

        return RunPropertyWizard(selector);
    }

    private static int RunManifest(string platform, string[] args)
    {
        if (args.Length == 0)
        {
            PrintManifestUsage(platform);
            return 1;
        }

        var subCommand = args[0].Trim();
        if (string.Equals(subCommand, "list", StringComparison.OrdinalIgnoreCase))
        {
            return RunManifestList(platform, args.Skip(1).ToArray());
        }

        if (string.Equals(subCommand, "set", StringComparison.OrdinalIgnoreCase))
        {
            return RunManifestSet(platform, args.Skip(1).ToArray());
        }

        PrintManifestUsage(platform);
        return 1;
    }

    private static void PrintManifestUsage(string platform)
    {
        if (string.Equals(platform, "ios", StringComparison.OrdinalIgnoreCase))
        {
            Console.Error.WriteLine("Usage: vx ios:mf list [!projectPattern] | vx ios:mf set [!projectPattern]");
        }
        else
        {
            Console.Error.WriteLine("Usage: vx droid:mf list [!projectPattern] | vx droid:mf set [!projectPattern]");
        }
    }

    private static int RunManifestList(string platform, string[] args)
    {
        string? selector = null;
        if (args.Length > 0)
        {
            selector = NormalizeProjectSelector(string.Join(" ", args));
        }

        return ListManifestProperties(platform, selector);
    }

    private static int RunManifestSet(string platform, string[] args)
    {
        string? selector = null;
        if (args.Length > 0)
        {
            selector = NormalizeProjectSelector(string.Join(" ", args));
        }

        return RunManifestWizard(platform, selector);
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

        if (selector.EndsWith("!", StringComparison.Ordinal))
        {
            selector = selector.TrimEnd('!').Trim();
        }

        return selector;
    }

    private static int ListProjectProperties(string selector)
    {
        var instances = VsRot.GetRunningDteInstances();
        try
        {
            OleMessageFilter.Register();
            if (!EnsureVsInstanceAvailable(instances))
            {
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
            if (solution == null || !TryGetValue(() => (bool)solution.IsOpen, false))
            {
                Console.Error.WriteLine("No solution is open in the active Visual Studio instance.");
                return 1;
            }

            var project = FindProjectBySelector(solution, selector);
            if (project == null)
            {
                Console.Error.WriteLine($"Project not found: {selector}");
                return 1;
            }

            var includeMsBuild = PrepareMsBuild();
            var pages = GetProjectPropertyPages((object?)project, includeMsBuild);
            if (pages.Count == 0)
            {
                Console.Error.WriteLine("No project properties found.");
                return 1;
            }

            var printablePages = pages.Where(p => p.Items.Count > 0).ToList();
            if (printablePages.Count == 0)
            {
                Console.Error.WriteLine("No project properties found. Ensure MSBuild evaluation is available.");
                return 1;
            }

            foreach (var page in printablePages)
            {
                Console.WriteLine($"*[{page.Name}]:");
                foreach (var item in page.Items)
                {
                    var value = SafeFormatPropertyValue(item);
                    Console.WriteLine($"-{item.Name}: {value}");
                }
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
            Console.Error.WriteLine($"Unexpected error: {FormatUnexpectedError(ex)}");
            return 1;
        }
        finally
        {
            OleMessageFilter.Revoke();
            VsRot.ReleaseInstances(instances);
        }
    }

    private static int RunPropertyWizard(string? selector)
    {
        var instances = VsRot.GetRunningDteInstances();
        try
        {
            OleMessageFilter.Register();
            if (!EnsureVsInstanceAvailable(instances))
            {
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
            if (solution == null || !TryGetValue(() => (bool)solution.IsOpen, false))
            {
                Console.Error.WriteLine("No solution is open in the active Visual Studio instance.");
                return 1;
            }

            var project = string.IsNullOrWhiteSpace(selector)
                ? SelectProjectInteractively(solution)
                : FindProjectBySelector(solution, selector);

            if (project == null)
            {
                Console.Error.WriteLine(string.IsNullOrWhiteSpace(selector)
                    ? "No project selected."
                    : $"Project not found: {selector}");
                return 1;
            }

            var includeMsBuild = PrepareMsBuild();
            return RunPropertyWizardForProject(project, includeMsBuild, null);
        }
        catch (COMException ex)
        {
            Console.Error.WriteLine($"Visual Studio automation error: {ex.Message}");
            return 1;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unexpected error: {FormatUnexpectedError(ex)}");
            return 1;
        }
        finally
        {
            OleMessageFilter.Revoke();
            VsRot.ReleaseInstances(instances);
        }
    }

    private static int ListManifestProperties(string platform, string? selector)
    {
        var instances = VsRot.GetRunningDteInstances();
        try
        {
            OleMessageFilter.Register();
            if (!EnsureVsInstanceAvailable(instances))
            {
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
            if (solution == null || !TryGetValue(() => (bool)solution.IsOpen, false))
            {
                Console.Error.WriteLine("No solution is open in the active Visual Studio instance.");
                return 1;
            }

            var project = ResolveManifestProject(solution, platform, selector);
            if (project == null)
            {
                Console.Error.WriteLine(string.IsNullOrWhiteSpace(selector)
                    ? $"No {platform} project found in the active solution."
                    : $"Project not found: {selector}");
                return 1;
            }

            var includeMsBuild = PrepareMsBuild();
            var pages = GetManifestPropertyPages((object?)project, platform, includeMsBuild, out var manifestPath, out var error);
            if (!string.IsNullOrWhiteSpace(error))
            {
                Console.Error.WriteLine(error);
                return 1;
            }

            if (pages.Count == 0)
            {
                Console.Error.WriteLine("No manifest properties found.");
                return 1;
            }

            foreach (var page in pages.Where(p => p.Items.Count > 0))
            {
                Console.WriteLine($"*[{page.Name}]:");
                foreach (var item in page.Items)
                {
                    var value = SafeFormatPropertyValue(item);
                    Console.WriteLine($"-{item.Name}: {value}");
                }
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
            Console.Error.WriteLine($"Unexpected error: {FormatUnexpectedError(ex)}");
            return 1;
        }
        finally
        {
            OleMessageFilter.Revoke();
            VsRot.ReleaseInstances(instances);
        }
    }

    private static int RunManifestWizard(string platform, string? selector)
    {
        var instances = VsRot.GetRunningDteInstances();
        try
        {
            OleMessageFilter.Register();
            if (!EnsureVsInstanceAvailable(instances))
            {
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
            if (solution == null || !TryGetValue(() => (bool)solution.IsOpen, false))
            {
                Console.Error.WriteLine("No solution is open in the active Visual Studio instance.");
                return 1;
            }

            var project = ResolveManifestProject(solution, platform, selector);
            if (project == null)
            {
                Console.Error.WriteLine(string.IsNullOrWhiteSpace(selector)
                    ? $"No {platform} project found in the active solution."
                    : $"Project not found: {selector}");
                return 1;
            }

            var includeMsBuild = PrepareMsBuild();
            var pages = GetManifestPropertyPages((object?)project, platform, includeMsBuild, out _, out var error);
            if (!string.IsNullOrWhiteSpace(error))
            {
                Console.Error.WriteLine(error);
                return 1;
            }

            return RunPropertyWizardForPages(pages);
        }
        catch (COMException ex)
        {
            Console.Error.WriteLine($"Visual Studio automation error: {ex.Message}");
            return 1;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unexpected error: {FormatUnexpectedError(ex)}");
            return 1;
        }
        finally
        {
            OleMessageFilter.Revoke();
            VsRot.ReleaseInstances(instances);
        }
    }

    private static int RunPropertyWizardForPages(List<PropertyPage> pages)
    {
        if (pages.Count == 0)
        {
            Console.WriteLine("No properties found.");
            return 1;
        }

        pages = pages.Where(p => p.Items.Count > 0).ToList();
        if (pages.Count == 0)
        {
            Console.WriteLine("No properties found.");
            return 1;
        }

        while (true)
        {
            Console.WriteLine("*select a manifest page:");
            for (var i = 0; i < pages.Count; i++)
            {
                Console.WriteLine($"{i + 1}. {pages[i].Name}");
            }

            Console.Write("Select page (b to go back, q to quit): ");
            var input = Console.ReadLine()?.Trim();
            if (string.IsNullOrWhiteSpace(input))
            {
                continue;
            }

            if (IsQuit(input))
            {
                return 0;
            }

            if (IsBack(input))
            {
                return 0;
            }

            if (!int.TryParse(input, out var pageIndex) || pageIndex < 1 || pageIndex > pages.Count)
            {
                Console.WriteLine("Invalid selection.");
                continue;
            }

            var page = pages[pageIndex - 1];
            RunPropertyWizardForPage(page, null);
        }
    }

    private static dynamic? SelectProjectInteractively(dynamic solution)
    {
        var projectCollection = TryGetValue(() => (object?)solution.Projects);
        var projects = EnumerateProjects(projectCollection)
            .Select(project => new ProjectEntry(project,
                TryGetValue(() => (string?)project.Name) ?? "(unknown)",
                TryGetValue(() => (string?)project.UniqueName)))
            .OrderBy(p => p.DisplayName, StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (projects.Count == 0)
        {
            Console.Error.WriteLine("No projects found in the active solution.");
            return null;
        }

        while (true)
        {
            Console.WriteLine("*select a project:");
            for (var i = 0; i < projects.Count; i++)
            {
                Console.WriteLine($"{i + 1}. {projects[i].DisplayName}");
            }

            Console.Write("Select project (q to quit): ");
            var input = Console.ReadLine()?.Trim();
            if (string.IsNullOrWhiteSpace(input))
            {
                continue;
            }

            if (IsQuit(input))
            {
                return null;
            }

            if (int.TryParse(input, out var index) && index > 0 && index <= projects.Count)
            {
                return projects[index - 1].Project;
            }

            Console.WriteLine("Invalid selection.");
        }
    }

    private static dynamic? ResolveManifestProject(dynamic solution, string platform, string? selector)
    {
        if (!string.IsNullOrWhiteSpace(selector))
        {
            return FindProjectBySelector(solution, selector);
        }

        var candidates = FindProjectsByPlatform(solution, platform).ToList();
        if (candidates.Count == 0)
        {
            return null;
        }

        if (candidates.Count == 1)
        {
            return candidates[0].Project;
        }

        return SelectProjectInteractivelyFromList(candidates, $"select a {platform} project:");
    }

    private static List<ProjectEntry> FindProjectsByPlatform(dynamic solution, string platform)
    {
        var projectCollection = TryGetValue(() => (object?)solution.Projects);
        var projects = EnumerateProjects(projectCollection)
            .Select(project => new ProjectEntry(project,
                TryGetValue(() => (string?)project.Name) ?? "(unknown)",
                TryGetValue(() => (string?)project.UniqueName)))
            .Where(entry => ProjectMatchesPlatform(entry.Project, platform))
            .OrderBy(entry => entry.DisplayName, StringComparer.OrdinalIgnoreCase)
            .ToList();

        return projects;
    }

    private static bool ProjectMatchesPlatform(dynamic project, string platform)
    {
        var projectPath = GetProjectPath(project);
        var fileValues = new Dictionary<string, (string Value, string FilePath)>(StringComparer.OrdinalIgnoreCase);
        if (!string.IsNullOrWhiteSpace(projectPath) && File.Exists(projectPath))
        {
            fileValues = CollectPropertyValuesFromProjectFiles(projectPath);
        }

        var tfms = GetProjectTargetFrameworks((object?)project, fileValues);
        if (tfms.Count == 0)
        {
            return false;
        }

        return tfms.Any(tfm =>
        {
            if (string.IsNullOrWhiteSpace(tfm))
            {
                return false;
            }

            if (string.Equals(platform, "ios", StringComparison.OrdinalIgnoreCase))
            {
                return tfm.Contains("ios", StringComparison.OrdinalIgnoreCase);
            }

            if (string.Equals(platform, "android", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(platform, "droid", StringComparison.OrdinalIgnoreCase))
            {
                return tfm.Contains("android", StringComparison.OrdinalIgnoreCase);
            }

            return false;
        });
    }

    private static dynamic? SelectProjectInteractivelyFromList(List<ProjectEntry> projects, string title)
    {
        if (projects.Count == 0)
        {
            return null;
        }

        while (true)
        {
            Console.WriteLine($"*{title}");
            for (var i = 0; i < projects.Count; i++)
            {
                Console.WriteLine($"{i + 1}. {projects[i].DisplayName}");
            }

            Console.Write("Select project (q to quit): ");
            var input = Console.ReadLine()?.Trim();
            if (string.IsNullOrWhiteSpace(input))
            {
                continue;
            }

            if (IsQuit(input))
            {
                return null;
            }

            if (int.TryParse(input, out var index) && index > 0 && index <= projects.Count)
            {
                return projects[index - 1].Project;
            }

            Console.WriteLine("Invalid selection.");
        }
    }

    private static int RunPropertyWizardForProject(dynamic project, bool includeMsBuild, UiaPropertySession? uiaSession)
    {
        var pages = GetProjectPropertyPages((object?)project, includeMsBuild);
        if (pages.Count == 0)
        {
            Console.WriteLine("No project properties found.");
            return 1;
        }

        pages = pages.Where(p => p.Items.Count > 0).ToList();
        if (pages.Count == 0)
        {
            Console.WriteLine("No project properties found.");
            return 1;
        }

        if (uiaSession != null)
        {
            OverlayUiaValues(pages, uiaSession);
        }

        while (true)
        {
            Console.WriteLine("*select a project property page:");
            for (var i = 0; i < pages.Count; i++)
            {
                Console.WriteLine($"{i + 1}. {pages[i].Name}");
            }

            Console.Write("Select page (b to go back, q to quit): ");
            var input = Console.ReadLine()?.Trim();
            if (string.IsNullOrWhiteSpace(input))
            {
                continue;
            }

            if (IsQuit(input))
            {
                return 0;
            }

            if (IsBack(input))
            {
                return 0;
            }

            if (!int.TryParse(input, out var pageIndex) || pageIndex < 1 || pageIndex > pages.Count)
            {
                Console.WriteLine("Invalid selection.");
                continue;
            }

            var page = pages[pageIndex - 1];
            if (RunPropertyWizardForPage(page, uiaSession) == 0)
            {
                continue;
            }
        }
    }

    private static int RunPropertyWizardForPage(PropertyPage page, UiaPropertySession? uiaSession)
    {
        var filter = string.Empty;
        while (true)
        {
            var items = string.IsNullOrWhiteSpace(filter)
                ? page.Items
                : page.Items.Where(p => p.Name.Contains(filter, StringComparison.OrdinalIgnoreCase)).ToList();

            Console.WriteLine($"*select a property ({page.Name}):");
            if (!string.IsNullOrWhiteSpace(filter))
            {
                Console.WriteLine($"  filter: {filter}");
            }

            if (items.Count == 0)
            {
                Console.WriteLine("  (no properties match the filter)");
            }
            else
            {
                for (var i = 0; i < items.Count; i++)
                {
                    var value = SafeFormatPropertyValue(items[i]);
                    Console.WriteLine($"{i + 1}. {items[i].Name} = {value}");
                }
            }

            Console.Write("Select property (/text to filter, * to clear, b to go back, q to quit): ");
            var input = Console.ReadLine()?.Trim();
            if (string.IsNullOrWhiteSpace(input))
            {
                continue;
            }

            if (IsQuit(input))
            {
                return 0;
            }

            if (IsBack(input))
            {
                return 1;
            }

            if (input == "*")
            {
                filter = string.Empty;
                continue;
            }

            if (input.StartsWith("/", StringComparison.Ordinal))
            {
                filter = input.Substring(1).Trim();
                continue;
            }

            if (!int.TryParse(input, out var propIndex) || propIndex < 1 || propIndex > items.Count)
            {
                Console.WriteLine("Invalid selection.");
                continue;
            }

            var item = items[propIndex - 1];
            UpdatePropertyValue(page, item, uiaSession);
        }
    }

    private static void UpdatePropertyValue(PropertyPage page, PropertyEntry entry, UiaPropertySession? uiaSession)
    {
        var currentValue = GetPropertyValue(entry);
        var formattedCurrent = FormatPropertyValue(currentValue);

        var isReadOnly = IsPropertyReadOnly(entry);

        if (isReadOnly)
        {
            Console.WriteLine($"Property '{entry.Name}' is read-only.");
            return;
        }

        Console.WriteLine($"*update {entry.Name} value:");
        if (!string.IsNullOrWhiteSpace(entry.Description))
        {
            Console.WriteLine($"-info: {entry.Description}");
        }

        if (entry.AllowedValues != null && entry.AllowedValues.Count > 0)
        {
            Console.WriteLine($"-allowed: {string.Join(", ", entry.AllowedValues)}");
        }

        Console.Write("-enter new value ");
        Console.Write($"<{formattedCurrent}> ");
        Console.Write("(use !delete to remove): ");
        var input = Console.ReadLine();
        if (input == null)
        {
            return;
        }

        if (IsQuit(input))
        {
            Environment.Exit(0);
        }

        if (IsBack(input))
        {
            return;
        }

        var newValue = input.Trim();
        if (IsDeleteCommand(newValue))
        {
            if (TryDeleteProperty(entry, page, uiaSession, out var deleteMessage))
            {
                Console.WriteLine(deleteMessage ?? "Property deleted successfully.");
            }
            else
            {
                Console.WriteLine(deleteMessage ?? "Failed to delete property.");
            }

            return;
        }
        if (string.IsNullOrEmpty(newValue) ||
            string.Equals(newValue, formattedCurrent, StringComparison.OrdinalIgnoreCase))
        {
            try
            {
                SetPropertyValue(entry, currentValue);
                Console.WriteLine("Value unchanged.");
            }
            catch (COMException ex)
            {
                Console.WriteLine($"Failed to update property: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to update property: {ex.Message}");
            }

            return;
        }

        if (!TryConvertPropertyValue(entry, newValue, currentValue, out object? converted, out string hint))
        {
            Console.WriteLine($"Invalid value. {hint}");
            return;
        }

        try
        {
            SetPropertyValue(entry, converted);
            Console.WriteLine("Property updated successfully.");
            if (entry.UiaValue != null)
            {
                entry.UiaValue = converted;
            }
        }
        catch (COMException ex)
        {
            if (TrySetPropertyViaUiAutomation(page, entry, converted, uiaSession, out var message))
            {
                Console.WriteLine(message ?? "Property updated successfully via UI Automation.");
                return;
            }

            Console.WriteLine($"Failed to update property: {ex.Message}");
        }
        catch (Exception ex)
        {
            if (TrySetPropertyViaUiAutomation(page, entry, converted, uiaSession, out var message))
            {
                Console.WriteLine(message ?? "Property updated successfully via UI Automation.");
                return;
            }

            Console.WriteLine($"Failed to update property: {ex.Message}");
        }
    }

    private static bool TryConvertPropertyValue(PropertyEntry entry, string input, object? currentValue, out object? converted, out string hint)
    {
        hint = string.Empty;
        converted = input;

        if (entry.AllowedValueMap != null)
        {
            var key = input.Trim();
            if (entry.AllowedValueMap.TryGetValue(key, out var mapped))
            {
                converted = mapped;
                return true;
            }

            if (entry.AllowedValues != null && entry.AllowedValues.Count > 0)
            {
                hint = $"Expected one of: {string.Join(", ", entry.AllowedValues)}";
                return false;
            }
        }

        var targetType = currentValue?.GetType() ?? entry.PropertyType ?? entry.Descriptor?.PropertyType;

        if (targetType == null || targetType == typeof(string))
        {
            return true;
        }

        if (targetType == typeof(bool))
        {
            if (TryParseBool(input, out var boolValue))
            {
                converted = boolValue;
                return true;
            }

            hint = "Expected true/false.";
            return false;
        }

        if (targetType.IsEnum)
        {
            try
            {
                converted = Enum.Parse(targetType, input, true);
                return true;
            }
            catch
            {
                hint = $"Expected one of: {string.Join(", ", Enum.GetNames(targetType))}";
                return false;
            }
        }

        var converter = entry.Descriptor?.Converter ?? TypeDescriptor.GetConverter(targetType);
        if (converter != null && converter.CanConvertFrom(typeof(string)))
        {
            try
            {
                converted = converter.ConvertFromInvariantString(input);
                return true;
            }
            catch
            {
                if (converter.GetStandardValuesSupported())
                {
                    var values = converter.GetStandardValues();
                    if (values != null)
                    {
                        hint = $"Expected one of: {string.Join(", ", values.Cast<object>())}";
                        return false;
                    }
                }

                hint = $"Expected {targetType.Name} value.";
                return false;
            }
        }

        try
        {
            converted = Convert.ChangeType(input, targetType);
            return true;
        }
        catch
        {
            hint = $"Expected {targetType.Name} value.";
            return false;
        }
    }

    private static bool TryParseBool(string input, out bool value)
    {
        var normalized = input.Trim().ToLowerInvariant();
        switch (normalized)
        {
            case "true":
            case "t":
            case "yes":
            case "y":
            case "1":
                value = true;
                return true;
            case "false":
            case "f":
            case "no":
            case "n":
            case "0":
                value = false;
                return true;
            default:
                value = false;
                return false;
        }
    }

    private static bool IsQuit(string input)
    {
        return string.Equals(input.Trim(), "q", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsBack(string input)
    {
        return string.Equals(input.Trim(), "b", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsDeleteCommand(string input)
    {
        if (string.IsNullOrWhiteSpace(input))
        {
            return false;
        }

        var normalized = input.Trim();
        return string.Equals(normalized, "!delete", StringComparison.OrdinalIgnoreCase)
            || string.Equals(normalized, "!del", StringComparison.OrdinalIgnoreCase);
    }

    private static bool MsBuildRegistrationAttempted;
    private static bool MsBuildRegistrationSucceeded;
    private static bool MsBuildRegistrationLogged;
    private static bool MsBuildEvaluationLogged;
    private static bool UiaUnavailableLogged;
    private static bool AssemblyResolversRegistered;
    private static bool DotnetSdkEnvironmentResolved;
    private static string? CachedDotnetSdksPath;
    private static string? CachedDotnetSdkVersion;
    private static string? CachedDotnetSdkRootPath;
    private static bool NuGetAssembliesEnsured;
    private static readonly Dictionary<string, List<RuleDefinition>> RuleDefinitionCache = new(StringComparer.OrdinalIgnoreCase);
    private static readonly object RuleDefinitionLock = new();

    private static readonly string[] KnownPropertyPages =
    {
        "Application",
        "Application/General",
        "Application/Win32 Resources",
        "Application/Dependencies",
        "Application/iOS Targets",
        "Application/Android Targets",
        "Application/Windows Targets",
        "Application/Tizen Targets",
        "Global Usings",
        "Global Usings/General",
        "Build",
        "Build/General",
        "Build/Errors and warnings",
        "Build/Output",
        "Build/Events",
        "Build/Publish",
        "Build/Strong naming",
        "Build/Advanced",
        "Package",
        "Package/General",
        "Package/License",
        "Package/Symbols",
        "Code Analysis",
        "Code Analysis/All analyzers",
        "Code Analysis/.NET analyzers",
        "Resources",
        "Resources/General",
        "MAUI Shared",
        "MAUI Shared/General",
        "iOS",
        "ios/Build",
        "ios/Bundle Signing",
        "ios/debug",
        "ios/IPA Options",
        "ios/Manifest",
        "ios/On Demand Resources",
        "ios/Run Options"
    };

    private static readonly string[] IosPropertyNameHints =
    {
        "CodesignKey",
        "CodesignKeychain",
        "CodesignProvision",
        "CodesignEntitlements",
        "CodesignExtraArgs",
        "DevelopmentTeamId",
        "ProvisioningType",
        "BuildIpa",
        "IpaPackageDir",
        "IpaPackageName",
        "IpaPackagePath",
        "ArchiveOnBuild",
        "MtouchArch",
        "MtouchLink",
        "MtouchDebug",
        "MtouchSdkVersion",
        "MtouchUseLlvm",
        "MtouchUseBitcode",
        "MtouchNoSymbolStrip",
        "MtouchExtraArgs",
        "MtouchEnableSGenConc",
        "MtouchProfiling",
        "MtouchFastDev",
        "MtouchFloat32",
        "MtouchHttpClientHandlerType",
        "MtouchTlsProvider",
        "MtouchInterpreter",
        "MtouchUseInterpreter",
        "MtouchEnableBitcode",
        "DeviceSpecificBuild",
        "EmbedOnDemandResources",
        "IOSDebuggerConnectOverUsb",
        "IOSDebuggerPort",
        "IOSDebugOverWiFi",
        "IpaIncludeArtwork",
        "OnDemandResourcesInitialInstallTags",
        "OnDemandResourcesPrefetchOrder",
        "OnDemandResourcesUrl"
    };

    private static readonly string[] AndroidPropertyNameHints =
    {
        "AndroidKeyStore",
        "AndroidSigningKeyStore",
        "AndroidSigningKeyAlias",
        "AndroidSigningKeyPass",
        "AndroidSigningStorePass",
        "AndroidUseAapt2",
        "AndroidPackageFormat",
        "AndroidManifest",
        "AndroidLinkMode",
        "AndroidEnableProfiledAot",
        "AndroidEnablePreloadAssemblies",
        "AndroidDexTool",
        "AndroidDexToolExecutable",
        "AndroidUseSharedRuntime",
        "AndroidSdkBuildToolsVersion",
        "AndroidSdkEmulatorVersion",
        "AndroidSdkPlatformToolsVersion",
        "AndroidSdkPlatformVersion",
        "AndroidNdkVersion",
        "AndroidUseApkSigner",
        "AndroidPackageNamingPolicy"
    };

    private static readonly HashSet<string> KnownPropertyNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "AssemblyName",
        "RootNamespace",
        "OutputType",
        "TargetFramework",
        "TargetFrameworks",
        "TargetFrameworkIdentifier",
        "TargetFrameworkVersion",
        "TargetFrameworkMoniker",
        "TargetPlatformIdentifier",
        "TargetPlatformVersion",
        "SupportedOSPlatformVersion",
        "ApplicationIcon",
        "ApplicationManifest",
        "StartupObject",
        "AssemblyVersion",
        "FileVersion",
        "Version",
        "VersionPrefix",
        "VersionSuffix",
        "Company",
        "Product",
        "Description",
        "Copyright",
        "NeutralLanguage",
        "GenerateAssemblyInfo",
        "GenerateDocumentationFile",
        "DocumentationFile",
        "DefineConstants",
        "Optimize",
        "DebugType",
        "DebugSymbols",
        "WarningLevel",
        "TreatWarningsAsErrors",
        "TreatSpecificWarningsAsErrors",
        "NoWarn",
        "Nullable",
        "LangVersion",
        "PlatformTarget",
        "Prefer32Bit",
        "AllowUnsafeBlocks",
        "Deterministic",
        "CheckForOverflowUnderflow",
        "IntermediateOutputPath",
        "OutputPath",
        "BaseIntermediateOutputPath",
        "BaseOutputPath",
        "RunAnalyzersDuringBuild",
        "RunAnalyzersDuringLiveAnalysis",
        "AnalysisLevel",
        "AnalysisMode",
        "EnableNETAnalyzers",
        "ImplicitUsings",
        "PackageId",
        "PackageVersion",
        "PackageDescription",
        "PackageTags",
        "PackageProjectUrl",
        "PackageReleaseNotes",
        "PackageIcon",
        "PackageLicenseExpression",
        "PackageLicenseFile",
        "PackageRequireLicenseAcceptance",
        "GeneratePackageOnBuild",
        "PackageOutputPath",
        "PackageReadmeFile",
        "RepositoryUrl",
        "RepositoryType",
        "IncludeSymbols",
        "SymbolPackageFormat",
        "SignAssembly",
        "DelaySign",
        "AssemblyOriginatorKeyFile",
        "UseMaui",
        "ApplicationTitle",
        "ApplicationId",
        "ApplicationVersion",
        "ApplicationDisplayVersion",
        "RuntimeIdentifier",
        "RuntimeIdentifiers",
        "CodesignKey",
        "CodesignProvision",
        "CodesignEntitlements",
        "DevelopmentTeamId",
        "AndroidKeyStore",
        "AndroidSigningKeyStore",
        "AndroidSigningKeyAlias",
        "AndroidSigningKeyPass",
        "AndroidSigningStorePass"
    };

    private static List<PropertyPage> GetProjectPropertyPages(object? project, bool includeMsBuild)
    {
        var pages = new List<PropertyPage>();
        foreach (var known in KnownPropertyPages)
        {
            pages.Add(new PropertyPage(known, new List<PropertyEntry>()));
        }

        dynamic? dynamicProject = project;
        MsBuildEvaluation? msbuildEvaluation = null;
        Dictionary<string, MsBuildPropertyValue>? msbuildMap = null;
        HashSet<string>? platformPropertyNames = null;
        HashSet<string>? additionalSchemaFiles = null;
        var projectPath = GetProjectPath(dynamicProject);
        var fileValues = new Dictionary<string, (string Value, string FilePath)>(StringComparer.OrdinalIgnoreCase);
        if (!string.IsNullOrWhiteSpace(projectPath) && File.Exists(projectPath))
        {
            fileValues = CollectPropertyValuesFromProjectFiles(projectPath);
        }

        var globalProperties = GetMsBuildGlobalProperties(dynamicProject);
        var targetFrameworks = GetProjectTargetFrameworks((object?)dynamicProject, fileValues);
        var primaryTargetFramework = targetFrameworks.FirstOrDefault();
        if (!string.IsNullOrWhiteSpace(primaryTargetFramework))
        {
            globalProperties["TargetFramework"] = primaryTargetFramework;
        }

        if (includeMsBuild)
        {
            msbuildEvaluation = TryCreateMsBuildEvaluation(dynamicProject, globalProperties);
            if (msbuildEvaluation != null)
            {
                platformPropertyNames = ExtractPlatformPropertyNames(msbuildEvaluation.Project);

                if (targetFrameworks.Count > 1)
                {
                    var additionalPlatforms = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    additionalSchemaFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                    foreach (var framework in targetFrameworks)
                    {
                        if (string.IsNullOrWhiteSpace(framework) ||
                            string.Equals(framework, primaryTargetFramework, StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        var globals = new Dictionary<string, string>(globalProperties, StringComparer.OrdinalIgnoreCase)
                        {
                            ["TargetFramework"] = framework
                        };

                        using var eval = TryCreateMsBuildEvaluation(dynamicProject, globals);
                        if (eval == null)
                        {
                            continue;
                        }

                        foreach (var schema in GetPropertyPageSchemaFiles(eval.Project))
                        {
                            additionalSchemaFiles.Add(schema);
                        }

                        foreach (var name in ExtractPlatformPropertyNames(eval.Project))
                        {
                            additionalPlatforms.Add(name);
                        }
                    }

                    if (additionalSchemaFiles.Count == 0)
                    {
                        additionalSchemaFiles = null;
                    }

                    if (additionalPlatforms.Count > 0)
                    {
                        if (platformPropertyNames == null)
                        {
                            platformPropertyNames = additionalPlatforms;
                        }
                        else
                        {
                            foreach (var name in additionalPlatforms)
                            {
                                platformPropertyNames.Add(name);
                            }
                        }
                    }
                }
            }
        }

        var ruleDefinitions = LoadRuleDefinitions(projectPath, msbuildEvaluation?.Project, additionalSchemaFiles);

        if (includeMsBuild && msbuildEvaluation != null)
        {
            var allowedNames = ShouldIncludeAllMsBuildProperties()
                ? null
                : BuildAllowedMsBuildPropertyNames(projectPath, ruleDefinitions, platformPropertyNames);
            msbuildMap = BuildMsBuildPropertyMap(msbuildEvaluation.Project, allowedNames);
        }

        AddPropertyPagesFromRules(pages, ruleDefinitions, msbuildMap, fileValues, projectPath, globalProperties);
        AddPropertyPagesFromPlatformDefinitions(pages, platformPropertyNames, msbuildMap, fileValues, projectPath, globalProperties);

        if (includeMsBuild && ShouldIncludeAllMsBuildProperties() && msbuildMap != null)
        {
            AddPropertyPagesFromMsBuild(pages, dynamicProject, msbuildMap);
        }
        else if (!string.IsNullOrWhiteSpace(projectPath) && File.Exists(projectPath))
        {
            AddPropertyPagesFromProjectFilesFallback(pages, projectPath);
        }

        AddPropertyPagesFromUserFile(pages, projectPath);

        pages = DeduplicatePropertyPages(pages);

        foreach (var page in pages)
        {
            page.Items.Sort((a, b) => string.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase));
        }

        var ordered = pages
            .OrderBy(p =>
            {
                var index = Array.IndexOf(KnownPropertyPages, p.Name);
                return index < 0 ? int.MaxValue : index;
            })
            .ThenBy(p => p.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();

        msbuildEvaluation?.Dispose();
        return ordered;
    }

    private static List<PropertyPage> GetManifestPropertyPages(object? project, string platform, bool includeMsBuild, out string? manifestPath, out string? error)
    {
        manifestPath = null;
        error = null;

        dynamic? dynamicProject = project;
        var projectPath = GetProjectPath(dynamicProject);
        var fileValues = new Dictionary<string, (string Value, string FilePath)>(StringComparer.OrdinalIgnoreCase);
        if (!string.IsNullOrWhiteSpace(projectPath) && File.Exists(projectPath))
        {
            fileValues = CollectPropertyValuesFromProjectFiles(projectPath);
        }

        var globalProperties = GetMsBuildGlobalProperties(dynamicProject);
        var targetFrameworks = GetProjectTargetFrameworks(project, fileValues);
        var targetFramework = SelectTargetFrameworkForPlatform(targetFrameworks, platform);
        if (!string.IsNullOrWhiteSpace(targetFramework))
        {
            globalProperties["TargetFramework"] = targetFramework;
        }

        MsBuildEvaluation? msbuildEvaluation = null;
        if (includeMsBuild)
        {
            msbuildEvaluation = TryCreateMsBuildEvaluation(dynamicProject, globalProperties);
        }

        manifestPath = ResolveManifestPath(platform, projectPath, fileValues, msbuildEvaluation);
        msbuildEvaluation?.Dispose();

        if (string.IsNullOrWhiteSpace(manifestPath))
        {
            error = platform.Equals("ios", StringComparison.OrdinalIgnoreCase)
                ? "Info.plist not found for the selected project."
                : "AndroidManifest.xml not found for the selected project.";
            return new List<PropertyPage>();
        }

        if (platform.Equals("ios", StringComparison.OrdinalIgnoreCase))
        {
            return BuildIosManifestPages(manifestPath, out error);
        }

        return BuildAndroidManifestPages(manifestPath, out error);
    }

    private static string? SelectTargetFrameworkForPlatform(List<string> targetFrameworks, string platform)
    {
        if (targetFrameworks.Count == 0)
        {
            return null;
        }

        if (string.Equals(platform, "ios", StringComparison.OrdinalIgnoreCase))
        {
            return targetFrameworks.FirstOrDefault(tfm => tfm.Contains("ios", StringComparison.OrdinalIgnoreCase))
                   ?? targetFrameworks.FirstOrDefault();
        }

        if (string.Equals(platform, "android", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(platform, "droid", StringComparison.OrdinalIgnoreCase))
        {
            return targetFrameworks.FirstOrDefault(tfm => tfm.Contains("android", StringComparison.OrdinalIgnoreCase))
                   ?? targetFrameworks.FirstOrDefault();
        }

        return targetFrameworks.FirstOrDefault();
    }

    private static string? ResolveManifestPath(
        string platform,
        string? projectPath,
        Dictionary<string, (string Value, string FilePath)> fileValues,
        MsBuildEvaluation? msbuildEvaluation)
    {
        if (string.IsNullOrWhiteSpace(projectPath) || !File.Exists(projectPath))
        {
            return null;
        }

        var projectDir = Path.GetDirectoryName(projectPath) ?? string.Empty;
        string? candidate = null;

        if (platform.Equals("ios", StringComparison.OrdinalIgnoreCase))
        {
            candidate = ResolveManifestPathFromProperty("InfoPlist", projectDir, fileValues, msbuildEvaluation)
                ?? ResolveManifestPathFromProperty("ApplicationManifest", projectDir, fileValues, msbuildEvaluation)
                ?? ResolveManifestPathFromProperty("AppManifest", projectDir, fileValues, msbuildEvaluation)
                ?? ResolveManifestPathFromProperty("IosAppManifest", projectDir, fileValues, msbuildEvaluation);

            if (string.IsNullOrWhiteSpace(candidate))
            {
                candidate = FindManifestPathInProject(projectPath, "Info.plist");
            }

            candidate ??= FindFirstExistingPath(projectDir, new[]
            {
                Path.Combine("Platforms", "iOS", "Info.plist"),
                Path.Combine("Platforms", "ios", "Info.plist"),
                Path.Combine("Properties", "Info.plist"),
                "Info.plist"
            });
        }
        else
        {
            candidate = ResolveManifestPathFromProperty("AndroidManifest", projectDir, fileValues, msbuildEvaluation)
                ?? ResolveManifestPathFromProperty("ApplicationManifest", projectDir, fileValues, msbuildEvaluation);

            if (string.IsNullOrWhiteSpace(candidate))
            {
                candidate = FindManifestPathInProject(projectPath, "AndroidManifest.xml");
            }

            candidate ??= FindFirstExistingPath(projectDir, new[]
            {
                Path.Combine("Platforms", "Android", "AndroidManifest.xml"),
                Path.Combine("Properties", "AndroidManifest.xml"),
                "AndroidManifest.xml"
            });
        }

        return candidate;
    }

    private static string? ResolveManifestPathFromProperty(
        string propertyName,
        string projectDir,
        Dictionary<string, (string Value, string FilePath)> fileValues,
        MsBuildEvaluation? msbuildEvaluation)
    {
        if (msbuildEvaluation != null)
        {
            var value = msbuildEvaluation.Project.GetPropertyValue(propertyName);
            var resolved = ResolveManifestCandidate(value, projectDir);
            if (!string.IsNullOrWhiteSpace(resolved))
            {
                return resolved;
            }
        }

        if (fileValues.TryGetValue(propertyName, out var fileValue))
        {
            return ResolveManifestCandidate(fileValue.Value, projectDir);
        }

        return null;
    }

    private static string? ResolveManifestCandidate(string? value, string projectDir)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        var trimmed = value.Trim();
        var path = Path.IsPathRooted(trimmed) ? trimmed : Path.Combine(projectDir, trimmed);
        if (File.Exists(path))
        {
            return path;
        }

        return null;
    }

    private static string? FindManifestPathInProject(string projectPath, string fileName)
    {
        try
        {
            var root = ProjectRootElement.Open(projectPath);
            foreach (var item in root.Items)
            {
                var include = item.Include;
                if (string.IsNullOrWhiteSpace(include))
                {
                    continue;
                }

                if (include.EndsWith(fileName, StringComparison.OrdinalIgnoreCase))
                {
                    var projectDir = Path.GetDirectoryName(projectPath) ?? string.Empty;
                    var candidate = Path.IsPathRooted(include) ? include : Path.Combine(projectDir, include);
                    if (File.Exists(candidate))
                    {
                        return candidate;
                    }
                }
            }
        }
        catch
        {
            return null;
        }

        return null;
    }

    private static string? FindFirstExistingPath(string projectDir, IEnumerable<string> relativePaths)
    {
        foreach (var relative in relativePaths)
        {
            var candidate = Path.Combine(projectDir, relative);
            if (File.Exists(candidate))
            {
                return candidate;
            }
        }

        return null;
    }

    private static List<PropertyPage> BuildIosManifestPages(string manifestPath, out string? error)
    {
        error = null;

        try
        {
            var entries = ReadPlistEntries(manifestPath, out error);
            if (!string.IsNullOrWhiteSpace(error))
            {
                return new List<PropertyPage>();
            }

            var page = new PropertyPage("ios/Manifest", new List<PropertyEntry>());
            foreach (var entry in entries)
            {
                var key = entry.Key;
                var valueKind = entry.Kind;
                object? cachedValue = entry.Value;
                var description = DescribePlistKind(valueKind);

                Action<object?> setter = newValue =>
                {
                    if (!TrySetPlistValue(manifestPath, key, newValue, valueKind, out var message))
                    {
                        throw new InvalidOperationException(message ?? "Failed to update Info.plist.");
                    }

                    cachedValue = newValue;
                };

                Func<object?> getter = () => cachedValue;
                var entryItem = new PropertyEntry(key, key, page.Name, getter, setter, GetPlistClrType(valueKind), PropertySource.MsBuild)
                {
                    Description = description
                };

                entryItem.DeleteAction = () =>
                {
                    var success = RemovePlistKey(manifestPath, key, out var message);
                    if (success)
                    {
                        cachedValue = null;
                    }

                    return (success, message);
                };

                page.Items.Add(entryItem);
            }

            return new List<PropertyPage> { page };
        }
        catch (Exception ex)
        {
            error = $"Failed to read Info.plist: {ex.Message}";
            return new List<PropertyPage>();
        }
    }

    private static List<PropertyPage> BuildAndroidManifestPages(string manifestPath, out string? error)
    {
        error = null;

        try
        {
            var doc = XDocument.Load(manifestPath, LoadOptions.PreserveWhitespace);
            var manifest = doc.Root;
            if (manifest == null || !string.Equals(manifest.Name.LocalName, "manifest", StringComparison.OrdinalIgnoreCase))
            {
                error = "AndroidManifest.xml root element not found.";
                return new List<PropertyPage>();
            }

            var pages = new List<PropertyPage>();
            var androidNs = XNamespace.Get("http://schemas.android.com/apk/res/android");

            var manifestPage = new PropertyPage("droid/Manifest", new List<PropertyEntry>());
            manifestPage.Items.AddRange(BuildAndroidAttributeEntries(manifestPath, "manifest", manifest, androidNs));
            pages.Add(manifestPage);

            var appElement = manifest.Elements().FirstOrDefault(el => string.Equals(el.Name.LocalName, "application", StringComparison.OrdinalIgnoreCase));
            if (appElement != null)
            {
                var appPage = new PropertyPage("droid/Application", new List<PropertyEntry>());
                appPage.Items.AddRange(BuildAndroidAttributeEntries(manifestPath, "application", appElement, androidNs));
                pages.Add(appPage);
            }

            var sdkElement = manifest.Elements().FirstOrDefault(el => string.Equals(el.Name.LocalName, "uses-sdk", StringComparison.OrdinalIgnoreCase));
            var sdkPage = new PropertyPage("droid/SDK", new List<PropertyEntry>());
            sdkPage.Items.AddRange(BuildAndroidSdkEntries(manifestPath, androidNs, sdkElement));
            if (sdkPage.Items.Count > 0)
            {
                pages.Add(sdkPage);
            }

            var permissionsPage = new PropertyPage("droid/Permissions", new List<PropertyEntry>());
            permissionsPage.Items.Add(BuildAndroidListEntry(
                manifestPath,
                "Uses Permissions",
                "uses-permission",
                "droid/Permissions",
                manifest,
                androidNs,
                "uses-permission",
                "name"));
            pages.Add(permissionsPage);

            var featuresPage = new PropertyPage("droid/Features", new List<PropertyEntry>());
            featuresPage.Items.Add(BuildAndroidListEntry(
                manifestPath,
                "Uses Features",
                "uses-feature",
                "droid/Features",
                manifest,
                androidNs,
                "uses-feature",
                "name"));
            pages.Add(featuresPage);

            return pages;
        }
        catch (Exception ex)
        {
            error = $"Failed to read AndroidManifest.xml: {ex.Message}";
            return new List<PropertyPage>();
        }
    }

    private static List<PropertyEntry> BuildAndroidAttributeEntries(string manifestPath, string elementName, XElement element, XNamespace androidNs)
    {
        var entries = new List<PropertyEntry>();
        foreach (var attribute in element.Attributes())
        {
            if (attribute.IsNamespaceDeclaration)
            {
                continue;
            }

            var name = attribute.Name.NamespaceName == androidNs.NamespaceName
                ? $"android:{attribute.Name.LocalName}"
                : attribute.Name.LocalName;

            object? cachedValue = attribute.Value;
            var propertyType = TryParseBool(attribute.Value, out _) ? typeof(bool) : typeof(string);

            Action<object?> setter = newValue =>
            {
                var textValue = newValue?.ToString() ?? string.Empty;
                SetAndroidAttribute(manifestPath, elementName, attribute.Name, textValue, androidNs);
                cachedValue = textValue;
            };

            Func<object?> getter = () => cachedValue;
            var entry = new PropertyEntry(name, name, elementName, getter, setter, propertyType, PropertySource.MsBuild)
            {
                Description = $"{elementName} attribute"
            };

            entry.DeleteAction = () =>
            {
                var success = RemoveAndroidAttribute(manifestPath, elementName, attribute.Name, androidNs, out var message);
                if (success)
                {
                    cachedValue = null;
                }

                return (success, message);
            };

            entries.Add(entry);
        }

        return entries;
    }

    private static List<PropertyEntry> BuildAndroidSdkEntries(string manifestPath, XNamespace androidNs, XElement? sdkElement)
    {
        var entries = new List<PropertyEntry>();
        var sdkAttributes = new[] { "minSdkVersion", "targetSdkVersion", "maxSdkVersion" };

        foreach (var attrName in sdkAttributes)
        {
            var xname = androidNs + attrName;
            object? cachedValue = sdkElement?.Attribute(xname)?.Value;

            Action<object?> setter = newValue =>
            {
                var textValue = newValue?.ToString() ?? string.Empty;
                SetAndroidSdkAttribute(manifestPath, xname, textValue, androidNs);
                cachedValue = textValue;
            };

            Func<object?> getter = () => cachedValue;
            var entry = new PropertyEntry($"android:{attrName}", $"android:{attrName}", "droid/SDK", getter, setter, typeof(string), PropertySource.MsBuild)
            {
                Description = "uses-sdk attribute"
            };

            entry.DeleteAction = () =>
            {
                var success = RemoveAndroidSdkAttribute(manifestPath, xname, androidNs, out var message);
                if (success)
                {
                    cachedValue = null;
                }

                return (success, message);
            };

            entries.Add(entry);
        }

        return entries;
    }

    private static PropertyEntry BuildAndroidListEntry(
        string manifestPath,
        string displayName,
        string key,
        string pageName,
        XElement manifest,
        XNamespace androidNs,
        string elementName,
        string attributeName)
    {
        object? cachedValue = GetAndroidListValue(manifest, androidNs, elementName, attributeName);
        Action<object?> setter = newValue =>
        {
            var textValue = newValue?.ToString() ?? string.Empty;
            SetAndroidListValue(manifestPath, androidNs, elementName, attributeName, textValue);
            cachedValue = textValue;
        };

        Func<object?> getter = () => cachedValue;
        return new PropertyEntry(displayName, key, pageName, getter, setter, typeof(string), PropertySource.MsBuild)
        {
            Description = $"Semicolon-separated list for {elementName}"
        };
    }

    private static void SetAndroidAttribute(string manifestPath, string elementName, XName attributeName, string value, XNamespace androidNs)
    {
        var doc = XDocument.Load(manifestPath, LoadOptions.PreserveWhitespace);
        var manifest = doc.Root;
        if (manifest == null)
        {
            throw new InvalidOperationException("AndroidManifest.xml root element not found.");
        }

        XElement? target = manifest;
        if (!string.Equals(elementName, "manifest", StringComparison.OrdinalIgnoreCase))
        {
            target = manifest.Elements().FirstOrDefault(el => string.Equals(el.Name.LocalName, elementName, StringComparison.OrdinalIgnoreCase));
        }

        if (target == null)
        {
            throw new InvalidOperationException($"Element '{elementName}' not found.");
        }

        target.SetAttributeValue(attributeName, value);
        doc.Save(manifestPath);
    }

    private static bool RemoveAndroidAttribute(string manifestPath, string elementName, XName attributeName, XNamespace androidNs, out string? message)
    {
        message = null;

        try
        {
            var doc = XDocument.Load(manifestPath, LoadOptions.PreserveWhitespace);
            var manifest = doc.Root;
            if (manifest == null)
            {
                message = "AndroidManifest.xml root element not found.";
                return false;
            }

            XElement? target = manifest;
            if (!string.Equals(elementName, "manifest", StringComparison.OrdinalIgnoreCase))
            {
                target = manifest.Elements().FirstOrDefault(el => string.Equals(el.Name.LocalName, elementName, StringComparison.OrdinalIgnoreCase));
            }

            if (target == null)
            {
                message = $"Element '{elementName}' not found.";
                return false;
            }

            var attr = target.Attribute(attributeName);
            if (attr == null)
            {
                message = "Attribute not found.";
                return false;
            }

            attr.Remove();
            doc.Save(manifestPath);
            return true;
        }
        catch (Exception ex)
        {
            message = ex.Message;
            return false;
        }
    }

    private static void SetAndroidSdkAttribute(string manifestPath, XName attributeName, string value, XNamespace androidNs)
    {
        var doc = XDocument.Load(manifestPath, LoadOptions.PreserveWhitespace);
        var manifest = doc.Root;
        if (manifest == null)
        {
            throw new InvalidOperationException("AndroidManifest.xml root element not found.");
        }

        var sdk = manifest.Elements().FirstOrDefault(el => string.Equals(el.Name.LocalName, "uses-sdk", StringComparison.OrdinalIgnoreCase));
        if (sdk == null)
        {
            sdk = new XElement("uses-sdk");
            manifest.AddFirst(sdk);
        }

        sdk.SetAttributeValue(attributeName, value);
        doc.Save(manifestPath);
    }

    private static bool RemoveAndroidSdkAttribute(string manifestPath, XName attributeName, XNamespace androidNs, out string? message)
    {
        message = null;

        try
        {
            var doc = XDocument.Load(manifestPath, LoadOptions.PreserveWhitespace);
            var manifest = doc.Root;
            if (manifest == null)
            {
                message = "AndroidManifest.xml root element not found.";
                return false;
            }

            var sdk = manifest.Elements().FirstOrDefault(el => string.Equals(el.Name.LocalName, "uses-sdk", StringComparison.OrdinalIgnoreCase));
            if (sdk == null)
            {
                message = "uses-sdk element not found.";
                return false;
            }

            var attr = sdk.Attribute(attributeName);
            if (attr == null)
            {
                message = "Attribute not found.";
                return false;
            }

            attr.Remove();
            doc.Save(manifestPath);
            return true;
        }
        catch (Exception ex)
        {
            message = ex.Message;
            return false;
        }
    }

    private static string GetAndroidListValue(XElement manifest, XNamespace androidNs, string elementName, string attributeName)
    {
        var values = manifest.Elements()
            .Where(el => string.Equals(el.Name.LocalName, elementName, StringComparison.OrdinalIgnoreCase))
            .Select(el => el.Attribute(androidNs + attributeName)?.Value)
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .ToList();

        return string.Join("; ", values);
    }

    private static void SetAndroidListValue(string manifestPath, XNamespace androidNs, string elementName, string attributeName, string value)
    {
        var doc = XDocument.Load(manifestPath, LoadOptions.PreserveWhitespace);
        var manifest = doc.Root;
        if (manifest == null)
        {
            throw new InvalidOperationException("AndroidManifest.xml root element not found.");
        }

        var items = SplitManifestList(value);
        var existing = manifest.Elements().Where(el => string.Equals(el.Name.LocalName, elementName, StringComparison.OrdinalIgnoreCase)).ToList();
        foreach (var node in existing)
        {
            node.Remove();
        }

        foreach (var item in items)
        {
            var element = new XElement(elementName);
            element.SetAttributeValue(androidNs + attributeName, item);
            manifest.Add(element);
        }

        doc.Save(manifestPath);
    }

    private static List<string> SplitManifestList(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return new List<string>();
        }

        return value
            .Split(new[] { ';', ',', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(item => item.Trim())
            .Where(item => !string.IsNullOrWhiteSpace(item))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private sealed class PlistEntry
    {
        public string Key { get; set; } = string.Empty;
        public object? Value { get; set; }
        public PlistValueKind Kind { get; set; }
    }

    private enum PlistValueKind
    {
        String,
        Integer,
        Real,
        Boolean,
        Date,
        Data,
        Array,
        Dictionary,
        Unknown
    }

    private static List<PlistEntry> ReadPlistEntries(string manifestPath, out string? error)
    {
        error = null;
        var entries = new List<PlistEntry>();

        var doc = XDocument.Load(manifestPath, LoadOptions.PreserveWhitespace);
        var dict = doc.Root?.Elements().FirstOrDefault(el => string.Equals(el.Name.LocalName, "dict", StringComparison.OrdinalIgnoreCase))
                   ?? doc.Descendants().FirstOrDefault(el => string.Equals(el.Name.LocalName, "dict", StringComparison.OrdinalIgnoreCase));
        if (dict == null)
        {
            error = "Info.plist dictionary root not found.";
            return entries;
        }

        var elements = dict.Elements().ToList();
        for (var i = 0; i < elements.Count; i++)
        {
            var element = elements[i];
            if (!string.Equals(element.Name.LocalName, "key", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var key = element.Value?.Trim();
            if (string.IsNullOrWhiteSpace(key))
            {
                continue;
            }

            if (i + 1 >= elements.Count)
            {
                continue;
            }

            var valueElement = elements[i + 1];
            var value = ParsePlistValueElement(valueElement, out var kind);
            entries.Add(new PlistEntry
            {
                Key = key,
                Value = value,
                Kind = kind
            });
        }

        return entries;
    }

    private static object? ParsePlistValueElement(XElement element, out PlistValueKind kind)
    {
        var name = element.Name.LocalName;
        switch (name)
        {
            case "string":
                kind = PlistValueKind.String;
                return element.Value;
            case "integer":
                kind = PlistValueKind.Integer;
                if (long.TryParse(element.Value, out var intValue))
                {
                    return intValue;
                }

                return element.Value;
            case "real":
                kind = PlistValueKind.Real;
                if (double.TryParse(element.Value, out var realValue))
                {
                    return realValue;
                }

                return element.Value;
            case "true":
                kind = PlistValueKind.Boolean;
                return true;
            case "false":
                kind = PlistValueKind.Boolean;
                return false;
            case "date":
                kind = PlistValueKind.Date;
                if (DateTime.TryParse(element.Value, out var dateValue))
                {
                    return dateValue;
                }

                return element.Value;
            case "data":
                kind = PlistValueKind.Data;
                return element.Value;
            case "array":
                kind = PlistValueKind.Array;
                return SerializePlistObjectToJson(ParsePlistArray(element));
            case "dict":
                kind = PlistValueKind.Dictionary;
                return SerializePlistObjectToJson(ParsePlistDictionary(element));
            default:
                kind = PlistValueKind.Unknown;
                return element.Value;
        }
    }

    private static object? ParsePlistArray(XElement element)
    {
        var list = new List<object?>();
        foreach (var child in element.Elements())
        {
            list.Add(ParsePlistValueElement(child, out _));
        }

        return list;
    }

    private static object? ParsePlistDictionary(XElement element)
    {
        var dict = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        var elements = element.Elements().ToList();
        for (var i = 0; i < elements.Count; i++)
        {
            if (!string.Equals(elements[i].Name.LocalName, "key", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var key = elements[i].Value;
            if (string.IsNullOrWhiteSpace(key) || i + 1 >= elements.Count)
            {
                continue;
            }

            var valueElement = elements[i + 1];
            dict[key] = ParsePlistValueElement(valueElement, out _);
        }

        return dict;
    }

    private static string SerializePlistObjectToJson(object? value)
    {
        try
        {
            return JsonSerializer.Serialize(value);
        }
        catch
        {
            return value?.ToString() ?? string.Empty;
        }
    }

    private static Type GetPlistClrType(PlistValueKind kind)
    {
        return kind switch
        {
            PlistValueKind.Boolean => typeof(bool),
            PlistValueKind.Integer => typeof(long),
            PlistValueKind.Real => typeof(double),
            PlistValueKind.Date => typeof(DateTime),
            _ => typeof(string)
        };
    }

    private static string? DescribePlistKind(PlistValueKind kind)
    {
        return kind switch
        {
            PlistValueKind.Array => "JSON array",
            PlistValueKind.Dictionary => "JSON object",
            PlistValueKind.Data => "Base64 data",
            PlistValueKind.Date => "ISO-8601 date",
            _ => null
        };
    }

    private static bool TrySetPlistValue(string manifestPath, string key, object? value, PlistValueKind kind, out string? message)
    {
        message = null;

        try
        {
            var doc = XDocument.Load(manifestPath, LoadOptions.PreserveWhitespace);
            var dict = doc.Root?.Elements().FirstOrDefault(el => string.Equals(el.Name.LocalName, "dict", StringComparison.OrdinalIgnoreCase))
                       ?? doc.Descendants().FirstOrDefault(el => string.Equals(el.Name.LocalName, "dict", StringComparison.OrdinalIgnoreCase));
            if (dict == null)
            {
                message = "Info.plist dictionary root not found.";
                return false;
            }

            XElement? keyElement = null;
            XElement? valueElement = null;
            var elements = dict.Elements().ToList();
            for (var i = 0; i < elements.Count; i++)
            {
                if (!string.Equals(elements[i].Name.LocalName, "key", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (string.Equals(elements[i].Value, key, StringComparison.OrdinalIgnoreCase))
                {
                    keyElement = elements[i];
                    if (i + 1 < elements.Count)
                    {
                        valueElement = elements[i + 1];
                    }

                    break;
                }
            }

            var newValueElement = CreatePlistValueElement(value, kind, out var errorMessage);
            if (newValueElement == null)
            {
                message = errorMessage ?? "Invalid plist value.";
                return false;
            }

            if (keyElement == null || valueElement == null)
            {
                dict.Add(new XElement("key", key));
                dict.Add(newValueElement);
            }
            else
            {
                valueElement.ReplaceWith(newValueElement);
            }

            doc.Save(manifestPath);
            return true;
        }
        catch (Exception ex)
        {
            message = ex.Message;
            return false;
        }
    }

    private static XElement? CreatePlistValueElement(object? value, PlistValueKind kind, out string? error)
    {
        error = null;
        try
        {
            if (kind == PlistValueKind.Unknown)
            {
                if (value is bool boolValue)
                {
                    return boolValue ? new XElement("true") : new XElement("false");
                }

                if (value is int or long)
                {
                    return new XElement("integer", value.ToString());
                }

                if (value is float or double or decimal)
                {
                    return new XElement("real", value.ToString());
                }

                if (value is Dictionary<string, object?> dictValue)
                {
                    return BuildPlistDictElement(dictValue);
                }

                if (value is List<object?> listValue)
                {
                    return BuildPlistArrayElement(listValue);
                }
            }

            if (kind == PlistValueKind.Array || kind == PlistValueKind.Dictionary)
            {
                var jsonText = value?.ToString() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(jsonText))
                {
                    return kind == PlistValueKind.Array ? new XElement("array") : new XElement("dict");
                }

                if (!TryParseJson(jsonText, out var jsonValue, out var jsonError))
                {
                    error = jsonError ?? "Invalid JSON.";
                    return null;
                }

                return kind == PlistValueKind.Array
                    ? BuildPlistArrayElement(jsonValue as List<object?>)
                    : BuildPlistDictElement(jsonValue as Dictionary<string, object?>);
            }

            if (kind == PlistValueKind.Boolean)
            {
                var boolValue = value is bool b ? b : TryParseBool(value?.ToString() ?? string.Empty, out var parsed) ? parsed : false;
                return boolValue ? new XElement("true") : new XElement("false");
            }

            if (kind == PlistValueKind.Integer)
            {
                var textValue = value?.ToString() ?? "0";
                return new XElement("integer", textValue);
            }

            if (kind == PlistValueKind.Real)
            {
                var textValue = value?.ToString() ?? "0";
                return new XElement("real", textValue);
            }

            if (kind == PlistValueKind.Date)
            {
                var textValue = value is DateTime date
                    ? date.ToString("o")
                    : value?.ToString() ?? string.Empty;
                return new XElement("date", textValue);
            }

            if (kind == PlistValueKind.Data)
            {
                return new XElement("data", value?.ToString() ?? string.Empty);
            }

            return new XElement("string", value?.ToString() ?? string.Empty);
        }
        catch (Exception ex)
        {
            error = ex.Message;
            return null;
        }
    }

    private static XElement BuildPlistArrayElement(List<object?>? values)
    {
        var array = new XElement("array");
        if (values == null)
        {
            return array;
        }

        foreach (var value in values)
        {
            array.Add(CreatePlistValueElement(value, PlistValueKind.Unknown, out _));
        }

        return array;
    }

    private static XElement BuildPlistDictElement(Dictionary<string, object?>? values)
    {
        var dict = new XElement("dict");
        if (values == null)
        {
            return dict;
        }

        foreach (var pair in values)
        {
            dict.Add(new XElement("key", pair.Key));
            dict.Add(CreatePlistValueElement(pair.Value, PlistValueKind.Unknown, out _));
        }

        return dict;
    }

    private static bool TryParseJson(string json, out object? value, out string? error)
    {
        value = null;
        error = null;

        try
        {
            using var doc = JsonDocument.Parse(json);
            value = ConvertJsonElement(doc.RootElement);
            return true;
        }
        catch (Exception ex)
        {
            error = ex.Message;
            return false;
        }
    }

    private static object? ConvertJsonElement(JsonElement element)
    {
        switch (element.ValueKind)
        {
            case JsonValueKind.String:
                return element.GetString();
            case JsonValueKind.Number:
                if (element.TryGetInt64(out var intValue))
                {
                    return intValue;
                }

                if (element.TryGetDouble(out var doubleValue))
                {
                    return doubleValue;
                }

                return element.GetRawText();
            case JsonValueKind.True:
                return true;
            case JsonValueKind.False:
                return false;
            case JsonValueKind.Object:
                var dict = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
                foreach (var prop in element.EnumerateObject())
                {
                    dict[prop.Name] = ConvertJsonElement(prop.Value);
                }

                return dict;
            case JsonValueKind.Array:
                var list = new List<object?>();
                foreach (var item in element.EnumerateArray())
                {
                    list.Add(ConvertJsonElement(item));
                }

                return list;
            default:
                return null;
        }
    }

    private static bool RemovePlistKey(string manifestPath, string key, out string? message)
    {
        message = null;

        try
        {
            var doc = XDocument.Load(manifestPath, LoadOptions.PreserveWhitespace);
            var dict = doc.Root?.Elements().FirstOrDefault(el => string.Equals(el.Name.LocalName, "dict", StringComparison.OrdinalIgnoreCase))
                       ?? doc.Descendants().FirstOrDefault(el => string.Equals(el.Name.LocalName, "dict", StringComparison.OrdinalIgnoreCase));
            if (dict == null)
            {
                message = "Info.plist dictionary root not found.";
                return false;
            }

            var elements = dict.Elements().ToList();
            for (var i = 0; i < elements.Count; i++)
            {
                if (!string.Equals(elements[i].Name.LocalName, "key", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (!string.Equals(elements[i].Value, key, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var keyElement = elements[i];
                XElement? valueElement = null;
                if (i + 1 < elements.Count)
                {
                    valueElement = elements[i + 1];
                }

                keyElement.Remove();
                valueElement?.Remove();
                doc.Save(manifestPath);
                return true;
            }

            message = "Key not found.";
            return false;
        }
        catch (Exception ex)
        {
            message = ex.Message;
            return false;
        }
    }

    private static bool ShouldIncludeAllMsBuildProperties()
    {
        var value = Environment.GetEnvironmentVariable("VX_PROPS_ALL");
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        return string.Equals(value, "1", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(value, "true", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(value, "yes", StringComparison.OrdinalIgnoreCase);
    }

    private static HashSet<string>? BuildAllowedMsBuildPropertyNames(string? projectPath, List<RuleDefinition> ruleDefinitions, HashSet<string>? platformPropertyNames)
    {
        var allowed = new HashSet<string>(KnownPropertyNames, StringComparer.OrdinalIgnoreCase);

        foreach (var rule in ruleDefinitions)
        {
            foreach (var property in rule.Properties)
            {
                if (!string.IsNullOrWhiteSpace(property.Name))
                {
                    allowed.Add(property.Name);
                }
            }
        }

        if (platformPropertyNames != null)
        {
            foreach (var name in platformPropertyNames)
            {
                if (!string.IsNullOrWhiteSpace(name))
                {
                    allowed.Add(name);
                }
            }
        }

        if (string.IsNullOrWhiteSpace(projectPath) || !File.Exists(projectPath))
        {
            return allowed;
        }

        foreach (var pair in CollectPropertyValuesFromProjectFiles(projectPath))
        {
            if (!string.IsNullOrWhiteSpace(pair.Key))
            {
                allowed.Add(pair.Key);
            }
        }

        var userFile = projectPath + ".user";
        foreach (var pair in ReadPropertyValuesFromMsBuildFile(userFile))
        {
            if (!string.IsNullOrWhiteSpace(pair.Key))
            {
                allowed.Add(pair.Key);
            }
        }

        return allowed;
    }

    private static List<RuleDefinition> LoadStaticRuleDefinitions(string? projectPath)
    {
        var cacheKey = BuildRuleDefinitionCacheKey(projectPath);
        lock (RuleDefinitionLock)
        {
            if (RuleDefinitionCache.TryGetValue(cacheKey, out var cached))
            {
                return cached;
            }
        }

        var rules = new List<RuleDefinition>();
        var ruleFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var targetFile in GetPropertyPageSchemaTargetFiles(projectPath))
        {
            foreach (var schemaFile in GetPropertyPageSchemaFiles(targetFile))
            {
                if (File.Exists(schemaFile))
                {
                    ruleFiles.Add(schemaFile);
                }
            }
        }

        foreach (var ruleFile in ruleFiles)
        {
            var rule = ParseRuleDefinition(ruleFile);
            if (rule != null)
            {
                rules.Add(rule);
            }
        }

        lock (RuleDefinitionLock)
        {
            RuleDefinitionCache[cacheKey] = rules;
        }

        return rules;
    }

    private static List<RuleDefinition> LoadRuleDefinitions(string? projectPath, Project? msbuildProject, IEnumerable<string>? extraRuleFiles)
    {
        if (msbuildProject == null && (extraRuleFiles == null || !extraRuleFiles.Any()))
        {
            return LoadStaticRuleDefinitions(projectPath);
        }

        var ruleFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var targetFile in GetPropertyPageSchemaTargetFiles(projectPath))
        {
            foreach (var schemaFile in GetPropertyPageSchemaFiles(targetFile))
            {
                if (File.Exists(schemaFile))
                {
                    ruleFiles.Add(schemaFile);
                }
            }
        }

        if (msbuildProject != null)
        {
            foreach (var schemaFile in GetPropertyPageSchemaFiles(msbuildProject))
            {
                if (File.Exists(schemaFile))
                {
                    ruleFiles.Add(schemaFile);
                }
            }
        }

        if (extraRuleFiles != null)
        {
            foreach (var schemaFile in extraRuleFiles)
            {
                if (File.Exists(schemaFile))
                {
                    ruleFiles.Add(schemaFile);
                }
            }
        }

        var rules = new List<RuleDefinition>();
        foreach (var ruleFile in ruleFiles)
        {
            var rule = ParseRuleDefinition(ruleFile);
            if (rule != null)
            {
                rules.Add(rule);
            }
        }

        return rules;
    }

    private static string BuildRuleDefinitionCacheKey(string? projectPath)
    {
        var language = GetProjectLanguage(projectPath);
        var vsRoot = GetVisualStudioRoot() ?? string.Empty;
        var sdkRoot = CachedDotnetSdkRootPath ?? string.Empty;
        return $"{language}|{vsRoot}|{sdkRoot}";
    }

    private static string GetProjectLanguage(string? projectPath)
    {
        if (string.IsNullOrWhiteSpace(projectPath))
        {
            return "Unknown";
        }

        var ext = Path.GetExtension(projectPath);
        if (string.Equals(ext, ".csproj", StringComparison.OrdinalIgnoreCase))
        {
            return "CSharp";
        }

        if (string.Equals(ext, ".vbproj", StringComparison.OrdinalIgnoreCase))
        {
            return "VisualBasic";
        }

        if (string.Equals(ext, ".fsproj", StringComparison.OrdinalIgnoreCase))
        {
            return "FSharp";
        }

        return "Unknown";
    }

    private static IEnumerable<string> GetPropertyPageSchemaTargetFiles(string? projectPath)
    {
        var vsRoot = GetVisualStudioRoot();
        if (!string.IsNullOrWhiteSpace(vsRoot))
        {
            var managedRoot = Path.Combine(vsRoot, "MSBuild", "Microsoft", "VisualStudio", "Managed");
            var managedTargets = Path.Combine(managedRoot, "Microsoft.Managed.DesignTime.targets");
            if (File.Exists(managedTargets))
            {
                yield return managedTargets;
            }

            var language = GetProjectLanguage(projectPath);
            if (string.Equals(language, "CSharp", StringComparison.OrdinalIgnoreCase))
            {
                var csharpTargets = Path.Combine(managedRoot, "Microsoft.CSharp.DesignTime.targets");
                if (File.Exists(csharpTargets))
                {
                    yield return csharpTargets;
                }
            }
            else if (string.Equals(language, "VisualBasic", StringComparison.OrdinalIgnoreCase))
            {
                var vbTargets = Path.Combine(managedRoot, "Microsoft.VisualBasic.DesignTime.targets");
                if (File.Exists(vbTargets))
                {
                    yield return vbTargets;
                }
            }
            else if (string.Equals(language, "FSharp", StringComparison.OrdinalIgnoreCase))
            {
                var fsTargets = Path.Combine(managedRoot, "Microsoft.FSharp.DesignTime.targets");
                if (File.Exists(fsTargets))
                {
                    yield return fsTargets;
                }
            }

            var mauiTargets = Path.Combine(vsRoot, "MSBuild", "Microsoft", "VisualStudio", "Maui", "Maui.DesignTime.targets");
            if (File.Exists(mauiTargets))
            {
                yield return mauiTargets;
            }

            var msbuildCurrentBin = Path.Combine(vsRoot, "MSBuild", "Current", "Bin");
            var commonTargets = Path.Combine(msbuildCurrentBin, "Microsoft.Common.CurrentVersion.targets");
            if (File.Exists(commonTargets))
            {
                yield return commonTargets;
            }

            var csharpCurrentTargets = Path.Combine(msbuildCurrentBin, "Microsoft.CSharp.CurrentVersion.targets");
            if (File.Exists(csharpCurrentTargets))
            {
                yield return csharpCurrentTargets;
            }

            var vbCurrentTargets = Path.Combine(msbuildCurrentBin, "Microsoft.VisualBasic.CurrentVersion.targets");
            if (File.Exists(vbCurrentTargets))
            {
                yield return vbCurrentTargets;
            }

            var fsCurrentTargets = Path.Combine(msbuildCurrentBin, "Microsoft.FSharp.CurrentVersion.targets");
            if (File.Exists(fsCurrentTargets))
            {
                yield return fsCurrentTargets;
            }
        }
    }

    private static IEnumerable<string> GetPropertyPageSchemaFiles(string targetFile)
    {
        ProjectRootElement? root;
        try
        {
            root = ProjectRootElement.Open(targetFile);
        }
        catch
        {
            yield break;
        }

        if (root == null)
        {
            yield break;
        }

        var properties = BuildPropertyPageSchemaProperties(targetFile);
        foreach (var item in root.Items)
        {
            if (!string.Equals(item.ItemType, "PropertyPageSchema", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (!ShouldIncludePropertyPageSchemaItem(item))
            {
                continue;
            }

            foreach (var include in item.Include.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var resolved = ExpandPropertyPageSchemaPath(include.Trim(), targetFile, properties);
                if (!string.IsNullOrWhiteSpace(resolved))
                {
                    yield return resolved;
                }
            }
        }
    }

    private static IEnumerable<string> GetPropertyPageSchemaFiles(Project msbuildProject)
    {
        if (msbuildProject == null)
        {
            yield break;
        }

        foreach (var item in msbuildProject.GetItems("PropertyPageSchema"))
        {
            if (!ShouldIncludePropertyPageSchemaItem(item))
            {
                continue;
            }

            var include = item.EvaluatedInclude;
            if (string.IsNullOrWhiteSpace(include))
            {
                continue;
            }

            foreach (var part in include.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var resolved = ResolvePropertyPageSchemaPath(part, msbuildProject.DirectoryPath);
                if (!string.IsNullOrWhiteSpace(resolved))
                {
                    yield return resolved;
                }
            }
        }
    }

    private static bool ShouldIncludePropertyPageSchemaItem(ProjectItem item)
    {
        var context = item.GetMetadataValue("Context");
        if (string.IsNullOrWhiteSpace(context))
        {
            return false;
        }

        var tokens = context
            .Split(';', StringSplitOptions.RemoveEmptyEntries)
            .Select(token => token.Trim())
            .Where(token => !string.IsNullOrWhiteSpace(token))
            .ToArray();

        if (tokens.Length == 0)
        {
            return false;
        }

        var hasProject = tokens.Any(token =>
            string.Equals(token, "Project", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(token, "ProjectConfiguration", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(token, "Configuration", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(token, "PropertySheet", StringComparison.OrdinalIgnoreCase));

        if (!hasProject)
        {
            return false;
        }

        if (tokens.Any(token =>
                token.Contains("BrowseObject", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(token, "File", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(token, "ProjectSubscriptionService", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(token, "ConfiguredBrowseObject", StringComparison.OrdinalIgnoreCase)))
        {
            return false;
        }

        return true;
    }

    private static bool ShouldIncludePropertyPageSchemaItem(ProjectItemElement item)
    {
        if (item == null)
        {
            return false;
        }

        var context = item.Metadata
            .FirstOrDefault(metadata => string.Equals(metadata.Name, "Context", StringComparison.OrdinalIgnoreCase))
            ?.Value;

        if (string.IsNullOrWhiteSpace(context))
        {
            return false;
        }

        var tokens = context
            .Split(';', StringSplitOptions.RemoveEmptyEntries)
            .Select(token => token.Trim())
            .Where(token => !string.IsNullOrWhiteSpace(token))
            .ToArray();

        if (tokens.Length == 0)
        {
            return false;
        }

        var hasProject = tokens.Any(token =>
            string.Equals(token, "Project", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(token, "ProjectConfiguration", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(token, "Configuration", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(token, "PropertySheet", StringComparison.OrdinalIgnoreCase));

        if (!hasProject)
        {
            return false;
        }

        if (tokens.Any(token =>
                token.Contains("BrowseObject", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(token, "File", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(token, "ProjectSubscriptionService", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(token, "ConfiguredBrowseObject", StringComparison.OrdinalIgnoreCase)))
        {
            return false;
        }

        return true;
    }

    private static string? ResolvePropertyPageSchemaPath(string include, string? baseDirectory)
    {
        if (string.IsNullOrWhiteSpace(include))
        {
            return null;
        }

        var trimmed = include.Trim();
        if (!Path.IsPathRooted(trimmed) && !string.IsNullOrWhiteSpace(baseDirectory))
        {
            trimmed = Path.Combine(baseDirectory, trimmed);
        }

        try
        {
            return Path.GetFullPath(trimmed);
        }
        catch
        {
            return trimmed;
        }
    }

    private static Dictionary<string, string> BuildPropertyPageSchemaProperties(string targetFile)
    {
        var properties = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["MSBuildThisFileDirectory"] = EnsureTrailingSlash(Path.GetDirectoryName(targetFile))
        };

        var directory = Path.GetDirectoryName(targetFile);
        if (!string.IsNullOrWhiteSpace(directory) &&
            directory.EndsWith(Path.Combine("VisualStudio", "Managed"), StringComparison.OrdinalIgnoreCase))
        {
            var neutral = EnsureTrailingSlash(directory);
            properties["ManagedXamlNeutralResourcesDirectory"] = neutral;
            properties["ManagedXamlResourcesDirectory"] = ResolveManagedXamlResourcesDirectory(neutral);
        }

        if (!string.IsNullOrWhiteSpace(CachedDotnetSdksPath))
        {
            properties["MSBuildSDKsPath"] = EnsureTrailingSlash(CachedDotnetSdksPath);
        }

        return properties;
    }

    private static string? ExpandPropertyPageSchemaPath(string include, string targetFile, Dictionary<string, string> properties)
    {
        if (string.IsNullOrWhiteSpace(include))
        {
            return null;
        }

        var expanded = ReplacePropertyTokens(include, properties, targetFile);
        if (string.IsNullOrWhiteSpace(expanded))
        {
            return null;
        }

        var trimmed = expanded.Trim();
        if (!Path.IsPathRooted(trimmed))
        {
            var baseDir = Path.GetDirectoryName(targetFile);
            if (!string.IsNullOrWhiteSpace(baseDir))
            {
                trimmed = Path.Combine(baseDir, trimmed);
            }
        }

        try
        {
            return Path.GetFullPath(trimmed);
        }
        catch
        {
            return trimmed;
        }
    }

    private static string ReplacePropertyTokens(string value, Dictionary<string, string> properties, string targetFile)
    {
        return Regex.Replace(value, "\\$\\(([^)]+)\\)", match =>
        {
            var key = match.Groups[1].Value;
            if (string.Equals(key, "MSBuildThisFileDirectory", StringComparison.OrdinalIgnoreCase))
            {
                return EnsureTrailingSlash(Path.GetDirectoryName(targetFile));
            }

            if (properties.TryGetValue(key, out var replacement))
            {
                return replacement;
            }

            return match.Value;
        }, RegexOptions.IgnoreCase);
    }

    private static string ResolveManagedXamlResourcesDirectory(string neutralDirectory)
    {
        if (string.IsNullOrWhiteSpace(neutralDirectory))
        {
            return string.Empty;
        }

        return EnsureTrailingSlash(neutralDirectory);
    }

    private static string EnsureTrailingSlash(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return string.Empty;
        }

        if (path.EndsWith(Path.DirectorySeparatorChar) || path.EndsWith(Path.AltDirectorySeparatorChar))
        {
            return path;
        }

        return path + Path.DirectorySeparatorChar;
    }

    private static RuleDefinition? ParseRuleDefinition(string ruleFile)
    {
        try
        {
            var doc = XDocument.Load(ruleFile);
            var ruleElement = doc.Root;
            if (ruleElement == null || !string.Equals(ruleElement.Name.LocalName, "Rule", StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            var name = ruleElement.Attribute("Name")?.Value;
            if (string.IsNullOrWhiteSpace(name))
            {
                name = Path.GetFileNameWithoutExtension(ruleFile);
            }

            var displayName = ruleElement.Attribute("DisplayName")?.Value;
            if (string.IsNullOrWhiteSpace(displayName))
            {
                displayName = name;
            }

            var rule = new RuleDefinition(name, displayName)
            {
                Description = ruleElement.Attribute("Description")?.Value,
                AppliesTo = ruleElement.Attribute("AppliesTo")?.Value,
                PageTemplate = ruleElement.Attribute("PageTemplate")?.Value
            };

            var categoriesElement = ruleElement.Elements()
                .FirstOrDefault(element => string.Equals(element.Name.LocalName, "Rule.Categories", StringComparison.OrdinalIgnoreCase));
            if (categoriesElement != null)
            {
                foreach (var categoryElement in categoriesElement.Elements()
                             .Where(element => string.Equals(element.Name.LocalName, "Category", StringComparison.OrdinalIgnoreCase)))
                {
                    var categoryName = categoryElement.Attribute("Name")?.Value ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(categoryName))
                    {
                        continue;
                    }

                    var categoryDisplay = categoryElement.Attribute("DisplayName")?.Value ?? categoryName;
                    var category = new RuleCategoryDefinition(categoryName, categoryDisplay)
                    {
                        Description = categoryElement.Attribute("Description")?.Value
                    };
                    rule.Categories[categoryName] = category;
                }
            }

            var ruleDataSource = ruleElement.Elements()
                .FirstOrDefault(element => string.Equals(element.Name.LocalName, "Rule.DataSource", StringComparison.OrdinalIgnoreCase));
            var dataSourceElement = ruleDataSource?.Elements()
                .FirstOrDefault(element => string.Equals(element.Name.LocalName, "DataSource", StringComparison.OrdinalIgnoreCase));
            if (dataSourceElement != null)
            {
                rule.DataSourcePersistence = dataSourceElement.Attribute("Persistence")?.Value;
                rule.DataSourceItemType = dataSourceElement.Attribute("ItemType")?.Value;
            }

            foreach (var propElement in ruleElement.Elements()
                         .Where(element => element.Name.LocalName.EndsWith("Property", StringComparison.OrdinalIgnoreCase)))
            {
                var propName = propElement.Attribute("Name")?.Value;
                if (string.IsNullOrWhiteSpace(propName))
                {
                    continue;
                }

                var propDisplayName = propElement.Attribute("DisplayName")?.Value;
                if (string.IsNullOrWhiteSpace(propDisplayName))
                {
                    propDisplayName = propName;
                }

                var propDef = new RulePropertyDefinition(propName, propDisplayName)
                {
                    Description = propElement.Attribute("Description")?.Value,
                    Category = propElement.Attribute("Category")?.Value,
                    PropertyKind = propElement.Name.LocalName
                };

                if (!string.IsNullOrWhiteSpace(propDef.Category) &&
                    rule.Categories.TryGetValue(propDef.Category, out var category))
                {
                    propDef.CategoryDisplayName = category.DisplayName;
                }
                else if (!string.IsNullOrWhiteSpace(propDef.Category))
                {
                    propDef.CategoryDisplayName = propDef.Category;
                }

                var visibleValue = propElement.Attribute("Visible")?.Value;
                if (!string.IsNullOrWhiteSpace(visibleValue) &&
                    string.Equals(visibleValue, "False", StringComparison.OrdinalIgnoreCase))
                {
                    propDef.Visible = false;
                }

                var propDataSource = propElement.Elements()
                    .FirstOrDefault(element => element.Name.LocalName.EndsWith(".DataSource", StringComparison.OrdinalIgnoreCase));
                var propDataSourceElement = propDataSource?.Elements()
                    .FirstOrDefault(element => string.Equals(element.Name.LocalName, "DataSource", StringComparison.OrdinalIgnoreCase));
                if (propDataSourceElement != null)
                {
                    propDef.DataSourcePersistence = propDataSourceElement.Attribute("Persistence")?.Value;
                    propDef.DataSourceItemType = propDataSourceElement.Attribute("ItemType")?.Value;
                    propDef.PersistedName = propDataSourceElement.Attribute("PersistedName")?.Value;
                }
                else
                {
                    propDef.DataSourcePersistence ??= rule.DataSourcePersistence;
                    propDef.DataSourceItemType ??= rule.DataSourceItemType;
                }

                foreach (var enumElement in propElement.Descendants()
                             .Where(element => string.Equals(element.Name.LocalName, "EnumValue", StringComparison.OrdinalIgnoreCase)))
                {
                    var enumName = enumElement.Attribute("Name")?.Value;
                    if (string.IsNullOrWhiteSpace(enumName))
                    {
                        continue;
                    }

                    var enumDisplay = enumElement.Attribute("DisplayName")?.Value ?? enumName;
                    propDef.EnumValues.Add(enumDisplay);
                    propDef.EnumValueMap[enumDisplay] = enumName;
                    propDef.EnumValueMap[enumName] = enumName;
                }

                rule.Properties.Add(propDef);
            }

            return rule;
        }
        catch
        {
            return null;
        }
    }

    private static void AddPropertyPagesFromRules(
        List<PropertyPage> pages,
        List<RuleDefinition> ruleDefinitions,
        Dictionary<string, MsBuildPropertyValue>? msbuildMap,
        Dictionary<string, (string Value, string FilePath)> fileValues,
        string? projectPath,
        IDictionary<string, string> globalProperties)
    {
        foreach (var rule in ruleDefinitions)
        {
            foreach (var prop in rule.Properties)
            {
                if (!prop.Visible && !ShouldIncludeAllMsBuildProperties())
                {
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(prop.DataSourceItemType))
                {
                    continue;
                }

                var pageName = GetRulePropertyPageName(rule, prop);
                var persistedName = GetPersistedPropertyName(prop);

                object? Getter()
                {
                    return GetPropertyValueFromMaps(persistedName, msbuildMap, fileValues);
                }

                Action<object?>? setter = BuildRulePropertySetter(prop, projectPath, globalProperties);
                var entry = new PropertyEntry(prop.DisplayName, persistedName, pageName, Getter, setter, InferPropertyTypeFromRule(prop), PropertySource.MsBuild)
                {
                    Description = prop.Description
                };

                if (prop.EnumValues.Count > 0)
                {
                    entry.AllowedValues = prop.EnumValues;
                    entry.AllowedValueMap = prop.EnumValueMap;
                }

                var deleteAction = BuildRulePropertyDeleteAction(prop, projectPath, globalProperties);
                if (deleteAction != null)
                {
                    entry.DeleteAction = deleteAction;
                }

                AddPropertyEntry(pages, pageName, entry);
            }
        }
    }

    private static void AddPropertyPagesFromPlatformDefinitions(
        List<PropertyPage> pages,
        HashSet<string>? platformPropertyNames,
        Dictionary<string, MsBuildPropertyValue>? msbuildMap,
        Dictionary<string, (string Value, string FilePath)> fileValues,
        string? projectPath,
        IDictionary<string, string> globalProperties)
    {
        if (platformPropertyNames == null || platformPropertyNames.Count == 0)
        {
            return;
        }

        foreach (var name in platformPropertyNames.OrderBy(n => n, StringComparer.OrdinalIgnoreCase))
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                continue;
            }

            var pageName = MapPropertyPageName(string.Empty, name);
            var persistedName = name;

            object? Getter()
            {
                return GetPropertyValueFromMaps(persistedName, msbuildMap, fileValues);
            }

            var setter = BuildDefaultPropertySetter(persistedName, projectPath, globalProperties);
            var rawValue = GetPropertyValueFromMaps(persistedName, msbuildMap, fileValues);
            var entry = new PropertyEntry(persistedName, persistedName, pageName, Getter, setter, InferMsBuildPropertyType(rawValue?.ToString()), PropertySource.MsBuild);

            var deleteAction = BuildDefaultPropertyDeleteAction(persistedName, projectPath, globalProperties);
            if (deleteAction != null)
            {
                entry.DeleteAction = deleteAction;
            }

            AddPropertyEntry(pages, pageName, entry);
        }
    }

    private static string GetRulePropertyPageName(RuleDefinition rule, RulePropertyDefinition prop)
    {
        var baseName = string.IsNullOrWhiteSpace(rule.DisplayName) ? rule.Name : rule.DisplayName;
        var category = prop.CategoryDisplayName ?? prop.Category;
        if (string.IsNullOrWhiteSpace(category))
        {
            return NormalizePageName(baseName);
        }

        return NormalizePageName($"{baseName}/{category}");
    }

    private static string GetPersistedPropertyName(RulePropertyDefinition prop)
    {
        if (!string.IsNullOrWhiteSpace(prop.PersistedName))
        {
            return prop.PersistedName;
        }

        return prop.Name;
    }

    private static object? GetPropertyValueFromMaps(
        string propertyName,
        Dictionary<string, MsBuildPropertyValue>? msbuildMap,
        Dictionary<string, (string Value, string FilePath)> fileValues)
    {
        if (string.IsNullOrWhiteSpace(propertyName))
        {
            return null;
        }

        var key = NormalizePropertyKey(propertyName);
        if (msbuildMap != null && msbuildMap.TryGetValue(key, out var value))
        {
            return value.Value;
        }

        if (fileValues.TryGetValue(propertyName, out var fileValue))
        {
            return fileValue.Value;
        }

        return null;
    }

    private static Action<object?>? BuildRulePropertySetter(RulePropertyDefinition prop, string? projectPath, IDictionary<string, string> globalProperties)
    {
        if (string.IsNullOrWhiteSpace(projectPath))
        {
            return null;
        }

        var persistence = prop.DataSourcePersistence ?? string.Empty;
        var propertyName = GetPersistedPropertyName(prop);

        if (string.Equals(persistence, "UserFile", StringComparison.OrdinalIgnoreCase))
        {
            var userFile = projectPath + ".user";
            return value =>
            {
                var text = value?.ToString() ?? string.Empty;
                SetPropertyInMsBuildFile(userFile, propertyName, text);
            };
        }

        return value => SetMsBuildProperty(projectPath, globalProperties, propertyName, value);
    }

    private static Func<(bool Success, string? Message)>? BuildRulePropertyDeleteAction(RulePropertyDefinition prop, string? projectPath, IDictionary<string, string> globalProperties)
    {
        if (string.IsNullOrWhiteSpace(projectPath))
        {
            return null;
        }

        var persistence = prop.DataSourcePersistence ?? string.Empty;
        var propertyName = GetPersistedPropertyName(prop);

        if (string.Equals(persistence, "UserFile", StringComparison.OrdinalIgnoreCase))
        {
            var userFile = projectPath + ".user";
            return () =>
            {
                var success = RemovePropertyFromMsBuildFile(userFile, propertyName, out var message);
                return (success, message);
            };
        }

        return () =>
        {
            var success = RemoveMsBuildProperty(projectPath, globalProperties, propertyName, out var message);
            return (success, message);
        };
    }

    private static Action<object?>? BuildDefaultPropertySetter(string propertyName, string? projectPath, IDictionary<string, string> globalProperties)
    {
        if (string.IsNullOrWhiteSpace(projectPath))
        {
            return null;
        }

        return value => SetMsBuildProperty(projectPath, globalProperties, propertyName, value);
    }

    private static Func<(bool Success, string? Message)>? BuildDefaultPropertyDeleteAction(string propertyName, string? projectPath, IDictionary<string, string> globalProperties)
    {
        if (string.IsNullOrWhiteSpace(projectPath))
        {
            return null;
        }

        return () =>
        {
            var success = RemoveMsBuildProperty(projectPath, globalProperties, propertyName, out var message);
            return (success, message);
        };
    }

    private static Type InferPropertyTypeFromRule(RulePropertyDefinition prop)
    {
        if (prop.PropertyKind == null)
        {
            return typeof(string);
        }

        var kind = prop.PropertyKind;
        if (kind.EndsWith("BoolProperty", StringComparison.OrdinalIgnoreCase))
        {
            return typeof(bool);
        }

        if (kind.EndsWith("IntProperty", StringComparison.OrdinalIgnoreCase))
        {
            return typeof(int);
        }

        return typeof(string);
    }

    private static HashSet<string> ExtractPlatformPropertyNames(Project msbuildProject)
    {
        var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var prop in msbuildProject.AllEvaluatedProperties)
        {
            var name = prop.Name;
            if (string.IsNullOrWhiteSpace(name))
            {
                continue;
            }

            if (prop.IsEnvironmentProperty || prop.IsReservedProperty)
            {
                continue;
            }

            if (!ShouldIncludeMsBuildProperty(name))
            {
                continue;
            }

            if (IsPlatformSpecificPropertyName(name))
            {
                result.Add(name);
            }
        }

        foreach (var hint in IosPropertyNameHints)
        {
            result.Add(hint);
        }

        foreach (var hint in AndroidPropertyNameHints)
        {
            result.Add(hint);
        }

        return result;
    }

    private static void AddPropertyPagesFromObject(List<PropertyPage> pages, object? target, Dictionary<string, MsBuildPropertyValue>? msbuildMap, MsBuildEvaluation? msbuildEvaluation)
    {
        if (target == null)
        {
            return;
        }

        PropertyDescriptorCollection? descriptors;
        try
        {
            descriptors = TypeDescriptor.GetProperties(target);
        }
        catch
        {
            return;
        }

        if (descriptors == null || descriptors.Count == 0)
        {
            return;
        }

        foreach (PropertyDescriptor descriptor in descriptors)
        {
            try
            {
                if (!descriptor.IsBrowsable)
                {
                    continue;
                }

                var descriptorName = string.IsNullOrWhiteSpace(descriptor.Name) ? descriptor.DisplayName : descriptor.Name;
                var displayName = string.IsNullOrWhiteSpace(descriptor.DisplayName) ? descriptorName : descriptor.DisplayName;
                var category = string.IsNullOrWhiteSpace(descriptor.Category) ? string.Empty : descriptor.Category;
                var pageName = MapPropertyPageName(category, displayName);
                var key = NormalizePropertyKey(descriptorName);
                object? fallbackValue = null;
                if (msbuildMap != null && msbuildMap.TryGetValue(key, out var msbuildValue))
                {
                    fallbackValue = msbuildValue.Value;
                }
                else if (msbuildEvaluation != null && !string.IsNullOrWhiteSpace(descriptorName))
                {
                    fallbackValue = msbuildEvaluation.Project.GetPropertyValue(descriptorName);
                }
                var entry = new PropertyEntry(displayName, descriptorName, target, descriptor, category)
                {
                    FallbackValue = fallbackValue
                };
                AddPropertyEntry(pages, pageName, entry);
            }
            catch
            {
                // Skip properties that throw during descriptor access.
            }
        }
    }

    private static void AddPropertyPageFromComProperties(List<PropertyPage> pages, dynamic properties, Dictionary<string, MsBuildPropertyValue>? msbuildMap, MsBuildEvaluation? msbuildEvaluation)
    {
        if (properties == null)
        {
            return;
        }

        foreach (var prop in EnumerateComCollection(properties))
        {
            if (prop == null)
            {
                continue;
            }

            var propName = TryGetValue(() => (string?)prop.Name);
            if (string.IsNullOrWhiteSpace(propName))
            {
                continue;
            }

            var pageName = MapPropertyPageName(string.Empty, propName);
            var key = NormalizePropertyKey(propName);
            object? fallbackValue = null;
            if (msbuildMap != null && msbuildMap.TryGetValue(key, out var msbuildValue))
            {
                fallbackValue = msbuildValue.Value;
            }
            else if (msbuildEvaluation != null && !string.IsNullOrWhiteSpace(propName))
            {
                fallbackValue = msbuildEvaluation.Project.GetPropertyValue(propName);
            }
            var entry = new PropertyEntry(propName, propName, prop, pageName)
            {
                FallbackValue = fallbackValue
            };
            AddPropertyEntry(pages, pageName, entry);
        }
    }

    private static void AddPropertyPagesFromMsBuild(List<PropertyPage> pages, dynamic project, Dictionary<string, MsBuildPropertyValue> msbuildMap)
    {
        var projectPath = GetProjectPath(project);

        if (string.IsNullOrWhiteSpace(projectPath) || !File.Exists(projectPath))
        {
            return;
        }

        var globalProperties = GetMsBuildGlobalProperties(project);

        foreach (var entryPair in msbuildMap.Values)
        {
            var name = entryPair.Name;
            var value = entryPair.Value;
            if (!ShouldIncludeMsBuildProperty(name))
            {
                continue;
            }

            var pageName = MapPropertyPageName(string.Empty, name);
            var category = pageName;
            object? cachedValue = value;
            var inferredType = InferMsBuildPropertyType(value?.ToString());

            Action<object?> setter = newValue =>
            {
                SetMsBuildProperty(projectPath, globalProperties, name, newValue);
                cachedValue = newValue;
            };

            Func<object?> getter = () => cachedValue;
            var entry = new PropertyEntry(name, name, category, getter, setter, inferredType, PropertySource.MsBuild);
            entry.DeleteAction = () =>
            {
                string? deleteMessage;
                var success = RemoveMsBuildProperty(projectPath, globalProperties, name, out deleteMessage);
                if (success)
                {
                    cachedValue = null;
                }

                return (success, deleteMessage);
            };
            AddPropertyEntry(pages, pageName, entry);
        }
    }

    private static void AddPropertyPagesFromUserFile(List<PropertyPage> pages, string? projectPath)
    {
        if (string.IsNullOrWhiteSpace(projectPath))
        {
            return;
        }

        var userFilePath = projectPath + ".user";
        if (!File.Exists(userFilePath))
        {
            return;
        }

        Dictionary<string, string> properties;
        try
        {
            properties = ReadPropertyValuesFromMsBuildFile(userFilePath);
        }
        catch
        {
            return;
        }

        foreach (var entryPair in properties)
        {
            var name = entryPair.Key;
            var value = entryPair.Value;
            if (string.IsNullOrWhiteSpace(name))
            {
                continue;
            }

            var pageName = MapPropertyPageName(string.Empty, name);
            var category = pageName;
            object? cachedValue = value;
            var inferredType = InferMsBuildPropertyType(value);

            Action<object?> setter = newValue =>
            {
                SetPropertyInMsBuildFile(userFilePath, name, newValue);
                cachedValue = newValue;
            };

            Func<object?> getter = () => cachedValue;
            var entry = new PropertyEntry(name, name, category, getter, setter, inferredType, PropertySource.UserFile);
            entry.DeleteAction = () =>
            {
                string? deleteMessage;
                var success = RemovePropertyFromMsBuildFile(userFilePath, name, out deleteMessage);
                if (success)
                {
                    cachedValue = null;
                }

                return (success, deleteMessage);
            };
            AddPropertyEntry(pages, pageName, entry);
        }
    }

    private static void AddPropertyPagesFromProjectFilesFallback(List<PropertyPage> pages, string projectPath)
    {
        Dictionary<string, (string Value, string FilePath)> properties;
        try
        {
            properties = CollectPropertyValuesFromProjectFiles(projectPath);
        }
        catch
        {
            return;
        }

        foreach (var entryPair in properties)
        {
            var name = entryPair.Key;
            var value = entryPair.Value.Value;
            var sourceFile = entryPair.Value.FilePath;
            if (string.IsNullOrWhiteSpace(name))
            {
                continue;
            }

            var pageName = MapPropertyPageName(string.Empty, name);
            var category = pageName;
            object? cachedValue = value;
            var inferredType = InferMsBuildPropertyType(value);

            Action<object?> setter = newValue =>
            {
                SetPropertyInMsBuildFile(sourceFile, name, newValue);
                cachedValue = newValue;
            };

            Func<object?> getter = () => cachedValue;
            var entry = new PropertyEntry(name, name, category, getter, setter, inferredType, PropertySource.MsBuild);
            entry.DeleteAction = () =>
            {
                string? deleteMessage;
                var success = RemovePropertyFromMsBuildFile(sourceFile, name, out deleteMessage);
                if (success)
                {
                    cachedValue = null;
                }

                return (success, deleteMessage);
            };
            AddPropertyEntry(pages, pageName, entry);
        }
    }

    private static Dictionary<string, (string Value, string FilePath)> CollectPropertyValuesFromProjectFiles(string projectPath)
    {
        var map = new Dictionary<string, (string Value, string FilePath)>(StringComparer.OrdinalIgnoreCase);
        var projectDir = Path.GetDirectoryName(projectPath);

        if (!string.IsNullOrWhiteSpace(projectDir))
        {
            foreach (var file in EnumerateDirectoryBuildFiles(projectDir, "Directory.Build.props", topDown: true))
            {
                MergePropertyFile(map, file);
            }

            foreach (var file in EnumerateDirectoryBuildFiles(projectDir, "Directory.Packages.props", topDown: true))
            {
                MergePropertyFile(map, file);
            }
        }

        MergePropertyFile(map, projectPath);

        if (!string.IsNullOrWhiteSpace(projectDir))
        {
            foreach (var file in EnumerateDirectoryBuildFiles(projectDir, "Directory.Build.targets", topDown: false))
            {
                MergePropertyFile(map, file);
            }
        }

        return map;
    }

    private static void MergePropertyFile(Dictionary<string, (string Value, string FilePath)> map, string filePath)
    {
        var properties = ReadPropertyValuesFromMsBuildFile(filePath);
        foreach (var entry in properties)
        {
            map[entry.Key] = (entry.Value, filePath);
        }
    }

    private static IEnumerable<string> EnumerateDirectoryBuildFiles(string startDir, string fileName, bool topDown)
    {
        var directories = new List<string>();
        var current = new DirectoryInfo(startDir);
        while (current != null)
        {
            directories.Add(current.FullName);
            current = current.Parent;
        }

        if (topDown)
        {
            directories.Reverse();
        }

        foreach (var dir in directories)
        {
            var path = Path.Combine(dir, fileName);
            if (File.Exists(path))
            {
                yield return path;
            }
        }
    }

    private static string? GetProjectPath(dynamic project)
    {
        var path = TryGetValue(() => (string?)project?.FullName)
            ?? TryGetValue(() => (string?)project?.FileName);

        if (string.IsNullOrWhiteSpace(path))
        {
            path = TryGetProjectPropertyValue(project, "FullPath")
                ?? TryGetProjectPropertyValue(project, "FullProjectFileName")
                ?? TryGetProjectPropertyValue(project, "ProjectFile");
        }

        if (string.IsNullOrWhiteSpace(path))
        {
            return null;
        }

        if (!Path.IsPathRooted(path))
        {
            var solutionPath = TryGetValue(() => (string?)project?.DTE?.Solution?.FullName);
            var solutionDir = string.IsNullOrWhiteSpace(solutionPath)
                ? null
                : Path.GetDirectoryName(solutionPath);
            if (!string.IsNullOrWhiteSpace(solutionDir))
            {
                path = Path.Combine(solutionDir, path);
            }
        }

        if (Directory.Exists(path))
        {
            var projectName = TryGetValue(() => (string?)project?.Name) ?? string.Empty;
            var candidate = Directory.EnumerateFiles(path, $"{projectName}.*proj", SearchOption.TopDirectoryOnly)
                .FirstOrDefault()
                ?? Directory.EnumerateFiles(path, "*.*proj", SearchOption.TopDirectoryOnly).FirstOrDefault();
            if (!string.IsNullOrWhiteSpace(candidate))
            {
                return candidate;
            }
        }

        return path;
    }

    private static List<string> GetProjectTargetFrameworks(object? project, Dictionary<string, (string Value, string FilePath)> fileValues)
    {
        var result = new List<string>();
        dynamic? dynamicProject = project;

        AddTargetFramework(result, TryGetActiveConfigurationPropertyValue(dynamicProject, "TargetFramework"));
        AddTargetFramework(result, TryGetProjectPropertyValue(dynamicProject, "TargetFramework"));
        AddTargetFramework(result, GetFilePropertyValue(fileValues, "TargetFramework"));

        var targetFrameworks = TryGetActiveConfigurationPropertyValue(dynamicProject, "TargetFrameworks")
            ?? TryGetProjectPropertyValue(dynamicProject, "TargetFrameworks")
            ?? GetFilePropertyValue(fileValues, "TargetFrameworks");

        AddTargetFrameworks(result, targetFrameworks);

        return result
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Select(value => value.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static void AddTargetFramework(List<string> list, string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return;
        }

        list.Add(value.Trim());
    }

    private static void AddTargetFrameworks(List<string> list, string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return;
        }

        foreach (var part in value.Split(';', StringSplitOptions.RemoveEmptyEntries))
        {
            var trimmed = part.Trim();
            if (!string.IsNullOrWhiteSpace(trimmed))
            {
                list.Add(trimmed);
            }
        }
    }

    private static string? GetFilePropertyValue(Dictionary<string, (string Value, string FilePath)> fileValues, string name)
    {
        if (fileValues.Count == 0)
        {
            return null;
        }

        if (fileValues.TryGetValue(name, out var entry))
        {
            return entry.Value;
        }

        return null;
    }

    private static string? TryGetActiveConfigurationPropertyValue(dynamic project, string propertyName)
    {
        var configManager = TryGetValue(() => (dynamic)project?.ConfigurationManager);
        var activeConfig = TryGetValue(() => (dynamic)configManager?.ActiveConfiguration);
        var props = TryGetValue(() => (dynamic)activeConfig?.Properties);
        var prop = TryGetValue(() => (dynamic)props?.Item(propertyName));
        if (prop == null)
        {
            return null;
        }

        return TryGetValue(() => (string?)prop.Value);
    }

    private static string? TryGetProjectPropertyValue(dynamic project, string propertyName)
    {
        var props = TryGetValue(() => (dynamic)project?.Properties);
        if (props == null)
        {
            return null;
        }

        var prop = TryGetValue(() => (dynamic)props.Item(propertyName));
        if (prop == null)
        {
            return null;
        }

        return TryGetValue(() => (string?)prop.Value);
    }

    private static Dictionary<string, string> ReadPropertyValuesFromMsBuildFile(string filePath)
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var root = ProjectRootElement.Open(filePath);
        foreach (var group in root.PropertyGroups)
        {
            foreach (var prop in group.Properties)
            {
                if (string.IsNullOrWhiteSpace(prop.Name))
                {
                    continue;
                }

                map[prop.Name] = prop.Value ?? string.Empty;
            }
        }

        return map;
    }

    private static void SetPropertyInMsBuildFile(string filePath, string propertyName, object? value)
    {
        var textValue = value?.ToString() ?? string.Empty;
        var root = LoadOrCreateMsBuildFile(filePath);
        var group = FindOrCreatePropertyGroup(root);

        var property = group.Properties
            .FirstOrDefault(p => string.Equals(p.Name, propertyName, StringComparison.OrdinalIgnoreCase));
        if (property == null)
        {
            group.AddProperty(propertyName, textValue);
        }
        else
        {
            property.Value = textValue;
        }

        root.Save();
    }

    private static bool RemovePropertyFromMsBuildFile(string filePath, string propertyName, out string? message)
    {
        message = null;
        if (!File.Exists(filePath))
        {
            message = "Property file not found.";
            return false;
        }

        var root = ProjectRootElement.Open(filePath);
        foreach (var group in root.PropertyGroups)
        {
            var match = group.Properties
                .FirstOrDefault(p => string.Equals(p.Name, propertyName, StringComparison.OrdinalIgnoreCase));
            if (match != null)
            {
                group.RemoveChild(match);
                root.Save();
                return true;
            }
        }

        message = "Property not found in file.";
        return false;
    }

    private static ProjectRootElement LoadOrCreateMsBuildFile(string filePath)
    {
        if (File.Exists(filePath))
        {
            return ProjectRootElement.Open(filePath);
        }

        var directory = Path.GetDirectoryName(filePath);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        var root = ProjectRootElement.Create(filePath);
        root.Save();
        return root;
    }

    private static ProjectPropertyGroupElement FindOrCreatePropertyGroup(ProjectRootElement root)
    {
        var group = root.PropertyGroups.FirstOrDefault(pg => string.IsNullOrWhiteSpace(pg.Condition));
        return group ?? root.AddPropertyGroup();
    }

    private static MsBuildEvaluation? TryCreateMsBuildEvaluation(dynamic project, IDictionary<string, string>? globalPropertiesOverride)
    {
        var projectPath = TryGetValue(() => (string?)project.FullName)
            ?? TryGetValue(() => (string?)project.FileName);

        if (string.IsNullOrWhiteSpace(projectPath) || !File.Exists(projectPath))
        {
            return null;
        }

        var globalProperties = globalPropertiesOverride != null
            ? new Dictionary<string, string>(globalPropertiesOverride, StringComparer.OrdinalIgnoreCase)
            : GetMsBuildGlobalProperties(project);

        try
        {
            var collection = new ProjectCollection(globalProperties);
            var loadSettings = ProjectLoadSettings.IgnoreMissingImports
                | ProjectLoadSettings.IgnoreInvalidImports
                | ProjectLoadSettings.IgnoreEmptyImports;
            var msbuildProject = new Project(projectPath, globalProperties, null, collection, loadSettings);
            return new MsBuildEvaluation(collection, msbuildProject);
        }
        catch (Exception ex)
        {
            LogMsBuildEvaluationFailure(ex.Message);
            return null;
        }
    }

    private static Dictionary<string, MsBuildPropertyValue> BuildMsBuildPropertyMap(Project msbuildProject, HashSet<string>? allowedNames)
    {
        var map = new Dictionary<string, MsBuildPropertyValue>(StringComparer.OrdinalIgnoreCase);

        foreach (var prop in msbuildProject.AllEvaluatedProperties)
        {
            var name = prop.Name;
            if (string.IsNullOrWhiteSpace(name))
            {
                continue;
            }

            if (prop.IsEnvironmentProperty || prop.IsReservedProperty)
            {
                continue;
            }

            if (!ShouldIncludeMsBuildPropertyName(name, allowedNames))
            {
                continue;
            }

            var key = NormalizePropertyKey(name);
            if (!map.ContainsKey(key))
            {
                map[key] = new MsBuildPropertyValue(prop.Name, prop.EvaluatedValue);
            }
        }

        return map;
    }

    private static bool EnsureMsBuildRegistered(out string? error)
    {
        error = null;

        if (MsBuildRegistrationAttempted)
        {
            return MsBuildRegistrationSucceeded;
        }

        MsBuildRegistrationAttempted = true;
        try
        {
            if (IsPackagedMsBuildAssembly())
            {
                MsBuildRegistrationSucceeded = true;
                return true;
            }

            if (!MSBuildLocator.IsRegistered)
            {
                if (!TryRegisterMsBuildViaDotnetSdk(out var dotnetError))
                {
                    var allowDotNetInstances = string.IsNullOrWhiteSpace(dotnetError)
                        || !dotnetError.Contains("No compatible dotnet SDK", StringComparison.OrdinalIgnoreCase);

                    if (!TryRegisterMsBuildViaVisualStudioInstances(allowDotNetInstances, out var instanceError))
                    {
                        if (!string.IsNullOrWhiteSpace(dotnetError))
                        {
                            throw new InvalidOperationException(dotnetError);
                        }

                        if (!string.IsNullOrWhiteSpace(instanceError))
                        {
                            throw new InvalidOperationException(instanceError);
                        }

                        throw new InvalidOperationException("No MSBuild instances found.");
                    }
                }
            }

            MsBuildRegistrationSucceeded = true;
            return true;
        }
        catch (Exception ex)
        {
            var fallbackError = ex.Message;
            if (TryRegisterMsBuildViaVswhere(out var vswhereError))
            {
                MsBuildRegistrationSucceeded = true;
                return true;
            }

            MsBuildRegistrationSucceeded = false;
            error = string.IsNullOrWhiteSpace(vswhereError)
                ? fallbackError
                : $"{fallbackError} {vswhereError}";
            return false;
        }
    }

    private static bool IsPackagedMsBuildAssembly()
    {
        try
        {
            var location = typeof(ProjectCollection).Assembly.Location;
            if (string.IsNullOrWhiteSpace(location))
            {
                return false;
            }

            return location.StartsWith(AppContext.BaseDirectory, StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    private static void RegisterAssemblyResolvers()
    {
        if (AssemblyResolversRegistered)
        {
            return;
        }

        AssemblyResolversRegistered = true;
        AssemblyLoadContext.Default.Resolving += (_, name) =>
        {
            if (!string.Equals(name.Name, "System.Collections.Immutable", StringComparison.OrdinalIgnoreCase))
            {
                if (!string.IsNullOrWhiteSpace(name.Name) &&
                    name.Name.StartsWith("NuGet.", StringComparison.OrdinalIgnoreCase))
                {
                    var sdkRoot = CachedDotnetSdkRootPath;
                    if (string.IsNullOrWhiteSpace(sdkRoot))
                    {
                        if (TryResolveDotnetSdkSdksPath(null, out var sdksPath, out var sdkVersion))
                        {
                            CachedDotnetSdksPath = sdksPath;
                            CachedDotnetSdkVersion = sdkVersion;
                            CachedDotnetSdkRootPath = Path.GetDirectoryName(sdksPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
                        }

                        sdkRoot = CachedDotnetSdkRootPath;
                    }

                    if (!string.IsNullOrWhiteSpace(sdkRoot))
                    {
                        var nugetCandidate = Path.Combine(sdkRoot, $"{name.Name}.dll");
                        if (File.Exists(nugetCandidate))
                        {
                            try
                            {
                                return AssemblyLoadContext.Default.LoadFromAssemblyPath(nugetCandidate);
                            }
                            catch
                            {
                                return null;
                            }
                        }
                    }
                }

                return null;
            }

            var candidate = Path.Combine(AppContext.BaseDirectory, "System.Collections.Immutable.dll");
            if (!File.Exists(candidate))
            {
                return null;
            }

            try
            {
                return AssemblyLoadContext.Default.LoadFromAssemblyPath(candidate);
            }
            catch
            {
                return null;
            }
        };
    }

    private static bool TryRegisterMsBuildViaDotnetSdk(out string? error)
    {
        error = null;
        if (MSBuildLocator.IsRegistered)
        {
            return true;
        }

        if (!TryFindDotnetSdkPath(out var sdkPath, out error))
        {
            return false;
        }

        try
        {
            MSBuildLocator.RegisterMSBuildPath(sdkPath);
            return true;
        }
        catch (Exception ex)
        {
            error = ex.Message;
            return false;
        }
    }

    private static bool TryRegisterMsBuildViaVisualStudioInstances(bool allowDotNetSdkInstances, out string? error)
    {
        error = null;
        if (MSBuildLocator.IsRegistered)
        {
            return true;
        }

        var instances = MSBuildLocator.QueryVisualStudioInstances().ToList();
        if (!allowDotNetSdkInstances)
        {
            instances = instances
                .Where(item => item.DiscoveryType != DiscoveryType.DotNetSdk)
                .ToList();
        }

        if (instances.Count == 0)
        {
            error = "No Visual Studio MSBuild instances found.";
            return false;
        }

        var ordered = instances
            .OrderBy(item => IsPrereleaseInstance(item))
            .ThenByDescending(item => item.DiscoveryType == DiscoveryType.VisualStudioSetup)
            .ThenByDescending(item => item.Version)
            .ToList();

        try
        {
            MSBuildLocator.RegisterInstance(ordered[0]);
            return true;
        }
        catch (Exception ex)
        {
            error = ex.Message;
            return false;
        }
    }

    private static bool IsPrereleaseInstance(VisualStudioInstance instance)
    {
        return IsPrereleasePath(instance.MSBuildPath) || IsPrereleasePath(instance.Name);
    }

    private static bool IsPrereleasePath(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        return value.Contains("Insiders", StringComparison.OrdinalIgnoreCase) ||
               value.Contains("Preview", StringComparison.OrdinalIgnoreCase);
    }

    private static bool TryFindDotnetSdkPath(out string? sdkPath, out string? error)
    {
        sdkPath = null;
        error = null;

        if (!TryGetDotnetSdks(out var sdks, out error))
        {
            return false;
        }

        var runtimeMajor = Environment.Version.Major;
        var candidate = sdks
            .Where(sdk => sdk.Version.Major <= runtimeMajor && File.Exists(Path.Combine(sdk.Path, "MSBuild.dll")))
            .OrderByDescending(sdk => sdk.Version)
            .FirstOrDefault();

        if (candidate == null)
        {
            error = $"No compatible dotnet SDK with MSBuild.dll found (<= {runtimeMajor}).";
            return false;
        }

        sdkPath = candidate.Path;
        return true;
    }

    private static bool TryGetDotnetSdks(out List<DotnetSdkInfo> sdks, out string? error)
    {
        sdks = new List<DotnetSdkInfo>();
        error = null;

        if (!TryRunProcess("dotnet", "--list-sdks", out var stdout, out var stderr))
        {
            error = string.IsNullOrWhiteSpace(stderr) ? "Failed to query dotnet SDKs." : stderr;
            return false;
        }

        foreach (var line in stdout.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
        {
            var trimmed = line.Trim();
            var bracketIndex = trimmed.IndexOf('[');
            if (bracketIndex <= 0)
            {
                continue;
            }

            var versionText = trimmed.Substring(0, bracketIndex).Trim();
            var endBracket = trimmed.IndexOf(']', bracketIndex + 1);
            if (endBracket <= bracketIndex)
            {
                continue;
            }

            var basePath = trimmed.Substring(bracketIndex + 1, endBracket - bracketIndex - 1).Trim();
            if (!Version.TryParse(versionText, out var version))
            {
                continue;
            }

            var candidate = Path.Combine(basePath, versionText);
            if (!Directory.Exists(candidate))
            {
                continue;
            }

            sdks.Add(new DotnetSdkInfo(version, candidate));
        }

        if (sdks.Count == 0)
        {
            error = "No dotnet SDKs found.";
            return false;
        }

        return true;
    }

    private static void EnsureDotnetSdkEnvironment(string? projectPath, string? solutionDir)
    {
        if (DotnetSdkEnvironmentResolved)
        {
            return;
        }

        DotnetSdkEnvironmentResolved = true;

        var globalJsonPath = FindGlobalJsonPath(projectPath, solutionDir);
        if (!TryResolveDotnetSdkSdksPath(globalJsonPath, out var sdksPath, out var sdkVersion))
        {
            return;
        }

        CachedDotnetSdksPath = sdksPath;
        CachedDotnetSdkVersion = sdkVersion;
        CachedDotnetSdkRootPath = Path.GetDirectoryName(sdksPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));

        Environment.SetEnvironmentVariable("MSBuildSDKsPath", sdksPath);
        Environment.SetEnvironmentVariable("DOTNET_MSBUILD_SDK_RESOLVER_SDKS_DIR", sdksPath);
        if (!string.IsNullOrWhiteSpace(sdkVersion))
        {
            Environment.SetEnvironmentVariable("DOTNET_MSBUILD_SDK_RESOLVER_SDKS_VER", sdkVersion);
        }
    }

    private static void EnsureNuGetAssembliesPresent()
    {
        if (NuGetAssembliesEnsured)
        {
            return;
        }

        NuGetAssembliesEnsured = true;
        var sdkRoot = CachedDotnetSdkRootPath;
        if (string.IsNullOrWhiteSpace(sdkRoot))
        {
            if (TryResolveDotnetSdkSdksPath(null, out var sdksPath, out var sdkVersion))
            {
                CachedDotnetSdksPath = sdksPath;
                CachedDotnetSdkVersion = sdkVersion;
                CachedDotnetSdkRootPath = Path.GetDirectoryName(sdksPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
            }

            sdkRoot = CachedDotnetSdkRootPath;
        }

        if (string.IsNullOrWhiteSpace(sdkRoot) || !Directory.Exists(sdkRoot))
        {
            return;
        }

        foreach (var dll in Directory.EnumerateFiles(sdkRoot, "NuGet*.dll"))
        {
            var destination = Path.Combine(AppContext.BaseDirectory, Path.GetFileName(dll));
            if (File.Exists(destination))
            {
                continue;
            }

            try
            {
                File.Copy(dll, destination, overwrite: false);
            }
            catch
            {
            }
        }
    }

    private static bool TryResolveDotnetSdkSdksPath(string? globalJsonPath, out string? sdksPath, out string? sdkVersion)
    {
        sdksPath = null;
        sdkVersion = null;

        if (!TryGetDotnetSdks(out var sdks, out _))
        {
            return false;
        }

        DotnetSdkInfo? chosen = null;
        var requestedVersion = TryReadGlobalJsonSdkVersion(globalJsonPath);
        if (!string.IsNullOrWhiteSpace(requestedVersion) && Version.TryParse(requestedVersion, out var requested))
        {
            chosen = sdks.FirstOrDefault(sdk => sdk.Version == requested);
        }

        chosen ??= sdks
            .OrderByDescending(sdk => sdk.Version)
            .FirstOrDefault();

        if (chosen == null)
        {
            return false;
        }

        var candidate = Path.Combine(chosen.Path, "Sdks");
        if (!Directory.Exists(candidate))
        {
            return false;
        }

        sdksPath = candidate;
        sdkVersion = chosen.Version.ToString();
        return true;
    }

    private static string? FindGlobalJsonPath(string? projectPath, string? solutionDir)
    {
        var startDir = !string.IsNullOrWhiteSpace(solutionDir)
            ? solutionDir
            : Path.GetDirectoryName(projectPath);

        if (string.IsNullOrWhiteSpace(startDir))
        {
            return null;
        }

        var current = startDir;
        while (!string.IsNullOrWhiteSpace(current))
        {
            var candidate = Path.Combine(current, "global.json");
            if (File.Exists(candidate))
            {
                return candidate;
            }

            var parent = Directory.GetParent(current);
            current = parent?.FullName;
        }

        return null;
    }

    private static string? TryReadGlobalJsonSdkVersion(string? globalJsonPath)
    {
        if (string.IsNullOrWhiteSpace(globalJsonPath) || !File.Exists(globalJsonPath))
        {
            return null;
        }

        try
        {
            using var stream = File.OpenRead(globalJsonPath);
            using var doc = JsonDocument.Parse(stream);
            if (doc.RootElement.TryGetProperty("sdk", out var sdkElement)
                && sdkElement.TryGetProperty("version", out var versionElement)
                && versionElement.ValueKind == JsonValueKind.String)
            {
                return versionElement.GetString();
            }
        }
        catch
        {
            return null;
        }

        return null;
    }

    private static bool PrepareMsBuild()
    {
        if (MsBuildRegistrationAttempted)
        {
            return MsBuildRegistrationSucceeded;
        }

        var success = EnsureMsBuildRegistered(out var error);
        if (!success && !MsBuildRegistrationLogged && !string.IsNullOrWhiteSpace(error))
        {
            Console.Error.WriteLine($"MSBuild evaluation unavailable: {error}");
            MsBuildRegistrationLogged = true;
        }

        return success;
    }

    private static void LogMsBuildEvaluationFailure(string message)
    {
        if (MsBuildEvaluationLogged)
        {
            return;
        }

        Console.Error.WriteLine($"MSBuild evaluation failed: {message}");
        MsBuildEvaluationLogged = true;
    }

    private static bool IsUiaDisabled()
    {
        var value = Environment.GetEnvironmentVariable("VX_DISABLE_UIA");
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        return string.Equals(value, "1", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(value, "true", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(value, "yes", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsUiaEnabledForList()
    {
        if (IsUiaDisabled())
        {
            return false;
        }

        var value = Environment.GetEnvironmentVariable("VX_UIA_LIST");
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        return string.Equals(value, "1", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(value, "true", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(value, "yes", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsUiaDebugEnabled()
    {
        var value = Environment.GetEnvironmentVariable("VX_UIA_DEBUG");
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        return string.Equals(value, "1", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(value, "true", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(value, "yes", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsDebugStackEnabled()
    {
        var value = Environment.GetEnvironmentVariable("VX_DEBUG_STACK");
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        return string.Equals(value, "1", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(value, "true", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(value, "yes", StringComparison.OrdinalIgnoreCase);
    }

    private static string FormatUnexpectedError(Exception ex)
    {
        return IsDebugStackEnabled() ? ex.ToString() : ex.Message;
    }

    private static int GetUiaTimeoutMs(string envName, int defaultValue)
    {
        var value = Environment.GetEnvironmentVariable(envName);
        if (int.TryParse(value, out var parsed) && parsed > 0)
        {
            return parsed;
        }

        return defaultValue;
    }

    private static void UiaDebug(string message)
    {
        if (!IsUiaDebugEnabled())
        {
            return;
        }

        Console.Error.WriteLine($"[uia] {message}");
    }

    private static void LogUiaUnavailable(string message)
    {
        if (UiaUnavailableLogged)
        {
            UiaDebug($"UI Automation unavailable: {message}");
            return;
        }

        Console.Error.WriteLine($"UI Automation unavailable: {message}");
        UiaUnavailableLogged = true;
    }

    private static bool TryRegisterMsBuildViaVswhere(out string? error)
    {
        error = null;
        if (MSBuildLocator.IsRegistered)
        {
            return true;
        }

        if (!TryFindMsBuildPath(out var msbuildPath, out error))
        {
            return false;
        }

        try
        {
            MSBuildLocator.RegisterMSBuildPath(msbuildPath);
            return true;
        }
        catch (Exception ex)
        {
            error = ex.Message;
            return false;
        }
    }

    private static bool TryFindMsBuildPath(out string? msbuildPath, out string? error)
    {
        msbuildPath = null;
        error = null;

        var vswherePath = GetVswherePath();
        if (!string.IsNullOrWhiteSpace(vswherePath))
        {
            var arguments = "-latest -products * -requires Microsoft.Component.MSBuild -find \"MSBuild\\**\\Bin\\MSBuild.exe\"";
            if (TryRunProcess(vswherePath, arguments, out var stdout, out var stderr))
            {
                var first = stdout
                    .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(line => line.Trim())
                    .FirstOrDefault(line => !string.IsNullOrWhiteSpace(line));

                if (!string.IsNullOrWhiteSpace(first) && File.Exists(first))
                {
                    msbuildPath = Path.GetDirectoryName(first);
                    if (!string.IsNullOrWhiteSpace(msbuildPath))
                    {
                        return true;
                    }
                }
            }

            if (!string.IsNullOrWhiteSpace(stderr))
            {
                error = stderr.Trim();
            }

            var prereleaseArguments = $"{arguments} -prerelease";
            if (TryRunProcess(vswherePath, prereleaseArguments, out var prereleaseOut, out var prereleaseErr))
            {
                var first = prereleaseOut
                    .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(line => line.Trim())
                    .FirstOrDefault(line => !string.IsNullOrWhiteSpace(line));

                if (!string.IsNullOrWhiteSpace(first) && File.Exists(first))
                {
                    msbuildPath = Path.GetDirectoryName(first);
                    if (!string.IsNullOrWhiteSpace(msbuildPath))
                    {
                        return true;
                    }
                }
            }

            if (!string.IsNullOrWhiteSpace(prereleaseErr))
            {
                error = string.IsNullOrWhiteSpace(error) ? prereleaseErr.Trim() : error;
            }
        }

        foreach (var root in GetCandidateVisualStudioRoots()
            .OrderBy(path => IsPrereleasePath(path)))
        {
            var candidate = Path.Combine(root, "MSBuild", "Current", "Bin");
            if (File.Exists(Path.Combine(candidate, "MSBuild.exe")))
            {
                msbuildPath = candidate;
                return true;
            }
        }

        error = error ?? "No MSBuild install found.";
        return false;
    }

    private static string? GetVswherePath()
    {
        var programFilesX86 = Environment.GetEnvironmentVariable("ProgramFiles(x86)")
            ?? Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);

        if (string.IsNullOrWhiteSpace(programFilesX86))
        {
            return null;
        }

        var candidate = Path.Combine(programFilesX86, "Microsoft Visual Studio", "Installer", "vswhere.exe");
        return File.Exists(candidate) ? candidate : null;
    }

    private static IEnumerable<string> GetCandidateVisualStudioRoots()
    {
        var roots = new List<string>();
        var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
        var programFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);

        if (!string.IsNullOrWhiteSpace(programFiles))
        {
            roots.Add(programFiles);
        }

        if (!string.IsNullOrWhiteSpace(programFilesX86) && !string.Equals(programFiles, programFilesX86, StringComparison.OrdinalIgnoreCase))
        {
            roots.Add(programFilesX86);
        }

        foreach (var root in roots.Distinct(StringComparer.OrdinalIgnoreCase))
        {
            var vsRoot = Path.Combine(root, "Microsoft Visual Studio", "2022");
            if (!Directory.Exists(vsRoot))
            {
                continue;
            }

            foreach (var edition in Directory.EnumerateDirectories(vsRoot))
            {
                yield return edition;
            }
        }
    }

    private static string? GetVisualStudioRoot()
    {
        return GetCandidateVisualStudioRoots()
            .OrderBy(path => IsPrereleasePath(path))
            .FirstOrDefault();
    }

    private static bool TryRunProcess(string fileName, string arguments, out string stdout, out string stderr)
    {
        stdout = string.Empty;
        stderr = string.Empty;

        try
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = fileName,
                Arguments = arguments,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            using var process = Process.Start(startInfo);
            if (process == null)
            {
                return false;
            }

            stdout = process.StandardOutput.ReadToEnd();
            stderr = process.StandardError.ReadToEnd();
            process.WaitForExit(3000);
            return process.ExitCode == 0;
        }
        catch
        {
            return false;
        }
    }

    private static bool IsVisualStudioProcessRunning()
    {
        try
        {
            return Process.GetProcessesByName("devenv").Length > 0;
        }
        catch
        {
            return false;
        }
    }

    private static bool EnsureVsInstanceAvailable(List<VsInstance> instances)
    {
        if (instances.Count > 0)
        {
            return true;
        }

        if (IsVisualStudioProcessRunning())
        {
            Console.Error.WriteLine("Visual Studio process detected but DTE is not accessible. Run vx as Administrator or restart Visual Studio without elevation.");
        }
        else
        {
            Console.Error.WriteLine("No running Visual Studio 2022 instance found.");
        }

        return false;
    }

    private static Dictionary<string, string> GetMsBuildGlobalProperties(dynamic project)
    {
        var globals = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["DesignTimeBuild"] = "true",
            ["BuildingInsideVisualStudio"] = "true"
        };

        var (configuration, platform) = GetProjectConfigurationAndPlatform((object?)project);
        if (!string.IsNullOrWhiteSpace(configuration))
        {
            globals["Configuration"] = configuration;
        }

        if (!string.IsNullOrWhiteSpace(platform))
        {
            globals["Platform"] = NormalizePlatformName(platform);
        }

        var projectPath = TryGetValue(() => (string?)project?.FullName)
            ?? TryGetValue(() => (string?)project?.FileName);

        var solutionPath = TryGetValue(() => (string?)project?.DTE?.Solution?.FullName);
        string? solutionDir = null;
        if (!string.IsNullOrWhiteSpace(solutionPath))
        {
            globals["SolutionPath"] = solutionPath;
            globals["SolutionName"] = Path.GetFileNameWithoutExtension(solutionPath);
            globals["SolutionFileName"] = Path.GetFileName(solutionPath);

            solutionDir = Path.GetDirectoryName(solutionPath);
            if (!string.IsNullOrWhiteSpace(solutionDir))
            {
                if (!solutionDir.EndsWith(Path.DirectorySeparatorChar))
                {
                    solutionDir += Path.DirectorySeparatorChar;
                }

                globals["SolutionDir"] = solutionDir;
            }
        }

        EnsureDotnetSdkEnvironment(projectPath, solutionDir);
        EnsureNuGetAssembliesPresent();
        if (!string.IsNullOrWhiteSpace(CachedDotnetSdksPath))
        {
            globals["MSBuildSDKsPath"] = CachedDotnetSdksPath;
        }

        if (!globals.ContainsKey("VisualStudioVersion"))
        {
            var vsVersion = TryGetValue(() => (string?)project?.DTE?.Version);
            var normalized = NormalizeVisualStudioVersion(vsVersion);
            globals["VisualStudioVersion"] = string.IsNullOrWhiteSpace(normalized) ? "17.0" : normalized;
        }

        return globals;
    }

    private static (string? Configuration, string? Platform) GetProjectConfigurationAndPlatform(object? project)
    {
        dynamic? dynamicProject = project;
        var configManager = TryGetValue(() => (dynamic)dynamicProject?.ConfigurationManager);
        var activeConfig = TryGetValue(() => (dynamic)configManager?.ActiveConfiguration);
        var configurationName = TryGetValue(() => (string?)activeConfig?.ConfigurationName);
        var platformName = TryGetValue(() => (string?)activeConfig?.PlatformName);

        if (string.IsNullOrWhiteSpace(configurationName))
        {
            var composite = TryGetValue(() => (string?)activeConfig?.Name);
            if (!string.IsNullOrWhiteSpace(composite))
            {
                var parts = composite.Split('|');
                if (parts.Length > 0)
                {
                    configurationName = parts[0].Trim();
                }

                if (parts.Length > 1 && string.IsNullOrWhiteSpace(platformName))
                {
                    platformName = parts[1].Trim();
                }
            }
        }

        return (configurationName, platformName);
    }

    private static string NormalizePlatformName(string platform)
    {
        if (string.Equals(platform, "Any CPU", StringComparison.OrdinalIgnoreCase))
        {
            return "AnyCPU";
        }

        return platform;
    }

    private static string? NormalizeVisualStudioVersion(string? version)
    {
        if (string.IsNullOrWhiteSpace(version))
        {
            return null;
        }

        var parts = version.Split('.', StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length >= 2)
        {
            return $"{parts[0]}.{parts[1]}";
        }

        if (parts.Length == 1)
        {
            return $"{parts[0]}.0";
        }

        return null;
    }

    private static bool ShouldIncludeMsBuildProperty(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            return false;
        }

        if (name.StartsWith("_", StringComparison.Ordinal))
        {
            return false;
        }

        if (name.StartsWith("MSBuild", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return true;
    }

    private static bool ShouldIncludeMsBuildPropertyName(string name, HashSet<string>? allowedNames)
    {
        if (!ShouldIncludeMsBuildProperty(name))
        {
            return false;
        }

        if (ShouldIncludeAllMsBuildProperties())
        {
            return true;
        }

        if (allowedNames != null && allowedNames.Contains(name))
        {
            return true;
        }

        return IsPlatformSpecificPropertyName(name);
    }

    private static Type InferMsBuildPropertyType(string? value)
    {
        if (!string.IsNullOrWhiteSpace(value) && bool.TryParse(value, out _))
        {
            return typeof(bool);
        }

        if (!string.IsNullOrWhiteSpace(value) && int.TryParse(value, out _))
        {
            return typeof(int);
        }

        return typeof(string);
    }

    private static void SetMsBuildProperty(string projectPath, IDictionary<string, string> globalProperties, string propertyName, object? value)
    {
        var textValue = value?.ToString() ?? string.Empty;

        using var collection = new ProjectCollection(globalProperties);
        var project = collection.LoadProject(projectPath);

        var property = project.GetProperty(propertyName);
        if (property != null && !property.IsImported && property.Xml != null)
        {
            property.UnevaluatedValue = textValue;
        }
        else
        {
            var condition = BuildConfigurationCondition(globalProperties);
            var group = FindOrCreatePropertyGroup(project, condition);
            group.SetProperty(propertyName, textValue);
        }

        project.Save();
        collection.UnloadAllProjects();
    }

    private static bool RemoveMsBuildProperty(string projectPath, IDictionary<string, string> globalProperties, string propertyName, out string? message)
    {
        message = null;

        using var collection = new ProjectCollection(globalProperties);
        var project = collection.LoadProject(projectPath);

        var property = project.GetProperty(propertyName);
        if (property == null)
        {
            message = "Property not found in project file.";
            return false;
        }

        if (property.IsImported)
        {
            message = "Property is imported and cannot be deleted. Consider overriding it instead.";
            return false;
        }

        if (property.Xml?.Parent == null)
        {
            message = "Property location could not be determined.";
            return false;
        }

        property.Xml.Parent.RemoveChild(property.Xml);
        project.Save();
        collection.UnloadAllProjects();
        return true;
    }

    private static string? BuildConfigurationCondition(IDictionary<string, string> globalProperties)
    {
        globalProperties.TryGetValue("Configuration", out var configuration);
        globalProperties.TryGetValue("Platform", out var platform);
        globalProperties.TryGetValue("TargetFramework", out var targetFramework);

        if (!string.IsNullOrWhiteSpace(configuration) &&
            !string.IsNullOrWhiteSpace(platform) &&
            !string.IsNullOrWhiteSpace(targetFramework))
        {
            return $"'$(Configuration)|$(Platform)|$(TargetFramework)'=='{configuration}|{platform}|{targetFramework}'";
        }

        if (!string.IsNullOrWhiteSpace(configuration) && !string.IsNullOrWhiteSpace(platform))
        {
            return $"'$(Configuration)|$(Platform)'=='{configuration}|{platform}'";
        }

        if (!string.IsNullOrWhiteSpace(configuration) && !string.IsNullOrWhiteSpace(targetFramework))
        {
            return $"'$(Configuration)|$(TargetFramework)'=='{configuration}|{targetFramework}'";
        }

        if (!string.IsNullOrWhiteSpace(platform) && !string.IsNullOrWhiteSpace(targetFramework))
        {
            return $"'$(Platform)|$(TargetFramework)'=='{platform}|{targetFramework}'";
        }

        if (!string.IsNullOrWhiteSpace(configuration))
        {
            return $"'$(Configuration)'=='{configuration}'";
        }

        if (!string.IsNullOrWhiteSpace(platform))
        {
            return $"'$(Platform)'=='{platform}'";
        }

        if (!string.IsNullOrWhiteSpace(targetFramework))
        {
            return $"'$(TargetFramework)'=='{targetFramework}'";
        }

        return null;
    }

    private static ProjectPropertyGroupElement FindOrCreatePropertyGroup(Project project, string? condition)
    {
        if (!string.IsNullOrWhiteSpace(condition))
        {
            var conditionedGroup = project.Xml.PropertyGroups
                .FirstOrDefault(group => string.Equals(group.Condition, condition, StringComparison.OrdinalIgnoreCase));

            if (conditionedGroup != null)
            {
                return conditionedGroup;
            }

            var created = project.Xml.AddPropertyGroup();
            created.Condition = condition;
            return created;
        }

        var existing = project.Xml.PropertyGroups.FirstOrDefault(group => string.IsNullOrWhiteSpace(group.Condition));
        return existing ?? project.Xml.AddPropertyGroup();
    }

    private static void AddPropertyEntry(List<PropertyPage> pages, string pageName, PropertyEntry entry)
    {
        var normalized = NormalizePageName(pageName);
        var page = pages.FirstOrDefault(p => string.Equals(p.Name, normalized, StringComparison.OrdinalIgnoreCase));
        if (page == null)
        {
            page = new PropertyPage(normalized, new List<PropertyEntry>());
            pages.Add(page);
        }

        var entryKey = NormalizePropertyKey(entry.Key);
        var existingIndex = page.Items.FindIndex(item => string.Equals(NormalizePropertyKey(item.Key), entryKey, StringComparison.OrdinalIgnoreCase));
        if (existingIndex >= 0)
        {
            if (entry.Source < page.Items[existingIndex].Source)
            {
                page.Items[existingIndex] = entry;
            }

            return;
        }

        page.Items.Add(entry);
    }

    private static string NormalizePropertyKey(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        var chars = value.Where(char.IsLetterOrDigit).ToArray();
        return new string(chars);
    }

    private static List<PropertyPage> DeduplicatePropertyPages(List<PropertyPage> pages)
    {
        var merged = new List<PropertyPage>();
        foreach (var page in pages)
        {
            foreach (var entry in page.Items)
            {
                AddPropertyEntry(merged, page.Name, entry);
            }
        }

        return merged;
    }

    private static void OverlayUiaValues(List<PropertyPage> pages, UiaPropertySession uiaSession)
    {
        var uiaValues = uiaSession.ReadAllPages();
        foreach (var pagePair in uiaValues)
        {
            var pageName = pagePair.Key;
            var page = pages.FirstOrDefault(p => string.Equals(p.Name, pageName, StringComparison.OrdinalIgnoreCase));
            if (page == null)
            {
                page = new PropertyPage(pageName, new List<PropertyEntry>());
                pages.Add(page);
            }

            foreach (var propertyPair in pagePair.Value)
            {
                var key = NormalizePropertyKey(propertyPair.Key);
                var existing = page.Items.FirstOrDefault(item =>
                    string.Equals(NormalizePropertyKey(item.Key), key, StringComparison.OrdinalIgnoreCase));

                if (existing != null)
                {
                    existing.UiaValue = propertyPair.Value;
                    continue;
                }

                var entry = new PropertyEntry(propertyPair.Key, propertyPair.Key, pageName, (object?)propertyPair.Value);
                page.Items.Add(entry);
            }
        }
    }

    private static bool TrySetPropertyViaUiAutomation(PropertyPage page, PropertyEntry entry, object? value, UiaPropertySession? uiaSession, out string? message)
    {
        message = null;
        if (uiaSession == null)
        {
            return false;
        }

        var textValue = value?.ToString() ?? string.Empty;
        if (uiaSession.TrySetProperty(page.Name, entry.Name, textValue, out message))
        {
            entry.UiaValue = textValue;
            return true;
        }

        return false;
    }

    private static bool TryDeleteProperty(PropertyEntry entry, PropertyPage page, UiaPropertySession? uiaSession, out string? message)
    {
        message = null;
        if (entry.DeleteAction != null)
        {
            var result = entry.DeleteAction();
            message = result.Message;
            return result.Success;
        }

        if (TrySetPropertyViaUiAutomation(page, entry, string.Empty, uiaSession, out message))
        {
            if (string.IsNullOrWhiteSpace(message))
            {
                message = "Property cleared via UI Automation.";
            }

            return true;
        }

        if (string.IsNullOrWhiteSpace(message))
        {
            message = "Delete is not supported for this property.";
        }

        return false;
    }

    private static string NormalizePageName(string pageName)
    {
        if (string.IsNullOrWhiteSpace(pageName))
        {
            return "Application/General";
        }

        if (string.Equals(pageName, "Application/Resources", StringComparison.OrdinalIgnoreCase))
        {
            return "Application/Win32 Resources";
        }

        foreach (var known in KnownPropertyPages)
        {
            if (string.Equals(known, pageName, StringComparison.OrdinalIgnoreCase))
            {
                return known;
            }
        }

        return pageName;
    }

    private static bool IsPlatformSpecificPropertyName(string name)
    {
        return IsIosPropertyName(name)
            || IsAndroidPropertyName(name)
            || IsWindowsPropertyName(name)
            || IsTizenPropertyName(name);
    }

    private static bool IsIosPropertyName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            return false;
        }

        return name.Contains("Mtouch", StringComparison.OrdinalIgnoreCase)
               || name.Contains("Codesign", StringComparison.OrdinalIgnoreCase)
               || name.Contains("Provision", StringComparison.OrdinalIgnoreCase)
               || name.Contains("Entitlement", StringComparison.OrdinalIgnoreCase)
               || name.Contains("Ipa", StringComparison.OrdinalIgnoreCase)
               || name.Contains("iOS", StringComparison.OrdinalIgnoreCase)
               || name.StartsWith("Ios", StringComparison.OrdinalIgnoreCase)
               || name.Contains("Xcode", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsAndroidPropertyName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            return false;
        }

        return name.Contains("Android", StringComparison.OrdinalIgnoreCase)
               || name.StartsWith("Aapt", StringComparison.OrdinalIgnoreCase)
               || name.StartsWith("Dex", StringComparison.OrdinalIgnoreCase)
               || name.Contains("D8", StringComparison.OrdinalIgnoreCase)
               || name.Contains("R8", StringComparison.OrdinalIgnoreCase)
               || name.Contains("Proguard", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsWindowsPropertyName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            return false;
        }

        return name.Contains("Windows", StringComparison.OrdinalIgnoreCase)
               || name.Contains("Uap", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsTizenPropertyName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            return false;
        }

        return name.Contains("Tizen", StringComparison.OrdinalIgnoreCase);
    }

    private static string MapIosPropertyPageName(string propertyName)
    {
        bool Contains(string value) => propertyName.Contains(value, StringComparison.OrdinalIgnoreCase);

        if (Contains("Codesign") || Contains("Provision") || Contains("Entitlement") || Contains("Signing") || Contains("Team"))
        {
            return "ios/Bundle Signing";
        }

        if (Contains("Ipa") || Contains("Archive"))
        {
            return "ios/IPA Options";
        }

        if (Contains("Manifest") || Contains("InfoPlist") || Contains("Info.plist"))
        {
            return "ios/Manifest";
        }

        if (Contains("OnDemand"))
        {
            return "ios/On Demand Resources";
        }

        if (Contains("Debug"))
        {
            return "ios/debug";
        }

        if (Contains("Run") || Contains("Launch"))
        {
            return "ios/Run Options";
        }

        return "ios/Build";
    }

    private static string MapAndroidPropertyPageName(string propertyName)
    {
        return "Application/Android Targets";
    }

    private static string MapPropertyPageName(string category, string propertyName)
    {
        if (!string.IsNullOrWhiteSpace(category) && category.Contains("/", StringComparison.Ordinal))
        {
            return NormalizePageName(category);
        }

        if (IsIosPropertyName(propertyName))
        {
            return NormalizePageName(MapIosPropertyPageName(propertyName));
        }

        if (IsAndroidPropertyName(propertyName))
        {
            return NormalizePageName(MapAndroidPropertyPageName(propertyName));
        }

        if (IsWindowsPropertyName(propertyName))
        {
            return "Application/Windows Targets";
        }

        if (IsTizenPropertyName(propertyName))
        {
            return "Application/Tizen Targets";
        }

        if (!string.IsNullOrWhiteSpace(category))
        {
            var normalizedCategory = category.Trim();
            foreach (var known in KnownPropertyPages)
            {
                if (string.Equals(known, normalizedCategory, StringComparison.OrdinalIgnoreCase))
                {
                    return known;
                }
            }

            if (string.Equals(normalizedCategory, "Application", StringComparison.OrdinalIgnoreCase))
            {
                return "Application/General";
            }

            if (string.Equals(normalizedCategory, "Build", StringComparison.OrdinalIgnoreCase))
            {
                return "Build/General";
            }

            if (string.Equals(normalizedCategory, "Package", StringComparison.OrdinalIgnoreCase))
            {
                return "Package/General";
            }

            if (string.Equals(normalizedCategory, "Code Analysis", StringComparison.OrdinalIgnoreCase))
            {
                return "Code Analysis/All analyzers";
            }

            if (string.Equals(normalizedCategory, "Resources", StringComparison.OrdinalIgnoreCase))
            {
                return "Resources/General";
            }

            if (string.Equals(normalizedCategory, "Global Usings", StringComparison.OrdinalIgnoreCase))
            {
                return "Global Usings/General";
            }

            if (string.Equals(normalizedCategory, "MAUI Shared", StringComparison.OrdinalIgnoreCase))
            {
                return "MAUI Shared/General";
            }

            if (string.Equals(normalizedCategory, "iOS", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(normalizedCategory, "ios", StringComparison.OrdinalIgnoreCase))
            {
                return "ios/Build";
            }
        }

        var key = $"{category} {propertyName}".Trim().ToLowerInvariant();

        bool Contains(string value) => key.Contains(value, StringComparison.OrdinalIgnoreCase);

        if (Contains("global using") || Contains("implicitusing") || Contains("globalusing"))
        {
            return "Global Usings/General";
        }

        if (Contains("maui"))
        {
            return "MAUI Shared/General";
        }

        if (Contains("package") || propertyName.StartsWith("Package", StringComparison.OrdinalIgnoreCase))
        {
            if (Contains("license"))
            {
                return "Package/License";
            }

            if (Contains("symbol"))
            {
                return "Package/Symbols";
            }

            return "Package/General";
        }

        if (Contains("analysis") || Contains("analyzer"))
        {
            if (Contains(".net") || Contains("net analyzer") || Contains("netanalyzer"))
            {
                return "Code Analysis/.NET analyzers";
            }

            return "Code Analysis/All analyzers";
        }

        if (Contains("resource"))
        {
            if (Contains("win32"))
            {
                return "Application/Win32 Resources";
            }

            return "Resources/General";
        }

        if (Contains("dependency") || Contains("reference"))
        {
            return "Application/Dependencies";
        }

        if (Contains("application") || Contains("startup") || Contains("assembly") || Contains("rootnamespace") ||
            Contains("outputtype") || Contains("targetframework") || Contains("icon"))
        {
            if (Contains("win32"))
            {
                return "Application/Win32 Resources";
            }

            if (Contains("ios"))
            {
                return "Application/iOS Targets";
            }

            if (Contains("android"))
            {
                return "Application/Android Targets";
            }

            if (Contains("windows"))
            {
                return "Application/Windows Targets";
            }

            if (Contains("tizen"))
            {
                return "Application/Tizen Targets";
            }

            return "Application/General";
        }

        if (Contains("build") || Contains("defineconstant") || Contains("optimize") || Contains("warning") ||
            Contains("nullable") || Contains("debug"))
        {
            if (Contains("error") || Contains("warning"))
            {
                return "Build/Errors and warnings";
            }

            if (Contains("output") || Contains("outdir") || Contains("outputpath") || Contains("documentation"))
            {
                return "Build/Output";
            }

            if (Contains("prebuild") || Contains("postbuild") || Contains("event"))
            {
                return "Build/Events";
            }

            if (Contains("publish"))
            {
                return "Build/Publish";
            }

            if (Contains("sign") || Contains("strongname") || Contains("snk"))
            {
                return "Build/Strong naming";
            }

            if (Contains("advanced") || Contains("deterministic") || Contains("debuggable"))
            {
                return "Build/Advanced";
            }

            return "Build/General";
        }

        if (Contains("ios"))
        {
            if (Contains("bundle") || Contains("sign"))
            {
                return "ios/Bundle Signing";
            }

            if (Contains("ipa"))
            {
                return "ios/IPA Options";
            }

            if (Contains("manifest"))
            {
                return "ios/Manifest";
            }

            if (Contains("demand"))
            {
                return "ios/On Demand Resources";
            }

            if (Contains("run"))
            {
                return "ios/Run Options";
            }

            if (Contains("debug"))
            {
                return "ios/debug";
            }

            return "ios/Build";
        }

        if (!string.IsNullOrWhiteSpace(category))
        {
            return NormalizePageName(category);
        }

        return "Application/General";
    }

    private static string SafeFormatPropertyValue(PropertyEntry entry)
    {
        try
        {
            if (entry.UiaValue != null)
            {
                return FormatPropertyValue(entry.UiaValue);
            }

            if (entry.Source != PropertySource.MsBuild && entry.FallbackValue != null)
            {
                return FormatPropertyValue(entry.FallbackValue);
            }

            var value = GetPropertyValue(entry);
            return FormatPropertyValue(value);
        }
        catch (NotImplementedException)
        {
            if (entry.FallbackValue != null)
            {
                return FormatPropertyValue(entry.FallbackValue);
            }

            return "<unavailable: The method or operation is not implemented.>";
        }
        catch (COMException ex)
        {
            if (IsNotImplementedComError(ex) && entry.FallbackValue != null)
            {
                return FormatPropertyValue(entry.FallbackValue);
            }

            return $"<unavailable: {ex.Message}>";
        }
        catch (Exception ex)
        {
            if (entry.FallbackValue != null &&
                ex.Message.Contains("not implemented", StringComparison.OrdinalIgnoreCase))
            {
                return FormatPropertyValue(entry.FallbackValue);
            }

            return $"<unavailable: {ex.Message}>";
        }
    }

    private static bool IsNotImplementedComError(COMException ex)
    {
        const int ENotImpl = unchecked((int)0x80004001);
        return ex.HResult == ENotImpl
            || ex.Message.Contains("not implemented", StringComparison.OrdinalIgnoreCase);
    }

    private static object? GetPropertyValue(PropertyEntry entry)
    {
        if (entry.UiaValue != null)
        {
            return entry.UiaValue;
        }

        if (entry.Source != PropertySource.MsBuild && entry.FallbackValue != null)
        {
            return entry.FallbackValue;
        }

        if (entry.Getter != null)
        {
            return entry.Getter();
        }

        if (entry.Descriptor != null && entry.Owner != null)
        {
            return entry.Descriptor.GetValue(entry.Owner);
        }

        if (entry.ComProperty != null)
        {
            return entry.ComProperty.Value;
        }

        return null;
    }

    private static void SetPropertyValue(PropertyEntry entry, object? value)
    {
        if (entry.Setter != null)
        {
            entry.Setter(value);
            return;
        }

        if (entry.Descriptor != null && entry.Owner != null)
        {
            entry.Descriptor.SetValue(entry.Owner, value);
            return;
        }

        if (entry.ComProperty != null)
        {
            entry.ComProperty.Value = value;
        }
    }

    private static bool IsPropertyReadOnly(PropertyEntry entry)
    {
        if (entry.Setter != null)
        {
            return false;
        }

        if (entry.Getter != null && entry.Setter == null)
        {
            return true;
        }

        if (entry.Descriptor != null)
        {
            return entry.Descriptor.IsReadOnly;
        }

        if (entry.ComProperty != null)
        {
            return TryGetValue(() => (bool)entry.ComProperty.IsReadOnly, false)
                || TryGetValue(() => (bool)entry.ComProperty.ReadOnly, false);
        }

        return true;
    }

    private static string FormatPropertyValue(object? value)
    {
        if (value == null)
        {
            return "(null)";
        }

        if (value is Array array)
        {
            var parts = new List<string>();
            foreach (var item in array)
            {
                parts.Add(item?.ToString() ?? "(null)");
            }

            return string.Join(", ", parts);
        }

        return value.ToString() ?? "(null)";
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
            if (!EnsureVsInstanceAvailable(instances))
            {
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
            Console.Error.WriteLine($"Unexpected error: {FormatUnexpectedError(ex)}");
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
            if (!EnsureVsInstanceAvailable(instances))
            {
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
            Console.Error.WriteLine($"Unexpected error: {FormatUnexpectedError(ex)}");
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

    private const int VsUiSelectionTypeSelect = 1;

    private static bool TrySelectProjectInSolutionExplorer(dynamic dte, dynamic project)
    {
        try
        {
            if (TryInvoke(() => dte.ActiveSolutionProjects = new object[] { project }, out _))
            {
                return true;
            }
        }
        catch
        {
            return false;
        }

        return false;
    }

    private static dynamic? FindProjectHierarchyItem(dynamic item, dynamic project)
    {
        if (item == null)
        {
            return null;
        }

        var itemObject = TryGetValue(() => (dynamic)item.Object);
        if (itemObject != null)
        {
            var itemUnique = TryGetValue(() => (string?)itemObject.UniqueName);
            var projectUnique = TryGetValue(() => (string?)project.UniqueName);
            if (!string.IsNullOrWhiteSpace(itemUnique) && !string.IsNullOrWhiteSpace(projectUnique) &&
                string.Equals(itemUnique, projectUnique, StringComparison.OrdinalIgnoreCase))
            {
                return item;
            }

            var itemName = TryGetValue(() => (string?)itemObject.Name);
            var projectName = TryGetValue(() => (string?)project.Name);
            if (!string.IsNullOrWhiteSpace(itemName) && !string.IsNullOrWhiteSpace(projectName) &&
                string.Equals(itemName, projectName, StringComparison.OrdinalIgnoreCase))
            {
                return item;
            }
        }

        var children = TryGetValue(() => (dynamic)item.UIHierarchyItems);
        foreach (var child in EnumerateComCollection(children))
        {
            var found = FindProjectHierarchyItem(child, project);
            if (found != null)
            {
                return found;
            }
        }

        return null;
    }

    private static bool TryOpenProjectProperties(dynamic dte)
    {
        if (TryInvoke(() => dte.ExecuteCommand("Project.Properties"), out _))
        {
            return true;
        }

        if (TryInvoke(() => dte.ExecuteCommand("Project.ProjectProperties"), out _))
        {
            return true;
        }

        if (TryInvoke(() => dte.ExecuteCommand("File.Properties"), out _))
        {
            return true;
        }

        return TryInvoke(() => dte.ExecuteCommand("View.PropertiesWindow"), out _);
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

        var patterns = ExpandProjectPatterns(selector);

        bool Match(string? candidate)
        {
            if (string.IsNullOrWhiteSpace(candidate))
            {
                return false;
            }

            foreach (var pattern in patterns)
            {
                if (IsWildcardMatch(candidate, pattern))
                {
                    return true;
                }
            }

            return false;
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

    private static IEnumerable<string> ExpandProjectPatterns(string selector)
    {
        var patterns = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var trimmed = selector.Trim();
        if (string.IsNullOrWhiteSpace(trimmed))
        {
            return patterns;
        }

        void Add(string value)
        {
            if (!string.IsNullOrWhiteSpace(value))
            {
                patterns.Add(value);
            }
        }

        Add(trimmed);

        if (trimmed.Contains('*'))
        {
            if (!trimmed.StartsWith("*", StringComparison.Ordinal))
            {
                Add("*" + trimmed);
            }

            if (!trimmed.EndsWith("*", StringComparison.Ordinal))
            {
                Add(trimmed + "*");
            }

            var core = trimmed.Trim('*').Trim();
            if (!string.IsNullOrWhiteSpace(core))
            {
                Add("*" + core + "*");
            }
        }
        else
        {
            Add("*" + trimmed + "*");
        }

        return patterns;
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

    private enum PropertySource
    {
        UserFile = -2,
        UiAutomation = -1,
        MsBuild = 0,
        DteDescriptor = 1,
        ComProperty = 2
    }

    private readonly struct MsBuildPropertyValue
    {
        public MsBuildPropertyValue(string name, object? value)
        {
            Name = name;
            Value = value;
        }

        public string Name { get; }
        public object? Value { get; }
    }

    private sealed class DotnetSdkInfo
    {
        public DotnetSdkInfo(Version version, string path)
        {
            Version = version;
            Path = path;
        }

        public Version Version { get; }
        public string Path { get; }
    }

    private sealed class RuleDefinition
    {
        public RuleDefinition(string name, string displayName)
        {
            Name = name;
            DisplayName = displayName;
            Categories = new Dictionary<string, RuleCategoryDefinition>(StringComparer.OrdinalIgnoreCase);
            Properties = new List<RulePropertyDefinition>();
        }

        public string Name { get; }
        public string DisplayName { get; }
        public string? Description { get; set; }
        public string? AppliesTo { get; set; }
        public string? PageTemplate { get; set; }
        public string? DataSourcePersistence { get; set; }
        public string? DataSourceItemType { get; set; }
        public Dictionary<string, RuleCategoryDefinition> Categories { get; }
        public List<RulePropertyDefinition> Properties { get; }
    }

    private sealed class RuleCategoryDefinition
    {
        public RuleCategoryDefinition(string name, string displayName)
        {
            Name = name;
            DisplayName = displayName;
        }

        public string Name { get; }
        public string DisplayName { get; }
        public string? Description { get; set; }
    }

    private sealed class RulePropertyDefinition
    {
        public RulePropertyDefinition(string name, string displayName)
        {
            Name = name;
            DisplayName = displayName;
            EnumValues = new List<string>();
            EnumValueMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        public string Name { get; }
        public string DisplayName { get; }
        public string? Description { get; set; }
        public string? Category { get; set; }
        public string? CategoryDisplayName { get; set; }
        public string? PropertyKind { get; set; }
        public bool Visible { get; set; } = true;
        public string? DataSourcePersistence { get; set; }
        public string? DataSourceItemType { get; set; }
        public string? PersistedName { get; set; }
        public List<string> EnumValues { get; }
        public Dictionary<string, string> EnumValueMap { get; }
    }

    private sealed class PlatformPropertyDefinition
    {
        public PlatformPropertyDefinition(string name, string pageName, Type propertyType)
        {
            Name = name;
            PageName = pageName;
            PropertyType = propertyType;
        }

        public string Name { get; }
        public string PageName { get; }
        public Type PropertyType { get; }
        public string? Description { get; set; }
    }

    private sealed class MsBuildEvaluation : IDisposable
    {
        public MsBuildEvaluation(ProjectCollection collection, Project project)
        {
            Collection = collection;
            Project = project;
        }

        public ProjectCollection Collection { get; }
        public Project Project { get; }

        public void Dispose()
        {
            Collection.UnloadAllProjects();
            Collection.Dispose();
        }
    }

    private readonly struct UiaRow
    {
        public UiaRow(string name, string value, AutomationElement? valueElement)
        {
            Name = name;
            Value = value;
            ValueElement = valueElement;
        }

        public string Name { get; }
        public string Value { get; }
        public AutomationElement? ValueElement { get; }
    }

    private sealed class UiaPropertySession : IDisposable
    {
        private const int UiDelayMilliseconds = 150;
        private readonly AutomationElement _vsRoot;
        private readonly AutomationElement _propertyRoot;
        private readonly AutomationElement _pageTree;
        private readonly Dictionary<string, AutomationElement> _pageItems;

        private UiaPropertySession(
            AutomationElement vsRoot,
            AutomationElement propertyRoot,
            AutomationElement pageTree,
            Dictionary<string, AutomationElement> pageItems)
        {
            _vsRoot = vsRoot;
            _propertyRoot = propertyRoot;
            _pageTree = pageTree;
            _pageItems = pageItems;
        }

        public static UiaPropertySession? TryCreate(dynamic dte, dynamic project)
        {
            try
            {
                UiaDebug("Starting UI Automation session.");
                TryInvoke(() => dte.MainWindow.Activate(), out _);
                var projectSelected = TrySelectProjectInSolutionExplorer(dte, project);
                if (!projectSelected)
                {
                    LogUiaUnavailable("Project could not be selected in Solution Explorer.");
                }
                else
                {
                    UiaDebug("Active project set.");
                }

                if (!TryOpenProjectProperties(dte))
                {
                    LogUiaUnavailable("Project properties command is unavailable.");
                }
                else
                {
                    UiaDebug("Project properties command invoked.");
                }

                var hwnd = TryGetValue(() => (int)dte.MainWindow.HWnd, 0);
                if (hwnd == 0)
                {
                    LogUiaUnavailable("Main window handle unavailable.");
                    return null;
                }

                var vsRoot = AutomationElement.FromHandle(new IntPtr(hwnd));
                if (vsRoot == null)
                {
                    LogUiaUnavailable("Unable to access Visual Studio UI Automation root.");
                    return null;
                }

                var processId = vsRoot.Current.ProcessId;
                var projectName = TryGetValue(() => (string?)project.Name) ?? string.Empty;
                AutomationElement? propertyRoot = null;
                var allowUnknownTree = false;
                var propertyWindowHandle = WaitForPropertyWindowHandle(dte, projectName);
                if (propertyWindowHandle != 0)
                {
                    UiaDebug("Checking DTE window handle for property pages.");
                    var windowRoot = AutomationElement.FromHandle(new IntPtr(propertyWindowHandle));
                    if (windowRoot != null && FindPropertyPagesTree(windowRoot, allowUnknownTree: true) != null)
                    {
                        propertyRoot = windowRoot;
                        allowUnknownTree = true;
                    }
                }

                if (propertyRoot == null)
                {
                    UiaDebug("Searching for property pages UI.");
                }
                propertyRoot ??= WaitForPropertyPagesRoot(vsRoot, projectName, processId);
                if (propertyRoot == null)
                {
                    LogUiaUnavailable("Project property pages UI not found.");
                    return null;
                }

                UiaDebug("Property pages UI found.");
                var pageTree = FindPropertyPagesTree(propertyRoot, allowUnknownTree);
                if (pageTree == null)
                {
                    UiaDebug("Property pages tree not found with known names; retrying without name checks.");
                    pageTree = FindPropertyPagesTree(propertyRoot, allowUnknownTree: true);
                }
                if (pageTree == null)
                {
                    LogUiaUnavailable("Property pages navigation tree not found.");
                    return null;
                }

                var pages = BuildPageIndex(pageTree);
                if (pages.Count == 0)
                {
                    LogUiaUnavailable("Property pages could not be enumerated.");
                    return null;
                }

                UiaDebug($"Property pages enumerated: {pages.Count}.");
                return new UiaPropertySession(vsRoot, propertyRoot, pageTree, pages);
            }
            catch (Exception ex)
            {
                LogUiaUnavailable(ex.Message);
                return null;
            }
        }

        public Dictionary<string, Dictionary<string, string>> ReadAllPages()
        {
            var result = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);
            var totalTimeoutMs = GetUiaTimeoutMs("VX_UIA_TIMEOUT_MS", 8000);
            var perPageTimeoutMs = GetUiaTimeoutMs("VX_UIA_PAGE_TIMEOUT_MS", 1500);
            var start = Environment.TickCount64;
            foreach (var page in _pageItems)
            {
                if (totalTimeoutMs > 0 && Environment.TickCount64 - start > totalTimeoutMs)
                {
                    UiaDebug("UI Automation time budget reached; stopping page traversal.");
                    break;
                }

                UiaDebug($"Reading page '{page.Key}'.");
                if (!TrySelectPage(page.Key))
                {
                    continue;
                }

                Thread.Sleep(UiDelayMilliseconds);
                var remaining = totalTimeoutMs > 0
                    ? (int)Math.Max(0, totalTimeoutMs - (Environment.TickCount64 - start))
                    : perPageTimeoutMs;
                var timeoutMs = Math.Min(perPageTimeoutMs, remaining);
                if (timeoutMs <= 0)
                {
                    UiaDebug("UI Automation time budget exhausted before reading page.");
                    break;
                }

                var rows = ReadPropertyRowsWithTimeout(_propertyRoot, timeoutMs);
                if (rows.Count == 0)
                {
                    continue;
                }

                var pageValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var row in rows)
                {
                    if (!pageValues.ContainsKey(row.Name))
                    {
                        pageValues[row.Name] = row.Value;
                    }
                }

                result[page.Key] = pageValues;
            }

            return result;
        }

        public bool TrySetProperty(string pageName, string propertyName, string newValue, out string? message)
        {
            message = null;
            if (!TrySelectPage(pageName))
            {
                message = "Property page could not be selected.";
                return false;
            }

            Thread.Sleep(UiDelayMilliseconds);
            var timeoutMs = GetUiaTimeoutMs("VX_UIA_PAGE_TIMEOUT_MS", 1500);
            var rows = ReadPropertyRowsWithTimeout(_propertyRoot, timeoutMs);
            var targetKey = NormalizePropertyKey(propertyName);
            var row = rows.FirstOrDefault(r => NormalizePropertyKey(r.Name) == targetKey);
            if (string.IsNullOrWhiteSpace(row.Name))
            {
                message = "Property not found in UI Automation view.";
                return false;
            }

            if (row.ValueElement == null)
            {
                message = "Property value control not found.";
                return false;
            }

            if (TrySetUiaValue(row.ValueElement, newValue, out message))
            {
                return true;
            }

            return false;
        }

        private bool TrySelectPage(string pageName)
        {
            if (_pageItems.TryGetValue(pageName, out var item))
            {
                var timeoutMs = GetUiaTimeoutMs("VX_UIA_SELECT_TIMEOUT_MS", 800);
                return TrySelectTreeItemWithTimeout(item, timeoutMs);
            }

            var match = _pageItems.Keys.FirstOrDefault(key =>
                key.EndsWith(pageName, StringComparison.OrdinalIgnoreCase) ||
                pageName.EndsWith(key, StringComparison.OrdinalIgnoreCase));

            if (match != null && _pageItems.TryGetValue(match, out var matchedItem))
            {
                var timeoutMs = GetUiaTimeoutMs("VX_UIA_SELECT_TIMEOUT_MS", 800);
                return TrySelectTreeItemWithTimeout(matchedItem, timeoutMs);
            }

            return false;
        }

        private static int WaitForPropertyWindowHandle(dynamic dte, string projectName)
        {
            for (var i = 0; i < 10; i++)
            {
                var handle = TryGetPropertyWindowHandle(dte, projectName);
                if (handle != 0)
                {
                    return handle;
                }

                Thread.Sleep(200);
            }

            return 0;
        }

        private static int TryGetPropertyWindowHandle(dynamic dte, string projectName)
        {
            var activeWindow = TryGetValue(() => (dynamic)dte.ActiveWindow);
            var activeHandle = TryGetValue(() => (int)activeWindow.HWnd, 0);
            var activeCaption = TryGetValue(() => (string?)activeWindow.Caption) ?? string.Empty;
            var fallbackHandle = activeHandle;
            if (activeHandle != 0 && WindowCaptionMatchesProject(activeCaption, projectName))
            {
                return activeHandle;
            }

            var windows = TryGetValue(() => (dynamic)dte.Windows);
            foreach (var window in EnumerateComCollection(windows))
            {
                var caption = TryGetValue(() => (string?)window.Caption) ?? string.Empty;
                if (!WindowCaptionMatchesProject(caption, projectName))
                {
                    continue;
                }

                var hwnd = TryGetValue(() => (int)window.HWnd, 0);
                if (hwnd != 0)
                {
                    return hwnd;
                }
            }

            return fallbackHandle;
        }

        private static bool WindowCaptionMatchesProject(string caption, string projectName)
        {
            if (string.IsNullOrWhiteSpace(caption))
            {
                return false;
            }

            if (!string.IsNullOrWhiteSpace(projectName) &&
                caption.Contains(projectName, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return caption.Contains("Properties", StringComparison.OrdinalIgnoreCase);
        }

        private static AutomationElement? WaitForPropertyPagesRoot(AutomationElement vsRoot, string projectName, int processId)
        {
            for (var i = 0; i < 25; i++)
            {
                var candidate = FindPropertyPagesRoot(vsRoot, projectName, processId);
                if (candidate != null)
                {
                    return candidate;
                }

                Thread.Sleep(200);
            }

            return null;
        }

        private static AutomationElement? FindPropertyPagesRoot(AutomationElement vsRoot, string projectName, int processId)
        {
            return FindPropertyPagesRootByTraversal(vsRoot);
        }

        private static AutomationElement? FindPropertyPagesRootByTraversal(AutomationElement root)
        {
            const int maxNodes = 5000;
            const int maxDepth = 8;
            var queue = new Queue<(AutomationElement Element, int Depth)>();
            queue.Enqueue((root, 0));
            var walker = TreeWalker.ControlViewWalker;
            var checkedRoots = new HashSet<int>();
            var visited = 0;

            while (queue.Count > 0)
            {
                var (element, depth) = queue.Dequeue();
                visited++;
                if (visited > maxNodes)
                {
                    break;
                }

                var controlType = element.Current.ControlType;
                if (controlType == ControlType.Tree)
                {
                    if (ContainsKnownPage(element))
                    {
                        return GetPropertyRootFromTree(element);
                    }

                    var candidateRoot = GetPropertyRootFromTree(element);
                    if (ReferenceEquals(candidateRoot, element))
                    {
                        continue;
                    }

                    var nativeHandle = 0;
                    try
                    {
                        nativeHandle = candidateRoot.Current.NativeWindowHandle;
                    }
                    catch
                    {
                        nativeHandle = 0;
                    }

                    if (nativeHandle != 0 && !checkedRoots.Add(nativeHandle))
                    {
                        continue;
                    }

                    return candidateRoot;
                }

                if (depth >= maxDepth)
                {
                    continue;
                }

                AutomationElement? child = null;
                try
                {
                    child = walker.GetFirstChild(element);
                }
                catch
                {
                    child = null;
                }

                while (child != null)
                {
                    queue.Enqueue((child, depth + 1));
                    try
                    {
                        child = walker.GetNextSibling(child);
                    }
                    catch
                    {
                        break;
                    }
                }
            }

            UiaDebug("UI Automation traversal did not locate property pages tree.");
            return null;
        }

        private static AutomationElement GetPropertyRootFromTree(AutomationElement tree)
        {
            var walker = TreeWalker.ControlViewWalker;
            var current = tree;
            AutomationElement? best = null;
            while (true)
            {
                AutomationElement? parent;
                try
                {
                    parent = walker.GetParent(current);
                }
                catch
                {
                    break;
                }

                if (parent == null)
                {
                    break;
                }

                var type = parent.Current.ControlType;
                if (type == ControlType.Window || type == ControlType.Pane || type == ControlType.Document)
                {
                    if (FindPropertyGrid(parent) != null)
                    {
                        best = parent;
                    }
                }

                current = parent;
            }

            return best ?? tree;
        }



        private static AutomationElement? FindPropertyPagesTree(AutomationElement root, bool allowUnknownTree = false)
        {
            var tree = root.FindFirst(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Tree));
            if (tree != null && (allowUnknownTree || ContainsKnownPage(tree)))
            {
                return tree;
            }

            var trees = root.FindAll(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Tree));
            foreach (AutomationElement candidate in trees)
            {
                if (allowUnknownTree || ContainsKnownPage(candidate))
                {
                    return candidate;
                }
            }

            return null;
        }

        private static bool ContainsKnownPage(AutomationElement tree)
        {
            var items = tree.FindAll(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TreeItem));

            foreach (AutomationElement item in items)
            {
                var name = item.Current.Name;
                if (string.IsNullOrWhiteSpace(name))
                {
                    continue;
                }

                if (KnownPropertyPages.Any(page => page.EndsWith(name, StringComparison.OrdinalIgnoreCase) ||
                                                   string.Equals(page, name, StringComparison.OrdinalIgnoreCase)))
                {
                    return true;
                }
            }

            return false;
        }

        private static Dictionary<string, AutomationElement> BuildPageIndex(AutomationElement tree)
        {
            var map = new Dictionary<string, AutomationElement>(StringComparer.OrdinalIgnoreCase);
            var rootItems = tree.FindAll(TreeScope.Children,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TreeItem));
            var remaining = 200;

            foreach (AutomationElement item in rootItems)
            {
                AddTreeItem(map, item, null, ref remaining);
                if (remaining <= 0)
                {
                    break;
                }
            }

            return map;
        }

        private static void AddTreeItem(Dictionary<string, AutomationElement> map, AutomationElement item, string? parentPath, ref int remaining)
        {
            if (remaining <= 0)
            {
                return;
            }

            var name = item.Current.Name;
            if (string.IsNullOrWhiteSpace(name))
            {
                return;
            }

            var path = string.IsNullOrWhiteSpace(parentPath) ? name : $"{parentPath}/{name}";
            map[path] = item;
            remaining--;
            if (remaining <= 0)
            {
                return;
            }

            if (TryExpand(item))
            {
                Thread.Sleep(UiDelayMilliseconds);
            }

            var children = item.FindAll(TreeScope.Children,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TreeItem));

            foreach (AutomationElement child in children)
            {
                AddTreeItem(map, child, path, ref remaining);
                if (remaining <= 0)
                {
                    return;
                }
            }
        }

        private static bool TrySelectTreeItem(AutomationElement item)
        {
            if (item.TryGetCurrentPattern(SelectionItemPattern.Pattern, out var selection))
            {
                ((SelectionItemPattern)selection).Select();
                return true;
            }

            if (item.TryGetCurrentPattern(InvokePattern.Pattern, out var invoke))
            {
                ((InvokePattern)invoke).Invoke();
                return true;
            }

            try
            {
                item.SetFocus();
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool TrySelectTreeItemWithTimeout(AutomationElement item, int timeoutMs)
        {
            if (timeoutMs <= 0)
            {
                return false;
            }

            var result = false;
            Exception? error = null;
            var thread = new Thread(() =>
            {
                try
                {
                    result = TrySelectTreeItem(item);
                }
                catch (Exception ex)
                {
                    error = ex;
                }
            })
            {
                IsBackground = true
            };

            try
            {
                thread.SetApartmentState(ApartmentState.STA);
            }
            catch
            {
                // Ignore STA failures.
            }

            thread.Start();
            if (!thread.Join(timeoutMs))
            {
                UiaDebug("UI Automation tree selection timed out.");
                return false;
            }

            if (error != null)
            {
                UiaDebug($"UI Automation tree selection failed: {error.Message}");
                return false;
            }

            return result;
        }

        private static bool TryExpand(AutomationElement item)
        {
            if (item.TryGetCurrentPattern(ExpandCollapsePattern.Pattern, out var pattern))
            {
                var expand = (ExpandCollapsePattern)pattern;
                if (expand.Current.ExpandCollapseState != ExpandCollapseState.Expanded)
                {
                    expand.Expand();
                }

                return true;
            }

            return false;
        }

        private static List<UiaRow> ReadPropertyRows(AutomationElement root)
        {
            var rows = new List<UiaRow>();
            var grid = FindPropertyGrid(root);
            if (grid == null)
            {
                UiaDebug("Property grid not found for current page.");
                return rows;
            }

            foreach (var row in EnumerateRowItems(grid))
            {
                if (TryParseRow(row, out var name, out var value, out var valueElement))
                {
                    rows.Add(new UiaRow(name, value, valueElement));
                }
            }

            return rows;
        }

        private static List<UiaRow> ReadPropertyRowsWithTimeout(AutomationElement root, int timeoutMs)
        {
            if (timeoutMs <= 0)
            {
                return new List<UiaRow>();
            }

            List<UiaRow>? result = null;
            Exception? error = null;
            var thread = new Thread(() =>
            {
                try
                {
                    result = ReadPropertyRows(root);
                }
                catch (Exception ex)
                {
                    error = ex;
                }
            })
            {
                IsBackground = true
            };

            try
            {
                thread.SetApartmentState(ApartmentState.STA);
            }
            catch
            {
                // Ignore STA failures.
            }

            thread.Start();
            if (!thread.Join(timeoutMs))
            {
                UiaDebug("Property grid read timed out.");
                return new List<UiaRow>();
            }

            if (error != null)
            {
                UiaDebug($"Property grid read failed: {error.Message}");
                return new List<UiaRow>();
            }

            return result ?? new List<UiaRow>();
        }

        private static IEnumerable<AutomationElement> EnumerateRowItems(AutomationElement grid)
        {
            const int maxRows = 400;
            const int maxNodes = 4000;
            var walker = TreeWalker.ControlViewWalker;
            var queue = new Queue<AutomationElement>();
            queue.Enqueue(grid);
            var visited = 0;
            var rows = 0;

            while (queue.Count > 0 && visited < maxNodes && rows < maxRows)
            {
                var current = queue.Dequeue();
                visited++;

                var controlType = current.Current.ControlType;
                if (controlType == ControlType.DataItem || controlType == ControlType.ListItem)
                {
                    rows++;
                    yield return current;
                }

                AutomationElement? child = null;
                try
                {
                    child = walker.GetFirstChild(current);
                }
                catch
                {
                    child = null;
                }

                while (child != null)
                {
                    queue.Enqueue(child);
                    try
                    {
                        child = walker.GetNextSibling(child);
                    }
                    catch
                    {
                        break;
                    }
                }
            }
        }

        private static AutomationElement? FindPropertyGrid(AutomationElement root)
        {
            var gridCondition = new OrCondition(
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.DataGrid),
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Table),
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.List));

            return root.FindFirst(TreeScope.Descendants, gridCondition);
        }

        private static bool TryParseRow(AutomationElement row, out string name, out string value, out AutomationElement? valueElement)
        {
            name = string.Empty;
            value = string.Empty;
            valueElement = null;

            var candidates = row.FindAll(TreeScope.Descendants, new OrCondition(
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text),
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit),
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ComboBox),
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.CheckBox)));

            if (candidates.Count == 0)
            {
                return false;
            }

            AutomationElement? nameElement = null;
            var nameRect = new Rect(double.MaxValue, 0, 0, 0);
            foreach (AutomationElement candidate in candidates)
            {
                if (candidate.Current.ControlType != ControlType.Text)
                {
                    continue;
                }

                var candidateName = candidate.Current.Name;
                if (string.IsNullOrWhiteSpace(candidateName))
                {
                    continue;
                }

                var rect = candidate.Current.BoundingRectangle;
                if (rect == Rect.Empty)
                {
                    continue;
                }

                if (rect.X < nameRect.X)
                {
                    nameElement = candidate;
                    nameRect = rect;
                }
            }

            if (nameElement == null)
            {
                return false;
            }

            name = nameElement.Current.Name;

            AutomationElement? valueCandidate = null;
            var valueRect = new Rect(double.MinValue, 0, 0, 0);
            foreach (AutomationElement candidate in candidates)
            {
                if (candidate == nameElement)
                {
                    continue;
                }

                var rect = candidate.Current.BoundingRectangle;
                if (rect == Rect.Empty)
                {
                    continue;
                }

                if (rect.X > valueRect.X)
                {
                    valueCandidate = candidate;
                    valueRect = rect;
                }
            }

            if (valueCandidate == null)
            {
                return false;
            }

            value = ReadUiaValue(valueCandidate);
            valueElement = valueCandidate;
            return true;
        }

        private static string ReadUiaValue(AutomationElement element)
        {
            if (element.TryGetCurrentPattern(ValuePattern.Pattern, out var valuePattern))
            {
                return ((ValuePattern)valuePattern).Current.Value ?? string.Empty;
            }

            if (element.TryGetCurrentPattern(TogglePattern.Pattern, out var togglePattern))
            {
                var state = ((TogglePattern)togglePattern).Current.ToggleState;
                return state == ToggleState.On ? "True" : "False";
            }

            if (element.TryGetCurrentPattern(SelectionPattern.Pattern, out var selectionPattern))
            {
                var selection = (SelectionPattern)selectionPattern;
                var selected = selection.Current.GetSelection();
                if (selected.Length > 0)
                {
                    return selected[0].Current.Name;
                }
            }

            return element.Current.Name ?? string.Empty;
        }

        private static bool TrySetUiaValue(AutomationElement element, string value, out string? message)
        {
            message = null;

            if (element.TryGetCurrentPattern(ValuePattern.Pattern, out var valuePattern))
            {
                var pattern = (ValuePattern)valuePattern;
                if (pattern.Current.IsReadOnly)
                {
                    message = "Value is read-only.";
                    return false;
                }

                pattern.SetValue(value);
                return true;
            }

            if (element.TryGetCurrentPattern(TogglePattern.Pattern, out var togglePattern))
            {
                if (!TryParseBool(value, out var boolValue))
                {
                    message = "Expected a boolean value (true/false).";
                    return false;
                }

                var pattern = (TogglePattern)togglePattern;
                var desired = boolValue ? ToggleState.On : ToggleState.Off;
                if (pattern.Current.ToggleState != desired)
                {
                    pattern.Toggle();
                }

                return true;
            }

            if (element.TryGetCurrentPattern(SelectionPattern.Pattern, out var selectionPattern))
            {
                var selection = (SelectionPattern)selectionPattern;
                var options = element.FindAll(TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem));

                foreach (AutomationElement option in options)
                {
                    if (string.Equals(option.Current.Name, value, StringComparison.OrdinalIgnoreCase) &&
                        option.TryGetCurrentPattern(SelectionItemPattern.Pattern, out var selectionItem))
                    {
                        ((SelectionItemPattern)selectionItem).Select();
                        return true;
                    }
                }

                message = "Selection value not found.";
                return false;
            }

            if (element.TryGetCurrentPattern(ExpandCollapsePattern.Pattern, out var expandPattern))
            {
                var pattern = (ExpandCollapsePattern)expandPattern;
                pattern.Expand();
                Thread.Sleep(UiDelayMilliseconds);

                var root = AutomationElement.RootElement;
                var listItems = root.FindAll(TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem));

                foreach (AutomationElement option in listItems)
                {
                    if (string.Equals(option.Current.Name, value, StringComparison.OrdinalIgnoreCase) &&
                        option.TryGetCurrentPattern(SelectionItemPattern.Pattern, out var selectionItem))
                    {
                        ((SelectionItemPattern)selectionItem).Select();
                        return true;
                    }
                }

                pattern.Collapse();
                message = "Selection value not found.";
                return false;
            }

            message = "Unsupported value control.";
            return false;
        }

        public void Dispose()
        {
        }
    }

    private sealed class PropertyPage
    {
        public PropertyPage(string name, List<PropertyEntry> items)
        {
            Name = name;
            Items = items;
        }

        public string Name { get; }
        public List<PropertyEntry> Items { get; }
    }

    private sealed class PropertyEntry
    {
        public PropertyEntry(string name, string key, dynamic comProperty, string category)
        {
            Name = name;
            Key = string.IsNullOrWhiteSpace(key) ? name : key;
            ComProperty = comProperty;
            Category = category;
            Source = PropertySource.ComProperty;
        }

        public PropertyEntry(string name, string key, object owner, PropertyDescriptor descriptor, string category)
        {
            Name = name;
            Key = string.IsNullOrWhiteSpace(key) ? name : key;
            Owner = owner;
            Descriptor = descriptor;
            Category = category;
            PropertyType = descriptor.PropertyType;
            Source = PropertySource.DteDescriptor;
        }

        public PropertyEntry(string name, string key, string category, Func<object?> getter, Action<object?>? setter, Type? propertyType, PropertySource source)
        {
            Name = name;
            Key = string.IsNullOrWhiteSpace(key) ? name : key;
            Category = category;
            Getter = getter;
            Setter = setter;
            PropertyType = propertyType;
            Source = source;
        }

        public PropertyEntry(string name, string key, string category, object? uiValue)
        {
            Name = name;
            Key = string.IsNullOrWhiteSpace(key) ? name : key;
            Category = category;
            UiaValue = uiValue;
            Source = PropertySource.UiAutomation;
        }

        public string Name { get; }
        public string Key { get; }
        public string Category { get; }
        public object? Owner { get; }
        public PropertyDescriptor? Descriptor { get; }
        public dynamic? ComProperty { get; }
        public Func<object?>? Getter { get; }
        public Action<object?>? Setter { get; }
        public Type? PropertyType { get; }
        public PropertySource Source { get; }
        public object? FallbackValue { get; set; }
        public object? UiaValue { get; set; }
        public Func<(bool Success, string? Message)>? DeleteAction { get; set; }
        public IReadOnlyList<string>? AllowedValues { get; set; }
        public Dictionary<string, string>? AllowedValueMap { get; set; }
        public string? Description { get; set; }
    }

    private sealed class ProjectEntry
    {
        public ProjectEntry(dynamic project, string displayName, string? uniqueName)
        {
            Project = project;
            DisplayName = displayName;
            UniqueName = uniqueName;
        }

        public dynamic Project { get; }
        public string DisplayName { get; }
        public string? UniqueName { get; }
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
