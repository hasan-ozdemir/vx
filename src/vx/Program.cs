using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text.Json;
using System.Text.RegularExpressions;
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
            Console.Error.WriteLine($"Unexpected error: {ex.Message}");
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
            return RunPropertyWizardForProject(project, includeMsBuild);
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

    private static int RunPropertyWizardForProject(dynamic project, bool includeMsBuild)
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
            if (RunPropertyWizardForPage(page) == 0)
            {
                continue;
            }
        }
    }

    private static int RunPropertyWizardForPage(PropertyPage page)
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
            UpdatePropertyValue(item);
        }
    }

    private static void UpdatePropertyValue(PropertyEntry entry)
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
        Console.Write($"-enter new value <{formattedCurrent}>: ");
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
        }
        catch (COMException ex)
        {
            Console.WriteLine($"Failed to update property: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to update property: {ex.Message}");
        }
    }

    private static bool TryConvertPropertyValue(PropertyEntry entry, string input, object? currentValue, out object? converted, out string hint)
    {
        hint = string.Empty;
        converted = input;

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

    private static bool MsBuildRegistrationAttempted;
    private static bool MsBuildRegistrationSucceeded;
    private static bool MsBuildRegistrationLogged;

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

    private static List<PropertyPage> GetProjectPropertyPages(object? project, bool includeMsBuild)
    {
        var pages = new List<PropertyPage>();
        foreach (var known in KnownPropertyPages)
        {
            pages.Add(new PropertyPage(known, new List<PropertyEntry>()));
        }

        dynamic? dynamicProject = project;
        var projectObject = TryGetValue(() => (object?)dynamicProject?.Object) ?? project;
        AddPropertyPagesFromObject(pages, projectObject);

        var configManager = TryGetValue(() => (dynamic)dynamicProject?.ConfigurationManager);
        var activeConfig = TryGetValue(() => (dynamic)configManager?.ActiveConfiguration);
        var configObject = TryGetValue(() => (object?)activeConfig?.Object) ?? (object?)activeConfig;
        AddPropertyPagesFromObject(pages, configObject);

        var projectProps = TryGetValue(() => (dynamic)dynamicProject?.Properties);
        AddPropertyPageFromComProperties(pages, projectProps);

        var configProps = TryGetValue(() => (dynamic)activeConfig?.Properties);
        AddPropertyPageFromComProperties(pages, configProps);

        if (includeMsBuild)
        {
            AddPropertyPagesFromMsBuild(pages, dynamicProject);
        }

        foreach (var page in pages)
        {
            page.Items.Sort((a, b) => string.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase));
        }

        return pages
            .OrderBy(p =>
            {
                var index = Array.IndexOf(KnownPropertyPages, p.Name);
                return index < 0 ? int.MaxValue : index;
            })
            .ThenBy(p => p.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static void AddPropertyPagesFromObject(List<PropertyPage> pages, object? target)
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

                var displayName = string.IsNullOrWhiteSpace(descriptor.DisplayName) ? descriptor.Name : descriptor.DisplayName;
                var category = string.IsNullOrWhiteSpace(descriptor.Category) ? string.Empty : descriptor.Category;
                var pageName = MapPropertyPageName(category, displayName);
                AddPropertyEntry(pages, pageName, new PropertyEntry(displayName, target, descriptor, category));
            }
            catch
            {
                // Skip properties that throw during descriptor access.
            }
        }
    }

    private static void AddPropertyPageFromComProperties(List<PropertyPage> pages, dynamic properties)
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
            AddPropertyEntry(pages, pageName, new PropertyEntry(propName, prop, pageName));
        }
    }

    private static void AddPropertyPagesFromMsBuild(List<PropertyPage> pages, dynamic project)
    {
        var projectPath = TryGetValue(() => (string?)project.FullName)
            ?? TryGetValue(() => (string?)project.FileName);

        if (string.IsNullOrWhiteSpace(projectPath) || !File.Exists(projectPath))
        {
            return;
        }

        var globalProperties = GetMsBuildGlobalProperties(project);

        ProjectCollection? collection = null;
        try
        {
            collection = new ProjectCollection(globalProperties);
            var msbuildProject = collection.LoadProject(projectPath);

            foreach (var prop in msbuildProject.AllEvaluatedProperties)
            {
                var name = prop.Name;
                if (!ShouldIncludeMsBuildProperty(name))
                {
                    continue;
                }

                var pageName = MapPropertyPageName(string.Empty, name);
                var category = pageName;
                object? cachedValue = prop.EvaluatedValue;
                var inferredType = InferMsBuildPropertyType(prop.EvaluatedValue);

                Action<object?> setter = newValue =>
                {
                    SetMsBuildProperty(projectPath, globalProperties, name, newValue);
                    cachedValue = newValue;
                };

                Func<object?> getter = () => cachedValue;
                var entry = new PropertyEntry(name, category, getter, setter, inferredType, PropertySource.MsBuild);
                AddPropertyEntry(pages, pageName, entry);
            }
        }
        catch
        {
            // Skip MSBuild evaluation issues.
        }
        finally
        {
            if (collection != null)
            {
                collection.UnloadAllProjects();
                collection.Dispose();
            }
        }
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
            if (!MSBuildLocator.IsRegistered)
            {
                var instances = MSBuildLocator.QueryVisualStudioInstances().ToList();
                if (instances.Count > 0)
                {
                    var instance = instances
                        .OrderByDescending(item => item.Version)
                        .First();
                    MSBuildLocator.RegisterInstance(instance);
                }
                else
                {
                    MSBuildLocator.RegisterDefaults();
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
            var arguments = "-latest -products * -requires Microsoft.Component.MSBuild -find \"MSBuild\\**\\Bin\\MSBuild.exe\" -prerelease";
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
        }

        foreach (var root in GetCandidateVisualStudioRoots())
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

        var solutionPath = TryGetValue(() => (string?)project?.DTE?.Solution?.FullName);
        if (!string.IsNullOrWhiteSpace(solutionPath))
        {
            globals["SolutionPath"] = solutionPath;
            globals["SolutionName"] = Path.GetFileNameWithoutExtension(solutionPath);
            globals["SolutionFileName"] = Path.GetFileName(solutionPath);

            var solutionDir = Path.GetDirectoryName(solutionPath);
            if (!string.IsNullOrWhiteSpace(solutionDir))
            {
                if (!solutionDir.EndsWith(Path.DirectorySeparatorChar))
                {
                    solutionDir += Path.DirectorySeparatorChar;
                }

                globals["SolutionDir"] = solutionDir;
            }
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

    private static string? BuildConfigurationCondition(IDictionary<string, string> globalProperties)
    {
        globalProperties.TryGetValue("Configuration", out var configuration);
        globalProperties.TryGetValue("Platform", out var platform);

        if (!string.IsNullOrWhiteSpace(configuration) && !string.IsNullOrWhiteSpace(platform))
        {
            return $"'$(Configuration)|$(Platform)'=='{configuration}|{platform}'";
        }

        if (!string.IsNullOrWhiteSpace(configuration))
        {
            return $"'$(Configuration)'=='{configuration}'";
        }

        if (!string.IsNullOrWhiteSpace(platform))
        {
            return $"'$(Platform)'=='{platform}'";
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

        var existingIndex = page.Items.FindIndex(item => string.Equals(item.Name, entry.Name, StringComparison.OrdinalIgnoreCase));
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

    private static string NormalizePageName(string pageName)
    {
        if (string.IsNullOrWhiteSpace(pageName))
        {
            return "Application/General";
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

    private static string MapPropertyPageName(string category, string propertyName)
    {
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
            var value = GetPropertyValue(entry);
            return FormatPropertyValue(value);
        }
        catch (COMException ex)
        {
            return $"<unavailable: {ex.Message}>";
        }
        catch (Exception ex)
        {
            return $"<unavailable: {ex.Message}>";
        }
    }

    private static object? GetPropertyValue(PropertyEntry entry)
    {
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
        MsBuild = 0,
        DteDescriptor = 1,
        ComProperty = 2
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
        public PropertyEntry(string name, dynamic comProperty, string category)
        {
            Name = name;
            ComProperty = comProperty;
            Category = category;
            Source = PropertySource.ComProperty;
        }

        public PropertyEntry(string name, object owner, PropertyDescriptor descriptor, string category)
        {
            Name = name;
            Owner = owner;
            Descriptor = descriptor;
            Category = category;
            PropertyType = descriptor.PropertyType;
            Source = PropertySource.DteDescriptor;
        }

        public PropertyEntry(string name, string category, Func<object?> getter, Action<object?>? setter, Type? propertyType, PropertySource source)
        {
            Name = name;
            Category = category;
            Getter = getter;
            Setter = setter;
            PropertyType = propertyType;
            Source = source;
        }

        public string Name { get; }
        public string Category { get; }
        public object? Owner { get; }
        public PropertyDescriptor? Descriptor { get; }
        public dynamic? ComProperty { get; }
        public Func<object?>? Getter { get; }
        public Action<object?>? Setter { get; }
        public Type? PropertyType { get; }
        public PropertySource Source { get; }
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
