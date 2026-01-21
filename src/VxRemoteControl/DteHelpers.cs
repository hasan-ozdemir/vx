using System.Collections.Generic;
using EnvDTE;
using EnvDTE80;

namespace VxRemoteControl
{
    internal static class DteHelpers
    {
        public static IEnumerable<Project> EnumerateProjects(Solution solution)
        {
            if (solution == null)
            {
                yield break;
            }

            foreach (Project project in solution.Projects)
            {
                foreach (var nested in EnumerateProjectNode(project))
                {
                    yield return nested;
                }
            }
        }

        private static IEnumerable<Project> EnumerateProjectNode(Project project)
        {
            if (project == null)
            {
                yield break;
            }

            if (project.Kind == ProjectKinds.vsProjectKindSolutionFolder)
            {
                if (project.ProjectItems == null)
                {
                    yield break;
                }

                foreach (ProjectItem item in project.ProjectItems)
                {
                    var subProject = item.SubProject;
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

        public static string GetDocumentText(Document document)
        {
            if (document == null)
            {
                return null;
            }

            if (document.Object("TextDocument") is TextDocument textDoc)
            {
                var start = textDoc.StartPoint.CreateEditPoint();
                return start.GetText(textDoc.EndPoint);
            }

            return null;
        }

        public static bool SetDocumentText(Document document, string text)
        {
            if (document == null)
            {
                return false;
            }

            if (document.Object("TextDocument") is TextDocument textDoc)
            {
                var start = textDoc.StartPoint.CreateEditPoint();
                start.Delete(textDoc.EndPoint);
                start.Insert(text ?? string.Empty);
                return true;
            }

            return false;
        }
    }
}
