using System;
using System.Collections.Generic;
using System.Threading;
using Community.VisualStudio.Toolkit;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace ExtensibilityEssentials
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration(Vsix.Name, Vsix.Description, Vsix.Version)]
    [ProvideAutoLoad("07ce51b0-5439-4c81-bbc6-dc629f343357", flags: PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideUIContextRule("07ce51b0-5439-4c81-bbc6-dc629f343357",
        name: "VSIX project",
        expression: "vsix",
        termNames: new[] { "vsix" },
        termValues: new[] { "SolutionHasProjectFlavor:{82b43b9b-a64c-4715-b499-d71e9ca2bd60}" })]
    public sealed partial class ExtensibilityEssentialsPackage : ToolkitPackage
    {
        private static readonly Guid _vsixProjectSubType = new("82b43b9b-a64c-4715-b499-d71e9ca2bd60");

#nullable disable
        private IVsSolution _solution;
#nullable restore

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            _solution = await VS.Solution.GetSolutionAsync();

            if (_solution is not null)
            {
                // Listen for solutions being opened and new projects being added to existing
                // solutions so that we can add a build property to each VSIX project.
                SolutionEvents events = VS.Events.SolutionEvents;

                if (events is not null)
                {
                    events.Opened += OnSolutionOpened;
                    events.ProjectAdded += OnProjectAdded;
                }

                // If a solution is already open, then act as though it has just been opened.
                // TODO: In newer versions of the toolkit, call _solution.IsOpen() instead;
                if (ErrorHandler.Succeeded(_solution.GetProperty((int)__VSPROPID.VSPROPID_IsSolutionOpen, out var value)))
                {
                    if (value is bool isOpen && isOpen)
                    {
                        OnSolutionOpened();
                    }
                }
            }
        }

        private void OnSolutionOpened()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            IEnumerable<IVsHierarchy> hierarchies = _solution.GetAllProjectHierarchys();

            foreach (IVsHierarchy hierarchy in hierarchies)
            {
                if (IsVsixProject(hierarchy))
                {
                    ApplyBuildPropertyToProject(hierarchy);
                }
            }
        }

        private void OnProjectAdded(Project project)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (ErrorHandler.Succeeded(_solution.GetProjectOfUniqueName(project.UniqueName, out IVsHierarchy hierarchy)))
            {
                if (IsVsixProject(hierarchy))
                {
                    ApplyBuildPropertyToProject(hierarchy);
                }
            }
        }

        // TODO: Replace with project.IsKind(VSIXGuid) in newer versions of the Tookit
        private bool IsVsixProject(IVsHierarchy hierarchy)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (hierarchy is IVsAggregatableProject aggregatable)
            {
                if (ErrorHandler.Succeeded(aggregatable.GetAggregateProjectTypeGuids(out var types)))
                {
                    foreach (var type in types.Split(';'))
                    {
                        if (Guid.TryParse(type, out Guid identifier) && _vsixProjectSubType.Equals(identifier))
                        {
                            return true;
                        }
                    }
                }
            }

            return false;
        }

        // TODO: Use Project.TrySetBuildProperty in newer versions of the Toolkit
        private void ApplyBuildPropertyToProject(IVsHierarchy hierarchy)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (hierarchy is IVsBuildPropertyStorage storage)
            {
                // Store the build property in the user file instead of the project
                // file, because we don't want to affect the real project file.
                storage.SetPropertyValue("ExtensibilityEssentialsInstalled", "", (uint)_PersistStorageType.PST_USER_FILE, "true");
            }
        }
    }
}
