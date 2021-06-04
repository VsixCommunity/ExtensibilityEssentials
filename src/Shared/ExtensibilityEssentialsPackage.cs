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
        name: "autoload",
        expression: "vsix & building",
        termNames: new[] { "vsix", "building" },
        termValues: new[] { "SolutionHasProjectFlavor:" + ProjectTypes.EXTENSIBILITY, VSConstants.UICONTEXT.SolutionBuilding_string })]
    public sealed partial class ExtensibilityEssentialsPackage : ToolkitPackage
    {
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            VS.Events!.SolutionEvents!.Opened += OnSolutionOpened;
            VS.Events!.SolutionEvents!.ProjectAdded += ApplyBuildPropertyToProject;

            IVsSolution _solution = await VS.Solution.GetSolutionAsync();

            if (_solution.IsOpen())
            {
                OnSolutionOpened();
            }
        }

        private void OnSolutionOpened()
        {
            JoinableTaskFactory.RunAsync(async delegate
            {
                await JoinableTaskFactory.SwitchToMainThreadAsync(DisposalToken);
                IEnumerable<Project> projects = await VS.Solution.GetAllProjectsInSolutionAsync();

                foreach (Project project in projects)
                {
                    ApplyBuildPropertyToProject(project);
                }
            });
        }

        private void ApplyBuildPropertyToProject(Project project)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (project.IsKind(ProjectTypes.EXTENSIBILITY))
            {
                project.TrySetBuildPropertyAsync("ExtensibilityEssentialsInstalled", "true").FireAndForget();
            }
        }
    }
}
