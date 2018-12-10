using System;
using System.ComponentModel.Design;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using EnvDTE;
using Microsoft;
using Microsoft.VisualStudio.Shell;
using Task = System.Threading.Tasks.Task;

namespace VSIXMigrator
{
    internal sealed class UpdateCommand
    {
        public const int CommandId = 0x0100;

        public static readonly Guid CommandSet = new Guid("0eadfacd-2154-4b4a-98f4-b7b401549556");

        private readonly AsyncPackage package;

        private UpdateCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        public static UpdateCommand Instance
        {
            get;
            private set;
        }

        private IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync((typeof(IMenuCommandService))) as OleMenuCommandService;
            Instance = new UpdateCommand(package, commandService);
        }

        private async void Execute(object sender, EventArgs e)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
            var _dteObject = (DTE)await ServiceProvider.GetServiceAsync(typeof(DTE));
            Assumes.Present(_dteObject);
            var uiHierarchyObject = (UIHierarchy)_dteObject.Windows.Item("{3AE79031-E1BC-11D0-8F78-00A0C9110057}").Object;
            var selectedItems = uiHierarchyObject.SelectedItems as object[];
            if (selectedItems != null && selectedItems.Any())
            {
                var item = selectedItems.Where(t => (t as UIHierarchyItem)?.Object is ProjectItem)
                                        .Select(t => ((ProjectItem)((UIHierarchyItem)t).Object))
                                        .FirstOrDefault();
                if (item != null)
                {
                    var fileInfo = new FileInfo(item.FileNames[1]);
                    var fileName = Path.GetFileNameWithoutExtension(fileInfo.Name);
                    var projectName = item.ProjectItems.ContainingProject.Name;

                    if (fileInfo.DirectoryName.EndsWith("Migrations") && 
                        !string.IsNullOrEmpty(fileName) &&
                        !string.IsNullOrEmpty(projectName))
                    {
                        var script =
                            "Update-Database -TargetMigration " + fileName +
                            " -ProjectName " + projectName +
                            " -StartUpProjectName " + projectName;
                        _dteObject.ExecuteCommand("View.PackageManagerConsole");
                        Clipboard.SetText(script);
                    }
                    else
                    {
                        _dteObject.ExecuteCommand("View.PackageManagerConsole");
                        Clipboard.SetText("Update-Database");
                    }
                }
            }
        }
    }
}