//------------------------------------------------------------------------------
// <copyright file="Update.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System.Linq;
using System.IO;
using EnvDTE;
using System.Collections;
using System.Threading;
using System.Windows.Forms;

namespace Migrator
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class Update
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("404cc6cd-c069-4edc-ac4c-3aa80af6354d");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="Update"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private Update(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
                commandService.AddCommand(menuItem);
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static Update Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(Package package)
        {
            Instance = new Update(package);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            var _dteObject = (DTE)ServiceProvider.GetService(typeof(DTE));
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
                    if (fileInfo.DirectoryName.EndsWith("Migrations"))
                    {
                        var fileName = Path.GetFileNameWithoutExtension(fileInfo.Name);
                        var projectName = item.ProjectItems.ContainingProject.Name;
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
