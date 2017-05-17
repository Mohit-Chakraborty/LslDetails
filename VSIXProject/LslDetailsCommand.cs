//------------------------------------------------------------------------------
// <copyright file="LslDetailsCommand.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

namespace LslDetails
{
    using System;
    using System.ComponentModel.Design;
    using Microsoft.VisualStudio;
    using Microsoft.VisualStudio.Shell;
    using Microsoft.VisualStudio.Shell.Interop;

    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class LslDetailsCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("5357888f-abfa-4963-9d64-ca8080211fef");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="LslDetailsCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private LslDetailsCommand(Package package)
        {
            this.package = package ?? throw new ArgumentNullException("package");

            if (this.ServiceProvider.GetService(typeof(IMenuCommandService)) is OleMenuCommandService commandService)
            {
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);

                //menuItem.Visible

                commandService.AddCommand(menuItem);
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static LslDetailsCommand Instance
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
            Instance = new LslDetailsCommand(package);
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
            //string message = string.Format(CultureInfo.CurrentCulture, "A solution is {0} open", IsSolutionOpen() ? string.Empty : "not");
            //string title = "Command1";

            if (this.IsSolutionLoadDeferred())
            {
                Guid g = Guid.Empty;
                var solution = this.ServiceProvider.GetService(typeof(IVsSolution)) as IVsSolution;
                solution.GetProjectEnum((uint)__VSENUMPROJFLAGS.EPF_ALLPROJECTS, ref g, out IEnumHierarchies enumHierarchies);
            }

            // Show a message box to prove we were here
            //VsShellUtilities.ShowMessageBox(
            //    this.ServiceProvider,
            //    message,
            //    title,
            //    OLEMSGICON.OLEMSGICON_INFO,
            //    OLEMSGBUTTON.OLEMSGBUTTON_OK,
            //    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }

        private bool IsSolutionOpen()
        {
            var solution = this.ServiceProvider.GetService(typeof(IVsSolution)) as IVsSolution;

            int hr = solution.GetProperty((int)__VSPROPID.VSPROPID_IsSolutionOpen, out object pVar);
            if (ErrorHandler.Succeeded(hr))
            {
                return (bool)pVar;
            }

            return false;
        }

        private bool IsSolutionLoadDeferred()
        {
            var solution = this.ServiceProvider.GetService(typeof(IVsSolution)) as IVsSolution7;

            return solution != null ? solution.IsSolutionLoadDeferred() : false;
        }
    }
}
