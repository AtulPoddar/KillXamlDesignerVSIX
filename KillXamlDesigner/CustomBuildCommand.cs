//------------------------------------------------------------------------------
// <copyright file="CustomBuildCommand.cs" company="Microsoft">
//     Copyright (c) Microsoft.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System.Diagnostics;

namespace KillXamlDesigner
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class CustomBuildCommand
    {
        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("f9eeb7aa-52cf-43bf-8e3d-ddaaa948961a");

        #region Command IDs

        /// <summary>
        /// Build Solution Command ID.
        /// </summary>
        public const int BuildSlnCommandId = 0x0100;

        /// <summary>
        /// ReBuild Solution Command ID.
        /// </summary>
        public const int ReBuildSlnCommandId = 0x0200;

        /// <summary>
        /// Clean Solution Command ID.
        /// </summary>
        public const int CleanSlnCommandId = 0x0300;

        #endregion

        #region Private Members

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        private OleMenuCommandService commandService = null;
        private MenuCommand buildSlnMenuItem = null;
        private MenuCommand rebuildSlnMenuItem = null;
        private MenuCommand cleanSlnMenuItem = null;

        #endregion

        /// <summary>
        /// Initializes a new instance of the <see cref="CustomBuildCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private CustomBuildCommand(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            commandService = ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var buildSlnMenuCommandID = new CommandID(CommandSet, BuildSlnCommandId);
                buildSlnMenuItem = new MenuCommand(BuildSlnMenuItemCallback, buildSlnMenuCommandID);
                commandService.AddCommand(buildSlnMenuItem);

                var rebuildSlnMenuCommandID = new CommandID(CommandSet, ReBuildSlnCommandId);
                rebuildSlnMenuItem = new MenuCommand(ReBuildSlnMenuItemCallback, rebuildSlnMenuCommandID);
                commandService.AddCommand(rebuildSlnMenuItem);

                var cleanSlnMenuCommandID = new CommandID(CommandSet, CleanSlnCommandId);
                cleanSlnMenuItem = new MenuCommand(CleanSlnMenuItemCallback, cleanSlnMenuCommandID);
                commandService.AddCommand(cleanSlnMenuItem);

                var dteObj = ServiceProvider.GetService(typeof(SDTE)) as EnvDTE80.DTE2;
                dteObj.Events.BuildEvents.OnBuildBegin += BuildEvents_OnBuildBegin;
                dteObj.Events.BuildEvents.OnBuildDone += BuildEvents_OnBuildDone;
                dteObj.Events.SolutionEvents.BeforeClosing += SolutionEvents_BeforeClosing;
                dteObj.Events.SolutionEvents.Opened += SolutionEvents_Opened;
            }
        }

        private void SolutionEvents_Opened()
        {
            AddCommand(buildSlnMenuItem);
            AddCommand(rebuildSlnMenuItem);
            AddCommand(cleanSlnMenuItem);
        }

        private void SolutionEvents_BeforeClosing()
        {
            RemoveCommand(buildSlnMenuItem);
            RemoveCommand(rebuildSlnMenuItem);
            RemoveCommand(cleanSlnMenuItem);
        }

        private void BuildEvents_OnBuildBegin(EnvDTE.vsBuildScope Scope, EnvDTE.vsBuildAction Action)
        {
            DisableCommand(buildSlnMenuItem);
            DisableCommand(rebuildSlnMenuItem);
            DisableCommand(cleanSlnMenuItem);
        }

        private void BuildEvents_OnBuildDone(EnvDTE.vsBuildScope Scope, EnvDTE.vsBuildAction Action)
        {
            EnableCommand(buildSlnMenuItem);
            EnableCommand(rebuildSlnMenuItem);
            EnableCommand(cleanSlnMenuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static CustomBuildCommand Instance
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
            Instance = new CustomBuildCommand(package);
        }

        private void BuildSlnMenuItemCallback(object sender, EventArgs e)
        {
            KillXamlDesignerProcesses();

            var dteObj = ServiceProvider.GetService(typeof(SDTE)) as EnvDTE80.DTE2;
            dteObj.ExecuteCommand("Build.BuildSolution");
        }

        private void ReBuildSlnMenuItemCallback(object sender, EventArgs e)
        {
            KillXamlDesignerProcesses();

            var dteObj = ServiceProvider.GetService(typeof(SDTE)) as EnvDTE80.DTE2;
            dteObj.ExecuteCommand("Build.ReBuildSolution");
        }

        private void CleanSlnMenuItemCallback(object sender, EventArgs e)
        {
            KillXamlDesignerProcesses();

            var dteObj = ServiceProvider.GetService(typeof(SDTE)) as EnvDTE80.DTE2;
            dteObj.ExecuteCommand("Build.CleanSolution");
        }

        private void AddCommand(MenuCommand commandItem)
        {
            if (commandService != null && commandItem != null && commandService.FindCommand(commandItem.CommandID) == null)
            {
                commandService.AddCommand(commandItem);
            }
        }

        private void RemoveCommand(MenuCommand commandItem)
        {
            if (commandService != null && commandItem != null && commandService.FindCommand(commandItem.CommandID) != null)
            {
                commandService.RemoveCommand(commandItem);
            }
        }

        private void EnableCommand(MenuCommand commandItem)
        {
            if (commandItem != null)
            {
                commandItem.Enabled = true;
            }
        }

        private void DisableCommand(MenuCommand commandItem)
        {
            if (commandItem != null)
            {
                commandItem.Enabled = false;
            }
        }

        private void KillXamlDesignerProcesses()
        {
            try
            {
                foreach (var process in Process.GetProcessesByName("XDesProc"))
                {
                    process.Kill();
                }
            }
            catch (Exception)
            { }
        }
    }
}
