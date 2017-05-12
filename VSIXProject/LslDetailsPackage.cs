//------------------------------------------------------------------------------
// <copyright file="LslDetailsPackage.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

namespace LslDetails
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics.CodeAnalysis;
    using System.Runtime.InteropServices;
    using Microsoft.VisualStudio;
    using Microsoft.VisualStudio.Shell;
    using Microsoft.VisualStudio.Shell.Interop;

    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [Guid(LslDetailsPackage.PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [ProvideAutoLoad(UIContextGuids.SolutionExists)]
    public sealed class LslDetailsPackage : Package
    {
        /// <summary>
        /// LslDetailsPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "d9b593e4-9503-44e8-bfa7-894e13d365b7";

        private const string DeferredProjectCaptionSuffix = "*";

        private const string LslwindowPaneCaption = "Lightweight solution load";
        private IVsOutputWindowPane windowPane;

        /// <summary>
        /// Initializes a new instance of the <see cref="LslDetailsPackage"/> class.
        /// </summary>
        public LslDetailsPackage()
        {
            // Inside this method you can place any initialization code that does not require
            // any Visual Studio service because at this point the package object is created but
            // not sited yet inside Visual Studio environment. The place to do all the other
            // initialization is the Initialize method.
        }

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            base.Initialize();

            KnownUIContexts.SolutionExistsAndFullyLoadedContext.UIContextChanged += this.SolutionExistsAndFullyLoadedContext_UIContextChanged;
        }

        private void SolutionExistsAndFullyLoadedContext_UIContextChanged(object sender, UIContextChangedEventArgs e)
        {
            // KnownUIContexts.SolutionExistsAndFullyLoadedContext.IsActive == true
            if (e.Activated)
            {
                this.UpdateCaptionOfDeferredProjects();
                this.EnumerateProjectsInSolution();
            }
        }

        private void UpdateCaptionOfDeferredProjects()
        {
            var solution = this.GetService(typeof(IVsSolution)) as IVsSolution;

            Guid g = Guid.Empty;
            int hr = solution.GetProjectEnum((uint)__VSENUMPROJFLAGS3.EPF_DEFERRED, ref g, out IEnumHierarchies enumHierarchies);

            if (ErrorHandler.Failed(hr))
            {
                return;
            }

            IVsHierarchy[] hierachies = new IVsHierarchy[1];
            hr = enumHierarchies.Next(1, hierachies, out uint fetchedCount);

            while (ErrorHandler.Succeeded(hr) && (fetchedCount == 1))
            {
                IVsHierarchy deferredProject = hierachies[0];

                hr = deferredProject.GetProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID.VSHPROPID_Caption, out object caption);

                if (ErrorHandler.Succeeded(hr))
                {
                    deferredProject.SetProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID.VSHPROPID_Caption, caption + DeferredProjectCaptionSuffix);
                }

                hr = enumHierarchies.Next(1, hierachies, out fetchedCount);
            }
        }

        private void EnumerateProjectsInSolution()
        {
            var solution = this.GetService(typeof(IVsSolution)) as IVsSolution;

            Guid g = Guid.Empty;
            int hr = solution.GetProjectEnum((uint)__VSENUMPROJFLAGS.EPF_LOADEDINSOLUTION, ref g, out IEnumHierarchies enumHierarchies);

            if (ErrorHandler.Failed(hr))
            {
                return;
            }

            var deferredProjects = new List<IVsHierarchy>();
            var nonDeferredProjects = new List<IVsHierarchy>();

            IVsHierarchy[] hierachies = new IVsHierarchy[1];
            hr = enumHierarchies.Next(1, hierachies, out uint fetchedCount);

            while (ErrorHandler.Succeeded(hr) && (fetchedCount == 1))
            {
                hr = hierachies[0].GetProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID9.VSHPROPID_IsDeferred, out object isDeferred);

                if (ErrorHandler.Succeeded(hr) && (bool)isDeferred)
                {
                    deferredProjects.Add(hierachies[0]);
                }
                else
                {
                    hr = hierachies[0].GetGuidProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID.VSHPROPID_TypeGuid, out Guid type);
                    if (ErrorHandler.Succeeded(hr) && (type != VSConstants.ItemTypeGuid.VirtualFolder_guid))
                    {
                        nonDeferredProjects.Add(hierachies[0]);
                    }
                }

                hr = enumHierarchies.Next(1, hierachies, out fetchedCount);
            }

            solution.GetProperty((int)__VSPROPID.VSPROPID_SolutionFileName, out object solutionFilePath);

            this.WriteToOutputPane("Solution '" + solutionFilePath + "' has " + (deferredProjects.Count + nonDeferredProjects.Count) + " projects.");

            this.WriteToOutputPane("* Deferred projects - " + deferredProjects.Count);
            foreach (var hierarchy in deferredProjects)
            {
                hr = hierarchy.GetCanonicalName((uint)VSConstants.VSITEMID.Root, out string name);

                if (ErrorHandler.Succeeded(hr))
                {
                    this.WriteToOutputPane(name);
                }
            }

            this.WriteToOutputPane("* Non deferred projects - " + nonDeferredProjects.Count);
            foreach (var hierarchy in nonDeferredProjects)
            {
                hr = hierarchy.GetCanonicalName((uint)VSConstants.VSITEMID.Root, out string name);

                if (ErrorHandler.Succeeded(hr))
                {
                    this.WriteToOutputPane(name);
                }
            }
        }

        private void WriteToOutputPane(string data)
        {
            if (this.windowPane == null)
            {
                this.windowPane = this.GetOutputPane(VSConstants.OutputWindowPaneGuid.GeneralPane_guid, LslwindowPaneCaption);
            }

            this.windowPane?.Activate();
            this.windowPane?.OutputString(data + System.Environment.NewLine);
        }
    }
}
