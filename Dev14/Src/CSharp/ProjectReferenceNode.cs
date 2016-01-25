/********************************************************************************************

Copyright (c) Microsoft Corporation 
All rights reserved. 

Microsoft Public License: 

This license governs use of the accompanying software. If you use the software, you 
accept this license. If you do not accept the license, do not use the software. 

1. Definitions 
The terms "reproduce," "reproduction," "derivative works," and "distribution" have the 
same meaning here as under U.S. copyright law. 
A "contribution" is the original software, or any additions or changes to the software. 
A "contributor" is any person that distributes its contribution under this license. 
"Licensed patents" are a contributor's patent claims that read directly on its contribution. 

2. Grant of Rights 
(A) Copyright Grant- Subject to the terms of this license, including the license conditions 
and limitations in section 3, each contributor grants you a non-exclusive, worldwide, 
royalty-free copyright license to reproduce its contribution, prepare derivative works of 
its contribution, and distribute its contribution or any derivative works that you create. 
(B) Patent Grant- Subject to the terms of this license, including the license conditions 
and limitations in section 3, each contributor grants you a non-exclusive, worldwide, 
royalty-free license under its licensed patents to make, have made, use, sell, offer for 
sale, import, and/or otherwise dispose of its contribution in the software or derivative 
works of the contribution in the software. 

3. Conditions and Limitations 
(A) No Trademark License- This license does not grant you rights to use any contributors' 
name, logo, or trademarks. 
(B) If you bring a patent claim against any contributor over patents that you claim are 
infringed by the software, your patent license from such contributor to the software ends 
automatically. 
(C) If you distribute any portion of the software, you must retain all copyright, patent, 
trademark, and attribution notices that are present in the software. 
(D) If you distribute any portion of the software in source code form, you may do so only 
under this license by including a complete copy of this license with your distribution. 
If you distribute any portion of the software in compiled or object code form, you may only 
do so under a license that complies with this license. 
(E) The software is licensed "as-is." You bear the risk of using it. The contributors give 
no express warranties, guarantees or conditions. You may have additional consumer rights 
under your local laws which this license cannot change. To the extent permitted under your 
local laws, the contributors exclude the implied warranties of merchantability, fitness for 
a particular purpose and non-infringement.

********************************************************************************************/

using System;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using VsTeXProject.VisualStudio.Project.Automation;
using Constants = EnvDTE.Constants;

namespace VsTeXProject.VisualStudio.Project
{
    [CLSCompliant(false), ComVisible(true)]
    public class ProjectReferenceNode : ReferenceNode
    {
        #region fieds

        /// <summary>
        ///     The name of the assembly this refernce represents
        /// </summary>
        private Guid referencedProjectGuid;

        private readonly string referencedProjectRelativePath = string.Empty;

        private readonly string referencedProjectFullPath = string.Empty;

        private readonly BuildDependency buildDependency;

        /// <summary>
        ///     This is a reference to the automation object for the referenced project.
        /// </summary>
        private EnvDTE.Project referencedProject;

        /// <summary>
        ///     This state is controlled by the solution events.
        ///     The state is set to false by OnBeforeUnloadProject.
        ///     The state is set to true by OnBeforeCloseProject event.
        /// </summary>
        private bool canRemoveReference = true;

        /// <summary>
        ///     Possibility for solution listener to update the state on the dangling reference.
        ///     It will be set in OnBeforeUnloadProject then the nopde is invalidated then it is reset to false.
        /// </summary>
        private bool isNodeValid;

        #endregion

        #region properties

        public override string Url
        {
            get { return referencedProjectFullPath; }
        }

        public override string Caption
        {
            get { return ReferencedProjectName; }
        }

        internal Guid ReferencedProjectGuid
        {
            get { return referencedProjectGuid; }
        }

        /// <summary>
        ///     Possiblity to shortcut and set the dangling project reference icon.
        ///     It is ussually manipulated by solution listsneres who handle reference updates.
        /// </summary>
        protected internal bool IsNodeValid
        {
            get { return isNodeValid; }
            set { isNodeValid = value; }
        }

        /// <summary>
        ///     Controls the state whether this reference can be removed or not. Think of the project unload scenario where the
        ///     project reference should not be deleted.
        /// </summary>
        internal bool CanRemoveReference
        {
            get { return canRemoveReference; }
            set { canRemoveReference = value; }
        }

        internal string ReferencedProjectName { get; } = string.Empty;

        /// <summary>
        ///     Gets the automation object for the referenced project.
        /// </summary>
        internal EnvDTE.Project ReferencedProjectObject
        {
            get
            {
                // If the referenced project is null then re-read.
                if (referencedProject == null)
                {
                    // Search for the project in the collection of the projects in the
                    // current solution.
                    var dte = (DTE) ProjectMgr.GetService(typeof (DTE));
                    if ((null == dte) || (null == dte.Solution))
                    {
                        return null;
                    }
                    foreach (EnvDTE.Project prj in dte.Solution.Projects)
                    {
                        //Skip this project if it is an umodeled project (unloaded)
                        if (
                            string.Compare(Constants.vsProjectKindUnmodeled, prj.Kind,
                                StringComparison.OrdinalIgnoreCase) == 0)
                        {
                            continue;
                        }

                        // Get the full path of the current project.
                        Property pathProperty = null;
                        try
                        {
                            if (prj.Properties == null)
                            {
                                continue;
                            }

                            pathProperty = prj.Properties.Item("FullPath");
                            if (null == pathProperty)
                            {
                                // The full path should alway be availabe, but if this is not the
                                // case then we have to skip it.
                                continue;
                            }
                        }
                        catch (ArgumentException)
                        {
                            continue;
                        }
                        var prjPath = pathProperty.Value.ToString();
                        Property fileNameProperty = null;
                        // Get the name of the project file.
                        try
                        {
                            fileNameProperty = prj.Properties.Item("FileName");
                            if (null == fileNameProperty)
                            {
                                // Again, this should never be the case, but we handle it anyway.
                                continue;
                            }
                        }
                        catch (ArgumentException)
                        {
                            continue;
                        }
                        prjPath = Path.Combine(prjPath, fileNameProperty.Value.ToString());

                        // If the full path of this project is the same as the one of this
                        // reference, then we have found the right project.
                        if (NativeMethods.IsSamePath(prjPath, referencedProjectFullPath))
                        {
                            referencedProject = prj;
                            break;
                        }
                    }
                }

                return referencedProject;
            }
            set { referencedProject = value; }
        }

        /// <summary>
        ///     Gets the full path to the assembly generated by this project.
        /// </summary>
        internal string ReferencedProjectOutputPath
        {
            get
            {
                // Make sure that the referenced project implements the automation object.
                if (null == ReferencedProjectObject)
                {
                    return null;
                }

                // Get the configuration manager from the project.
                var confManager = ReferencedProjectObject.ConfigurationManager;
                if (null == confManager)
                {
                    return null;
                }

                // Get the active configuration.
                var config = confManager.ActiveConfiguration;
                if (null == config)
                {
                    return null;
                }

                // Get the output path for the current configuration.
                var outputPathProperty = config.Properties.Item("OutputPath");
                if (null == outputPathProperty)
                {
                    return null;
                }

                var outputPath = outputPathProperty.Value.ToString();

                // Ususally the output path is relative to the project path, but it is possible
                // to set it as an absolute path. If it is not absolute, then evaluate its value
                // based on the project directory.
                if (!Path.IsPathRooted(outputPath))
                {
                    var projectDir = Path.GetDirectoryName(referencedProjectFullPath);
                    outputPath = Path.Combine(projectDir, outputPath);
                }

                // Now get the name of the assembly from the project.
                // Some project system throw if the property does not exist. We expect an ArgumentException.
                Property assemblyNameProperty = null;
                try
                {
                    assemblyNameProperty = ReferencedProjectObject.Properties.Item("OutputFileName");
                }
                catch (ArgumentException)
                {
                }

                if (null == assemblyNameProperty)
                {
                    return null;
                }
                // build the full path adding the name of the assembly to the output path.
                outputPath = Path.Combine(outputPath, assemblyNameProperty.Value.ToString());

                return outputPath;
            }
        }

        private OAProjectReference projectReference;

        internal override object Object
        {
            get
            {
                if (null == projectReference)
                {
                    projectReference = new OAProjectReference(this);
                }
                return projectReference;
            }
        }

        #endregion

        #region ctors

        /// <summary>
        ///     Constructor for the ReferenceNode. It is called when the project is reloaded, when the project element representing
        ///     the refernce exists.
        /// </summary>
        [SuppressMessage("Microsoft.Usage", "CA2234:PassSystemUriObjectsInsteadOfStrings")]
        public ProjectReferenceNode(ProjectNode root, ProjectElement element)
            : base(root, element)
        {
            referencedProjectRelativePath = ItemNode.GetMetadata(ProjectFileConstants.Include);
            Debug.Assert(!string.IsNullOrEmpty(referencedProjectRelativePath),
                "Could not retrive referenced project path form project file");

            var guidString = ItemNode.GetMetadata(ProjectFileConstants.Project);

            // Continue even if project setttings cannot be read.
            try
            {
                referencedProjectGuid = new Guid(guidString);

                buildDependency = new BuildDependency(ProjectMgr, referencedProjectGuid);
                ProjectMgr.AddBuildDependency(buildDependency);
            }
            finally
            {
                Debug.Assert(referencedProjectGuid != Guid.Empty,
                    "Could not retrive referenced project guidproject file");

                ReferencedProjectName = ItemNode.GetMetadata(ProjectFileConstants.Name);

                Debug.Assert(!string.IsNullOrEmpty(ReferencedProjectName),
                    "Could not retrive referenced project name form project file");
            }

            var uri = new Uri(ProjectMgr.BaseURI.Uri, referencedProjectRelativePath);

            if (uri != null)
            {
                referencedProjectFullPath = Microsoft.VisualStudio.Shell.Url.Unescape(uri.LocalPath, true);
            }
        }

        /// <summary>
        ///     constructor for the ProjectReferenceNode
        /// </summary>
        public ProjectReferenceNode(ProjectNode root, string referencedProjectName, string projectPath,
            string projectReference)
            : base(root)
        {
            Debug.Assert(
                root != null && !string.IsNullOrEmpty(referencedProjectName) && !string.IsNullOrEmpty(projectReference)
                && !string.IsNullOrEmpty(projectPath),
                "Can not add a reference because the input for adding one is invalid.");

            if (projectReference == null)
            {
                throw new ArgumentNullException("projectReference");
            }

            ReferencedProjectName = referencedProjectName;

            var indexOfSeparator = projectReference.IndexOf('|');


            var fileName = string.Empty;

            // Unfortunately we cannot use the path part of the projectReference string since it is not resolving correctly relative pathes.
            if (indexOfSeparator != -1)
            {
                var projectGuid = projectReference.Substring(0, indexOfSeparator);
                referencedProjectGuid = new Guid(projectGuid);
                if (indexOfSeparator + 1 < projectReference.Length)
                {
                    var remaining = projectReference.Substring(indexOfSeparator + 1);
                    indexOfSeparator = remaining.IndexOf('|');

                    if (indexOfSeparator == -1)
                    {
                        fileName = remaining;
                    }
                    else
                    {
                        fileName = remaining.Substring(0, indexOfSeparator);
                    }
                }
            }

            Debug.Assert(!string.IsNullOrEmpty(fileName),
                "Can not add a project reference because the input for adding one is invalid.");

            // Did we get just a file or a relative path?
            var uri = new Uri(projectPath);

            var referenceDir = PackageUtilities.GetPathDistance(ProjectMgr.BaseURI.Uri, uri);

            Debug.Assert(!string.IsNullOrEmpty(referenceDir),
                "Can not add a project reference because the input for adding one is invalid.");

            var justTheFileName = Path.GetFileName(fileName);
            referencedProjectRelativePath = Path.Combine(referenceDir, justTheFileName);

            referencedProjectFullPath = Path.Combine(projectPath, justTheFileName);

            buildDependency = new BuildDependency(ProjectMgr, referencedProjectGuid);
        }

        #endregion

        #region methods

        protected override NodeProperties CreatePropertiesObject()
        {
            return new ProjectReferencesProperties(this);
        }

        /// <summary>
        ///     The node is added to the hierarchy and then updates the build dependency list.
        /// </summary>
        public override void AddReference()
        {
            if (ProjectMgr == null)
            {
                return;
            }
            base.AddReference();
            ProjectMgr.AddBuildDependency(buildDependency);
        }

        /// <summary>
        ///     Overridden method. The method updates the build dependency list before removing the node from the hierarchy.
        /// </summary>
        public override void Remove(bool removeFromStorage)
        {
            if (ProjectMgr == null || !CanRemoveReference)
            {
                return;
            }
            ProjectMgr.RemoveBuildDependency(buildDependency);
            base.Remove(removeFromStorage);
        }

        /// <summary>
        ///     Links a reference node to the project file.
        /// </summary>
        protected override void BindReferenceData()
        {
            Debug.Assert(!string.IsNullOrEmpty(ReferencedProjectName),
                "The referencedProjectName field has not been initialized");
            Debug.Assert(referencedProjectGuid != Guid.Empty, "The referencedProjectName field has not been initialized");

            ItemNode = new ProjectElement(ProjectMgr, referencedProjectRelativePath,
                ProjectFileConstants.ProjectReference);

            ItemNode.SetMetadata(ProjectFileConstants.Name, ReferencedProjectName);
            ItemNode.SetMetadata(ProjectFileConstants.Project, referencedProjectGuid.ToString("B"));
            ItemNode.SetMetadata(ProjectFileConstants.Private, true.ToString());
        }

        /// <summary>
        ///     Defines whether this node is valid node for painting the refererence icon.
        /// </summary>
        /// <returns></returns>
        protected override bool CanShowDefaultIcon()
        {
            if (referencedProjectGuid == Guid.Empty || ProjectMgr == null || ProjectMgr.IsClosed || isNodeValid)
            {
                return false;
            }

            IVsHierarchy hierarchy = null;

            hierarchy = VsShellUtilities.GetHierarchy(ProjectMgr.Site, referencedProjectGuid);

            if (hierarchy == null)
            {
                return false;
            }

            //If the Project is unloaded return false
            if (ReferencedProjectObject == null)
            {
                return false;
            }

            return !string.IsNullOrEmpty(referencedProjectFullPath) && File.Exists(referencedProjectFullPath);
        }

        /// <summary>
        ///     Checks if a project reference can be added to the hierarchy. It calls base to see if the reference is not already
        ///     there, then checks for circular references.
        /// </summary>
        /// <param name="errorHandler">The error handler delegate to return</param>
        /// <returns></returns>
        protected override bool CanAddReference(out CannotAddReferenceErrorMessage errorHandler)
        {
            // When this method is called this refererence has not yet been added to the hierarchy, only instantiated.
            if (!base.CanAddReference(out errorHandler))
            {
                return false;
            }

            errorHandler = null;
            if (IsThisProjectReferenceInCycle())
            {
                errorHandler = ShowCircularReferenceErrorMessage;
                return false;
            }

            return true;
        }

        private bool IsThisProjectReferenceInCycle()
        {
            return IsReferenceInCycle(referencedProjectGuid);
        }

        private void ShowCircularReferenceErrorMessage()
        {
            var message = string.Format(CultureInfo.CurrentCulture,
                SR.GetString(SR.ProjectContainsCircularReferences, CultureInfo.CurrentUICulture), ReferencedProjectName);
            var title = string.Empty;
            var icon = OLEMSGICON.OLEMSGICON_CRITICAL;
            var buttons = OLEMSGBUTTON.OLEMSGBUTTON_OK;
            var defaultButton = OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST;
            VsShellUtilities.ShowMessageBox(ProjectMgr.Site, title, message, icon, buttons, defaultButton);
        }

        /// <summary>
        ///     Checks whether a reference added to a given project would introduce a circular dependency.
        /// </summary>
        private bool IsReferenceInCycle(Guid projectGuid)
        {
            var referencedHierarchy = VsShellUtilities.GetHierarchy(ProjectMgr.Site, projectGuid);

            var solutionBuildManager =
                ProjectMgr.Site.GetService(typeof (SVsSolutionBuildManager)) as IVsSolutionBuildManager2;
            if (solutionBuildManager == null)
            {
                throw new InvalidOperationException("Cannot find the IVsSolutionBuildManager2 service.");
            }

            int circular;
            Marshal.ThrowExceptionForHR(solutionBuildManager.CalculateProjectDependencies());
            Marshal.ThrowExceptionForHR(solutionBuildManager.QueryProjectDependency(referencedHierarchy,
                ProjectMgr.InteropSafeIVsHierarchy, out circular));

            return circular != 0;
        }

        #endregion
    }
}