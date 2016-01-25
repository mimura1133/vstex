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
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using VsTeXProject.VisualStudio.Project.Automation;
using Constants = EnvDTE.Constants;
using MSBuild = Microsoft.Build.Evaluation;

namespace VsTeXProject.VisualStudio.Project
{
    [CLSCompliant(false), ComVisible(true)]
    public abstract class ProjectContainerNode : ProjectNode,
        IVsParentProject,
        IBuildDependencyOnProjectContainer
    {
        #region ctors

        #endregion

        #region properties

        /// <summary>
        ///     Returns teh object that handles listening to file changes on the nested project files.
        /// </summary>
        internal FileChangeManager NestedProjectNodeReloader
        {
            get
            {
                if (nestedProjectNodeReloader == null)
                {
                    nestedProjectNodeReloader = new FileChangeManager(Site);
                    nestedProjectNodeReloader.FileChangedOnDisk += OnNestedProjectFileChangedOnDisk;
                }

                return nestedProjectNodeReloader;
            }
        }

        #endregion

        #region overridden properties

        /// <summary>
        ///     This is the object that will be returned by EnvDTE.Project.Object for this project
        /// </summary>
        internal override object Object
        {
            get { return new OASolutionFolder<ProjectContainerNode>(this); }
        }

        #endregion

        #region fields

        /// <summary>
        ///     Setting this flag to true will build all nested project when building this project
        /// </summary>
        private bool buildNestedProjectsOnBuild = true;

        private ProjectElement nestedProjectElement;

        /// <summary>
        ///     Defines the listener that would listen on file changes on the nested project node.
        /// </summary>
        /// <devremark>
        ///     This might need a refactoring when nested projects can be added and removed by demand.
        /// </devremark>
        private FileChangeManager nestedProjectNodeReloader;

        #endregion

        #region public overridden methods

        /// <summary>
        ///     Gets the nested hierarchy.
        /// </summary>
        /// <param name="itemId">The item id.</param>
        /// <param name="iidHierarchyNested">Identifier of the interface to be returned in ppHierarchyNested.</param>
        /// <param name="ppHierarchyNested">Pointer to the interface whose identifier was passed in iidHierarchyNested.</param>
        /// <param name="pItemId">Pointer to an item identifier of the root node of the nested hierarchy.</param>
        /// <returns>
        ///     If the method succeeds, it returns S_OK. If it fails, it returns an error code. If ITEMID is not a nested
        ///     hierarchy, this method returns E_FAIL.
        /// </returns>
        [CLSCompliant(false)]
        public override int GetNestedHierarchy(uint itemId, ref Guid iidHierarchyNested, out IntPtr ppHierarchyNested,
            out uint pItemId)
        {
            pItemId = VSConstants.VSITEMID_ROOT;
            ppHierarchyNested = IntPtr.Zero;
            if (FirstChild != null)
            {
                for (var n = FirstChild; n != null; n = n.NextSibling)
                {
                    var p = n as NestedProjectNode;

                    if (p != null && p.ID == itemId)
                    {
                        if (p.NestedHierarchy != null)
                        {
                            var iunknownPtr = IntPtr.Zero;
                            var returnValue = VSConstants.S_OK;
                            try
                            {
                                iunknownPtr = Marshal.GetIUnknownForObject(p.NestedHierarchy);
                                Marshal.QueryInterface(iunknownPtr, ref iidHierarchyNested, out ppHierarchyNested);
                            }
                            catch (COMException e)
                            {
                                returnValue = e.ErrorCode;
                            }
                            finally
                            {
                                if (iunknownPtr != IntPtr.Zero)
                                {
                                    Marshal.Release(iunknownPtr);
                                }
                            }

                            return returnValue;
                        }
                        break;
                    }
                }
            }

            return VSConstants.E_FAIL;
        }

        public override int IsItemDirty(uint itemId, IntPtr punkDocData, out int pfDirty)
        {
            var hierNode = NodeFromItemId(itemId);
            Debug.Assert(hierNode != null, "Hierarchy node not found");
            if (hierNode != this)
            {
                return ErrorHandler.ThrowOnFailure(hierNode.IsItemDirty(itemId, punkDocData, out pfDirty));
            }
            return ErrorHandler.ThrowOnFailure(base.IsItemDirty(itemId, punkDocData, out pfDirty));
        }

        public override int SaveItem(VSSAVEFLAGS dwSave, string silentSaveAsName, uint itemid, IntPtr punkDocData,
            out int pfCancelled)
        {
            var hierNode = NodeFromItemId(itemid);
            Debug.Assert(hierNode != null, "Hierarchy node not found");
            if (hierNode != this)
            {
                return
                    ErrorHandler.ThrowOnFailure(hierNode.SaveItem(dwSave, silentSaveAsName, itemid, punkDocData,
                        out pfCancelled));
            }
            return
                ErrorHandler.ThrowOnFailure(base.SaveItem(dwSave, silentSaveAsName, itemid, punkDocData, out pfCancelled));
        }

        protected override bool FilterItemTypeToBeAddedToHierarchy(string itemType)
        {
            if (string.Compare(itemType, ProjectFileConstants.SubProject, StringComparison.OrdinalIgnoreCase) == 0)
            {
                return true;
            }
            return base.FilterItemTypeToBeAddedToHierarchy(itemType);
        }

        /// <summary>
        ///     Called to reload a project item.
        ///     Reloads a project and its nested project nodes.
        /// </summary>
        /// <param name="itemId">Specifies itemid from VSITEMID.</param>
        /// <param name="reserved">Reserved.</param>
        /// <returns>If the method succeeds, it returns S_OK. If it fails, it returns an error code. </returns>
        public override int ReloadItem(uint itemId, uint reserved)
        {
            #region precondition

            if (IsClosed)
            {
                return VSConstants.E_FAIL;
            }

            #endregion

            var node = NodeFromItemId(itemId) as NestedProjectNode;

            if (node != null)
            {
                var propertyAsObject = node.GetProperty((int) __VSHPROPID.VSHPROPID_HandlesOwnReload);

                if (propertyAsObject != null && (bool) propertyAsObject)
                {
                    node.ReloadItem(reserved);
                }
                else
                {
                    ReloadNestedProjectNode(node);
                }

                return VSConstants.S_OK;
            }

            return base.ReloadItem(itemId, reserved);
        }

        /// <summary>
        ///     Reloads a project and its nested project nodes.
        /// </summary>
        protected override void Reload()
        {
            base.Reload();
            CreateNestedProjectNodes();
        }

        #endregion

        #region IVsParentProject

        public virtual int OpenChildren()
        {
            var solution = GetService(typeof (IVsSolution)) as IVsSolution;

            Debug.Assert(solution != null, "Could not retrieve the solution from the services provided by this project");
            if (solution == null)
            {
                return VSConstants.E_FAIL;
            }

            var iUnKnownForSolution = IntPtr.Zero;
            var returnValue = VSConstants.S_OK; // be optimistic.

            try
            {
                DisableQueryEdit = true;
                EventTriggeringFlag = EventTriggering.DoNotTriggerHierarchyEvents |
                                      EventTriggering.DoNotTriggerTrackerEvents;
                iUnKnownForSolution = Marshal.GetIUnknownForObject(solution);

                // notify SolutionEvents listeners that we are about to add children
                var fireSolutionEvents =
                    Marshal.GetTypedObjectForIUnknown(iUnKnownForSolution, typeof (IVsFireSolutionEvents)) as
                        IVsFireSolutionEvents;
                ErrorHandler.ThrowOnFailure(fireSolutionEvents.FireOnBeforeOpeningChildren(this));

                AddVirtualProjects();

                ErrorHandler.ThrowOnFailure(fireSolutionEvents.FireOnAfterOpeningChildren(this));
            }
            catch (Exception e)
            {
                // Exceptions are digested by the caller but we want then to be shown if not a ComException and if not in automation.
                if (!(e is COMException) && !Utilities.IsInAutomationFunction(Site))
                {
                    string title = null;
                    var icon = OLEMSGICON.OLEMSGICON_CRITICAL;
                    var buttons = OLEMSGBUTTON.OLEMSGBUTTON_OK;
                    var defaultButton = OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST;
                    VsShellUtilities.ShowMessageBox(Site, title, e.Message, icon, buttons, defaultButton);
                }

                Trace.WriteLine("Exception : " + e.Message);
                throw;
            }
            finally
            {
                DisableQueryEdit = false;

                if (iUnKnownForSolution != IntPtr.Zero)
                {
                    Marshal.Release(iUnKnownForSolution);
                }

                EventTriggeringFlag = EventTriggering.TriggerAll;
            }

            return returnValue;
        }

        public virtual int CloseChildren()
        {
            var returnValue = VSConstants.S_OK; // be optimistic.

            var solution = GetService(typeof (IVsSolution)) as IVsSolution;
            Debug.Assert(solution != null, "Could not retrieve the solution from the services provided by this project");

            if (solution == null)
            {
                return VSConstants.E_FAIL;
            }

            var iUnKnownForSolution = IntPtr.Zero;

            try
            {
                iUnKnownForSolution = Marshal.GetIUnknownForObject(solution);

                // notify SolutionEvents listeners that we are about to close children
                var fireSolutionEvents =
                    Marshal.GetTypedObjectForIUnknown(iUnKnownForSolution, typeof (IVsFireSolutionEvents)) as
                        IVsFireSolutionEvents;
                ErrorHandler.ThrowOnFailure(fireSolutionEvents.FireOnBeforeClosingChildren(this));

                // If the removal crashes we never fire the close children event. IS that a problem?
                RemoveNestedProjectNodes();

                ErrorHandler.ThrowOnFailure(fireSolutionEvents.FireOnAfterClosingChildren(this));
            }
            finally
            {
                if (iUnKnownForSolution != IntPtr.Zero)
                {
                    Marshal.Release(iUnKnownForSolution);
                }
            }

            return returnValue;
        }

        #endregion

        #region IBuildDependencyOnProjectContainerNode

        /// <summary>
        ///     Defines whether nested projects should be build with the parent project
        /// </summary>
        public virtual bool BuildNestedProjectsOnBuild
        {
            get { return buildNestedProjectsOnBuild; }
            set { buildNestedProjectsOnBuild = value; }
        }

        /// <summary>
        ///     Enumerates the nested hierachies that should be added to the build dependency list.
        /// </summary>
        /// <returns></returns>
        public virtual IVsHierarchy[] EnumNestedHierachiesForBuildDependency()
        {
            var nestedProjectList = new List<IVsHierarchy>();
            // Add all nested project among projectContainer child nodes
            for (var child = FirstChild; child != null; child = child.NextSibling)
            {
                var nestedProjectNode = child as NestedProjectNode;
                if (nestedProjectNode != null)
                {
                    nestedProjectList.Add(nestedProjectNode.NestedHierarchy);
                }
            }

            return nestedProjectList.ToArray();
        }

        #endregion

        #region helper methods

        protected internal void RemoveNestedProjectNodes()
        {
            for (var n = FirstChild; n != null; n = n.NextSibling)
            {
                var p = n as NestedProjectNode;
                if (p != null)
                {
                    p.CloseNestedProjectNode();
                }
            }

            // We do not care of file changes after this.
            NestedProjectNodeReloader.FileChangedOnDisk -= OnNestedProjectFileChangedOnDisk;
            NestedProjectNodeReloader.Dispose();
        }

        /// <summary>
        ///     This is used when loading the project to loop through all the items
        ///     and for each SubProject it finds, it create the project and a node
        ///     in our Hierarchy to hold the project.
        /// </summary>
        protected internal void CreateNestedProjectNodes()
        {
            // 1. Create a ProjectElement with the found item and then Instantiate a new Nested project with this ProjectElement.
            // 2. Link into the hierarchy.			
            // Read subprojects from from msbuildmodel
            var creationFlags = __VSCREATEPROJFLAGS.CPF_NOTINSLNEXPLR | __VSCREATEPROJFLAGS.CPF_SILENT;

            if (IsNewProject)
            {
                creationFlags |= __VSCREATEPROJFLAGS.CPF_CLONEFILE;
            }
            else
            {
                creationFlags |= __VSCREATEPROJFLAGS.CPF_OPENFILE;
            }

            foreach (var item in BuildProject.Items)
            {
                if (
                    string.Compare(item.ItemType, ProjectFileConstants.SubProject, StringComparison.OrdinalIgnoreCase) ==
                    0)
                {
                    nestedProjectElement = new ProjectElement(this, item, false);

                    if (!IsNewProject)
                    {
                        AddExistingNestedProject(null, creationFlags);
                    }
                    else
                    {
                        // If we are creating the subproject from a vstemplate/vsz file
                        var isVsTemplate = Utilities.IsTemplateFile(GetProjectTemplatePath(null));
                        if (isVsTemplate)
                        {
                            RunVsTemplateWizard(null, true);
                        }
                        else
                        {
                            // We are cloning the specified project file
                            AddNestedProjectFromTemplate(null, creationFlags);
                        }
                    }
                }
            }

            nestedProjectElement = null;
        }

        /// <summary>
        ///     Add an existing project as a nested node of our hierarchy.
        ///     This is used while loading the project and can also be used
        ///     to add an existing project to our hierarchy.
        /// </summary>
        protected internal virtual NestedProjectNode AddExistingNestedProject(ProjectElement element,
            __VSCREATEPROJFLAGS creationFlags)
        {
            var elementToUse = element == null ? nestedProjectElement : element;

            if (elementToUse == null)
            {
                throw new ArgumentNullException("element");
            }

            var filename = elementToUse.GetFullPathForElement();
            // Delegate to AddNestedProjectFromTemplate. Because we pass flags that specify open project rather then clone, this will works.
            Debug.Assert((creationFlags & __VSCREATEPROJFLAGS.CPF_OPENFILE) == __VSCREATEPROJFLAGS.CPF_OPENFILE,
                "__VSCREATEPROJFLAGS.CPF_OPENFILE should have been specified, did you mean to call AddNestedProjectFromTemplate?");
            return AddNestedProjectFromTemplate(filename, Path.GetDirectoryName(filename), Path.GetFileName(filename),
                elementToUse, creationFlags);
        }

        /// <summary>
        ///     Let the wizard code execute and provide us the information we need.
        ///     Our SolutionFolder automation object should be called back with the
        ///     details at which point it will call back one of our method to add
        ///     a nested project.
        ///     If you are trying to add a new subproject this is probably the
        ///     method you want to call. If you are just trying to clone a template
        ///     project file, then AddNestedProjectFromTemplate is what you want.
        /// </summary>
        /// <param name="element">The project item to use as the base of the nested project.</param>
        /// <param name="silent">true if the wizard should run silently, otherwise false.</param>
        [SuppressMessage("Microsoft.Naming", "CA1709:IdentifiersShouldBeCasedCorrectly", MessageId = "Vs")]
        protected internal void RunVsTemplateWizard(ProjectElement element, bool silent)
        {
            var elementToUse = element == null ? nestedProjectElement : element;

            if (elementToUse == null)
            {
                throw new ArgumentNullException("element");
            }
            nestedProjectElement = elementToUse;

            var oaProject = GetAutomationObject() as OAProject;
            if (oaProject == null || oaProject.ProjectItems == null)
                throw new InvalidOperationException(SR.GetString(SR.InvalidAutomationObject,
                    CultureInfo.CurrentUICulture));
            Debug.Assert(oaProject.Object != null,
                "The project automation object should have set the Object to the SolutionFolder");
            var folder = oaProject.Object as OASolutionFolder<ProjectContainerNode>;

            // Prepare the parameters to pass to RunWizardFile
            var destination = elementToUse.GetFullPathForElement();
            var template = GetProjectTemplatePath(elementToUse);

            var wizParams = new object[7];
            wizParams[0] = Constants.vsWizardAddSubProject;
            wizParams[1] = Path.GetFileNameWithoutExtension(destination);
            wizParams[2] = oaProject.ProjectItems;
            wizParams[3] = Path.GetDirectoryName(destination);
            wizParams[4] = Path.GetFileNameWithoutExtension(destination);
            wizParams[5] = Path.GetDirectoryName(folder.DTE.FullName); //VS install dir
            wizParams[6] = silent;

            var wizardTrust = GetService(typeof (SVsDetermineWizardTrust)) as IVsDetermineWizardTrust;
            if (wizardTrust != null)
            {
                var guidProjectAdding = Guid.Empty;

                // In case of a project template an empty guid should be added as the guid parameter. See env\msenv\core\newtree.h IsTrustedTemplate method definition.
                ErrorHandler.ThrowOnFailure(wizardTrust.OnWizardInitiated(template, ref guidProjectAdding));
            }

            try
            {
                // Make the call to execute the wizard. This should cause AddNestedProjectFromTemplate to be
                // called back with the correct set of parameters.
                var extensibilityService = (IVsExtensibility) GetService(typeof (IVsExtensibility));
                var result = extensibilityService.RunWizardFile(template, 0, ref wizParams);
                if (result == wizardResult.wizardResultFailure)
                    throw new COMException();
            }
            finally
            {
                if (wizardTrust != null)
                {
                    ErrorHandler.ThrowOnFailure(wizardTrust.OnWizardCompleted());
                }
            }
        }

        /// <summary>
        ///     This will clone a template project file and add it as a
        ///     subproject to our hierarchy.
        ///     If you want to create a project for which there exist a
        ///     vstemplate, consider using RunVsTemplateWizard instead.
        /// </summary>
        protected internal virtual NestedProjectNode AddNestedProjectFromTemplate(ProjectElement element,
            __VSCREATEPROJFLAGS creationFlags)
        {
            var elementToUse = element == null ? nestedProjectElement : element;

            if (elementToUse == null)
            {
                throw new ArgumentNullException("element");
            }
            var destination = elementToUse.GetFullPathForElement();
            var template = GetProjectTemplatePath(elementToUse);

            return AddNestedProjectFromTemplate(template, Path.GetDirectoryName(destination),
                Path.GetFileName(destination), elementToUse, creationFlags);
        }

        /// <summary>
        ///     This can be called directly or through RunVsTemplateWizard.
        ///     This will clone a template project file and add it as a
        ///     subproject to our hierarchy.
        ///     If you want to create a project for which there exist a
        ///     vstemplate, consider using RunVsTemplateWizard instead.
        /// </summary>
        protected internal virtual NestedProjectNode AddNestedProjectFromTemplate(string fileName, string destination,
            string projectName, ProjectElement element, __VSCREATEPROJFLAGS creationFlags)
        {
            // If this is project creation and the template specified a subproject in its project file, this.nestedProjectElement will be used 
            var elementToUse = element == null ? nestedProjectElement : element;

            if (elementToUse == null)
            {
                // If this is null, this means MSBuild does not know anything about our subproject so add an MSBuild item for it
                elementToUse = new ProjectElement(this, fileName, ProjectFileConstants.SubProject);
            }

            var node = CreateNestedProjectNode(elementToUse);
            node.Init(fileName, destination, projectName, creationFlags);

            // In case that with did not have an existing element, or the nestedProjectelement was null 
            //  and since Init computes the final path, set the MSBuild item to that path
            if (nestedProjectElement == null)
            {
                var relativePath = node.Url;
                if (Path.IsPathRooted(relativePath))
                {
                    relativePath = ProjectFolder;
                    if (!relativePath.EndsWith("/\\", StringComparison.Ordinal))
                    {
                        relativePath += Path.DirectorySeparatorChar;
                    }

                    relativePath = new Url(relativePath).MakeRelative(new Url(node.Url));
                }

                elementToUse.Rename(relativePath);
            }

            AddChild(node);
            return node;
        }

        /// <summary>
        ///     Override this method if you want to provide your own type of nodes.
        ///     This would be the case if you derive a class from NestedProjectNode
        /// </summary>
        protected virtual NestedProjectNode CreateNestedProjectNode(ProjectElement element)
        {
            return new NestedProjectNode(this, element);
        }

        /// <summary>
        ///     Links the nested project nodes to the solution. The default implementation parses all nested project nodes and
        ///     calles AddVirtualProjectEx on them.
        /// </summary>
        protected virtual void AddVirtualProjects()
        {
            for (var child = FirstChild; child != null; child = child.NextSibling)
            {
                var nestedProjectNode = child as NestedProjectNode;
                if (nestedProjectNode != null)
                {
                    nestedProjectNode.AddVirtualProject();
                }
            }
        }

        /// <summary>
        ///     Based on the Template and TypeGuid properties of the
        ///     element, generate the full template path.
        ///     TypeGuid should be the Guid of a registered project factory.
        ///     Template can be a full path, a project template (for projects
        ///     that support VsTemplates) or a relative path (for other projects).
        /// </summary>
        protected virtual string GetProjectTemplatePath(ProjectElement element)
        {
            var elementToUse = element == null ? nestedProjectElement : element;

            if (elementToUse == null)
            {
                throw new ArgumentNullException("element");
            }

            var templateFile = elementToUse.GetMetadata(ProjectFileConstants.Template);
            Debug.Assert(!string.IsNullOrEmpty(templateFile),
                "No template file has been specified in the template attribute in the project file");

            var fullPath = templateFile;
            if (!Path.IsPathRooted(templateFile))
            {
                var registeredProjectType = GetRegisteredProject(elementToUse);

                // This is not a full path
                Debug.Assert(
                    registeredProjectType != null &&
                    (!string.IsNullOrEmpty(registeredProjectType.DefaultProjectExtensionValue) ||
                     !string.IsNullOrEmpty(registeredProjectType.WizardTemplatesDirValue)),
                    " Registered wizard directory value not set in the registry.");

                // See if this specify a VsTemplate file
                fullPath = registeredProjectType.GetVsTemplateFile(templateFile);
                if (string.IsNullOrEmpty(fullPath))
                {
                    // Default to using the WizardTemplateDir to calculate the absolute path
                    fullPath = Path.Combine(registeredProjectType.WizardTemplatesDirValue, templateFile);
                }
            }

            return fullPath;
        }

        /// <summary>
        ///     Get information from the registry based for the project
        ///     factory corresponding to the TypeGuid of the element
        /// </summary>
        private RegisteredProjectType GetRegisteredProject(ProjectElement element)
        {
            var elementToUse = element == null ? nestedProjectElement : element;

            if (elementToUse == null)
            {
                throw new ArgumentNullException("element");
            }

            // Get the project type guid from project elementToUse				
            var typeGuidString = elementToUse.GetMetadataAndThrow(ProjectFileConstants.TypeGuid, new Exception());
            var projectFactoryGuid = new Guid(typeGuidString);

            var dte = ProjectMgr.Site.GetService(typeof (DTE)) as DTE;
            Debug.Assert(dte != null, "Could not get the automation object from the services exposed by this project");

            if (dte == null)
                throw new InvalidOperationException();

            var registeredProjectType = RegisteredProjectType.CreateRegisteredProjectType(projectFactoryGuid);
            Debug.Assert(registeredProjectType != null,
                "Could not read the registry setting associated to this project.");
            if (registeredProjectType == null)
            {
                throw new InvalidOperationException();
            }
            return registeredProjectType;
        }

        /// <summary>
        ///     Reloads a nested project node by deleting it and readding it.
        /// </summary>
        /// <param name="node">The node to reload.</param>
        protected virtual void ReloadNestedProjectNode(NestedProjectNode node)
        {
            if (node == null)
            {
                throw new ArgumentNullException("node");
            }

            var solution = GetService(typeof (IVsSolution)) as IVsSolution;

            if (solution == null)
            {
                throw new InvalidOperationException();
            }

            NestedProjectNode newNode = null;
            try
            {
                // (VS 2005 UPDATE) When deleting and re-adding the nested project,
                // we do not want SCC to see this as a delete and add operation. 
                EventTriggeringFlag = EventTriggering.DoNotTriggerTrackerEvents;

                // notify SolutionEvents listeners that we are about to add children
                var fireSolutionEvents = solution as IVsFireSolutionEvents;

                if (fireSolutionEvents == null)
                {
                    throw new InvalidOperationException();
                }

                ErrorHandler.ThrowOnFailure(fireSolutionEvents.FireOnBeforeUnloadProject(node.NestedHierarchy));

                var isDirtyAsInt = 0;
                IsDirty(out isDirtyAsInt);

                var isDirty = isDirtyAsInt == 0 ? false : true;

                var element = node.ItemNode;
                node.CloseNestedProjectNode();

                // Remove from the solution
                RemoveChild(node);

                // Now readd it                
                try
                {
                    var flags = __VSCREATEPROJFLAGS.CPF_NOTINSLNEXPLR | __VSCREATEPROJFLAGS.CPF_SILENT |
                                __VSCREATEPROJFLAGS.CPF_OPENFILE;
                    newNode = AddExistingNestedProject(element, flags);
                    newNode.AddVirtualProject();
                }
                catch (Exception e)
                {
                    // We get a System.Exception if anything failed, thus we have no choice but catch it. 
                    // Exceptions are digested by VS. Show the error if not in automation.
                    if (!Utilities.IsInAutomationFunction(Site))
                    {
                        var message = string.IsNullOrEmpty(e.Message)
                            ? SR.GetString(SR.NestedProjectFailedToReload, CultureInfo.CurrentUICulture)
                            : e.Message;
                        var title = string.Empty;
                        var icon = OLEMSGICON.OLEMSGICON_CRITICAL;
                        var buttons = OLEMSGBUTTON.OLEMSGBUTTON_OK;
                        var defaultButton = OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST;
                        VsShellUtilities.ShowMessageBox(Site, title, message, icon, buttons, defaultButton);
                    }

                    // Do not digest exception. let the caller handle it. If in a later stage this exception is not digested then the above messagebox is not needed.
                    throw;
                }

#if DEBUG
                IVsHierarchy nestedHierarchy;
                ErrorHandler.ThrowOnFailure(solution.GetProjectOfUniqueName(newNode.GetMkDocument(), out nestedHierarchy));
                Debug.Assert(
                    nestedHierarchy != null && Utilities.IsSameComObject(nestedHierarchy, newNode.NestedHierarchy),
                    "The nested hierrachy was not reloaded correctly.");
#endif
                SetProjectFileDirty(isDirty);

                ErrorHandler.ThrowOnFailure(fireSolutionEvents.FireOnAfterLoadProject(newNode.NestedHierarchy));
            }
            finally
            {
                // In this scenario the nested project failed to unload or reload the nested project. We will unload the whole project, otherwise the nested project is lost.
                // This is similar to the scenario when one wants to open a project and the nested project cannot be loaded because for example the project file has xml errors.
                // We should note that we rely here that if the unload fails then exceptions are not digested and are shown to the user.
                if (newNode == null || newNode.NestedHierarchy == null)
                {
                    ErrorHandler.ThrowOnFailure(
                        solution.CloseSolutionElement(
                            (uint) __VSSLNCLOSEOPTIONS.SLNCLOSEOPT_UnloadProject |
                            (uint) __VSSLNSAVEOPTIONS.SLNSAVEOPT_ForceSave, InteropSafeIVsHierarchy, 0));
                }
                else
                {
                    EventTriggeringFlag = EventTriggering.TriggerAll;
                }
            }
        }

        /// <summary>
        ///     Event callback. Called when one of the nested project files is changed.
        /// </summary>
        /// <param name="sender">The FileChangeManager object.</param>
        /// <param name="e">Event args containing the file name that was updated.</param>
        private void OnNestedProjectFileChangedOnDisk(object sender, FileChangedOnDiskEventArgs e)
        {
            #region Pre-condition validation

            Debug.Assert(e != null, "No event args specified for the FileChangedOnDisk event");

            // We care only about time change for reload.
            if ((e.FileChangeFlag & _VSFILECHANGEFLAGS.VSFILECHG_Time) == 0)
            {
                return;
            }

            // test if we actually have a document for this id.
            string moniker;
            GetMkDocument(e.ItemID, out moniker);
            Debug.Assert(NativeMethods.IsSamePath(moniker, e.FileName),
                " The file + " + e.FileName +
                " has changed but we could not retrieve the path for the item id associated to the path.");

            #endregion

            var reload = true;
            if (!Utilities.IsInAutomationFunction(Site))
            {
                // Prompt to reload the nested project file. We use the moniker here since the filename from the event arg is canonicalized.
                var message = string.Format(CultureInfo.CurrentCulture,
                    SR.GetString(SR.QueryReloadNestedProject, CultureInfo.CurrentUICulture), moniker);
                var title = string.Empty;
                var icon = OLEMSGICON.OLEMSGICON_INFO;
                var buttons = OLEMSGBUTTON.OLEMSGBUTTON_YESNO;
                var defaultButton = OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST;
                reload = VsShellUtilities.ShowMessageBox(Site, message, title, icon, buttons, defaultButton) ==
                         NativeMethods.IDYES;
            }

            if (reload)
            {
                // We have to use here the interface method call, since it might be that specialized project nodes like the project container item
                // is owerwriting the default functionality.
                ReloadItem(e.ItemID, 0);
            }
        }

        #endregion
    }
}