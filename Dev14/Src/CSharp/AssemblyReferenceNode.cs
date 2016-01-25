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
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using VsTeXProject.VisualStudio.Project.Automation;
using MSBuild = Microsoft.Build.Evaluation;
using MSBuildExecution = Microsoft.Build.Execution;

namespace VsTeXProject.VisualStudio.Project
{
    [CLSCompliant(false)]
    [ComVisible(true)]
    public class AssemblyReferenceNode : ReferenceNode
    {
        #region fieds

        /// <summary>
        ///     The name of the assembly this refernce represents
        /// </summary>
        private AssemblyName assemblyName;

        private string assemblyPath = string.Empty;

        /// <summary>
        ///     Defines the listener that would listen on file changes on the nested project node.
        /// </summary>
        private FileChangeManager fileChangeListener;

        /// <summary>
        ///     A flag for specifying if the object was disposed.
        /// </summary>
        private bool isDisposed;

        #endregion

        #region properties

        /// <summary>
        ///     The name of the assembly this reference represents.
        /// </summary>
        /// <value></value>
        internal AssemblyName AssemblyName
        {
            get { return assemblyName; }
        }

        /// <summary>
        ///     Returns the name of the assembly this reference refers to on this specific
        ///     machine. It can be different from the AssemblyName property because it can
        ///     be more specific.
        /// </summary>
        internal AssemblyName ResolvedAssembly { get; private set; }

        public override string Url
        {
            get { return assemblyPath; }
        }

        public override string Caption
        {
            get { return assemblyName.Name; }
        }

        private OAAssemblyReference assemblyRef;

        internal override object Object
        {
            get
            {
                if (null == assemblyRef)
                {
                    assemblyRef = new OAAssemblyReference(this);
                }
                return assemblyRef;
            }
        }

        #endregion

        #region ctors

        /// <summary>
        ///     Constructor for the ReferenceNode
        /// </summary>
        public AssemblyReferenceNode(ProjectNode root, ProjectElement element)
            : base(root, element)
        {
            GetPathNameFromProjectFile();

            InitializeFileChangeEvents();

            var include = ItemNode.GetMetadata(ProjectFileConstants.Include);

            CreateFromAssemblyName(new AssemblyName(include));
        }

        /// <summary>
        ///     Constructor for the AssemblyReferenceNode
        /// </summary>
        public AssemblyReferenceNode(ProjectNode root, string assemblyPath)
            : base(root)
        {
            // Validate the input parameters.
            if (null == root)
            {
                throw new ArgumentNullException("root");
            }
            if (string.IsNullOrEmpty(assemblyPath))
            {
                throw new ArgumentNullException("assemblyPath");
            }

            InitializeFileChangeEvents();

            // The assemblyPath variable can be an actual path on disk or a generic assembly name.
            if (File.Exists(assemblyPath))
            {
                // The assemblyPath parameter is an actual file on disk; try to load it.
                assemblyName = AssemblyName.GetAssemblyName(assemblyPath);
                this.assemblyPath = assemblyPath;

                // We register with listeningto chnages onteh path here. The rest of teh cases will call into resolving the assembly and registration is done there.
                fileChangeListener.ObserveItem(this.assemblyPath);
            }
            else
            {
                // The file does not exist on disk. This can be because the file / path is not
                // correct or because this is not a path, but an assembly name.
                // Try to resolve the reference as an assembly name.
                CreateFromAssemblyName(new AssemblyName(assemblyPath));
            }
        }

        #endregion

        #region methods

        /// <summary>
        ///     Closes the node.
        /// </summary>
        /// <returns></returns>
        public override int Close()
        {
            try
            {
                Dispose(true);
            }
            finally
            {
                base.Close();
            }

            return VSConstants.S_OK;
        }

        /// <summary>
        ///     Links a reference node to the project and hierarchy.
        /// </summary>
        protected override void BindReferenceData()
        {
            Debug.Assert(assemblyName != null, "The AssemblyName field has not been initialized");

            // If the item has not been set correctly like in case of a new reference added it now.
            // The constructor for the AssemblyReference node will create a default project item. In that case the Item is null.
            // We need to specify here the correct project element. 
            if (ItemNode == null || ItemNode.Item == null)
            {
                ItemNode = new ProjectElement(ProjectMgr, assemblyName.FullName, ProjectFileConstants.Reference);
            }

            // Set the basic information we know about
            ItemNode.SetMetadata(ProjectFileConstants.Name, assemblyName.Name);
            if (!string.IsNullOrEmpty(assemblyPath))
            {
                ItemNode.SetMetadata(ProjectFileConstants.AssemblyName, Path.GetFileName(assemblyPath));
            }
            else
            {
                ItemNode.SetMetadata(ProjectFileConstants.AssemblyName, null);
            }

            SetReferenceProperties();
        }

        /// <summary>
        ///     Disposes the node
        /// </summary>
        /// <param name="disposing"></param>
        protected override void Dispose(bool disposing)
        {
            if (isDisposed)
            {
                return;
            }

            try
            {
                UnregisterFromFileChangeService();
            }
            finally
            {
                base.Dispose(disposing);
                isDisposed = true;
            }
        }

        private void CreateFromAssemblyName(AssemblyName name)
        {
            assemblyName = name;

            // Use MsBuild to resolve the assemblyname 
            ResolveAssemblyReference();

            if (string.IsNullOrEmpty(assemblyPath) && (null != ItemNode.Item))
            {
                // Try to get the assembly name from the hintpath.
                GetPathNameFromProjectFile();
                if (assemblyPath == null)
                {
                    // Try to get the assembly name from the path
                    assemblyName = AssemblyName.GetAssemblyName(assemblyPath);
                }
            }
            if (null == ResolvedAssembly)
            {
                ResolvedAssembly = assemblyName;
            }
        }

        /// <summary>
        ///     Checks if an assembly is already added. The method parses all references and compares the full assemblynames, or
        ///     the location of the assemblies to decide whether two assemblies are the same.
        /// </summary>
        /// <returns>true if the assembly has already been added.</returns>
        protected internal override bool IsAlreadyAdded(out ReferenceNode existingReference)
        {
            var referencesFolder =
                ProjectMgr.FindChild(ReferenceContainerNode.ReferencesNodeVirtualName) as ReferenceContainerNode;
            Debug.Assert(referencesFolder != null, "Could not find the References node");
            var shouldCheckPath = !string.IsNullOrEmpty(Url);

            for (var n = referencesFolder.FirstChild; n != null; n = n.NextSibling)
            {
                var assemblyReferenceNode = n as AssemblyReferenceNode;
                if (null != assemblyReferenceNode)
                {
                    // We will check if the full assemblynames are the same or if the Url of the assemblies is the same.
                    if (
                        string.Compare(assemblyReferenceNode.AssemblyName.FullName, assemblyName.FullName,
                            StringComparison.OrdinalIgnoreCase) == 0 ||
                        (shouldCheckPath && NativeMethods.IsSamePath(assemblyReferenceNode.Url, Url)))
                    {
                        existingReference = assemblyReferenceNode;
                        return true;
                    }
                }
            }

            existingReference = null;
            return false;
        }

        /// <summary>
        ///     Determines if this is node a valid node for painting the default reference icon.
        /// </summary>
        /// <returns></returns>
        protected override bool CanShowDefaultIcon()
        {
            if (string.IsNullOrEmpty(assemblyPath) || !File.Exists(assemblyPath))
            {
                return false;
            }

            return true;
        }

        private void GetPathNameFromProjectFile()
        {
            var result = ItemNode.GetMetadata(ProjectFileConstants.HintPath);
            if (string.IsNullOrEmpty(result))
            {
                result = ItemNode.GetMetadata(ProjectFileConstants.AssemblyName);
                if (string.IsNullOrEmpty(result))
                {
                    assemblyPath = string.Empty;
                }
                else if (!result.EndsWith(".dll", StringComparison.OrdinalIgnoreCase))
                {
                    result += ".dll";
                    assemblyPath = result;
                }
            }
            else
            {
                assemblyPath = GetFullPathFromPath(result);
            }
        }

        private string GetFullPathFromPath(string path)
        {
            if (Path.IsPathRooted(path))
            {
                return path;
            }
            var uri = new Uri(ProjectMgr.BaseURI.Uri, path);

            if (uri != null)
            {
                return Microsoft.VisualStudio.Shell.Url.Unescape(uri.LocalPath, true);
            }

            return string.Empty;
        }

        protected override void ResolveReference()
        {
            ResolveAssemblyReference();
        }

        private void SetHintPathAndPrivateValue()
        {
            // Remove the HintPath, we will re-add it below if it is needed
            if (!string.IsNullOrEmpty(assemblyPath))
            {
                ItemNode.SetMetadata(ProjectFileConstants.HintPath, null);
            }

            // Get the list of items which require HintPath
            IEnumerable<MSBuild.ProjectItem> references =
                ProjectMgr.BuildProject.GetItems(MsBuildGeneratedItemType.ReferenceCopyLocalPaths);

            // Now loop through the generated References to find the corresponding one
            foreach (var reference in references)
            {
                var fileName = Path.GetFileNameWithoutExtension(reference.EvaluatedInclude);
                if (string.Compare(fileName, assemblyName.Name, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    // We found it, now set some properties based on this.
                    var hintPath = reference.GetMetadataValue(ProjectFileConstants.HintPath);
                    SetHintPathAndPrivateValue(hintPath);
                    break;
                }
            }
        }

        /// <summary>
        ///     Sets the hint path to the provided value.
        ///     It also sets the private value to true if it has not been already provided through the associated project element.
        /// </summary>
        /// <param name="hintPath">The hint path to set.</param>
        private void SetHintPathAndPrivateValue(string hintPath)
        {
            if (string.IsNullOrEmpty(hintPath))
            {
                return;
            }

            if (Path.IsPathRooted(hintPath))
            {
                hintPath = PackageUtilities.GetPathDistance(ProjectMgr.BaseURI.Uri, new Uri(hintPath));
            }

            ItemNode.SetMetadata(ProjectFileConstants.HintPath, hintPath);

            // Private means local copy; we want to know if it is already set to not override the default
            var privateValue = ItemNode != null ? ItemNode.GetMetadata(ProjectFileConstants.Private) : null;

            // If this is not already set, we default to true
            if (string.IsNullOrEmpty(privateValue))
            {
                ItemNode.SetMetadata(ProjectFileConstants.Private, true.ToString());
            }
        }

        /// <summary>
        ///     This function ensures that some properties of the reference are set.
        /// </summary>
        private void SetReferenceProperties()
        {
            // If there is an assembly path then just set the hint path
            if (!string.IsNullOrEmpty(assemblyPath))
            {
                SetHintPathAndPrivateValue(assemblyPath);
                return;
            }

            // Set a default HintPath for msbuild to be able to resolve the reference.
            ItemNode.SetMetadata(ProjectFileConstants.HintPath, assemblyPath);

            // Resolve assembly referernces. This is needed to make sure that properties like the full path
            // to the assembly or the hint path are set.
            if (ProjectMgr.Build(MsBuildTarget.ResolveAssemblyReferences) != MSBuildResult.Successful)
            {
                return;
            }

            // Check if we have to resolve again the path to the assembly.			
            ResolveReference();

            // Make sure that the hint path if set (if needed).
            SetHintPathAndPrivateValue();
        }

        /// <summary>
        ///     Does the actual job of resolving an assembly reference. We need a private method that does not violate
        ///     calling virtual method from the constructor.
        /// </summary>
        private void ResolveAssemblyReference()
        {
            if (ProjectMgr == null || ProjectMgr.IsClosed)
            {
                return;
            }

            var group = ProjectMgr.CurrentConfig.GetItems(ProjectFileConstants.ReferencePath);
            foreach (var item in group)
            {
                var fullPath = GetFullPathFromPath(item.EvaluatedInclude);

                var name = AssemblyName.GetAssemblyName(fullPath);

                // Try with full assembly name and then with weak assembly name.
                if (string.Equals(name.FullName, assemblyName.FullName, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name.Name, assemblyName.Name, StringComparison.OrdinalIgnoreCase))
                {
                    if (!NativeMethods.IsSamePath(fullPath, assemblyPath))
                    {
                        // set the full path now.
                        assemblyPath = fullPath;

                        // We have a new item to listen too, since the assembly reference is resolved from a different place.
                        fileChangeListener.ObserveItem(assemblyPath);
                    }

                    ResolvedAssembly = name;

                    // No hint path is needed since the assembly path will always be resolved.
                    return;
                }
            }
        }

        /// <summary>
        ///     Registers with File change events
        /// </summary>
        private void InitializeFileChangeEvents()
        {
            fileChangeListener = new FileChangeManager(ProjectMgr.Site);
            fileChangeListener.FileChangedOnDisk += OnAssemblyReferenceChangedOnDisk;
        }

        /// <summary>
        ///     Unregisters this node from file change notifications.
        /// </summary>
        private void UnregisterFromFileChangeService()
        {
            fileChangeListener.FileChangedOnDisk -= OnAssemblyReferenceChangedOnDisk;
            fileChangeListener.Dispose();
        }

        /// <summary>
        ///     Event callback. Called when one of the assembly file is changed.
        /// </summary>
        /// <param name="sender">The FileChangeManager object.</param>
        /// <param name="e">Event args containing the file name that was updated.</param>
        private void OnAssemblyReferenceChangedOnDisk(object sender, FileChangedOnDiskEventArgs e)
        {
            Debug.Assert(e != null, "No event args specified for the FileChangedOnDisk event");

            // We only care about file deletes, so check for one before enumerating references.			
            if ((e.FileChangeFlag & _VSFILECHANGEFLAGS.VSFILECHG_Del) == 0)
            {
                return;
            }


            if (NativeMethods.IsSamePath(e.FileName, assemblyPath))
            {
                OnInvalidateItems(Parent);
            }
        }

        #endregion
    }
}