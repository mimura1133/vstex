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
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;

namespace VsTeXProject.VisualStudio.Project.Automation
{
    [ComVisible(true), CLSCompliant(false)]
    public class OASolutionFolder<T> : SolutionFolder
        where T : HierarchyNode
    {
        private readonly T node;

        public OASolutionFolder(T associatedNode)
        {
            if (associatedNode == null)
            {
                throw new ArgumentNullException("associatedNode");
            }

            Debug.Assert(associatedNode.ProjectMgr is ProjectContainerNode,
                "Expecting obejct of type" + typeof (ProjectContainerNode).Name);

            if (!(associatedNode.ProjectMgr is ProjectContainerNode))
                throw new ArgumentException(SR.GetString(SR.InvalidParameter, CultureInfo.CurrentUICulture),
                    "associatedNode");

            node = associatedNode;
        }

        #region SolutionFolder Members

        public virtual EnvDTE.Project AddFromFile(string fileName)
        {
            var projectContainer = (ProjectContainerNode) node.ProjectMgr;
            var newElement = new ProjectElement(projectContainer, fileName, ProjectFileConstants.SubProject);
            var newNode = projectContainer.AddExistingNestedProject(newElement,
                __VSCREATEPROJFLAGS.CPF_NOTINSLNEXPLR | __VSCREATEPROJFLAGS.CPF_SILENT |
                __VSCREATEPROJFLAGS.CPF_OPENFILE);
            if (newNode == null)
                return null;
            // Now that the sub project was created, get its extensibility object so we can return it
            object newProject = null;
            if (
                ErrorHandler.Succeeded(newNode.NestedHierarchy.GetProperty(VSConstants.VSITEMID_ROOT,
                    (int) __VSHPROPID.VSHPROPID_ExtObject, out newProject)))
                return newProject as EnvDTE.Project;
            return null;
        }

        public virtual EnvDTE.Project AddFromTemplate(string fileName, string destination, string projectName)
        {
            var isVSTemplate = Utilities.IsTemplateFile(fileName);

            NestedProjectNode newNode = null;
            if (isVSTemplate)
            {
                // Get the wizard to run, we will get called again and use the alternate code path
                var newElement = new ProjectElement(node.ProjectMgr, Path.Combine(destination, projectName),
                    ProjectFileConstants.SubProject);
                newElement.SetMetadata(ProjectFileConstants.Template, fileName);
                ((ProjectContainerNode) node.ProjectMgr).RunVsTemplateWizard(newElement, false);
            }
            else
            {
                if (string.IsNullOrEmpty(Path.GetExtension(projectName)))
                {
                    var targetExtension = Path.GetExtension(fileName);
                    projectName = Path.ChangeExtension(projectName, targetExtension);
                }

                var projectContainer = (ProjectContainerNode) node.ProjectMgr;
                newNode = projectContainer.AddNestedProjectFromTemplate(fileName, destination, projectName, null,
                    __VSCREATEPROJFLAGS.CPF_NOTINSLNEXPLR | __VSCREATEPROJFLAGS.CPF_SILENT |
                    __VSCREATEPROJFLAGS.CPF_CLONEFILE);
            }
            if (newNode == null)
                return null;

            // Now that the sub project was created, get its extensibility object so we can return it
            object newProject = null;
            if (
                ErrorHandler.Succeeded(newNode.NestedHierarchy.GetProperty(VSConstants.VSITEMID_ROOT,
                    (int) __VSHPROPID.VSHPROPID_ExtObject, out newProject)))
                return newProject as EnvDTE.Project;
            return null;
        }

        public virtual EnvDTE.Project AddSolutionFolder(string Name)
        {
            throw new NotImplementedException();
        }

        public virtual EnvDTE.Project Parent
        {
            get { throw new NotImplementedException(); }
        }

        public virtual bool Hidden { get; set; } = false;

        public virtual DTE DTE
        {
            get { return (DTE) node.ProjectMgr.Site.GetService(typeof (DTE)); }
        }

        #endregion
    }
}