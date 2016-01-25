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
using System.Diagnostics.CodeAnalysis;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace VsTeXProject.VisualStudio.Project
{
    /// <summary>
    ///     Gets registry settings from for a project.
    /// </summary>
    internal class RegisteredProjectType
    {
        internal const string DefaultProjectExtension = "DefaultProjectExtension";
        internal const string WizardsTemplatesDir = "WizardsTemplatesDir";
        internal const string ProjectTemplatesDir = "ProjectTemplatesDir";
        internal const string Package = "Package";


        internal string DefaultProjectExtensionValue { get; set; }

        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        internal string ProjectTemplatesDirValue { get; set; }

        internal string WizardTemplatesDirValue { get; set; }

        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        internal Guid PackageGuidValue { get; set; }

        /// <summary>
        ///     If the project support VsTemplates, returns the path to
        ///     the vstemplate file corresponding to the requested template
        ///     You can pass in a string such as: "Windows\Console Application"
        /// </summary>
        internal string GetVsTemplateFile(string templateFile)
        {
            // First see if this use the vstemplate model
            if (!string.IsNullOrEmpty(DefaultProjectExtensionValue))
            {
                var dte = Microsoft.VisualStudio.Shell.Package.GetGlobalService(typeof (DTE)) as DTE2;
                if (dte != null)
                {
                    var solution = dte.Solution as Solution2;
                    if (solution != null)
                    {
                        var fullPath = solution.GetProjectTemplate(templateFile, DefaultProjectExtensionValue);
                        // The path returned by GetProjectTemplate can be in the format "path|FrameworkVersion=x.y|Language=xxx"
                        // where the framework version and language sections are optional.
                        // Here we are interested only in the full path, so we have to remove all the other sections.
                        var pipePos = fullPath.IndexOf('|');
                        if (0 == pipePos)
                        {
                            return null;
                        }
                        if (pipePos > 0)
                        {
                            fullPath = fullPath.Substring(0, pipePos);
                        }
                        return fullPath;
                    }
                }
            }
            return null;
        }

        internal static RegisteredProjectType CreateRegisteredProjectType(Guid projectTypeGuid)
        {
            RegisteredProjectType registederedProjectType = null;

            using (var rootKey = VSRegistry.RegistryRoot(__VsLocalRegistryType.RegType_Configuration))
            {
                if (rootKey == null)
                {
                    return null;
                }

                var projectPath = "Projects\\" + projectTypeGuid.ToString("B");
                using (var projectKey = rootKey.OpenSubKey(projectPath))
                {
                    if (projectKey == null)
                    {
                        return null;
                    }

                    registederedProjectType = new RegisteredProjectType();
                    registederedProjectType.DefaultProjectExtensionValue =
                        projectKey.GetValue(DefaultProjectExtension) as string;
                    registederedProjectType.ProjectTemplatesDirValue =
                        projectKey.GetValue(ProjectTemplatesDir) as string;
                    registederedProjectType.WizardTemplatesDirValue = projectKey.GetValue(WizardsTemplatesDir) as string;
                    registederedProjectType.PackageGuidValue = new Guid(projectKey.GetValue(Package) as string);
                }
            }

            return registederedProjectType;
        }
    }
}