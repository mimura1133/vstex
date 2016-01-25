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
using System.Runtime.InteropServices;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Win32;
using IServiceProvider = System.IServiceProvider;

namespace VsTeXProject.VisualStudio.Project
{
    /// <summary>
    ///     Provides implementation IVsSingleFileGeneratorFactory for
    /// </summary>
    public class SingleFileGeneratorFactory : IVsSingleFileGeneratorFactory
    {
        #region ctors

        /// <summary>
        ///     Constructor for SingleFileGeneratorFactory
        /// </summary>
        /// <param name="projectGuid">The project type guid of the associated project.</param>
        /// <param name="serviceProvider">A service provider.</param>
        public SingleFileGeneratorFactory(Guid projectType, IServiceProvider serviceProvider)
        {
            this.projectType = projectType;
            this.serviceProvider = serviceProvider;
        }

        #endregion

        #region nested types

        private class GeneratorMetaData
        {
            #region ctor

            #endregion

            #region fields

            #endregion

            #region Public Properties

            /// <summary>
            ///     Generator instance
            /// </summary>
            public object Generator { get; set; }

            /// <summary>
            ///     GeneratesDesignTimeSource reg value name under
            ///     HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\VisualStudio\[VsVer]\Generators\[ProjFacGuid]\[GeneratorProgId]
            /// </summary>
            public int GeneratesDesignTimeSource { get; set; } = -1;

            /// <summary>
            ///     GeneratesSharedDesignTimeSource reg value name under
            ///     HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\VisualStudio\[VsVer]\Generators\[ProjFacGuid]\[GeneratorProgId]
            /// </summary>
            public int GeneratesSharedDesignTimeSource { get; set; } = -1;

            /// <summary>
            ///     UseDesignTimeCompilationFlag reg value name under
            ///     HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\VisualStudio\[VsVer]\Generators\[ProjFacGuid]\[GeneratorProgId]
            /// </summary>
            public int UseDesignTimeCompilationFlag { get; set; } = -1;

            /// <summary>
            ///     Generator Class ID.
            /// </summary>
            public Guid GeneratorClsid { get; set; } = Guid.Empty;

            #endregion
        }

        #endregion

        #region fields

        /// <summary>
        ///     Base generator registry key for MPF based project
        /// </summary>
        private RegistryKey baseGeneratorRegistryKey;

        /// <summary>
        ///     CLSID reg value name under
        ///     HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\VisualStudio\[VsVer]\Generators\[ProjFacGuid]\[GeneratorProgId]
        /// </summary>
        private readonly string GeneratorClsid = "CLSID";

        /// <summary>
        ///     GeneratesDesignTimeSource reg value name under
        ///     HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\VisualStudio\[VsVer]\Generators\[ProjFacGuid]\[GeneratorProgId]
        /// </summary>
        private readonly string GeneratesDesignTimeSource = "GeneratesDesignTimeSource";

        /// <summary>
        ///     GeneratesSharedDesignTimeSource reg value name under
        ///     HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\VisualStudio\[VsVer]\Generators\[ProjFacGuid]\[GeneratorProgId]
        /// </summary>
        private readonly string GeneratesSharedDesignTimeSource = "GeneratesSharedDesignTimeSource";

        /// <summary>
        ///     UseDesignTimeCompilationFlag reg value name under
        ///     HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\VisualStudio\[VsVer]\Generators\[ProjFacGuid]\[GeneratorProgId]
        /// </summary>
        private readonly string UseDesignTimeCompilationFlag = "UseDesignTimeCompilationFlag";

        /// <summary>
        ///     Caches all the generators registered for the project type.
        /// </summary>
        private readonly Dictionary<string, GeneratorMetaData> generatorsMap =
            new Dictionary<string, GeneratorMetaData>();

        /// <summary>
        ///     The project type guid of the associated project.
        /// </summary>
        private Guid projectType;

        /// <summary>
        ///     A service provider
        /// </summary>
        private IServiceProvider serviceProvider;

        #endregion

        #region properties

        /// <summary>
        ///     Defines the project type guid of the associated project.
        /// </summary>
        public Guid ProjectGuid
        {
            get { return projectType; }
            set { projectType = value; }
        }

        /// <summary>
        ///     Defines an associated service provider.
        /// </summary>
        public IServiceProvider ServiceProvider
        {
            get { return serviceProvider; }
            set { serviceProvider = value; }
        }

        #endregion

        #region IVsSingleFileGeneratorFactory Helpers

        /// <summary>
        ///     Returns the project generator key under [VS-ConfigurationRoot]]\Generators
        /// </summary>
        private RegistryKey BaseGeneratorsKey
        {
            get
            {
                if (baseGeneratorRegistryKey == null)
                {
                    using (var root = VSRegistry.RegistryRoot(__VsLocalRegistryType.RegType_Configuration))
                    {
                        if (null != root)
                        {
                            var regPath = "Generators\\" + ProjectGuid.ToString("B");
                            baseGeneratorRegistryKey = root.OpenSubKey(regPath);
                        }
                    }
                }

                return baseGeneratorRegistryKey;
            }
        }

        /// <summary>
        ///     Returns the local registry instance
        /// </summary>
        private ILocalRegistry LocalRegistry
        {
            get { return serviceProvider.GetService(typeof (SLocalRegistry)) as ILocalRegistry; }
        }

        #endregion

        #region IVsSingleFileGeneratorFactory Members

        /// <summary>
        ///     Creates an instance of the single file generator requested
        /// </summary>
        /// <param name="progId">
        ///     prog id of the generator to be created. For e.g
        ///     HKLM\SOFTWARE\Microsoft\VisualStudio\9.0Exp\Generators\[prjfacguid]\[wszProgId]
        /// </param>
        /// <param name="generatesDesignTimeSource">GeneratesDesignTimeSource key value</param>
        /// <param name="generatesSharedDesignTimeSource">GeneratesSharedDesignTimeSource key value</param>
        /// <param name="useTempPEFlag">UseDesignTimeCompilationFlag key value</param>
        /// <param name="generate">IVsSingleFileGenerator interface</param>
        /// <returns>S_OK if succesful</returns>
        public virtual int CreateGeneratorInstance(string progId, out int generatesDesignTimeSource,
            out int generatesSharedDesignTimeSource, out int useTempPEFlag, out IVsSingleFileGenerator generate)
        {
            Guid genGuid;
            ErrorHandler.ThrowOnFailure(GetGeneratorInformation(progId, out generatesDesignTimeSource,
                out generatesSharedDesignTimeSource, out useTempPEFlag, out genGuid));

            //Create the single file generator and pass it out. Check to see if it is in the cache
            if (!generatorsMap.ContainsKey(progId) || (generatorsMap[progId].Generator == null))
            {
                var riid = VSConstants.IID_IUnknown;
                var dwClsCtx = (uint) CLSCTX.CLSCTX_INPROC_SERVER;
                var genIUnknown = IntPtr.Zero;
                //create a new one.
                ErrorHandler.ThrowOnFailure(LocalRegistry.CreateInstance(genGuid, null, ref riid, dwClsCtx,
                    out genIUnknown));
                if (genIUnknown != IntPtr.Zero)
                {
                    try
                    {
                        var generator = Marshal.GetObjectForIUnknown(genIUnknown);
                        //Build the generator meta data object and cache it.
                        var genData = new GeneratorMetaData();
                        genData.GeneratesDesignTimeSource = generatesDesignTimeSource;
                        genData.GeneratesSharedDesignTimeSource = generatesSharedDesignTimeSource;
                        genData.UseDesignTimeCompilationFlag = useTempPEFlag;
                        genData.GeneratorClsid = genGuid;
                        genData.Generator = generator;
                        generatorsMap[progId] = genData;
                    }
                    finally
                    {
                        Marshal.Release(genIUnknown);
                    }
                }
            }

            generate = generatorsMap[progId].Generator as IVsSingleFileGenerator;

            return VSConstants.S_OK;
        }

        /// <summary>
        ///     Gets the default generator based on the file extension.
        ///     HKLM\Software\Microsoft\VS\9.0\Generators\[prjfacguid]\.extension
        /// </summary>
        /// <param name="filename">File name with extension</param>
        /// <param name="progID">The generator prog ID</param>
        /// <returns>S_OK if successful</returns>
        public virtual int GetDefaultGenerator(string filename, out string progID)
        {
            progID = "";
            return VSConstants.E_NOTIMPL;
        }

        /// <summary>
        ///     Gets the generator information.
        /// </summary>
        /// <param name="progId">
        ///     prog id of the generator to be created. For e.g
        ///     HKLM\SOFTWARE\Microsoft\VisualStudio\9.0Exp\Generators\[prjfacguid]\[wszProgId]
        /// </param>
        /// <param name="generatesDesignTimeSource">GeneratesDesignTimeSource key value</param>
        /// <param name="generatesSharedDesignTimeSource">GeneratesSharedDesignTimeSource key value</param>
        /// <param name="useTempPEFlag">UseDesignTimeCompilationFlag key value</param>
        /// <param name="guiddGenerator">CLSID key value</param>
        /// <returns>S_OK if succesful</returns>
        public virtual int GetGeneratorInformation(string progId, out int generatesDesignTimeSource,
            out int generatesSharedDesignTimeSource, out int useTempPEFlag, out Guid guidGenerator)
        {
            RegistryKey genKey;
            generatesDesignTimeSource = -1;
            generatesSharedDesignTimeSource = -1;
            useTempPEFlag = -1;
            guidGenerator = Guid.Empty;
            if (string.IsNullOrEmpty(progId))
                return VSConstants.S_FALSE;

            //Create the single file generator and pass it out.
            if (!generatorsMap.ContainsKey(progId))
            {
                // We have to check whether the BaseGeneratorkey returns null.
                var tempBaseGeneratorKey = BaseGeneratorsKey;
                if (tempBaseGeneratorKey == null || (genKey = tempBaseGeneratorKey.OpenSubKey(progId)) == null)
                {
                    return VSConstants.S_FALSE;
                }

                //Get the CLSID
                var guid = (string) genKey.GetValue(GeneratorClsid, "");
                if (string.IsNullOrEmpty(guid))
                    return VSConstants.S_FALSE;

                var genData = new GeneratorMetaData();

                genData.GeneratorClsid = guidGenerator = new Guid(guid);
                //Get the GeneratesDesignTimeSource flag. Assume 0 if not present.
                genData.GeneratesDesignTimeSource =
                    generatesDesignTimeSource = (int) genKey.GetValue(GeneratesDesignTimeSource, 0);
                //Get the GeneratesSharedDesignTimeSource flag. Assume 0 if not present.
                genData.GeneratesSharedDesignTimeSource =
                    generatesSharedDesignTimeSource = (int) genKey.GetValue(GeneratesSharedDesignTimeSource, 0);
                //Get the UseDesignTimeCompilationFlag flag. Assume 0 if not present.
                genData.UseDesignTimeCompilationFlag =
                    useTempPEFlag = (int) genKey.GetValue(UseDesignTimeCompilationFlag, 0);
                generatorsMap.Add(progId, genData);
            }
            else
            {
                var genData = generatorsMap[progId];
                generatesDesignTimeSource = genData.GeneratesDesignTimeSource;
                //Get the GeneratesSharedDesignTimeSource flag. Assume 0 if not present.
                generatesSharedDesignTimeSource = genData.GeneratesSharedDesignTimeSource;
                //Get the UseDesignTimeCompilationFlag flag. Assume 0 if not present.
                useTempPEFlag = genData.UseDesignTimeCompilationFlag;
                //Get the CLSID
                guidGenerator = genData.GeneratorClsid;
            }

            return VSConstants.S_OK;
        }

        #endregion
    }
}