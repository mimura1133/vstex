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
using System.Text;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
//#define ConfigTrace
using MSBuild = Microsoft.Build.Evaluation;
using MSBuildExecution = Microsoft.Build.Execution;
using MSBuildConstruction = Microsoft.Build.Construction;

namespace VsTeXProject.VisualStudio.Project
{
    [CLSCompliant(false), ComVisible(true)]
    public class ProjectConfig :
        IVsCfg,
        IVsProjectCfg,
        IVsProjectCfg2,
        IVsProjectFlavorCfg,
        IVsDebuggableProjectCfg,
        ISpecifyPropertyPages,
        IVsSpecifyProjectDesignerPages,
        IVsCfgBrowseObject
    {
        #region ctors

        public ProjectConfig(ProjectNode project, string configuration)
        {
            ProjectMgr = project;
            ConfigName = configuration;

            // Because the project can be aggregated by a flavor, we need to make sure
            // we get the outer most implementation of that interface (hence: project --> IUnknown --> Interface)
            var projectUnknown = Marshal.GetIUnknownForObject(ProjectMgr);
            try
            {
                var flavorCfgProvider =
                    (IVsProjectFlavorCfgProvider)
                        Marshal.GetTypedObjectForIUnknown(projectUnknown, typeof (IVsProjectFlavorCfgProvider));
                ErrorHandler.ThrowOnFailure(flavorCfgProvider.CreateProjectFlavorCfg(this, out flavoredCfg));
                if (flavoredCfg == null)
                    throw new COMException();
            }
            finally
            {
                if (projectUnknown != IntPtr.Zero)
                    Marshal.Release(projectUnknown);
            }
            // if the flavored object support XML fragment, initialize it
            var persistXML = flavoredCfg as IPersistXMLFragment;
            if (null != persistXML)
            {
                ProjectMgr.LoadXmlFragment(persistXML, DisplayName);
            }
        }

        #endregion

        #region IVsSpecifyPropertyPages

        public void GetPages(CAUUID[] pages)
        {
            GetCfgPropertyPages(pages);
        }

        #endregion

        #region IVsSpecifyProjectDesignerPages

        /// <summary>
        ///     Implementation of the IVsSpecifyProjectDesignerPages. It will retun the pages that are configuration dependent.
        /// </summary>
        /// <param name="pages">The pages to return.</param>
        /// <returns>VSConstants.S_OK</returns>
        public virtual int GetProjectDesignerPages(CAUUID[] pages)
        {
            GetCfgPropertyPages(pages);
            return VSConstants.S_OK;
        }

        #endregion

        #region constants

        internal const string Debug = "Debug";
        internal const string Release = "Release";
        internal const string AnyCPU = "AnyCPU";

        #endregion

        #region fields

        private MSBuildExecution.ProjectInstance currentConfig;
        private List<OutputGroup> outputGroups;
        private IProjectConfigProperties configurationProperties;
        private readonly IVsProjectFlavorCfg flavoredCfg;
        private BuildableProjectConfig buildableCfg;

        #endregion

        #region properties

        public ProjectNode ProjectMgr { get; }

        public string ConfigName { get; set; }

        public virtual object ConfigurationProperties
        {
            get
            {
                if (configurationProperties == null)
                {
                    configurationProperties = new ProjectConfigProperties(this);
                }
                return configurationProperties;
            }
        }

        protected IList<OutputGroup> OutputGroups
        {
            get
            {
                if (null == outputGroups)
                {
                    // Initialize output groups
                    outputGroups = new List<OutputGroup>();

                    // Get the list of group names from the project.
                    // The main reason we get it from the project is to make it easier for someone to modify
                    // it by simply overriding that method and providing the correct MSBuild target(s).
                    var groupNames = ProjectMgr.GetOutputGroupNames();

                    if (groupNames != null)
                    {
                        // Populate the output array
                        foreach (var group in groupNames)
                        {
                            var outputGroup = CreateOutputGroup(ProjectMgr, group);
                            outputGroups.Add(outputGroup);
                        }
                    }
                }
                return outputGroups;
            }
        }

        #endregion

        #region methods

        protected virtual OutputGroup CreateOutputGroup(ProjectNode project, KeyValuePair<string, string> group)
        {
            var outputGroup = new OutputGroup(group.Key, group.Value, project, this);
            return outputGroup;
        }

        public void PrepareBuild(bool clean)
        {
            ProjectMgr.PrepareBuild(ConfigName, clean);
        }

        public virtual string GetConfigurationProperty(string propertyName, bool resetCache)
        {
            var property = GetMsBuildProperty(propertyName, resetCache);
            if (property == null)
                return null;

            return property.EvaluatedValue;
        }

        public virtual void SetConfigurationProperty(string propertyName, string propertyValue)
        {
            if (!ProjectMgr.QueryEditProjectFile(false))
            {
                throw Marshal.GetExceptionForHR(VSConstants.OLE_E_PROMPTSAVECANCELLED);
            }

            var condition = string.Format(CultureInfo.InvariantCulture, ConfigProvider.configString, ConfigName);

            SetPropertyUnderCondition(propertyName, propertyValue, condition);

            // property cache will need to be updated
            currentConfig = null;

            // Signal the output groups that something is changed
            foreach (var group in OutputGroups)
            {
                group.InvalidateGroup();
            }
            ProjectMgr.SetProjectFileDirty(true);
        }

        /// <summary>
        ///     Emulates the behavior of SetProperty(name, value, condition) on the old MSBuild object model.
        ///     This finds a property group with the specified condition (or creates one if necessary) then sets the property in
        ///     there.
        /// </summary>
        private void SetPropertyUnderCondition(string propertyName, string propertyValue, string condition)
        {
            var conditionTrimmed = condition == null ? string.Empty : condition.Trim();

            if (conditionTrimmed.Length == 0)
            {
                ProjectMgr.BuildProject.SetProperty(propertyName, propertyValue);
                return;
            }

            // New OM doesn't have a convenient equivalent for setting a property with a particular property group condition. 
            // So do it ourselves.
            MSBuildConstruction.ProjectPropertyGroupElement newGroup = null;

            foreach (var group in ProjectMgr.BuildProject.Xml.PropertyGroups)
            {
                if (string.Equals(group.Condition.Trim(), conditionTrimmed, StringComparison.OrdinalIgnoreCase))
                {
                    newGroup = group;
                    break;
                }
            }

            if (newGroup == null)
            {
                newGroup = ProjectMgr.BuildProject.Xml.AddPropertyGroup();
                    // Adds after last existing PG, else at start of project
                newGroup.Condition = condition;
            }

            foreach (var property in newGroup.PropertiesReversed) // If there's dupes, pick the last one so we win
            {
                if (string.Equals(property.Name, propertyName, StringComparison.OrdinalIgnoreCase) &&
                    property.Condition.Length == 0)
                {
                    property.Value = propertyValue;
                    return;
                }
            }

            newGroup.AddProperty(propertyName, propertyValue);
        }

        /// <summary>
        ///     If flavored, and if the flavor config can be dirty, ask it if it is dirty
        /// </summary>
        /// <param name="storageType">Project file or user file</param>
        /// <returns>0 = not dirty</returns>
        internal int IsFlavorDirty(_PersistStorageType storageType)
        {
            var isDirty = 0;
            if (flavoredCfg != null && flavoredCfg is IPersistXMLFragment)
            {
                ErrorHandler.ThrowOnFailure(((IPersistXMLFragment) flavoredCfg).IsFragmentDirty((uint) storageType,
                    out isDirty));
            }
            return isDirty;
        }

        /// <summary>
        ///     If flavored, ask the flavor if it wants to provide an XML fragment
        /// </summary>
        /// <param name="flavor">Guid of the flavor</param>
        /// <param name="storageType">Project file or user file</param>
        /// <param name="fragment">Fragment that the flavor wants to save</param>
        /// <returns>HRESULT</returns>
        internal int GetXmlFragment(Guid flavor, _PersistStorageType storageType, out string fragment)
        {
            fragment = null;
            var hr = VSConstants.S_OK;
            if (flavoredCfg != null && flavoredCfg is IPersistXMLFragment)
            {
                var flavorGuid = flavor;
                hr = ((IPersistXMLFragment) flavoredCfg).Save(ref flavorGuid, (uint) storageType, out fragment, 1);
            }
            return hr;
        }

        #endregion

        #region IVsCfg methods

        /// <summary>
        ///     The display name is a two part item
        ///     first part is the config name, 2nd part is the platform name
        /// </summary>
        public virtual int get_DisplayName(out string name)
        {
            name = DisplayName;
            return VSConstants.S_OK;
        }

        private string DisplayName
        {
            get
            {
                string name;
                var platform = new string[1];
                var actual = new uint[1];
                name = ConfigName;
                // currently, we only support one platform, so just add it..
                IVsCfgProvider provider;
                ErrorHandler.ThrowOnFailure(ProjectMgr.GetCfgProvider(out provider));
                ErrorHandler.ThrowOnFailure(((IVsCfgProvider2) provider).GetPlatformNames(1, platform, actual));
                if (!string.IsNullOrEmpty(platform[0]))
                {
                    name += "|" + platform[0];
                }
                return name;
            }
        }

        public virtual int get_IsDebugOnly(out int fDebug)
        {
            fDebug = 0;
            if (ConfigName == "Debug")
            {
                fDebug = 1;
            }
            return VSConstants.S_OK;
        }

        public virtual int get_IsReleaseOnly(out int fRelease)
        {
            CCITracing.TraceCall();
            fRelease = 0;
            if (ConfigName == "Release")
            {
                fRelease = 1;
            }
            return VSConstants.S_OK;
        }

        #endregion

        #region IVsProjectCfg methods

        public virtual int EnumOutputs(out IVsEnumOutputs eo)
        {
            CCITracing.TraceCall();
            eo = null;
            return VSConstants.E_NOTIMPL;
        }

        public virtual int get_BuildableProjectCfg(out IVsBuildableProjectCfg pb)
        {
            CCITracing.TraceCall();
            if (buildableCfg == null)
                buildableCfg = new BuildableProjectConfig(this);
            pb = buildableCfg;
            return VSConstants.S_OK;
        }

        public virtual int get_CanonicalName(out string name)
        {
            return ((IVsCfg) this).get_DisplayName(out name);
        }

        public virtual int get_IsPackaged(out int pkgd)
        {
            CCITracing.TraceCall();
            pkgd = 0;
            return VSConstants.S_OK;
        }

        public virtual int get_IsSpecifyingOutputSupported(out int f)
        {
            CCITracing.TraceCall();
            f = 1;
            return VSConstants.S_OK;
        }

        public virtual int get_Platform(out Guid platform)
        {
            CCITracing.TraceCall();
            platform = Guid.Empty;
            return VSConstants.E_NOTIMPL;
        }

        public virtual int get_ProjectCfgProvider(out IVsProjectCfgProvider p)
        {
            CCITracing.TraceCall();
            p = null;
            IVsCfgProvider cfgProvider = null;
            ProjectMgr.GetCfgProvider(out cfgProvider);
            if (cfgProvider != null)
            {
                p = cfgProvider as IVsProjectCfgProvider;
            }

            return null == p ? VSConstants.E_NOTIMPL : VSConstants.S_OK;
        }

        public virtual int get_RootURL(out string root)
        {
            CCITracing.TraceCall();
            root = null;
            return VSConstants.S_OK;
        }

        public virtual int get_TargetCodePage(out uint target)
        {
            CCITracing.TraceCall();
            target = (uint) Encoding.Default.CodePage;
            return VSConstants.S_OK;
        }

        public virtual int get_UpdateSequenceNumber(ULARGE_INTEGER[] li)
        {
            if (li == null)
            {
                throw new ArgumentNullException("li");
            }

            CCITracing.TraceCall();
            li[0] = new ULARGE_INTEGER();
            li[0].QuadPart = 0;
            return VSConstants.S_OK;
        }

        public virtual int OpenOutput(string name, out IVsOutput output)
        {
            CCITracing.TraceCall();
            output = null;
            return VSConstants.E_NOTIMPL;
        }

        #endregion

        #region IVsDebuggableProjectCfg methods

        /// <summary>
        ///     Called by the vs shell to start debugging (managed or unmanaged).
        ///     Override this method to support other debug engines.
        /// </summary>
        /// <param name="grfLaunch">
        ///     A flag that determines the conditions under which to start the debugger. For valid grfLaunch
        ///     values, see __VSDBGLAUNCHFLAGS
        /// </param>
        /// <returns>If the method succeeds, it returns S_OK. If it fails, it returns an error code</returns>
        public virtual int DebugLaunch(uint grfLaunch)
        {
            CCITracing.TraceCall();

            try
            {
                var info = new VsDebugTargetInfo();
                info.cbSize = (uint) Marshal.SizeOf(info);
                info.dlo = DEBUG_LAUNCH_OPERATION.DLO_CreateProcess;

                // On first call, reset the cache, following calls will use the cached values
                var property = GetConfigurationProperty("StartProgram", true);
                if (string.IsNullOrEmpty(property))
                {
                    info.bstrExe = ProjectMgr.GetOutputAssembly(ConfigName);
                }
                else
                {
                    info.bstrExe = property;
                }

                property = GetConfigurationProperty("WorkingDirectory", false);
                if (string.IsNullOrEmpty(property))
                {
                    info.bstrCurDir = Path.GetDirectoryName(info.bstrExe);
                }
                else
                {
                    info.bstrCurDir = property;
                }

                property = GetConfigurationProperty("CmdArgs", false);
                if (!string.IsNullOrEmpty(property))
                {
                    info.bstrArg = property;
                }

                property = GetConfigurationProperty("RemoteDebugMachine", false);
                if (property != null && property.Length > 0)
                {
                    info.bstrRemoteMachine = property;
                }

                info.fSendStdoutToOutputWindow = 0;

                property = GetConfigurationProperty("EnableUnmanagedDebugging", false);
                if (property != null && string.Compare(property, "true", StringComparison.OrdinalIgnoreCase) == 0)
                {
                    //Set the unmanged debugger
                    //TODO change to vsconstant when it is available in VsConstants (guidNativeOnlyEng was the old name, maybe it has got a new name)
                    info.clsidCustom = new Guid("{3B476D35-A401-11D2-AAD4-00C04F990171}");
                }
                else
                {
                    //Set the managed debugger
                    info.clsidCustom = VSConstants.CLSID_ComPlusOnlyDebugEngine;
                }
                info.grfLaunch = grfLaunch;
                VsShellUtilities.LaunchDebugger(ProjectMgr.Site, info);
            }
            catch (Exception e)
            {
                Trace.WriteLine("Exception : " + e.Message);

                return Marshal.GetHRForException(e);
            }

            return VSConstants.S_OK;
        }

        /// <summary>
        ///     Determines whether the debugger can be launched, given the state of the launch flags.
        /// </summary>
        /// <param name="flags">
        ///     Flags that determine the conditions under which to launch the debugger.
        ///     For valid grfLaunch values, see __VSDBGLAUNCHFLAGS or __VSDBGLAUNCHFLAGS2.
        /// </param>
        /// <param name="fCanLaunch">true if the debugger can be launched, otherwise false</param>
        /// <returns>S_OK if the method succeeds, otherwise an error code</returns>
        public virtual int QueryDebugLaunch(uint flags, out int fCanLaunch)
        {
            CCITracing.TraceCall();
            var assembly = ProjectMgr.GetAssemblyName(ConfigName);
            fCanLaunch = assembly != null &&
                         assembly.ToUpperInvariant().EndsWith(".exe", StringComparison.OrdinalIgnoreCase)
                ? 1
                : 0;
            if (fCanLaunch == 0)
            {
                var property = GetConfigurationProperty("StartProgram", true);
                fCanLaunch = property != null && property.Length > 0 ? 1 : 0;
            }
            return VSConstants.S_OK;
        }

        #endregion

        #region IVsProjectCfg2 Members

        public virtual int OpenOutputGroup(string szCanonicalName, out IVsOutputGroup ppIVsOutputGroup)
        {
            ppIVsOutputGroup = null;
            // Search through our list of groups to find the one they are looking forgroupName
            foreach (var group in OutputGroups)
            {
                string groupName;
                group.get_CanonicalName(out groupName);
                if (string.Compare(groupName, szCanonicalName, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    ppIVsOutputGroup = group;
                    break;
                }
            }
            return ppIVsOutputGroup != null ? VSConstants.S_OK : VSConstants.E_FAIL;
        }

        public virtual int OutputsRequireAppRoot(out int pfRequiresAppRoot)
        {
            pfRequiresAppRoot = 0;
            return VSConstants.E_NOTIMPL;
        }

        public virtual int get_CfgType(ref Guid iidCfg, out IntPtr ppCfg)
        {
            // Delegate to the flavored configuration (to enable a flavor to take control)
            // Since we can be asked for Configuration we don't support, avoid throwing and return the HRESULT directly
            var hr = flavoredCfg.get_CfgType(ref iidCfg, out ppCfg);

            return hr;
        }

        public virtual int get_IsPrivate(out int pfPrivate)
        {
            pfPrivate = 0;
            return VSConstants.S_OK;
        }

        public virtual int get_OutputGroups(uint celt, IVsOutputGroup[] rgpcfg, uint[] pcActual)
        {
            // Are they only asking for the number of groups?
            if (celt == 0)
            {
                if ((null == pcActual) || (0 == pcActual.Length))
                {
                    throw new ArgumentNullException("pcActual");
                }
                pcActual[0] = (uint) OutputGroups.Count;
                return VSConstants.S_OK;
            }

            // Check that the array of output groups is not null
            if ((null == rgpcfg) || (rgpcfg.Length == 0))
            {
                throw new ArgumentNullException("rgpcfg");
            }

            // Fill the array with our output groups
            uint count = 0;
            foreach (var group in OutputGroups)
            {
                if (rgpcfg.Length > count && celt > count && group != null)
                {
                    rgpcfg[count] = group;
                    ++count;
                }
            }

            if (pcActual != null && pcActual.Length > 0)
                pcActual[0] = count;

            // If the number asked for does not match the number returned, return S_FALSE
            return count == celt ? VSConstants.S_OK : VSConstants.S_FALSE;
        }

        public virtual int get_VirtualRoot(out string pbstrVRoot)
        {
            pbstrVRoot = null;
            return VSConstants.E_NOTIMPL;
        }

        #endregion

        #region IVsCfgBrowseObject

        /// <summary>
        ///     Maps back to the configuration corresponding to the browse object.
        /// </summary>
        /// <param name="cfg">The IVsCfg object represented by the browse object</param>
        /// <returns>If the method succeeds, it returns S_OK. If it fails, it returns an error code. </returns>
        public virtual int GetCfg(out IVsCfg cfg)
        {
            cfg = this;
            return VSConstants.S_OK;
        }

        /// <summary>
        ///     Maps back to the hierarchy or project item object corresponding to the browse object.
        /// </summary>
        /// <param name="hier">Reference to the hierarchy object.</param>
        /// <param name="itemid">Reference to the project item.</param>
        /// <returns>If the method succeeds, it returns S_OK. If it fails, it returns an error code. </returns>
        public virtual int GetProjectItem(out IVsHierarchy hier, out uint itemid)
        {
            if (ProjectMgr == null || ProjectMgr.NodeProperties == null)
            {
                throw new InvalidOperationException();
            }
            return ProjectMgr.NodeProperties.GetProjectItem(out hier, out itemid);
        }

        #endregion

        #region helper methods

        /// <summary>
        ///     Splits the canonical configuration name into platform and configuration name.
        /// </summary>
        /// <param name="canonicalName">The canonicalName name.</param>
        /// <param name="configName">The name of the configuration.</param>
        /// <param name="platformName">The name of the platform.</param>
        /// <returns>true if successfull.</returns>
        internal static bool TrySplitConfigurationCanonicalName(string canonicalName, out string configName,
            out string platformName)
        {
            configName = string.Empty;
            platformName = string.Empty;

            if (string.IsNullOrEmpty(canonicalName))
            {
                return false;
            }

            var splittedCanonicalName = canonicalName.Split(new[] {'|'}, StringSplitOptions.RemoveEmptyEntries);

            if (splittedCanonicalName == null ||
                (splittedCanonicalName.Length != 1 && splittedCanonicalName.Length != 2))
            {
                return false;
            }

            configName = splittedCanonicalName[0];
            if (splittedCanonicalName.Length == 2)
            {
                platformName = splittedCanonicalName[1];
            }

            return true;
        }

        private MSBuildExecution.ProjectPropertyInstance GetMsBuildProperty(string propertyName, bool resetCache)
        {
            if (resetCache || currentConfig == null)
            {
                // Get properties for current configuration from project file and cache it
                ProjectMgr.SetConfiguration(ConfigName);
                ProjectMgr.BuildProject.ReevaluateIfNecessary();
                // Create a snapshot of the evaluated project in its current state
                currentConfig = ProjectMgr.BuildProject.CreateProjectInstance();

                // Restore configuration
                ProjectMgr.SetCurrentConfiguration();
            }

            if (currentConfig == null)
                throw new Exception("Failed to retrieve properties");

            // return property asked for
            return currentConfig.GetProperty(propertyName);
        }

        /// <summary>
        ///     Retrieves the configuration dependent property pages.
        /// </summary>
        /// <param name="pages">The pages to return.</param>
        private void GetCfgPropertyPages(CAUUID[] pages)
        {
            // We do not check whether the supportsProjectDesigner is set to true on the ProjectNode.
            // We rely that the caller knows what to call on us.
            if (pages == null)
            {
                throw new ArgumentNullException("pages");
            }

            if (pages.Length == 0)
            {
                throw new ArgumentException(SR.GetString(SR.InvalidParameter, CultureInfo.CurrentUICulture), "pages");
            }

            // Retrive the list of guids from hierarchy properties.
            // Because a flavor could modify that list we must make sure we are calling the outer most implementation of IVsHierarchy
            var guidsList = string.Empty;
            var hierarchy = ProjectMgr.InteropSafeIVsHierarchy;
            object variant = null;
            ErrorHandler.ThrowOnFailure(
                hierarchy.GetProperty(VSConstants.VSITEMID_ROOT, (int) __VSHPROPID2.VSHPROPID_CfgPropertyPagesCLSIDList,
                    out variant), VSConstants.DISP_E_MEMBERNOTFOUND, VSConstants.E_NOTIMPL);
            guidsList = (string) variant;

            var guids = Utilities.GuidsArrayFromSemicolonDelimitedStringOfGuids(guidsList);
            if (guids == null || guids.Length == 0)
            {
                pages[0] = new CAUUID();
                pages[0].cElems = 0;
            }
            else
            {
                pages[0] = PackageUtilities.CreateCAUUIDFromGuidArray(guids);
            }
        }

        #endregion

        #region IVsProjectFlavorCfg Members

        /// <summary>
        ///     This is called to let the flavored config let go
        ///     of any reference it may still be holding to the base config
        /// </summary>
        /// <returns></returns>
        int IVsProjectFlavorCfg.Close()
        {
            // This is used to release the reference the flavored config is holding
            // on the base config, but in our scenario these 2 are the same object
            // so we have nothing to do here.
            return VSConstants.S_OK;
        }

        /// <summary>
        ///     Actual implementation of get_CfgType.
        ///     When not flavored or when the flavor delegate to use
        ///     we end up creating the requested config if we support it.
        /// </summary>
        /// <param name="iidCfg">IID representing the type of config object we should create</param>
        /// <param name="ppCfg">Config object that the method created</param>
        /// <returns>HRESULT</returns>
        int IVsProjectFlavorCfg.get_CfgType(ref Guid iidCfg, out IntPtr ppCfg)
        {
            ppCfg = IntPtr.Zero;

            // See if this is an interface we support
            if (iidCfg == typeof (IVsDebuggableProjectCfg).GUID)
                ppCfg = Marshal.GetComInterfaceForObject(this, typeof (IVsDebuggableProjectCfg));
            else if (iidCfg == typeof (IVsBuildableProjectCfg).GUID)
            {
                IVsBuildableProjectCfg buildableConfig;
                get_BuildableProjectCfg(out buildableConfig);
                ppCfg = Marshal.GetComInterfaceForObject(buildableConfig, typeof (IVsBuildableProjectCfg));
            }

            // If not supported
            if (ppCfg == IntPtr.Zero)
                return VSConstants.E_NOINTERFACE;

            return VSConstants.S_OK;
        }

        #endregion
    }

    //=============================================================================
    // NOTE: advises on out of proc build execution to maximize
    // future cross-platform targeting capabilities of the VS tools.

    [CLSCompliant(false)]
    [ComVisible(true)]
    [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Buildable")]
    public class BuildableProjectConfig : IVsBuildableProjectCfg
    {
        #region ctors

        public BuildableProjectConfig(ProjectConfig config)
        {
            this.config = config;
        }

        #endregion

        #region fields

        private readonly ProjectConfig config;
        private readonly EventSinkCollection callbacks = new EventSinkCollection();

        #endregion

        #region IVsBuildableProjectCfg methods

        public virtual int AdviseBuildStatusCallback(IVsBuildStatusCallback callback, out uint cookie)
        {
            CCITracing.TraceCall();

            cookie = callbacks.Add(callback);
            return VSConstants.S_OK;
        }

        public virtual int get_ProjectCfg(out IVsProjectCfg p)
        {
            CCITracing.TraceCall();

            p = config;
            return VSConstants.S_OK;
        }

        public virtual int QueryStartBuild(uint options, int[] supported, int[] ready)
        {
            CCITracing.TraceCall();
            if (supported != null && supported.Length > 0)
                supported[0] = 1;
            if (ready != null && ready.Length > 0)
                ready[0] = config.ProjectMgr.BuildInProgress ? 0 : 1;
            return VSConstants.S_OK;
        }

        public virtual int QueryStartClean(uint options, int[] supported, int[] ready)
        {
            CCITracing.TraceCall();
            config.PrepareBuild(false);
            if (supported != null && supported.Length > 0)
                supported[0] = 1;
            if (ready != null && ready.Length > 0)
                ready[0] = config.ProjectMgr.BuildInProgress ? 0 : 1;
            return VSConstants.S_OK;
        }

        public virtual int QueryStartUpToDateCheck(uint options, int[] supported, int[] ready)
        {
            CCITracing.TraceCall();
            config.PrepareBuild(false);
            if (supported != null && supported.Length > 0)
                supported[0] = 0; // TODO:
            if (ready != null && ready.Length > 0)
                ready[0] = config.ProjectMgr.BuildInProgress ? 0 : 1;
            return VSConstants.S_OK;
        }

        public virtual int QueryStatus(out int done)
        {
            CCITracing.TraceCall();

            done = config.ProjectMgr.BuildInProgress ? 0 : 1;
            return VSConstants.S_OK;
        }

        public virtual int StartBuild(IVsOutputWindowPane pane, uint options)
        {
            CCITracing.TraceCall();
            config.PrepareBuild(false);

            // Current version of MSBuild wish to be called in an STA
            var flags = VSConstants.VS_BUILDABLEPROJECTCFGOPTS_REBUILD;

            // If we are not asked for a rebuild, then we build the default target (by passing null)
            Build(options, pane, (options & flags) != 0 ? MsBuildTarget.Rebuild : null);

            return VSConstants.S_OK;
        }

        public virtual int StartClean(IVsOutputWindowPane pane, uint options)
        {
            CCITracing.TraceCall();
            config.PrepareBuild(true);
            // Current version of MSBuild wish to be called in an STA
            Build(options, pane, MsBuildTarget.Clean);
            return VSConstants.S_OK;
        }

        public virtual int StartUpToDateCheck(IVsOutputWindowPane pane, uint options)
        {
            CCITracing.TraceCall();

            return VSConstants.E_NOTIMPL;
        }

        public virtual int Stop(int fsync)
        {
            CCITracing.TraceCall();

            return VSConstants.S_OK;
        }

        public virtual int UnadviseBuildStatusCallback(uint cookie)
        {
            CCITracing.TraceCall();


            callbacks.RemoveAt(cookie);
            return VSConstants.S_OK;
        }

        public virtual int Wait(uint ms, int fTickWhenMessageQNotEmpty)
        {
            CCITracing.TraceCall();

            return VSConstants.E_NOTIMPL;
        }

        #endregion

        #region helpers

        [SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes")]
        private bool NotifyBuildBegin()
        {
            var shouldContinue = 1;
            foreach (IVsBuildStatusCallback cb in callbacks)
            {
                try
                {
                    ErrorHandler.ThrowOnFailure(cb.BuildBegin(ref shouldContinue));
                    if (shouldContinue == 0)
                    {
                        return false;
                    }
                }
                catch (Exception e)
                {
                    // If those who ask for status have bugs in their code it should not prevent the build/notification from happening
                    Debug.Fail(string.Format(CultureInfo.CurrentCulture,
                        SR.GetString(SR.BuildEventError, CultureInfo.CurrentUICulture), e.Message));
                }
            }

            return true;
        }

        [SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes")]
        private void NotifyBuildEnd(MSBuildResult result, string buildTarget)
        {
            var success = result == MSBuildResult.Successful ? 1 : 0;

            foreach (IVsBuildStatusCallback cb in callbacks)
            {
                try
                {
                    ErrorHandler.ThrowOnFailure(cb.BuildEnd(success));
                }
                catch (Exception e)
                {
                    // If those who ask for status have bugs in their code it should not prevent the build/notification from happening
                    Debug.Fail(string.Format(CultureInfo.CurrentCulture,
                        SR.GetString(SR.BuildEventError, CultureInfo.CurrentUICulture), e.Message));
                }
                finally
                {
                    // We want to refresh the references if we are building with the Build or Rebuild target or if the project was opened for browsing only.
                    var shouldRepaintReferences = buildTarget == null || buildTarget == MsBuildTarget.Build ||
                                                  buildTarget == MsBuildTarget.Rebuild;

                    // Now repaint references if that is needed. 
                    // We hardly rely here on the fact the ResolveAssemblyReferences target has been run as part of the build.
                    // One scenario to think at is when an assembly reference is renamed on disk thus becomming unresolvable, 
                    // but msbuild can actually resolve it.
                    // Another one if the project was opened only for browsing and now the user chooses to build or rebuild.
                    if (shouldRepaintReferences && (result == MSBuildResult.Successful))
                    {
                        RefreshReferences();
                    }
                }
            }
        }

        private void Build(uint options, IVsOutputWindowPane output, string target)
        {
            if (!NotifyBuildBegin())
            {
                return;
            }

            try
            {
                config.ProjectMgr.BuildAsync(options, config.ConfigName, output, target,
                    (result, buildTarget) => NotifyBuildEnd(result, buildTarget));
            }
            catch (Exception e)
            {
                Trace.WriteLine("Exception : " + e.Message);
                ErrorHandler.ThrowOnFailure(output.OutputStringThreadSafe("Unhandled Exception:" + e.Message + "\n"));
                NotifyBuildEnd(MSBuildResult.Failed, target);
                throw;
            }
            finally
            {
                ErrorHandler.ThrowOnFailure(output.FlushToTaskList());
            }
        }

        /// <summary>
        ///     Refreshes references and redraws them correctly.
        /// </summary>
        private void RefreshReferences()
        {
            // Refresh the reference container node for assemblies that could be resolved.
            var referenceContainer = config.ProjectMgr.GetReferenceContainer();
            foreach (var referenceNode in referenceContainer.EnumReferences())
            {
                referenceNode.RefreshReference();
            }
        }

        #endregion
    }
}