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
using Microsoft.Build.Construction;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using MSBuild = Microsoft.Build.Evaluation;

/* This file provides a basefunctionallity for IVsCfgProvider2.
   Instead of using the IVsProjectCfgEventsHelper object we have our own little sink and call our own helper methods
   similiar to the interface. But there is no real benefit in inheriting from the interface in the first place. 
   Using the helper object seems to be:  
    a) undocumented
    b) not really wise in the managed world
*/

namespace VsTeXProject.VisualStudio.Project
{
    [CLSCompliant(false)]
    [ComVisible(true)]
    public class ConfigProvider : IVsCfgProvider2, IVsProjectCfgProvider, IVsExtensibleObject
    {
        #region ctors

        public ConfigProvider(ProjectNode manager)
        {
            ProjectMgr = manager;
        }

        #endregion

        #region IVsExtensibleObject Members

        /// <summary>
        ///     Proved access to an IDispatchable object being a list of configuration properties
        /// </summary>
        /// <param name="configurationName">Combined Name and Platform for the configuration requested</param>
        /// <param name="configurationProperties">The IDispatchcable object</param>
        /// <returns>S_OK if successful</returns>
        public virtual int GetAutomationObject(string configurationName, out object configurationProperties)
        {
            //Init out param
            configurationProperties = null;

            string name, platform;
            if (!ProjectConfig.TrySplitConfigurationCanonicalName(configurationName, out name, out platform))
            {
                return VSConstants.E_INVALIDARG;
            }

            // Get the configuration
            IVsCfg cfg;
            ErrorHandler.ThrowOnFailure(GetCfgOfName(name, platform, out cfg));

            // Get the properties of the configuration
            configurationProperties = ((ProjectConfig) cfg).ConfigurationProperties;

            return VSConstants.S_OK;
        }

        #endregion

        /// <summary>
        ///     Get all the configurations in the project.
        /// </summary>
        private string[] GetPropertiesConditionedOn(string constant)
        {
            List<string> configurations = null;
            ProjectMgr.BuildProject.ReevaluateIfNecessary();
            ProjectMgr.BuildProject.ConditionedProperties.TryGetValue(constant, out configurations);

            return configurations == null ? new string[] {} : configurations.ToArray();
        }

        #region fields

        internal const string configString = " '$(Configuration)' == '{0}' ";
        internal const string AnyCPUPlatform = "Any CPU";
        internal const string x86Platform = "x86";

        private readonly EventSinkCollection cfgEventSinks = new EventSinkCollection();
        private readonly Dictionary<string, ProjectConfig> configurationsList = new Dictionary<string, ProjectConfig>();

        #endregion

        #region Properties

        /// <summary>
        ///     The associated project.
        /// </summary>
        protected ProjectNode ProjectMgr { get; }

        /// <summary>
        ///     If the project system wants to add custom properties to the property group then
        ///     they provide us with this data.
        ///     Returns/sets the [(<propName, propCondition>) <propValue>] collection
        /// </summary>
        [SuppressMessage("Microsoft.Design", "CA1002:DoNotExposeGenericLists"),
         SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual List<KeyValuePair<KeyValuePair<string, string>, string>> NewConfigProperties { get; set; } =
            new List<KeyValuePair<KeyValuePair<string, string>, string>>();

        #endregion

        #region methods

        /// <summary>
        ///     Creates new Project Configuartion objects based on the configuration name.
        /// </summary>
        /// <param name="configName">The name of the configuration</param>
        /// <returns>An instance of a ProjectConfig object.</returns>
        protected ProjectConfig GetProjectConfiguration(string configName)
        {
            // if we already created it, return the cached one
            if (configurationsList.ContainsKey(configName))
            {
                return configurationsList[configName];
            }

            var requestedConfiguration = CreateProjectConfiguration(configName);
            configurationsList.Add(configName, requestedConfiguration);

            return requestedConfiguration;
        }

        protected virtual ProjectConfig CreateProjectConfiguration(string configName)
        {
            return new ProjectConfig(ProjectMgr, configName);
        }

        #endregion

        #region IVsProjectCfgProvider methods

        /// <summary>
        ///     Provides access to the IVsProjectCfg interface implemented on a project's configuration object.
        /// </summary>
        /// <param name="projectCfgCanonicalName">The canonical name of the configuration to access.</param>
        /// <param name="projectCfg">The IVsProjectCfg interface of the configuration identified by szProjectCfgCanonicalName.</param>
        /// <returns>If the method succeeds, it returns S_OK. If it fails, it returns an error code. </returns>
        public virtual int OpenProjectCfg(string projectCfgCanonicalName, out IVsProjectCfg projectCfg)
        {
            if (projectCfgCanonicalName == null)
            {
                throw new ArgumentNullException("projectCfgCanonicalName");
            }

            projectCfg = null;

            // Be robust in release
            if (projectCfgCanonicalName == null)
            {
                return VSConstants.E_INVALIDARG;
            }


            Debug.Assert(ProjectMgr != null && ProjectMgr.BuildProject != null);

            var configs = GetPropertiesConditionedOn(ProjectFileConstants.Configuration);

            foreach (var config in configs)
            {
                if (string.Compare(config, projectCfgCanonicalName, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    projectCfg = GetProjectConfiguration(config);
                    if (projectCfg != null)
                    {
                        return VSConstants.S_OK;
                    }
                    return VSConstants.E_FAIL;
                }
            }

            return VSConstants.E_INVALIDARG;
        }

        /// <summary>
        ///     Checks whether or not this configuration provider uses independent configurations.
        /// </summary>
        /// <param name="usesIndependentConfigurations">
        ///     true if independent configurations are used, false if they are not used. By
        ///     default returns true.
        /// </param>
        /// <returns>If the method succeeds, it returns S_OK. If it fails, it returns an error code.</returns>
        public virtual int get_UsesIndependentConfigurations(out int usesIndependentConfigurations)
        {
            usesIndependentConfigurations = 1;
            return VSConstants.S_OK;
        }

        #endregion

        #region IVsCfgProvider2 methods

        /// <summary>
        ///     Copies an existing configuration name or creates a new one.
        /// </summary>
        /// <param name="name">The name of the new configuration.</param>
        /// <param name="cloneName">
        ///     the name of the configuration to copy, or a null reference, indicating that AddCfgsOfCfgName
        ///     should create a new configuration.
        /// </param>
        /// <param name="fPrivate">
        ///     Flag indicating whether or not the new configuration is private. If fPrivate is set to true, the
        ///     configuration is private. If set to false, the configuration is public. This flag can be ignored.
        /// </param>
        /// <returns>If the method succeeds, it returns S_OK. If it fails, it returns an error code. </returns>
        public virtual int AddCfgsOfCfgName(string name, string cloneName, int fPrivate)
        {
            // We need to QE/QS the project file
            if (!ProjectMgr.QueryEditProjectFile(false))
            {
                throw Marshal.GetExceptionForHR(VSConstants.OLE_E_PROMPTSAVECANCELLED);
            }

            // First create the condition that represent the configuration we want to clone
            var condition = cloneName == null
                ? string.Empty
                : string.Format(CultureInfo.InvariantCulture, configString, cloneName).Trim();

            // Get all configs
            var configGroup = new List<ProjectPropertyGroupElement>(ProjectMgr.BuildProject.Xml.PropertyGroups);
            ProjectPropertyGroupElement configToClone = null;

            if (cloneName != null)
            {
                // Find the configuration to clone
                foreach (var currentConfig in configGroup)
                {
                    // Only care about conditional property groups
                    if (currentConfig.Condition == null || currentConfig.Condition.Length == 0)
                        continue;

                    // Skip if it isn't the group we want
                    if (string.Compare(currentConfig.Condition.Trim(), condition, StringComparison.OrdinalIgnoreCase) !=
                        0)
                        continue;

                    configToClone = currentConfig;
                }
            }

            ProjectPropertyGroupElement newConfig = null;
            if (configToClone != null)
            {
                // Clone the configuration settings
                newConfig = ProjectMgr.ClonePropertyGroup(configToClone);
                //Will be added later with the new values to the path

                foreach (var property in newConfig.Properties)
                {
                    if (property.Name.Equals("OutputPath", StringComparison.OrdinalIgnoreCase))
                    {
                        property.Parent.RemoveChild(property);
                    }
                }
            }
            else
            {
                // no source to clone from, lets just create a new empty config
                newConfig = ProjectMgr.BuildProject.Xml.AddPropertyGroup();
                // Get the list of property name, condition value from the config provider
                IList<KeyValuePair<KeyValuePair<string, string>, string>> propVals = NewConfigProperties;
                foreach (var data in propVals)
                {
                    var propData = data.Key;
                    var value = data.Value;
                    var newProperty = newConfig.AddProperty(propData.Key, value);
                    if (!string.IsNullOrEmpty(propData.Value))
                        newProperty.Condition = propData.Value;
                }
            }


            //add the output path
            var outputBasePath = ProjectMgr.OutputBaseRelativePath;
            if (outputBasePath.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal))
                outputBasePath = Path.GetDirectoryName(outputBasePath);
            newConfig.AddProperty("OutputPath", Path.Combine(outputBasePath, name) + Path.DirectorySeparatorChar);

            // Set the condition that will define the new configuration
            var newCondition = string.Format(CultureInfo.InvariantCulture, configString, name);
            newConfig.Condition = newCondition;

            NotifyOnCfgNameAdded(name);
            return VSConstants.S_OK;
        }

        /// <summary>
        ///     Copies an existing platform name or creates a new one.
        /// </summary>
        /// <param name="platformName">The name of the new platform.</param>
        /// <param name="clonePlatformName">
        ///     The name of the platform to copy, or a null reference, indicating that
        ///     AddCfgsOfPlatformName should create a new platform.
        /// </param>
        /// <returns>If the method succeeds, it returns S_OK. If it fails, it returns an error code.</returns>
        public virtual int AddCfgsOfPlatformName(string platformName, string clonePlatformName)
        {
            return VSConstants.E_NOTIMPL;
        }

        /// <summary>
        ///     Deletes a specified configuration name.
        /// </summary>
        /// <param name="name">The name of the configuration to be deleted.</param>
        /// <returns>If the method succeeds, it returns S_OK. If it fails, it returns an error code. </returns>
        public virtual int DeleteCfgsOfCfgName(string name)
        {
            // We need to QE/QS the project file
            if (!ProjectMgr.QueryEditProjectFile(false))
            {
                throw Marshal.GetExceptionForHR(VSConstants.OLE_E_PROMPTSAVECANCELLED);
            }

            if (name == null)
            {
                Debug.Fail(string.Format(CultureInfo.CurrentCulture,
                    "Name of the configuration should not be null if you want to delete it from project: {0}",
                    ProjectMgr.BuildProject.FullPath));
                // The configuration " '$(Configuration)' ==  " does not exist, so technically the goal
                // is achieved so return S_OK
                return VSConstants.S_OK;
            }
            // Verify that this config exist
            var configs = GetPropertiesConditionedOn(ProjectFileConstants.Configuration);
            foreach (var config in configs)
            {
                if (string.Compare(config, name, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    // Create condition of config to remove
                    var condition = string.Format(CultureInfo.InvariantCulture, configString, config);

                    foreach (var element in ProjectMgr.BuildProject.Xml.PropertyGroups)
                    {
                        if (string.Equals(element.Condition, condition, StringComparison.OrdinalIgnoreCase))
                        {
                            element.Parent.RemoveChild(element);
                        }
                    }

                    NotifyOnCfgNameDeleted(name);
                }
            }

            return VSConstants.S_OK;
        }

        /// <summary>
        ///     Deletes a specified platform name.
        /// </summary>
        /// <param name="platName">The platform name to delet.</param>
        /// <returns>If the method succeeds, it returns S_OK. If it fails, it returns an error code.</returns>
        public virtual int DeleteCfgsOfPlatformName(string platName)
        {
            return VSConstants.E_NOTIMPL;
        }

        /// <summary>
        ///     Returns the existing configurations stored in the project file.
        /// </summary>
        /// <param name="celt">Specifies the requested number of property names. If this number is unknown, celt can be zero.</param>
        /// <param name="names">
        ///     On input, an allocated array to hold the number of configuration property names specified by celt. This parameter
        ///     can also be a null reference if the celt parameter is zero.
        ///     On output, names contains configuration property names.
        /// </param>
        /// <param name="actual">The actual number of property names returned.</param>
        /// <returns>If the method succeeds, it returns S_OK. If it fails, it returns an error code.</returns>
        public virtual int GetCfgNames(uint celt, string[] names, uint[] actual)
        {
            // get's called twice, once for allocation, then for retrieval            
            var i = 0;

            var configList = GetPropertiesConditionedOn(ProjectFileConstants.Configuration);

            if (names != null)
            {
                foreach (var config in configList)
                {
                    names[i++] = config;
                    if (i == celt)
                        break;
                }
            }
            else
                i = configList.Length;

            if (actual != null)
            {
                actual[0] = (uint) i;
            }

            return VSConstants.S_OK;
        }

        /// <summary>
        ///     Returns the configuration associated with a specified configuration or platform name.
        /// </summary>
        /// <param name="name">The name of the configuration to be returned.</param>
        /// <param name="platName">The name of the platform for the configuration to be returned.</param>
        /// <param name="cfg">The implementation of the IVsCfg interface.</param>
        /// <returns>If the method succeeds, it returns S_OK. If it fails, it returns an error code.</returns>
        public virtual int GetCfgOfName(string name, string platName, out IVsCfg cfg)
        {
            cfg = null;
            cfg = GetProjectConfiguration(name);

            return VSConstants.S_OK;
        }

        /// <summary>
        ///     Returns a specified configuration property.
        /// </summary>
        /// <param name="propid">
        ///     Specifies the property identifier for the property to return. For valid propid values, see
        ///     __VSCFGPROPID.
        /// </param>
        /// <param name="var">The value of the property.</param>
        /// <returns>If the method succeeds, it returns S_OK. If it fails, it returns an error code.</returns>
        public virtual int GetCfgProviderProperty(int propid, out object var)
        {
            var = false;
            switch ((__VSCFGPROPID) propid)
            {
                case __VSCFGPROPID.VSCFGPROPID_SupportsCfgAdd:
                    var = true;
                    break;

                case __VSCFGPROPID.VSCFGPROPID_SupportsCfgDelete:
                    var = true;
                    break;

                case __VSCFGPROPID.VSCFGPROPID_SupportsCfgRename:
                    var = true;
                    break;

                case __VSCFGPROPID.VSCFGPROPID_SupportsPlatformAdd:
                    var = false;
                    break;

                case __VSCFGPROPID.VSCFGPROPID_SupportsPlatformDelete:
                    var = false;
                    break;
            }
            return VSConstants.S_OK;
        }

        /// <summary>
        ///     Returns the per-configuration objects for this object.
        /// </summary>
        /// <param name="celt">
        ///     Number of configuration objects to be returned or zero, indicating a request for an unknown number
        ///     of objects.
        /// </param>
        /// <param name="a">
        ///     On input, pointer to an interface array or a null reference. On output, this parameter points to an
        ///     array of IVsCfg interfaces belonging to the requested configuration objects.
        /// </param>
        /// <param name="actual">
        ///     The number of configuration objects actually returned or a null reference, if this information is
        ///     not necessary.
        /// </param>
        /// <param name="flags">
        ///     Flags that specify settings for project configurations, or a null reference (Nothing in Visual
        ///     Basic) if no additional flag settings are required. For valid prgrFlags values, see __VSCFGFLAGS.
        /// </param>
        /// <returns>If the method succeeds, it returns S_OK. If it fails, it returns an error code.</returns>
        public virtual int GetCfgs(uint celt, IVsCfg[] a, uint[] actual, uint[] flags)
        {
            if (flags != null)
                flags[0] = 0;

            var i = 0;
            var configList = GetPropertiesConditionedOn(ProjectFileConstants.Configuration);

            if (a != null)
            {
                foreach (var configName in configList)
                {
                    a[i] = GetProjectConfiguration(configName);

                    i++;
                    if (i == celt)
                        break;
                }
            }
            else
                i = configList.Length;

            if (actual != null)
                actual[0] = (uint) i;

            return VSConstants.S_OK;
        }

        /// <summary>
        ///     Returns one or more platform names.
        /// </summary>
        /// <param name="celt">Specifies the requested number of platform names. If this number is unknown, celt can be zero.</param>
        /// <param name="names">
        ///     On input, an allocated array to hold the number of platform names specified by celt. This parameter
        ///     can also be a null reference if the celt parameter is zero. On output, names contains platform names.
        /// </param>
        /// <param name="actual">The actual number of platform names returned.</param>
        /// <returns>If the method succeeds, it returns S_OK. If it fails, it returns an error code.</returns>
        public virtual int GetPlatformNames(uint celt, string[] names, uint[] actual)
        {
            var platforms = GetPlatformsFromProject();
            return GetPlatforms(celt, names, actual, platforms);
        }

        /// <summary>
        ///     Returns the set of platforms that are installed on the user's machine.
        /// </summary>
        /// <param name="celt">
        ///     Specifies the requested number of supported platform names. If this number is unknown, celt can be
        ///     zero.
        /// </param>
        /// <param name="names">
        ///     On input, an allocated array to hold the number of names specified by celt. This parameter can also
        ///     be a null reference (Nothing in Visual Basic)if the celt parameter is zero. On output, names contains the names of
        ///     supported platforms
        /// </param>
        /// <param name="actual">The actual number of platform names returned.</param>
        /// <returns>If the method succeeds, it returns S_OK. If it fails, it returns an error code.</returns>
        public virtual int GetSupportedPlatformNames(uint celt, string[] names, uint[] actual)
        {
            var platforms = GetSupportedPlatformsFromProject();
            return GetPlatforms(celt, names, actual, platforms);
        }

        /// <summary>
        ///     Assigns a new name to a configuration.
        /// </summary>
        /// <param name="old">The old name of the target configuration.</param>
        /// <param name="newname">The new name of the target configuration.</param>
        /// <returns>If the method succeeds, it returns S_OK. If it fails, it returns an error code.</returns>
        public virtual int RenameCfgsOfCfgName(string old, string newname)
        {
            // First create the condition that represent the configuration we want to rename
            var condition = string.Format(CultureInfo.InvariantCulture, configString, old).Trim();

            foreach (var config in ProjectMgr.BuildProject.Xml.PropertyGroups)
            {
                // Only care about conditional property groups
                if (config.Condition == null || config.Condition.Length == 0)
                    continue;

                // Skip if it isn't the group we want
                if (string.Compare(config.Condition.Trim(), condition, StringComparison.OrdinalIgnoreCase) != 0)
                    continue;

                // Change the name 
                config.Condition = string.Format(CultureInfo.InvariantCulture, configString, newname);
                // Update the name in our config list
                if (configurationsList.ContainsKey(old))
                {
                    var configuration = configurationsList[old];
                    configurationsList.Remove(old);
                    configurationsList.Add(newname, configuration);
                    // notify the configuration of its new name
                    configuration.ConfigName = newname;
                }

                NotifyOnCfgNameRenamed(old, newname);
            }

            return VSConstants.S_OK;
        }

        /// <summary>
        ///     Cancels a registration for configuration event notification.
        /// </summary>
        /// <param name="cookie">The cookie used for registration.</param>
        /// <returns>If the method succeeds, it returns S_OK. If it fails, it returns an error code.</returns>
        public virtual int UnadviseCfgProviderEvents(uint cookie)
        {
            cfgEventSinks.RemoveAt(cookie);
            return VSConstants.S_OK;
        }

        /// <summary>
        ///     Registers the caller for configuration event notification.
        /// </summary>
        /// <param name="sink">
        ///     Reference to the IVsCfgProviderEvents interface to be called to provide notification of
        ///     configuration events.
        /// </param>
        /// <param name="cookie">Reference to a token representing the completed registration</param>
        /// <returns>If the method succeeds, it returns S_OK. If it fails, it returns an error code.</returns>
        public virtual int AdviseCfgProviderEvents(IVsCfgProviderEvents sink, out uint cookie)
        {
            cookie = cfgEventSinks.Add(sink);
            return VSConstants.S_OK;
        }

        #endregion

        #region helper methods

        /// <summary>
        ///     Called when a new configuration name was added.
        /// </summary>
        /// <param name="name">The name of configuration just added.</param>
        private void NotifyOnCfgNameAdded(string name)
        {
            foreach (IVsCfgProviderEvents sink in cfgEventSinks)
            {
                ErrorHandler.ThrowOnFailure(sink.OnCfgNameAdded(name));
            }
        }

        /// <summary>
        ///     Called when a config name was deleted.
        /// </summary>
        /// <param name="name">The name of the configuration.</param>
        private void NotifyOnCfgNameDeleted(string name)
        {
            foreach (IVsCfgProviderEvents sink in cfgEventSinks)
            {
                ErrorHandler.ThrowOnFailure(sink.OnCfgNameDeleted(name));
            }
        }

        /// <summary>
        ///     Called when a config name was renamed
        /// </summary>
        /// <param name="oldName">Old configuration name</param>
        /// <param name="newName">New configuration name</param>
        private void NotifyOnCfgNameRenamed(string oldName, string newName)
        {
            foreach (IVsCfgProviderEvents sink in cfgEventSinks)
            {
                ErrorHandler.ThrowOnFailure(sink.OnCfgNameRenamed(oldName, newName));
            }
        }

        /// <summary>
        ///     Called when a platform name was added
        /// </summary>
        /// <param name="platformName">The name of the platform.</param>
        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        private void NotifyOnPlatformNameAdded(string platformName)
        {
            foreach (IVsCfgProviderEvents sink in cfgEventSinks)
            {
                ErrorHandler.ThrowOnFailure(sink.OnPlatformNameAdded(platformName));
            }
        }

        /// <summary>
        ///     Called when a platform name was deleted
        /// </summary>
        /// <param name="platformName">The name of the platform.</param>
        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        private void NotifyOnPlatformNameDeleted(string platformName)
        {
            foreach (IVsCfgProviderEvents sink in cfgEventSinks)
            {
                ErrorHandler.ThrowOnFailure(sink.OnPlatformNameDeleted(platformName));
            }
        }

        /// <summary>
        ///     Gets all the platforms defined in the project
        /// </summary>
        /// <returns>An array of platform names.</returns>
        private string[] GetPlatformsFromProject()
        {
            var platforms = GetPropertiesConditionedOn(ProjectFileConstants.Platform);

            if (platforms == null || platforms.Length == 0)
            {
                return new[] {x86Platform, AnyCPUPlatform};
            }

            for (var i = 0; i < platforms.Length; i++)
            {
                platforms[i] = ConvertPlatformToVsProject(platforms[i]);
            }

            return platforms;
        }

        /// <summary>
        ///     Return the supported platform names.
        /// </summary>
        /// <returns>An array of supported platform names.</returns>
        private string[] GetSupportedPlatformsFromProject()
        {
            var platforms = ProjectMgr.BuildProject.GetPropertyValue(ProjectFileConstants.AvailablePlatforms);

            if (platforms == null)
            {
                return new string[] {};
            }

            if (platforms.Contains(","))
            {
                return platforms.Split(',');
            }

            return new[] {platforms};
        }

        /// <summary>
        ///     Helper function to convert AnyCPU to Any CPU.
        /// </summary>
        /// <param name="oldName">The oldname.</param>
        /// <returns>The new name.</returns>
        private static string ConvertPlatformToVsProject(string oldPlatformName)
        {
            if (string.Compare(oldPlatformName, ProjectFileValues.AnyCPU, StringComparison.OrdinalIgnoreCase) == 0)
            {
                return AnyCPUPlatform;
            }

            return oldPlatformName;
        }

        /// <summary>
        ///     Common method for handling platform names.
        /// </summary>
        /// <param name="celt">Specifies the requested number of platform names. If this number is unknown, celt can be zero.</param>
        /// <param name="names">
        ///     On input, an allocated array to hold the number of platform names specified by celt. This parameter
        ///     can also be null if the celt parameter is zero. On output, names contains platform names
        /// </param>
        /// <param name="actual">A count of the actual number of platform names returned.</param>
        /// <param name="platforms">An array of available platform names</param>
        /// <returns>A count of the actual number of platform names returned.</returns>
        /// <devremark>The platforms array is never null. It is assured by the callers.</devremark>
        private static int GetPlatforms(uint celt, string[] names, uint[] actual, string[] platforms)
        {
            Debug.Assert(platforms != null, "The plaforms array should never be null");
            if (names == null)
            {
                if (actual == null || actual.Length == 0)
                {
                    throw new ArgumentException(SR.GetString(SR.InvalidParameter, CultureInfo.CurrentUICulture),
                        "actual");
                }

                actual[0] = (uint) platforms.Length;
                return VSConstants.S_OK;
            }

            //Degenarate case
            if (celt == 0)
            {
                if (actual != null && actual.Length != 0)
                {
                    actual[0] = (uint) platforms.Length;
                }

                return VSConstants.S_OK;
            }

            uint returned = 0;
            for (var i = 0; i < platforms.Length && names.Length > returned; i++)
            {
                names[returned] = platforms[i];
                returned++;
            }

            if (actual != null && actual.Length != 0)
            {
                actual[0] = returned;
            }

            if (celt > returned)
            {
                return VSConstants.S_FALSE;
            }

            return VSConstants.S_OK;
        }

        #endregion
    }
}