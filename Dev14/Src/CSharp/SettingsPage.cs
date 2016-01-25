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
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.Windows.Forms;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Designer.Interfaces;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell.Interop;

namespace VsTeXProject.VisualStudio.Project
{
    /// <summary>
    ///     The base class for property pages.
    /// </summary>
    [CLSCompliant(false), ComVisible(true)]
    public abstract class SettingsPage :
        LocalizableProperties,
        IPropertyPage,
        IDisposable
    {
        #region IDisposable Members

        /// <summary>
        ///     Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        #endregion

        private void Dispose(bool disposing)
        {
            if (!isDisposed)
            {
                lock (Mutex)
                {
                    if (disposing)
                    {
                        ThePanel.Dispose();
                    }

                    isDisposed = true;
                }
            }
        }

        #region fields

        private bool active;
        private bool dirty;
        private IPropertyPageSite site;
        private ProjectConfig[] projectConfigs;
        private static volatile object Mutex = new object();
        private bool isDisposed;

        #endregion

        #region properties

        [Browsable(false)]
        [AutomationBrowsable(false)]
        public string Name { get; set; }

        [Browsable(false)]
        [AutomationBrowsable(false)]
        public ProjectNode ProjectMgr { get; private set; }

        protected IVSMDPropertyGrid Grid { get; private set; }

        protected bool IsDirty
        {
            get { return dirty; }
            set
            {
                if (dirty != value)
                {
                    dirty = value;
                    if (site != null)
                        site.OnStatusChange((uint) (dirty ? PropPageStatus.Dirty : PropPageStatus.Clean));
                }
            }
        }

        protected Panel ThePanel { get; private set; }

        #endregion

        #region abstract methods

        protected abstract void BindProperties();
        protected abstract int ApplyChanges();

        #endregion

        #region public methods

        public object GetTypedConfigProperty(string name, Type type)
        {
            var value = GetConfigProperty(name);
            if (string.IsNullOrEmpty(value)) return null;

            var tc = TypeDescriptor.GetConverter(type);
            return tc.ConvertFromInvariantString(value);
        }

        public object GetTypedProperty(string name, Type type)
        {
            var value = GetProperty(name);
            if (string.IsNullOrEmpty(value)) return null;

            var tc = TypeDescriptor.GetConverter(type);
            return tc.ConvertFromInvariantString(value);
        }

        public string GetProperty(string propertyName)
        {
            if (ProjectMgr != null)
            {
                string property;
                var found = ProjectMgr.BuildProject.GlobalProperties.TryGetValue(propertyName, out property);

                if (found)
                {
                    return property;
                }
            }

            return string.Empty;
        }

        // relative to active configuration.
        public string GetConfigProperty(string propertyName)
        {
            if (ProjectMgr != null)
            {
                string unifiedResult = null;
                var cacheNeedReset = true;

                for (var i = 0; i < projectConfigs.Length; i++)
                {
                    var config = projectConfigs[i];
                    var property = config.GetConfigurationProperty(propertyName, cacheNeedReset);
                    cacheNeedReset = false;

                    if (property != null)
                    {
                        var text = property.Trim();

                        if (i == 0)
                            unifiedResult = text;
                        else if (unifiedResult != text)
                            return ""; // tristate value is blank then
                    }
                }

                return unifiedResult;
            }

            return string.Empty;
        }

        /// <summary>
        ///     Sets the value of a configuration dependent property.
        ///     If the attribute does not exist it is created.
        ///     If value is null it will be set to an empty string.
        /// </summary>
        /// <param name="name">property name.</param>
        /// <param name="value">value of property</param>
        public void SetConfigProperty(string name, string value)
        {
            CCITracing.TraceCall();
            if (value == null)
            {
                value = string.Empty;
            }

            if (ProjectMgr != null)
            {
                for (int i = 0, n = projectConfigs.Length; i < n; i++)
                {
                    var config = projectConfigs[i];

                    config.SetConfigurationProperty(name, value);
                }

                ProjectMgr.SetProjectFileDirty(true);
            }
        }

        #endregion

        #region IPropertyPage methods.

        public virtual void Activate(IntPtr parent, RECT[] pRect, int bModal)
        {
            if (ThePanel == null)
            {
                if (pRect == null)
                {
                    throw new ArgumentNullException("pRect");
                }

                ThePanel = new Panel();
                ThePanel.Size = new Size(pRect[0].right - pRect[0].left, pRect[0].bottom - pRect[0].top);
                ThePanel.Text = SR.GetString(SR.Settings, CultureInfo.CurrentUICulture);
                ThePanel.Visible = false;
                ThePanel.Size = new Size(550, 300);
                ThePanel.CreateControl();
                NativeMethods.SetParent(ThePanel.Handle, parent);
            }

            if (Grid == null && ProjectMgr != null && ProjectMgr.Site != null)
            {
                var pb = ProjectMgr.Site.GetService(typeof (IVSMDPropertyBrowser)) as IVSMDPropertyBrowser;
                Grid = pb.CreatePropertyGrid();
            }

            if (Grid != null)
            {
                active = true;


                var cGrid = Control.FromHandle(new IntPtr(Grid.Handle));

                cGrid.Parent = Control.FromHandle(parent); //this.panel;
                cGrid.Size = new Size(544, 294);
                cGrid.Location = new Point(3, 3);
                cGrid.Visible = true;
                Grid.SetOption(_PROPERTYGRIDOPTION.PGOPT_TOOLBAR, false);
                Grid.GridSort = _PROPERTYGRIDSORT.PGSORT_CATEGORIZED | _PROPERTYGRIDSORT.PGSORT_ALPHABETICAL;
                NativeMethods.SetParent(new IntPtr(Grid.Handle), ThePanel.Handle);
                UpdateObjects();
            }
        }

        public virtual int Apply()
        {
            if (IsDirty)
            {
                return ApplyChanges();
            }
            return VSConstants.S_OK;
        }

        public virtual void Deactivate()
        {
            if (null != ThePanel)
            {
                ThePanel.Dispose();
                ThePanel = null;
            }
            active = false;
        }

        public virtual void GetPageInfo(PROPPAGEINFO[] arrInfo)
        {
            if (arrInfo == null)
            {
                throw new ArgumentNullException("arrInfo");
            }

            var info = new PROPPAGEINFO();

            info.cb = (uint) Marshal.SizeOf(typeof (PROPPAGEINFO));
            info.dwHelpContext = 0;
            info.pszDocString = null;
            info.pszHelpFile = null;
            info.pszTitle = Name;
            info.SIZE.cx = 550;
            info.SIZE.cy = 300;
            arrInfo[0] = info;
        }

        public virtual void Help(string helpDir)
        {
        }

        public virtual int IsPageDirty()
        {
            // Note this returns an HRESULT not a Bool.
            return IsDirty ? VSConstants.S_OK : VSConstants.S_FALSE;
        }

        public virtual void Move(RECT[] arrRect)
        {
            if (arrRect == null)
            {
                throw new ArgumentNullException("arrRect");
            }

            var r = arrRect[0];

            ThePanel.Location = new Point(r.left, r.top);
            ThePanel.Size = new Size(r.right - r.left, r.bottom - r.top);
        }

        public virtual void SetObjects(uint count, object[] punk)
        {
            if (punk == null)
            {
                return;
            }

            if (count > 0)
            {
                if (punk[0] is ProjectConfig)
                {
                    var configs = new ArrayList();

                    for (var i = 0; i < count; i++)
                    {
                        var config = (ProjectConfig) punk[i];

                        if (ProjectMgr == null || (ProjectMgr != (punk[0] as ProjectConfig).ProjectMgr))
                        {
                            ProjectMgr = config.ProjectMgr;
                        }

                        configs.Add(config);
                    }

                    projectConfigs = (ProjectConfig[]) configs.ToArray(typeof (ProjectConfig));
                }
                else if (punk[0] is NodeProperties)
                {
                    if (ProjectMgr == null || (ProjectMgr != (punk[0] as NodeProperties).Node.ProjectMgr))
                    {
                        ProjectMgr = (punk[0] as NodeProperties).Node.ProjectMgr;
                    }

                    var configsMap = new Dictionary<string, ProjectConfig>();

                    for (var i = 0; i < count; i++)
                    {
                        var property = (NodeProperties) punk[i];
                        IVsCfgProvider provider;
                        ErrorHandler.ThrowOnFailure(property.Node.ProjectMgr.GetCfgProvider(out provider));
                        var expected = new uint[1];
                        ErrorHandler.ThrowOnFailure(provider.GetCfgs(0, null, expected, null));
                        if (expected[0] > 0)
                        {
                            var configs = new ProjectConfig[expected[0]];
                            var actual = new uint[1];
                            ErrorHandler.ThrowOnFailure(provider.GetCfgs(expected[0], configs, actual, null));

                            foreach (var config in configs)
                            {
                                if (!configsMap.ContainsKey(config.ConfigName))
                                {
                                    configsMap.Add(config.ConfigName, config);
                                }
                            }
                        }
                    }

                    if (configsMap.Count > 0)
                    {
                        if (projectConfigs == null)
                        {
                            projectConfigs = new ProjectConfig[configsMap.Keys.Count];
                        }
                        configsMap.Values.CopyTo(projectConfigs, 0);
                    }
                }
            }
            else
            {
                ProjectMgr = null;
            }

            if (active && ProjectMgr != null)
            {
                UpdateObjects();
            }
        }


        public virtual void SetPageSite(IPropertyPageSite theSite)
        {
            site = theSite;
        }

        public virtual void Show(uint cmd)
        {
            ThePanel.Visible = true; // TODO: pass SW_SHOW* flags through      
            ThePanel.Show();
        }

        public virtual int TranslateAccelerator(MSG[] arrMsg)
        {
            if (arrMsg == null)
            {
                throw new ArgumentNullException("arrMsg");
            }

            var msg = arrMsg[0];

            if ((msg.message < NativeMethods.WM_KEYFIRST || msg.message > NativeMethods.WM_KEYLAST) &&
                (msg.message < NativeMethods.WM_MOUSEFIRST || msg.message > NativeMethods.WM_MOUSELAST))
                return 1;

            return NativeMethods.IsDialogMessageA(ThePanel.Handle, ref msg) ? 0 : 1;
        }

        #endregion

        #region helper methods

        protected ProjectConfig[] GetProjectConfigurations()
        {
            return projectConfigs;
        }

        protected void UpdateObjects()
        {
            if (projectConfigs != null && ProjectMgr != null)
            {
                // Demand unmanaged permissions in order to access unmanaged memory.
                new SecurityPermission(SecurityPermissionFlag.UnmanagedCode).Demand();

                var p = Marshal.GetIUnknownForObject(this);
                var ppUnk = Marshal.AllocCoTaskMem(Marshal.SizeOf(typeof (IntPtr)));
                try
                {
                    Marshal.WriteIntPtr(ppUnk, p);
                    BindProperties();
                    // BUGBUG -- this is really bad casting a pointer to "int"...
                    Grid.SetSelectedObjects(1, ppUnk.ToInt32());
                    Grid.Refresh();
                }
                finally
                {
                    if (ppUnk != IntPtr.Zero)
                    {
                        Marshal.FreeCoTaskMem(ppUnk);
                    }
                    if (p != IntPtr.Zero)
                    {
                        Marshal.Release(p);
                    }
                }
            }
        }

        #endregion
    }
}