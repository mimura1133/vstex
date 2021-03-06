using System;
using System.Runtime.InteropServices;
using VsTeXProject.VisualStudio.Project;
using IOleServiceProvider = Microsoft.VisualStudio.OLE.Interop.IServiceProvider;

namespace VsTeXProject
{
    [Guid("58FF5D7A-0041-4898-AF00-EDFC288382FD")]
    public class TeXProjectFactory : ProjectFactory
    {
        #region Fields
        private CustomProjectPackage package;
        #endregion

        #region Constructors
        /// <summary>
        /// Explicit default constructor.
        /// </summary>
        /// <param name="package">Value of the project package for initialize internal package field.</param>
        public TeXProjectFactory(CustomProjectPackage package)
            : base(package)
        {
            this.package = package;
        }
        #endregion

        #region Overriden implementation
        /// <summary>
        /// Creates a new project by cloning an existing template project.
        /// </summary>
        /// <returns></returns>
        protected override ProjectNode CreateProject()
        {
            TeXProjectNode project = new TeXProjectNode(this.package);
            project.SetSite((IOleServiceProvider)((IServiceProvider)this.package).GetService(typeof(IOleServiceProvider)));
            return project;
        }
        #endregion
    }
}
