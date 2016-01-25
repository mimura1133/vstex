using System;
using VsTeXProject.VisualStudio.Project;
using VsTeXProject.VisualStudio.Project.Automation;

namespace VsTeXProject
{
    /// <summary>
    /// This class extends the FileNode in order to represent a file 
    /// within the hierarchy.
    /// </summary>
    public class TeXProjectFileNode : FileNode
    {
        #region Fields
        private OATeXProjectFileItem automationObject;
        #endregion

        #region Constructors
        /// <summary>
        /// Initializes a new instance of the <see cref="TeXProjectFileNode"/> class.
        /// </summary>
        /// <param name="root">The project node.</param>
        /// <param name="e">The project element node.</param>
        internal TeXProjectFileNode(ProjectNode root, ProjectElement e)
            : base(root, e)
        {
        }
        #endregion

        #region Overriden implementation
        /// <summary>
        /// Gets the automation object for the file node.
        /// </summary>
        /// <returns></returns>
        public override object GetAutomationObject()
        {
            if(automationObject == null)
            {
                automationObject = new OATeXProjectFileItem(this.ProjectMgr.GetAutomationObject() as OAProject, this);
            }

            return automationObject;
        }
        #endregion

        #region Private implementation
        internal OleServiceProvider.ServiceCreatorCallback ServiceCreator
        {
            get { return new OleServiceProvider.ServiceCreatorCallback(this.CreateServices); }
        }

        private object CreateServices(Type serviceType)
        {
            object service = null;
            if(typeof(EnvDTE.ProjectItem) == serviceType)
            {
                service = GetAutomationObject();
            }
            return service;
        }
        #endregion
    }
}
