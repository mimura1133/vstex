using System.Runtime.InteropServices;
using VsTeXProject.VisualStudio.Project;
using VsTeXProject.VisualStudio.Project.Automation;

namespace VsTeXProject
{
    [ComVisible(true)]
    public class OaTeXProject : OAProject
    {
        #region Constructors
        /// <summary>
        /// Public constructor.
        /// </summary>
        /// <param name="project">Custom project.</param>
        public OaTeXProject(TeXProjectNode project)
            : base(project)
        {
        }
        #endregion
    }

    [ComVisible(true)]
    [Guid("A5B66A93-D986-4016-B055-947FA912E2EB")]
    public class OATeXProjectFileItem : OAFileItem
    {
        #region Constructors
        /// <summary>
        /// Public constructor.
        /// </summary>
        /// <param name="project">Automation project.</param>
        /// <param name="node">Custom file node.</param>
        public OATeXProjectFileItem(OAProject project, FileNode node)
            : base(project, node)
        {
        }
        #endregion
    }
}
