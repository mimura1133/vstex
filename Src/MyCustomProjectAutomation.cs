using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Project;
using Microsoft.VisualStudio.Project.Automation;

namespace VsTeXProject
{
    [ComVisible(true)]
    public class OAMyCustomProject : OAProject
    {
        #region Constructors
        /// <summary>
        /// Public constructor.
        /// </summary>
        /// <param name="project">Custom project.</param>
        public OAMyCustomProject(MyCustomProjectNode project)
            : base(project)
        {
        }
        #endregion
    }

    [ComVisible(true)]
    [Guid("A5B66A93-D986-4016-B055-947FA912E2EB")]
    public class OAMyCustomProjectFileItem : OAFileItem
    {
        #region Constructors
        /// <summary>
        /// Public constructor.
        /// </summary>
        /// <param name="project">Automation project.</param>
        /// <param name="node">Custom file node.</param>
        public OAMyCustomProjectFileItem(OAProject project, FileNode node)
            : base(project, node)
        {
        }
        #endregion
    }
}
