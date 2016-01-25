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
            try
            {
                if (node != null)
                {
                    var prop = (node.NodeProperties as FileNodeProperties);
                    switch (prop.Extension.ToLower())
                    {
                        case ".tex":
                            prop.BuildAction = BuildAction.Compile;
                            break;
                        case ".jpg":
                        case ".png":
                        case ".jpeg":
                        case ".bmp":
                        case ".gif":
                            prop.BuildAction = BuildAction.Picture;
                            break;
                        default:
                            prop.BuildAction = BuildAction.Content;
                            break;

                    }
                }
            }
            catch { }
        }
        #endregion
    }
}
