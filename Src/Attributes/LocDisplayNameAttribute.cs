using System;
using System.ComponentModel;

namespace VsTeXProject
{
    /// <summary>
    /// Specifies the display name for a property, event, 
    /// or public void method which takes no arguments.
    /// </summary>
    public sealed class LocDisplayNameAttribute : DisplayNameAttribute
    {
        #region Fields
        private string name;
        #endregion Fields

        #region Constructors
        /// <summary>
        /// Public constructor.
        /// </summary>
        /// <param name="name">Attribute display name.</param>
        public LocDisplayNameAttribute(string name)
        {
            this.name = name;
        }
        #endregion

        #region Overriden Implementation
        /// <summary>
        /// Gets attribute display name.
        /// </summary>
        public override string DisplayName
        {
            get
            {
                string result = Resources.GetString(this.name);

                if(result == null)
                {
                    result = this.name;
                }

                return result;
            }
        }
        #endregion
    }
}
