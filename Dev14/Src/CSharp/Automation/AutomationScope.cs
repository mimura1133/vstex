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
using System.Diagnostics.CodeAnalysis;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;

namespace VsTeXProject.VisualStudio.Project.Automation
{
    /// <summary>
    ///     Helper class that handle the scope of an automation function.
    ///     It should be used inside a "using" directive to define the scope of the
    ///     automation function and make sure that the ExitAutomation method is called.
    /// </summary>
    internal class AutomationScope : IDisposable
    {
        private static volatile object Mutex;
        private bool inAutomation;
        private bool isDisposed;

        /// <summary>
        ///     Initializes the <see cref="AutomationScope" /> class.
        /// </summary>
        [SuppressMessage("Microsoft.Performance", "CA1810:InitializeReferenceTypeStaticFieldsInline")]
        static AutomationScope()
        {
            Mutex = new object();
        }

        /// <summary>
        ///     Defines the beginning of the scope of an automation function. This constuctor
        ///     calls EnterAutomationFunction to signal the Shell that the current function is
        ///     changing the status of the automation objects.
        /// </summary>
        public AutomationScope(IServiceProvider provider)
        {
            if (null == provider)
            {
                throw new ArgumentNullException("provider");
            }
            Extensibility = provider.GetService(typeof (IVsExtensibility)) as IVsExtensibility3;
            if (null == Extensibility)
            {
                throw new InvalidOperationException();
            }
            ErrorHandler.ThrowOnFailure(Extensibility.EnterAutomationFunction());
            inAutomation = true;
        }

        /// <summary>
        ///     Gets the IVsExtensibility3 interface used in the automation function.
        /// </summary>
        public IVsExtensibility3 Extensibility { get; }

        /// <summary>
        ///     Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        ///     Ends the scope of the automation function. This function is also called by the
        ///     Dispose method.
        /// </summary>
        public void ExitAutomation()
        {
            if (inAutomation)
            {
                ErrorHandler.ThrowOnFailure(Extensibility.ExitAutomationFunction());
                inAutomation = false;
            }
        }

        #region IDisposable Members

        private void Dispose(bool disposing)
        {
            if (!isDisposed)
            {
                lock (Mutex)
                {
                    if (disposing)
                    {
                        ExitAutomation();
                    }

                    isDisposed = true;
                }
            }
        }

        #endregion
    }
}