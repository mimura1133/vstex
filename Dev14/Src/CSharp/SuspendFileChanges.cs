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
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using ShellConstants = Microsoft.VisualStudio.Shell.Interop.Constants;

namespace VsTeXProject.VisualStudio.Project
{
    /// <summary>
    ///     helper to make the editor ignore external changes
    /// </summary>
    internal class SuspendFileChanges
    {
        private readonly string documentFileName;

        private IVsDocDataFileChangeControl fileChangeControl;

        private bool isSuspending;

        private readonly IServiceProvider site;

        public SuspendFileChanges(IServiceProvider site, string document)
        {
            this.site = site;
            documentFileName = document;
        }


        [SuppressMessage("Microsoft.Performance", "CA1800:DoNotCastUnnecessarily")]
        public void Suspend()
        {
            if (isSuspending)
                return;

            var docData = IntPtr.Zero;
            try
            {
                var rdt = site.GetService(typeof (SVsRunningDocumentTable)) as IVsRunningDocumentTable;

                IVsHierarchy hierarchy;
                uint itemId;
                uint docCookie;
                IVsFileChangeEx fileChange;


                if (rdt == null) return;

                ErrorHandler.ThrowOnFailure(rdt.FindAndLockDocument((uint) _VSRDTFLAGS.RDT_NoLock, documentFileName,
                    out hierarchy, out itemId, out docData, out docCookie));

                if ((docCookie == (uint) ShellConstants.VSDOCCOOKIE_NIL) || docData == IntPtr.Zero)
                    return;

                fileChange = site.GetService(typeof (SVsFileChangeEx)) as IVsFileChangeEx;

                if (fileChange != null)
                {
                    isSuspending = true;
                    ErrorHandler.ThrowOnFailure(fileChange.IgnoreFile(0, documentFileName, 1));
                    if (docData != IntPtr.Zero)
                    {
                        IVsPersistDocData persistDocData = null;

                        // if interface is not supported, return null
                        var unknown = Marshal.GetObjectForIUnknown(docData);
                        if (unknown is IVsPersistDocData)
                        {
                            persistDocData = (IVsPersistDocData) unknown;
                            if (persistDocData is IVsDocDataFileChangeControl)
                            {
                                fileChangeControl = (IVsDocDataFileChangeControl) persistDocData;
                                if (fileChangeControl != null)
                                {
                                    ErrorHandler.ThrowOnFailure(fileChangeControl.IgnoreFileChanges(1));
                                }
                            }
                        }
                    }
                }
            }
            catch (InvalidCastException e)
            {
                Trace.WriteLine("Exception" + e.Message);
            }
            finally
            {
                if (docData != IntPtr.Zero)
                {
                    Marshal.Release(docData);
                }
            }
        }

        public void Resume()
        {
            if (!isSuspending)
                return;
            IVsFileChangeEx fileChange;
            fileChange = site.GetService(typeof (SVsFileChangeEx)) as IVsFileChangeEx;
            if (fileChange != null)
            {
                isSuspending = false;
                ErrorHandler.ThrowOnFailure(fileChange.IgnoreFile(0, documentFileName, 0));
                if (fileChangeControl != null)
                {
                    ErrorHandler.ThrowOnFailure(fileChangeControl.IgnoreFileChanges(0));
                }
            }
        }
    }
}