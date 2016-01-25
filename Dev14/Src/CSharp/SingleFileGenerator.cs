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
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.TextManager.Interop;

namespace VsTeXProject.VisualStudio.Project
{
    /// <summary>
    ///     Provides support for single file generator.
    /// </summary>
    internal class SingleFileGenerator : ISingleFileGenerator, IVsGeneratorProgress
    {
        #region ctors

        /// <summary>
        ///     Overloadde ctor.
        /// </summary>
        /// <param name="ProjectNode">The associated project</param>
        internal SingleFileGenerator(ProjectNode projectMgr)
        {
            this.projectMgr = projectMgr;
        }

        #endregion

        #region ISingleFileGenerator

        /// <summary>
        ///     Runs the generator on the current project item.
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        public virtual void RunGenerator(string document)
        {
            // Go run the generator on that node, but only if the file is dirty
            // in the running document table.  Otherwise there is no need to rerun
            // the generator because if the original document is not dirty then
            // the generated output should be already up to date.
            var itemid = VSConstants.VSITEMID_NIL;
            IVsHierarchy hier = projectMgr;
            if (document != null && hier != null &&
                ErrorHandler.Succeeded(hier.ParseCanonicalName(document, out itemid)))
            {
                IVsHierarchy rdtHier;
                IVsPersistDocData perDocData;
                uint cookie;
                if (VerifyFileDirtyInRdt(document, out rdtHier, out perDocData, out cookie))
                {
                    // Run the generator on the indicated document
                    var node = (FileNode) projectMgr.NodeFromItemId(itemid);
                    InvokeGenerator(node);
                }
            }
        }

        #endregion

        #region QueryEditQuerySave helpers

        /// <summary>
        ///     This function asks to the QueryEditQuerySave service if it is possible to
        ///     edit the file.
        /// </summary>
        private bool CanEditFile(string documentMoniker)
        {
            Trace.WriteLine(string.Format(CultureInfo.CurrentCulture, "\t**** CanEditFile called ****"));

            // Check the status of the recursion guard
            if (gettingCheckoutStatus)
            {
                return false;
            }

            try
            {
                // Set the recursion guard
                gettingCheckoutStatus = true;

                // Get the QueryEditQuerySave service
                var queryEditQuerySave = (IVsQueryEditQuerySave2) projectMgr.GetService(typeof (SVsQueryEditQuerySave));

                // Now call the QueryEdit method to find the edit status of this file
                string[] documents = {documentMoniker};
                uint result;
                uint outFlags;

                // Note that this function can popup a dialog to ask the user to checkout the file.
                // When this dialog is visible, it is possible to receive other request to change
                // the file and this is the reason for the recursion guard.
                var hr = queryEditQuerySave.QueryEditFiles(
                    0, // Flags
                    1, // Number of elements in the array
                    documents, // Files to edit
                    null, // Input flags
                    null, // Input array of VSQEQS_FILE_ATTRIBUTE_DATA
                    out result, // result of the checkout
                    out outFlags // Additional flags
                    );

                if (ErrorHandler.Succeeded(hr) && (result == (uint) tagVSQueryEditResult.QER_EditOK))
                {
                    // In this case (and only in this case) we can return true from this function.
                    return true;
                }
            }
            finally
            {
                gettingCheckoutStatus = false;
            }

            return false;
        }

        #endregion

        #region fields

        private bool gettingCheckoutStatus;
        private bool runningGenerator;
        private readonly ProjectNode projectMgr;

        #endregion

        #region IVsGeneratorProgress Members

        public virtual int GeneratorError(int warning, uint level, string err, uint line, uint col)
        {
            return VSConstants.E_NOTIMPL;
        }

        public virtual int Progress(uint complete, uint total)
        {
            return VSConstants.E_NOTIMPL;
        }

        #endregion

        #region virtual methods

        /// <summary>
        ///     Invokes the specified generator
        /// </summary>
        /// <param name="fileNode">The node on which to invoke the generator.</param>
        protected internal virtual void InvokeGenerator(FileNode fileNode)
        {
            if (fileNode == null)
            {
                throw new ArgumentNullException("fileNode");
            }

            var nodeproperties = fileNode.NodeProperties as SingleFileGeneratorNodeProperties;
            if (nodeproperties == null)
            {
                throw new InvalidOperationException();
            }

            var customToolProgID = nodeproperties.CustomTool;
            if (string.IsNullOrEmpty(customToolProgID))
            {
                return;
            }

            try
            {
                if (!runningGenerator)
                {
                    //Get the buffer contents for the current node
                    var moniker = fileNode.GetMkDocument();

                    runningGenerator = true;

                    //Get the generator
                    IVsSingleFileGenerator generator;
                    int generateDesignTimeSource;
                    int generateSharedDesignTimeSource;
                    int generateTempPE;
                    var factory = new SingleFileGeneratorFactory(projectMgr.ProjectGuid, projectMgr.Site);
                    ErrorHandler.ThrowOnFailure(factory.CreateGeneratorInstance(customToolProgID,
                        out generateDesignTimeSource, out generateSharedDesignTimeSource, out generateTempPE,
                        out generator));

                    //Check to see if the generator supports siting
                    var objWithSite = generator as IObjectWithSite;
                    if (objWithSite != null)
                    {
                        objWithSite.SetSite(fileNode.OleServiceProvider);
                    }

                    //Run the generator
                    var output = new IntPtr[1];
                    output[0] = IntPtr.Zero;
                    uint outPutSize;
                    string extension;
                    ErrorHandler.ThrowOnFailure(generator.DefaultExtension(out extension));

                    //Find if any dependent node exists
                    var dependentNodeName = Path.GetFileNameWithoutExtension(fileNode.FileName) + extension;
                    var dependentNode = fileNode.FirstChild;
                    while (dependentNode != null)
                    {
                        if (
                            string.Compare(dependentNode.ItemNode.GetMetadata(ProjectFileConstants.DependentUpon),
                                fileNode.FileName, StringComparison.OrdinalIgnoreCase) == 0)
                        {
                            dependentNodeName = ((FileNode) dependentNode).FileName;
                            break;
                        }

                        dependentNode = dependentNode.NextSibling;
                    }

                    //If you found a dependent node. 
                    if (dependentNode != null)
                    {
                        //Then check out the node and dependent node from SCC
                        if (!CanEditFile(dependentNode.GetMkDocument()))
                        {
                            throw Marshal.GetExceptionForHR(VSConstants.OLE_E_PROMPTSAVECANCELLED);
                        }
                    }
                    else //It is a new node to be added to the project
                    {
                        // Check out the project file if necessary.
                        if (!projectMgr.QueryEditProjectFile(false))
                        {
                            throw Marshal.GetExceptionForHR(VSConstants.OLE_E_PROMPTSAVECANCELLED);
                        }
                    }
                    IVsTextStream stream;
                    var inputFileContents = GetBufferContents(moniker, out stream);

                    ErrorHandler.ThrowOnFailure(generator.Generate(moniker, inputFileContents, "", output,
                        out outPutSize, this));
                    var data = new byte[outPutSize];

                    if (output[0] != IntPtr.Zero)
                    {
                        Marshal.Copy(output[0], data, 0, (int) outPutSize);
                        Marshal.FreeCoTaskMem(output[0]);
                    }

                    //Todo - Create a file and add it to the Project
                    UpdateGeneratedCodeFile(fileNode, data, (int) outPutSize, dependentNodeName);
                }
            }
            finally
            {
                runningGenerator = false;
            }
        }

        /// <summary>
        ///     Computes the names space based on the folder for the ProjectItem. It just replaces DirectorySeparatorCharacter
        ///     with "." for the directory in which the file is located.
        /// </summary>
        /// <returns>Returns the computed name space</returns>
        protected virtual string ComputeNamespace(string projectItemPath)
        {
            if (string.IsNullOrEmpty(projectItemPath))
            {
                throw new ArgumentException(
                    SR.GetString(SR.ParameterCannotBeNullOrEmpty, CultureInfo.CurrentUICulture), "projectItemPath");
            }


            var nspace = "";
            var filePath = Path.GetDirectoryName(projectItemPath);
            var toks = filePath.Split(':', '\\');
            foreach (var tok in toks)
            {
                if (!string.IsNullOrEmpty(tok))
                {
                    var temp = tok.Replace(" ", "");
                    nspace += temp + ".";
                }
            }
            nspace = nspace.Remove(nspace.LastIndexOf(".", StringComparison.Ordinal), 1);
            return nspace;
        }

        /// <summary>
        ///     This is called after the single file generator has been invoked to create or update the code file.
        /// </summary>
        /// <param name="fileNode">The node associated to the generator</param>
        /// <param name="data">data to update the file with</param>
        /// <param name="size">size of the data</param>
        /// <param name="fileName">Name of the file to update or create</param>
        /// <returns>full path of the file</returns>
        protected virtual string UpdateGeneratedCodeFile(FileNode fileNode, byte[] data, int size, string fileName)
        {
            var filePath = Path.Combine(Path.GetDirectoryName(fileNode.GetMkDocument()), fileName);
            var rdt = projectMgr.GetService(typeof (SVsRunningDocumentTable)) as IVsRunningDocumentTable;

            // (kberes) Shouldn't this be an InvalidOperationException instead with some not to annoying errormessage to the user?
            if (rdt == null)
            {
                ErrorHandler.ThrowOnFailure(VSConstants.E_FAIL);
            }

            IVsHierarchy hier;
            uint cookie;
            uint itemid;
            var docData = IntPtr.Zero;
            ErrorHandler.ThrowOnFailure(rdt.FindAndLockDocument((uint) _VSRDTFLAGS.RDT_NoLock, filePath, out hier,
                out itemid, out docData, out cookie));
            if (docData != IntPtr.Zero)
            {
                Marshal.Release(docData);
                IVsTextStream srpStream = null;
                if (srpStream != null)
                {
                    var oldLen = 0;
                    var hr = srpStream.GetSize(out oldLen);
                    if (ErrorHandler.Succeeded(hr))
                    {
                        var dest = IntPtr.Zero;
                        try
                        {
                            dest = Marshal.AllocCoTaskMem(data.Length);
                            Marshal.Copy(data, 0, dest, data.Length);
                            ErrorHandler.ThrowOnFailure(srpStream.ReplaceStream(0, oldLen, dest, size/2));
                        }
                        finally
                        {
                            if (dest != IntPtr.Zero)
                            {
                                Marshal.Release(dest);
                            }
                        }
                    }
                }
            }
            else
            {
                using (var generatedFileStream = File.Open(filePath, FileMode.OpenOrCreate))
                {
                    generatedFileStream.Write(data, 0, size);
                }

                var projectItem = fileNode.GetAutomationObject() as ProjectItem;
                if (projectItem != null && (projectMgr.FindChild(fileNode.FileName) == null))
                {
                    projectItem.ProjectItems.AddFromFile(filePath);
                }
            }
            return filePath;
        }

        #endregion

        #region helpers

        /// <summary>
        ///     Returns the buffer contents for a moniker.
        /// </summary>
        /// <returns>Buffer contents</returns>
        private string GetBufferContents(string fileName, out IVsTextStream srpStream)
        {
            var CLSID_VsTextBuffer = new Guid("{8E7B96A8-E33D-11d0-A6D5-00C04FB67F6A}");
            var bufferContents = "";
            srpStream = null;

            var rdt = projectMgr.GetService(typeof (SVsRunningDocumentTable)) as IVsRunningDocumentTable;
            if (rdt != null)
            {
                IVsHierarchy hier;
                IVsPersistDocData persistDocData;
                uint itemid, cookie;
                var docInRdt = true;
                var docData = IntPtr.Zero;
                var hr = NativeMethods.E_FAIL;
                try
                {
                    //Getting a read lock on the document. Must be released later.
                    hr = rdt.FindAndLockDocument((uint) _VSRDTFLAGS.RDT_ReadLock, fileName, out hier, out itemid,
                        out docData, out cookie);
                    if (ErrorHandler.Failed(hr) || docData == IntPtr.Zero)
                    {
                        var iid = VSConstants.IID_IUnknown;
                        cookie = 0;
                        docInRdt = false;
                        var localReg = projectMgr.GetService(typeof (SLocalRegistry)) as ILocalRegistry;
                        ErrorHandler.ThrowOnFailure(localReg.CreateInstance(CLSID_VsTextBuffer, null, ref iid,
                            (uint) CLSCTX.CLSCTX_INPROC_SERVER, out docData));
                    }

                    persistDocData = Marshal.GetObjectForIUnknown(docData) as IVsPersistDocData;
                }
                finally
                {
                    if (docData != IntPtr.Zero)
                    {
                        Marshal.Release(docData);
                    }
                }

                //Try to get the Text lines
                var srpTextLines = persistDocData as IVsTextLines;
                if (srpTextLines == null)
                {
                    // Try getting a text buffer provider first
                    var srpTextBufferProvider = persistDocData as IVsTextBufferProvider;
                    if (srpTextBufferProvider != null)
                    {
                        hr = srpTextBufferProvider.GetTextBuffer(out srpTextLines);
                    }
                }

                if (ErrorHandler.Succeeded(hr))
                {
                    srpStream = srpTextLines as IVsTextStream;
                    if (srpStream != null)
                    {
                        // QI for IVsBatchUpdate and call FlushPendingUpdates if they support it
                        var srpBatchUpdate = srpStream as IVsBatchUpdate;
                        if (srpBatchUpdate != null)
                            ErrorHandler.ThrowOnFailure(srpBatchUpdate.FlushPendingUpdates(0));

                        var lBufferSize = 0;
                        hr = srpStream.GetSize(out lBufferSize);

                        if (ErrorHandler.Succeeded(hr))
                        {
                            var dest = IntPtr.Zero;
                            try
                            {
                                // Note that GetStream returns Unicode to us so we don't need to do any conversions
                                dest = Marshal.AllocCoTaskMem((lBufferSize + 1)*2);
                                ErrorHandler.ThrowOnFailure(srpStream.GetStream(0, lBufferSize, dest));
                                //Get the contents
                                bufferContents = Marshal.PtrToStringUni(dest);
                            }
                            finally
                            {
                                if (dest != IntPtr.Zero)
                                    Marshal.FreeCoTaskMem(dest);
                            }
                        }
                    }
                }
                // Unlock the document in the RDT if necessary
                if (docInRdt && rdt != null)
                {
                    ErrorHandler.ThrowOnFailure(
                        rdt.UnlockDocument((uint) (_VSRDTFLAGS.RDT_ReadLock | _VSRDTFLAGS.RDT_Unlock_NoSave), cookie));
                }

                if (ErrorHandler.Failed(hr))
                {
                    // If this failed then it's probably not a text file.  In that case,
                    // we just read the file as a binary
                    bufferContents = File.ReadAllText(fileName);
                }
            }
            return bufferContents;
        }

        /// <summary>
        ///     Returns TRUE if open and dirty. Note that documents can be open without a
        ///     window frame so be careful. Returns the DocData and doc cookie if requested
        /// </summary>
        /// <param name="document">document path</param>
        /// <param name="pHier">hierarchy</param>
        /// <param name="ppDocData">doc data associated with document</param>
        /// <param name="cookie">item cookie</param>
        /// <returns>True if FIle is dirty</returns>
        private bool VerifyFileDirtyInRdt(string document, out IVsHierarchy pHier, out IVsPersistDocData ppDocData,
            out uint cookie)
        {
            var ret = 0;
            pHier = null;
            ppDocData = null;
            cookie = 0;

            var rdt = projectMgr.GetService(typeof (IVsRunningDocumentTable)) as IVsRunningDocumentTable;
            if (rdt != null)
            {
                IntPtr docData;
                uint dwCookie = 0;
                IVsHierarchy srpHier;
                var itemid = VSConstants.VSITEMID_NIL;

                ErrorHandler.ThrowOnFailure(rdt.FindAndLockDocument((uint) _VSRDTFLAGS.RDT_NoLock, document, out srpHier,
                    out itemid, out docData, out dwCookie));
                var srpIVsPersistHierarchyItem = srpHier as IVsPersistHierarchyItem;
                if (srpIVsPersistHierarchyItem != null)
                {
                    // Found in the RDT. See if it is dirty
                    try
                    {
                        ErrorHandler.ThrowOnFailure(srpIVsPersistHierarchyItem.IsItemDirty(itemid, docData, out ret));
                        cookie = dwCookie;
                        ppDocData = Marshal.GetObjectForIUnknown(docData) as IVsPersistDocData;
                    }
                    finally
                    {
                        if (docData != IntPtr.Zero)
                        {
                            Marshal.Release(docData);
                        }

                        pHier = srpHier;
                    }
                }
            }
            return ret == 1;
        }

        #endregion
    }
}