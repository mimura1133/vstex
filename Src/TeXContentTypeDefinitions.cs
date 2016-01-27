using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Utilities;

namespace VsTeXProject
{

    internal static class TeXContentTypeDefinitions
    {
        [Export] [Name("vstex")] [BaseDefinition("text")] internal static ContentTypeDefinition
            vstexContentTypeDefinition = null;

        [Export] [FileExtension(".tex")] [ContentType("vstex")] internal static FileExtensionToContentTypeDefinition
            vstex_texfile_FileExtensionDefinition = null;

        [Export]
        [FileExtension(".sty")]
        [ContentType("vstex")]
        internal static FileExtensionToContentTypeDefinition
            vstex_styfile_FileExtensionDefinition = null;

    }
}