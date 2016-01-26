using System.ComponentModel.Composition;
using Microsoft.VisualStudio.Text.Classification;
using Microsoft.VisualStudio.Utilities;

namespace VsTeXProject
{
    class TeXClassifierType
    {
        [Export(typeof (ClassificationTypeDefinition))] [Name("TeXClassifierNormalFormat")] internal static
            ClassificationTypeDefinition TeXClassifierNormalFormat = null;

        [Export(typeof (ClassificationTypeDefinition))] [Name("TeXClassifierCommentOutFormat")] internal static
            ClassificationTypeDefinition TeXClassifierCommentOutFormat = null;

        [Export(typeof (ClassificationTypeDefinition))] [Name("TeXClassifierBeginEndFormat")] internal static
            ClassificationTypeDefinition TeXClassifierBeginEndFormat = null;

        [Export(typeof (ClassificationTypeDefinition))] [Name("TeXClassifierFunctionFormat")] internal static
            ClassificationTypeDefinition TeXClassifierFunctionFormat = null;

        [Export(typeof (ClassificationTypeDefinition))] [Name("TeXClassifierBraceFormat")] internal static
            ClassificationTypeDefinition TeXClassifierBraceFormat = null;
    }
}
