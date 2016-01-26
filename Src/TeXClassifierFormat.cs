using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;
using Microsoft.VisualStudio.Text.Classification;
using Microsoft.VisualStudio.Utilities;

namespace VsTeXProject
{
    class TeXClassifierFormat
    {
        [Export(typeof(EditorFormatDefinition))]
        [ClassificationType(ClassificationTypeNames = "TeXClassifierNormalFormat")]
        [Name("TeXClassifierNormalFormat")]
        [UserVisible(true)] //this should be visible to the end user
        [Order(Before = Priority.Low)] //set the priority to be after the default classifiers
        internal sealed class TeXClassifierNormalFormat : ClassificationFormatDefinition
        {
            public TeXClassifierNormalFormat()
            {
                DisplayName = "TeXClassifierNormalFormat"; //human readable version of the name
                ForegroundColor = Colors.Black;
            }
        }

        [Export(typeof(EditorFormatDefinition))]
        [ClassificationType(ClassificationTypeNames = "TeXClassifierCommentOutFormat")]
        [Name("TeXClassifierCommentOutFormat")]
        [UserVisible(true)] //this should be visible to the end user
        [Order(Before = Priority.High)] //set the priority to be after the default classifiers
        internal sealed class TeXClassifierCommentOutFormat : ClassificationFormatDefinition
        {
            public TeXClassifierCommentOutFormat()
            {
                DisplayName = "TeXClassifierCommentOutFormat"; //human readable version of the name
                ForegroundColor = Colors.Green;
            }
        }

        [Export(typeof(EditorFormatDefinition))]
        [ClassificationType(ClassificationTypeNames = "TeXClassifierBeginEndFormat")]
        [Name("TeXClassifierBeginEndFormat")]
        [UserVisible(true)] //this should be visible to the end user
        [Order(Before = Priority.Default)] //set the priority to be after the default classifiers
        internal sealed class TeXClassifierBeginEndFormat : ClassificationFormatDefinition
        {
            public TeXClassifierBeginEndFormat()
            {
                DisplayName = "TeXClassifierBeginEndFormat"; //human readable version of the name
                ForegroundColor = Colors.DarkRed;
            }
        }

        [Export(typeof(EditorFormatDefinition))]
        [ClassificationType(ClassificationTypeNames = "TeXClassifierFunctionFormat")]
        [Name("TeXClassifierFunctionFormat")]
        [UserVisible(true)] //this should be visible to the end user
        [Order(Before = Priority.Default)] //set the priority to be after the default classifiers
        internal sealed class TeXClassifierFunctionFormat : ClassificationFormatDefinition
        {
            public TeXClassifierFunctionFormat()
            {
                DisplayName = "TeXClassifierFunctionFormat"; //human readable version of the name
                ForegroundColor = Colors.Blue;
            }
        }

        [Export(typeof(EditorFormatDefinition))]
        [ClassificationType(ClassificationTypeNames = "TeXClassifierBraceFormat")]
        [Name("TeXClassifierBraceFormat")]
        [UserVisible(true)] //this should be visible to the end user
        [Order(Before = Priority.High)] //set the priority to be after the default classifiers
        internal sealed class TeXClassifierBraceFormat : ClassificationFormatDefinition
        {
            public TeXClassifierBraceFormat()
            {
                DisplayName = "TeXClassifierBraceFormat"; //human readable version of the name
                ForegroundColor = Colors.PaleVioletRed;
            }
        }
    }
}
