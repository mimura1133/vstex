using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Classification;
using Microsoft.VisualStudio.Text.Tagging;
using Microsoft.VisualStudio.Utilities;

namespace VsTeXProject
{
    [Export(typeof(IClassifierProvider))]
    [ContentType("code")]
    class TeXClassifierProvider : IClassifierProvider
    {
        [Export(typeof(ClassificationTypeDefinition))]
        [Name("VsTeXClassifier")]
        internal ClassificationTypeDefinition ToDoClassificationType = null;

        [Import]
        internal IClassificationTypeRegistryService ClassificationRegistry = null;

        [Import]
        internal IBufferTagAggregatorFactoryService _tagAggregatorFactory = null;

        public IClassifier GetClassifier(ITextBuffer buffer)
        {
            IClassificationType classificationType = ClassificationRegistry.GetClassificationType("VsTeXClassifier");

            var tagAggregator = _tagAggregatorFactory.CreateTagAggregator<TeXTag>(buffer);
            return new TeXClassifier(tagAggregator, classificationType);
        }
    }

    [Export(typeof(EditorFormatDefinition))]
    [ClassificationType(ClassificationTypeNames = "VsTeXClassifier")]
    [Name("VsTeXFormatDefinition")]
    [UserVisible(true)]
    [Order(After = Priority.High)]
    internal sealed class ToDoFormat : ClassificationFormatDefinition
    {
        public ToDoFormat()
        {
            DisplayName = "VsTeX Classifier"; //human readable version of the name
            BackgroundOpacity = 1;
            BackgroundColor = Colors.Orange;
            ForegroundColor = Colors.OrangeRed;
        }
    }
}
