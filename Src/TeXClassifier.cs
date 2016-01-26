using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Classification;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Tagging;
using Microsoft.VisualStudio.Utilities;

namespace VsTeXProject
{

    [Export(typeof(IClassifierProvider))]
    [ContentType("vstex")]
    internal class TeXClassifierProvider : IClassifierProvider
    {
        /// <summary>
        ///     Import the classification registry to be used for getting a reference
        ///     to the custom classification type later.
        /// </summary>
        [Import]
        internal IClassificationTypeRegistryService classificationRegistry = null;

        [Import]
        internal IBufferTagAggregatorFactoryService tagAggregatorFactory = null;

        public TeXClassifierProvider()
        {
        }

        public IClassifier GetClassifier(ITextBuffer buffer)
        {
            return
                buffer.Properties.GetOrCreateSingletonProperty(
                    delegate { return new TeXClassifier(buffer, classificationRegistry, tagAggregatorFactory); });
        }
    }

    internal class TeXClassifier : IClassifier
    {
        private readonly IClassificationType _commentOutType;
        private readonly IClassificationType _beginEndType;
        private readonly IClassificationType _functionType;
        private readonly IClassificationType _braceType;

        private readonly ITagAggregator<TeXClassifierCommentOutFormatTag> _commentOutTag;
        private readonly ITagAggregator<TeXClassifierBeginEndFormatTag> _beginEndTag;
        private readonly ITagAggregator<TeXClassifierFunctionFormatTag> _functionTag;
        private readonly ITagAggregator<TeXClassifierBraceFormatTag> _braceTag;

        internal TeXClassifier(ITextBuffer buffer,IClassificationTypeRegistryService registry, IBufferTagAggregatorFactoryService factory)
        {
            _commentOutType = registry.GetClassificationType("TeXClassifierCommentOutFormat");
            _beginEndType = registry.GetClassificationType("TeXClassifierBeginEndFormat");
            _functionType = registry.GetClassificationType("TeXClassifierFunctionFormat");
            _braceType = registry.GetClassificationType("TeXClassifierBraceFormat");

            _commentOutTag = factory.CreateTagAggregator<TeXClassifierCommentOutFormatTag>(buffer);
            _beginEndTag = factory.CreateTagAggregator<TeXClassifierBeginEndFormatTag>(buffer);
            _functionTag = factory.CreateTagAggregator<TeXClassifierFunctionFormatTag>(buffer);
            _braceTag = factory.CreateTagAggregator<TeXClassifierBraceFormatTag>(buffer);
        }

        public IList<ClassificationSpan> GetClassificationSpans(SnapshotSpan span)
        {
            var classifiedSpans = new List<ClassificationSpan>();
            classifiedSpans.AddRange(GetClassificationSpansPerType(span, _commentOutType, _commentOutTag));
            classifiedSpans.AddRange(GetClassificationSpansPerType(span, _beginEndType, _beginEndTag));
            classifiedSpans.AddRange(GetClassificationSpansPerType(span, _functionType, _functionTag));
            classifiedSpans.AddRange(GetClassificationSpansPerType(span, _braceType, _braceTag));

            return classifiedSpans;
        }

        private List<ClassificationSpan> GetClassificationSpansPerType<T>(SnapshotSpan span, IClassificationType classtype,
            ITagAggregator<T> tagger)
            where T: IGlyphTag
        {
            var classifiedSpans = new List<ClassificationSpan>();
            var tags = tagger.GetTags(span);

            foreach (IMappingTagSpan<T> tagSpan in tags)
            {
                SnapshotSpan todoSpan = tagSpan.Span.GetSpans(span.Snapshot).First();
                classifiedSpans.Add(new ClassificationSpan(todoSpan, classtype));
            }

            return classifiedSpans;
        }

        /// <summary>
        /// Create an event for when the Classification changes
        /// </summary>
#pragma warning disable 67
        public event EventHandler<ClassificationChangedEventArgs> ClassificationChanged;
#pragma warning restore 67
    }
}
