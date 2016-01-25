using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Classification;
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
        [Import] internal IClassificationTypeRegistryService ClassificationRegistry = null;

        public TeXClassifierProvider()
        {
        }

        public IClassifier GetClassifier(ITextBuffer buffer)
        {
            return
                buffer.Properties.GetOrCreateSingletonProperty(
                    delegate { return new TeXClassifier(ClassificationRegistry); });
        }
    }

    internal class TeXClassifier : IClassifier
    {
        private readonly IClassificationType _normalType;
        private readonly IClassificationType _commentOutType;
        private readonly IClassificationType _beginEndType;
        private readonly IClassificationType _functionType;
        private readonly IClassificationType _braceType;

        internal TeXClassifier(IClassificationTypeRegistryService registry)
        {
            _normalType = registry.GetClassificationType("TeXClassifierNormalFormat");
            _commentOutType = registry.GetClassificationType("TeXClassifierCommentOutFormat");
            _beginEndType = registry.GetClassificationType("TeXClassifierBeginEndFormat");
            _functionType = registry.GetClassificationType("TeXClassifierFunctionFormat");
            _braceType = registry.GetClassificationType("TeXClassifierBraceFormat");
        }

        /// <summary>
        ///     This method scans the given SnapshotSpan for potential matches for this classification.
        ///     In this instance, it classifies everything and returns each span as a new ClassificationSpan.
        /// </summary>
        /// <param name="trackingSpan">The span currently being classified</param>
        /// <returns>A list of ClassificationSpans that represent spans identified to be of this classification</returns>
        public IList<ClassificationSpan> GetClassificationSpans(SnapshotSpan span)
        {
            var classifications = new List<ClassificationSpan>();
            var text = span.GetText();

            if (text == "")
            {
                return classifications;
            }

            for (int pt = span.Start; pt < span.End; pt++)
            {

                if (text == "")
                    return classifications;

                if (text[0] == '%')
                {
                    classifications.Add(
                        new ClassificationSpan(new SnapshotSpan(span.Snapshot, new Span(pt, text.Length)),
                            this._commentOutType));
                    return classifications;
                }

                if (text[0] == '\\')
                {
                    var color = this._functionType;
                    var len = 0;
                    var paramnest = 0;
                    var maxlines = 10;

                    if (text.StartsWith("\\begin") || text.StartsWith("\\end"))
                    {
                        color = this._beginEndType;
                    }

                    for (len = 1; len < text.Length; len++)
                    {
                        if (text[len] == '{' || text[len] == '[' || text[len] == '}' || text[len] == ']' ||
                            text[len] == ' ')
                            break;
                    }

                    var snaptext = text;
                    var line = span.Snapshot.GetLineNumberFromPosition(pt);
                    var start = pt;
                    var snapshotline = span.Snapshot.GetLineFromLineNumber(line);

                    do
                    {
                        
                        for (; len < snaptext.Length; len++)
                        {
                            if (snaptext[len] == '{' || snaptext[len] == '[') paramnest++;
                            if (snaptext[len] == '}' || snaptext[len] == ']') paramnest--;
                            if (snaptext[len] == '\\') len++;

                            if (paramnest == 0 && snaptext[len] != ' ' && snaptext[len] != '{' && snaptext[len] != '[' &&
                                snaptext[len] != '}' && snaptext[len] != ']') break;
                        }
                        classifications.Add(new ClassificationSpan(new SnapshotSpan(snapshotline.Snapshot, start, len),color));

                        maxlines--;
                        if (maxlines <= 0) break;
                        if (paramnest == 0) break;

                        line++;
                        if (line < span.Snapshot.Lines.Count())
                        {
                            snapshotline = span.Snapshot.GetLineFromLineNumber(line);
                            snaptext = snapshotline.GetTextIncludingLineBreak();
                            start = snapshotline.Start;
                            len = 0;
                        }
                    } while (line < span.Snapshot.Lines.Count());
                }

                text = text.Substring(1);
            }

            return classifications;
        }
#pragma warning disable 67
        public event EventHandler<ClassificationChangedEventArgs> ClassificationChanged;
#pragma warning restore 67
    }
}
