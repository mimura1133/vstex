using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Tagging;
using Microsoft.VisualStudio.Utilities;

namespace VsTeXProject
{
    [Export(typeof (ITaggerProvider))]
    [ContentType("vstex")]
    [TagType(typeof (TeXClassifierCommentOutFormatTag))]
    class TeXClassifierCommentOutFormatTagTaggerProvider : ITaggerProvider
    {
        public ITagger<T> CreateTagger<T>(ITextBuffer buffer) where T : ITag
        {
            if (buffer == null)
                throw new ArgumentNullException("buffer");

            return new TeXClassifierCommentOutFormatTagTagger() as ITagger<T>;
        }
    }

    [Export(typeof(ITaggerProvider))]
    [ContentType("vstex")]
    [TagType(typeof(TeXClassifierBeginEndFormatTag))]
    class TeXClassifierBeginEndFormatTagTaggerProvider : ITaggerProvider
    {
        public ITagger<T> CreateTagger<T>(ITextBuffer buffer) where T : ITag
        {
            if (buffer == null)
                throw new ArgumentNullException("buffer");

            return new TeXClassifierBeginEndFormatTagTagger() as ITagger<T>;
        }
    }

    [Export(typeof(ITaggerProvider))]
    [ContentType("vstex")]
    [TagType(typeof(TeXClassifierFunctionFormatTag))]
    class TeXClassifierFunctionFormatTagTaggerProvider : ITaggerProvider
    {
        public ITagger<T> CreateTagger<T>(ITextBuffer buffer) where T : ITag
        {
            if (buffer == null)
                throw new ArgumentNullException("buffer");

            return new TeXClassifierFunctionFormatTagTagger() as ITagger<T>;
        }
    }

    [Export(typeof(ITaggerProvider))]
    [ContentType("vstex")]
    [TagType(typeof(TeXClassifierBraceFormatTag))]
    class TeXClassifierBraceFormatTagTaggerProvider : ITaggerProvider
    {
        public ITagger<T> CreateTagger<T>(ITextBuffer buffer) where T : ITag
        {
            if (buffer == null)
                throw new ArgumentNullException("buffer");

            return new TeXClassifierBraceFormatTagTagger() as ITagger<T>;
        }
    }
}
