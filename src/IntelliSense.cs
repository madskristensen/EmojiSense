using System.Collections.Generic;
using System.Globalization;
using System.Collections.Immutable;
using System.ComponentModel.Composition;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Core.Imaging;
using Microsoft.VisualStudio.Imaging;
using Microsoft.VisualStudio.Language.Intellisense.AsyncCompletion;
using Microsoft.VisualStudio.Language.Intellisense.AsyncCompletion.Data;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Adornments;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Utilities;
using Microsoft.VisualStudio.Text.Classification;

namespace EmojiSense
{
    [Export(typeof(IAsyncCompletionSourceProvider))]
    [ContentType("text")]
    [Name(nameof(IntelliSense))]
    public class IntelliSense : IAsyncCompletionSourceProvider
    {
        [Import]
        public IClassifierAggregatorService ClassifierAggregatorService { get; set; }

        public IAsyncCompletionSource GetOrCreate(ITextView textView)
        {
            IClassifier classifier = ClassifierAggregatorService.GetClassifier(textView.TextBuffer);
            return textView.Properties.GetOrCreateSingletonProperty(() => new AsyncCompletionSource(classifier));
        }
    }

    public class AsyncCompletionSource : IAsyncCompletionSource
    {
        private static string[] _supportedClassifications = new[] { "string", "comment", "CSS Comment", "CSS String Value", "HTML Comment", "XML Comment" };

        private readonly IClassifier _classifier;
        private static ImmutableArray<CompletionItem> _cache;
        private static readonly ImageElement _icon = new(KnownMonikers.GlyphRight.ToImageId(), "Variable");
        private static readonly Regex _regex = new(@"(?:\s|^):([^:\s]*):?", RegexOptions.Compiled);

        public static CompletionFilter PeopleFilter = new("People", "P", new ImageElement(KnownMonikers.FeedbackSmile.ToImageId()));
        public static CompletionFilter NatureFilter = new("Nature", "N", new ImageElement(KnownMonikers.BlankWebSite.ToImageId()));
        public static CompletionFilter ObjectFilter = new("Objects", "O", new ImageElement(KnownMonikers.NotificationAlertMute.ToImageId()));
        public static CompletionFilter PlacesFilter = new("Places", "L", new ImageElement(KnownMonikers.FlagGroup.ToImageId()));
        public static CompletionFilter SymbolsFilter = new("Symbols", "S", new ImageElement(KnownMonikers.LineArrow.ToImageId()));

        public AsyncCompletionSource(IClassifier classifier)
        {
            _classifier = classifier;
        }

        public Task<CompletionContext> GetCompletionContextAsync(IAsyncCompletionSession session, CompletionTrigger trigger, SnapshotPoint triggerLocation, SnapshotSpan applicableToSpan, CancellationToken token)
        {
            if (_cache == null)
            {
                var people = Emojis.People.Select(e => CreateCompletionItem(e, PeopleFilter));
                var nature = Emojis.Nature.Select(e => CreateCompletionItem(e, NatureFilter));
                var objects = Emojis.Objects.Select(e => CreateCompletionItem(e, ObjectFilter));
                var places = Emojis.Places.Select(e => CreateCompletionItem(e, PlacesFilter));
                var symbols = Emojis.Symbols.Select(e => CreateCompletionItem(e, SymbolsFilter));
                _cache = people.Union(nature).Union(objects).Union(places).Union(symbols).ToImmutableArray();
            }

            return Task.FromResult(new CompletionContext(_cache));
        }

        private CompletionItem CreateCompletionItem(KeyValuePair<string, string> emojiPair, CompletionFilter compFilter)
        {
            string name = emojiPair.Key;
            string displayName = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(name.Trim(':'));
            string value = emojiPair.Value;
            ImmutableArray<CompletionFilter> filter = new List<CompletionFilter>() { compFilter }.ToImmutableArray();
            ImmutableArray<ImageElement> icons = ImmutableArray<ImageElement>.Empty;

            return new CompletionItem(displayName, this, _icon, filter, value, value, displayName, displayName, icons);
        }

        public Task<object> GetDescriptionAsync(IAsyncCompletionSession session, CompletionItem item, CancellationToken token)
        {
            return Task.FromResult<object>(null);
        }

        public CompletionStartData InitializeCompletion(CompletionTrigger trigger, SnapshotPoint triggerLocation, CancellationToken token)
        {
            if (trigger.Character == ':' &&
                (triggerLocation == triggerLocation.Snapshot.Length ||
                (triggerLocation.GetChar() != ':' &&
                !triggerLocation.GetChar().IsWhiteSpaceOrZero())))
            {
                return CompletionStartData.DoesNotParticipateInCompletion;
            }

            if (trigger.Character != ':')
            {
                return CompletionStartData.DoesNotParticipateInCompletion;
            }

            ITextSnapshotLine line = triggerLocation.GetContainingLine();

            var classifications = _classifier.GetClassificationSpans(line.Extent);
            var triggerClassification = classifications.LastOrDefault(c => c.Span.Contains(triggerLocation - 1));

            if (triggerClassification == null || !_supportedClassifications.Any(c => triggerClassification.ClassificationType.IsOfType(c)))
            {
                return CompletionStartData.DoesNotParticipateInCompletion;
            }

            string lineText = line.GetText();

            foreach (Match match in _regex.Matches(lineText))
            {
                int offset = match.Value[0].IsWhiteSpaceOrZero() ? 1 : 0;
                int start = match.Index + line.Start + offset;
                int end = start + match.Length - offset;

                if (triggerLocation >= start && triggerLocation <= end)
                {
                    SnapshotSpan span = new(triggerLocation.Snapshot, Span.FromBounds(start, end));
                    return new CompletionStartData(CompletionParticipation.ProvidesItems, span);
                }
            }
            return CompletionStartData.DoesNotParticipateInCompletion;
        }
    }
}
