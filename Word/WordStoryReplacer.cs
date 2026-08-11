using Word = Microsoft.Office.Interop.Word;

public static class WordStoryReplacer
{
    public static void Replace(
        Word.Document document,
        string placeholder,
        string value)
    {
        Word.StoryRanges stories =
            document.StoryRanges;

        try
        {
            for (int i = 1; i <= stories.Count; i++)
            {
                Word.Range? story = null;

                try
                {
                    story = stories[i];

                    ReplaceRecursive(
                        story,
                        placeholder,
                        value);
                }
                finally
                {
                    Com.Release(story);
                }
            }
        }
        finally
        {
            Com.Release(stories);
        }
    }

    private static void ReplaceRecursive(
        Word.Range? story,
        string placeholder,
        string value)
    {
        if (story == null)
            return;

        Word.Find? find = null;

        try
        {
            find = story.Find;

            find.ClearFormatting();
            find.Replacement.ClearFormatting();

            find.Text = placeholder;
            find.Replacement.Text = value;

            find.Forward = true;
            find.Wrap = Word.WdFindWrap.wdFindStop;
            find.Format = false;

            find.Execute(
                Replace: Word.WdReplace.wdReplaceAll);
        }
        finally
        {
            Com.Release(find);
        }

        Word.Range? next = null;

        try
        {
            next = story.NextStoryRange;

            ReplaceRecursive(
                next,
                placeholder,
                value);
        }
        finally
        {
            Com.Release(next);
        }
    }
}
