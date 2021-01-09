namespace DescriptionUpdater
{
    public static class StringExtractor
    {
        public static string GetInteger(this string target)
        {
            string result = string.Empty;
            foreach (char sym in target)
            {
                if (char.IsDigit(sym) || sym == '.')
                {
                    result += sym;
                }
            }
            return result;
        }
    }
}
