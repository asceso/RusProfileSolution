namespace DescriptionUpdater
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length > 0)
            {
                Updater updater = new Updater();
                updater.RunUpdater(args);
            }
        }
    }
}
