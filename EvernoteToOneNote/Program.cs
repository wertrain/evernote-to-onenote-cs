using System;

namespace EvernoteToOneNote
{
    class Program
    {
        static void Main(string[] args)
        {
            var accessToken = args[0];
            OneNoteApi oneNote = new OneNoteApi(accessToken);

            var id = oneNote.CreateNotebook("Temp Notebook");
            oneNote.CreatePage(id, "Test title", "this is test page!");
        }
    }
}
