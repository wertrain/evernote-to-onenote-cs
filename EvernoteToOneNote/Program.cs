using System;

namespace EvernoteToOneNote
{
    class Program
    {
        static void Main(string[] args)
        {
            var accessToken = args[0];
            OneNoteApi oneNote = new OneNoteApi(accessToken);

            var noteBookId = oneNote.CreateNotebook("Temp Notebook");
            var sectionId = oneNote.CreateSection(noteBookId, "Temp Section");
            oneNote.CreatePage(sectionId, new OneNoteApi.PageParameter(){
                Title = "Test title",
                Body = "this is test page!",
                DateTime = DateTime.Now
            });
        }
    }
}
