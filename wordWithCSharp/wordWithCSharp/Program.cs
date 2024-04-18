using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
namespace word
{
    class main
    {
        static void Main(string[] args)
        {
            object oMissing = System.Reflection.Missing.Value;
            object oEndOfDoc = "\\endofdoc"; 
            Word._Application oWord;
            Word._Document oDoc;
            oWord = new Word.Application();
            oWord.Visible = true;
            oDoc = oWord.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);
        }
    }
}