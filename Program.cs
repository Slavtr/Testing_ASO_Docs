using ASO_Docs;
using Word = Microsoft.Office.Interop.Word;

namespace ASO_Test
{
    class Program
    {
        public static void Main(string[] args)
        {
            WarrantRecord tst = new WarrantRecord();
            tst.OpenDock(@"D:\Slava\Программы\Проекты\ASO_Docs\Testing.docx");
            List<string> list = new List<string>{ "Проверка", "Связи" };
            tst.SaveDock(tst.DoThings(list), @"D:\Slava\Testing.docx");
        }
    }
}