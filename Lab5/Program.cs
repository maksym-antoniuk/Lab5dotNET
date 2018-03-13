using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lab5
{
    public class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Opening txt....");
            DocumentWorker documentWorker = new DocumentWorker("C:\\EnglishTest.txt");

            Console.WriteLine("Init Word document");
            documentWorker.Init();

            Console.WriteLine("Swapping marks to sentence");
            documentWorker.SwapText("@@Q1", "@@a1", "@@b1", "@@c1", "@@Q2", "@@a2", "@@b2", "@@c2", "@@Q3",
                "@@a3", "@@b3", "@@c3", "@@Q4", "@@a4", "@@b4", "@@c4", "@@Q5", "@@a5", "@@b5", "@@c5");

            Console.WriteLine("Are you ready to save file Result.docx on D disk?");
            Console.ReadKey();

            documentWorker.Save("D:", "Result");
        }
    }
}
