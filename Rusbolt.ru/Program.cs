using System;

namespace Rusbolt.ru
{
    public class Program
    {
        static void Main(string[] args)
        {
            var rusbolt = new Rusbolt();
            var boltType = Console.ReadLine();
            rusbolt.Parse(boltType);

            Console.WriteLine("All is ok!");
            Console.ReadKey();
        }
    }
}
