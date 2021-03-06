using System;
using RequirementsDocGenerator;
namespace RequirementsDocGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            if(args.Length < 2)
            {
                Console.WriteLine("Usage: Program <INPUT FILE> <OUTPUT FILE>");
                return;
            }
            var gen = new DocGenerator();
            Console.WriteLine($"Input: {args[0]}");
            Console.WriteLine($"Output: {args[1]}");
            gen.Generate(args[0], args[1]);
            Console.WriteLine("Document generated");
        }
    }
}
