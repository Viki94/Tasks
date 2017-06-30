using System;
using System.Diagnostics;

namespace Task2
{
    class Factoriel
    {
        static void Main()
        {
            Stopwatch sw = Stopwatch.StartNew();
            var random = new Random();
            var number = 1;
            int result = 1;

            for (int i = 0; i < 1000; i++)
            {
                for (int j = 1; j <= number; j++)
                {
                    if (number == 0)
                    {
                        result = 1;
                        return;
                    }
                    result = result * j;
                }

                Console.WriteLine(number + "! = " + result);
                number = random.Next(0, 11);
                result = 1;
            }

            Console.WriteLine(sw.ElapsedMilliseconds + "ms");

            sw.Stop();
        }
    }
}
