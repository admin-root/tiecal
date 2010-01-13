using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using TieCal;

namespace Tester
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 1)
            {
                Console.Error.WriteLine("Usage: Tester.exe text-file");
                Environment.Exit(-1);
            }
            try
            {
                string file = args[0];
                Console.WriteLine("Opening {0} for reading", file);
                using (TextReader reader = new StreamReader(file))
                {
                    string line;
                    int i = 0;
                    while ((line = reader.ReadLine()) != null)
                    {
                        i++;
                        if (line.Trim().Length == 0)
                            continue;
                        var items = line.Split(';');
                        if (items.Length < 5)
                        {
                            Console.WriteLine("Line {0} doesn't contain occurrences ({1})", i, items.Length > 1 ? items[1] : "(no subject)");
                            continue;
                        }
                        List<DateTime> occurrences = new List<DateTime>();
                        for (int col = 5; col < items.Length; col++)
                        {
                            var dt = DateTime.Parse(items[col]);
                            occurrences.Add(dt);
                        }
                        try
                        {
                            var pattern = RepeatPattern.CreateFromOccurrences(occurrences);
                            Console.WriteLine("Test SUCCEEDED ({0})", items[1]);
                            Console.WriteLine("Repeat pattern created from {0} occurrences: {1}", occurrences.Count, pattern.ToString());
                        }
                        catch (ArgumentException)
                        {
                            Console.WriteLine("Test FAILED ({0})", items[1]);
                            Console.WriteLine("Failed to create repeat pattern from {0} occurrences", occurrences.Count);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("Test failed");
                Console.Error.WriteLine(ex);
                Environment.Exit(-1);
            }
            Console.WriteLine("All done");
        }
    }
}
