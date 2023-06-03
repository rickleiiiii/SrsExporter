using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using TfsExporter;
using System.Runtime.CompilerServices;

namespace SrsExporter
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            SrsExporter se = new SrsExporter();
            await se.ListEpics();

            Console.ReadKey();
        }
    }
}
