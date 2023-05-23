using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

using Microsoft.Extensions.Configuration;

namespace SrsExporter
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var config = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json")
                .Build();

            var username = config["Username"];
            var password = config["Password"];
            var uri = config["Uri"];
            var collection = config["Collection"];
            var project = config["Project"];

            Console.ReadKey();
        }
    }
}
