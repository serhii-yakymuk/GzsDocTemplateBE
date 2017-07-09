using Microsoft.Owin.Hosting;
using System;

namespace GzsDocTempalteBE
{
    public class Program
    {
        static void Main()
        {
            string baseAddress = "http://*:1433/";

            using (WebApp.Start<Startup>(url: baseAddress))
            {
                Console.WriteLine("Server started at " + baseAddress);
                Console.ReadLine();
            }
        }
    }
}
