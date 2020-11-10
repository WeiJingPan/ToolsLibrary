using System;
using System.Text.RegularExpressions;

namespace Regular
{
    class Program
    {
        static void Main(string[] args)
        {
            var regex = new Regex("Name:(\\w+)\\s*Age:(\\d{1,3})");
            var matcher = regex.Match("Name:Aurora    Age:18");
            if (matcher.Success)
            {
                var group = matcher.Groups;
                var g1 = group[1];
                var g2 = group[2];
                Console.WriteLine("g1 = {0}", g1);
                Console.WriteLine("g2 = {0}", g2);
            }
        }
    }
}