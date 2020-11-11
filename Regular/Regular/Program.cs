using System;
using System.Collections;
using System.Linq;
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

            
            PrintStringSplit("二分，回溯，递归，分治".Split("[，；\\s]+"));
            PrintStringSplit("搜索，差找，旋转，遍历".Split("[，；\\s]+"));
            PrintStringSplit("数论，图论，逻辑，概率".Split("[，；\\s]+"));
            
            Console.WriteLine("二分,回溯,递归,分治".Replace("[,;\\s]+", ";"));
            Console.WriteLine("搜索;差找;旋转;遍历".Replace("[,;\\s]+", ";"));
            Console.WriteLine("数论 图论 逻辑 概率".Replace("[,;\\s]+", ";"));
        }

        private static void PrintStringSplit(string[] strs)
        {
            foreach (var str in strs)
            {
                Console.WriteLine(str);
            }
        }
    }
}
