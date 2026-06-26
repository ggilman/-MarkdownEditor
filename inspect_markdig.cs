using System;
using System.Reflection;
using System.Linq;

class Program {
    static void Main() {
        var asm = Assembly.LoadFrom("bin/Debug/net9.0-windows/Markdig.dll");
        foreach(var t in asm.GetTypes().Where(t => t.Name.Contains("Task"))) {
            Console.WriteLine(t.FullName);
            foreach(var p in t.GetProperties()) Console.WriteLine("  " + p.Name);
        }
    }
}
