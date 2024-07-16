using System.Text.Json;

namespace DocFiller
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            string templatePath = args[0];
            Console.WriteLine(templatePath);
            var json= JsonSerializer.Deserialize<Person>(args[1]);
            Console.WriteLine(json.name); 
        }
    }
    public class Person
    {
        public string name { get; set; }
        public int age { get; set; }
    }
}
