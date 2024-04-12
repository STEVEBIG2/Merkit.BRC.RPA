using System;
using System.ComponentModel;
using System.IO;
using Newtonsoft.Json;

namespace UnitTest
{
    public class JsonRepo<T>
    {
        readonly string path;
        public JsonRepo(string path)
        {
            this.path = path;
        }

        public T Read()
        {
            return JsonConvert.DeserializeObject<T>(File.ReadAllText(path));
        }

        public T Write(T value)
        {
            File.WriteAllText(path, JsonConvert.SerializeObject(value, Formatting.Indented));
            return value;
        }

        public T Append(T value, string desc)
        {
            File.AppendAllLines(path, new string[] { String.Format("---------- Date: {0}, Description: {1} ----------", DateTime.Now.ToString("yyyy.MM.dd HH:mm.ss"), desc) });
            File.AppendAllText(path, JsonConvert.SerializeObject(value, Formatting.Indented));
            File.AppendAllLines(path, new string[] { "" });
            return value;
        }
    }
}
