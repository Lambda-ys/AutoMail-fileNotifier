using System.IO;

namespace Sendmes
{
    public class Htmlsupport
    {
        public static string LoadTemplate(string templatePath)
        {
            return File.ReadAllText(templatePath);
        }

        public static string FillTemplate(string template, string content)
        {
            return template.Replace("{{FileList}}", content);
        }
    }
}