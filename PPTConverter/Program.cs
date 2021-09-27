using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using Spire.Presentation;
using System.Windows.Forms;



namespace PPTConverter
{
    class Program
    {
        [STAThreadAttribute]
        static void Main(string[] args)
        {
            Stardcmd();
        }
        public static void Stardcmd()
        {
            Console.Title = "PPTConvert" + Assembly.GetExecutingAssembly().GetName().Version.ToString();
            Console.WriteLine($"PPTConverter> ");
            string line = Console.ReadLine();
            while (RunCmdss(line))
            {
                Console.Write($"PPTConvert> ");
                line = Console.ReadLine();
            }
        }
        public static Boolean RunCmdss(string line)
        {
            if (line.ToLower().Contains("convert"))
            {
                try
                {
                    line = Regex.Replace(line, @"//.*", "");
                    line = Regex.Replace(line, @";", "");
                    line = Regex.Replace(line, @"\s+$", "");
                    line = Regex.Replace(line, @"^s+", "");
                    if (string.IsNullOrWhiteSpace(line))
                        return true;
                    List<string> arg = new List<string>();
                    var ms = Regex.Matches(line, @"([\w-=_:\\\\./]+|""([^""])+"")");
                    foreach (Match m in ms)
                    {
                        arg.Add(m.Value.Replace("\"", ""));
                    }
                    string[] args = arg.ToArray();
                    if (File.Exists(args[1]))
                    {
                        Console.WriteLine("Found file!");
                        FileInfo fi = new FileInfo(args[1]);
                        if (fi.Extension == ".pptx")
                        {
                            Console.WriteLine("It is a powerpoint!");
                            ProcessTextFromPowerpoint(args[1]);
                        }
                        if (fi.Extension == ".ppt")
                        {
                            Console.WriteLine("Powerpoint 1997-2003");
                        }
                    }
                    else
                    {
                        Console.WriteLine("File not exist!");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            else
            {
                Console.WriteLine("Cannot find command, please enter again");
            }
            return true;
        }

        public static string ProcessTextFromPowerpoint(string filePath)
        {
            string long_text = "";
            try
            {
                Presentation ppt = new Presentation();
                ppt.LoadFromFile(filePath);
                List<IShape> shapelist = new List<IShape>();
                List<IShape> bodylist = new List<IShape>();
                foreach (ISlide slide in ppt.Slides)
                {
                    foreach (IShape shape in slide.Shapes)
                    {
                        if (shape.Placeholder != null)
                        {
                            switch (shape.Placeholder.Type)
                            {
                                case PlaceholderType.Title:
                                    shapelist.Add(shape);
                                    break;
                                case PlaceholderType.CenteredTitle:
                                    shapelist.Add(shape);
                                    break;
                                case PlaceholderType.Subtitle:
                                    shapelist.Add(shape);
                                    break;
                                case PlaceholderType.Object:
                                    shapelist.Add(shape);
                                    break;
                                case PlaceholderType.None:
                                    shapelist.Add(shape);
                                    break;
                                case PlaceholderType.Body:
                                    shapelist.Add(shape);
                                    break;
                                case PlaceholderType.Media:
                                    shapelist.Add(shape);
                                    break;
                                case PlaceholderType.Table:
                                    shapelist.Add(shape);
                                    break;
                            }
                        }

                    }
                }
                Console.OutputEncoding = System.Text.Encoding.UTF8;
                for (int i = 0; i < shapelist.Count; i++)
                {
                    IAutoShape shape1 = shapelist[i] as IAutoShape;
                    long_text += shape1.TextFrame.Text;
                }
                Clipboard.SetText(long_text);
                Console.WriteLine("Copied text to Clipboard");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return long_text;
        }
    }
}
