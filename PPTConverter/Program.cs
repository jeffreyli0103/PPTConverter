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
                StringBuilder body = new StringBuilder();
                StringBuilder title = new StringBuilder();
                StringBuilder verse = new StringBuilder();
                foreach (ISlide slide in ppt.Slides)
                {

                    foreach (IShape shape in slide.Shapes)
                    {
                        
                        Console.WriteLine($"slides title: {slide.Name}");
                        Console.WriteLine($"slides title: {slide.Title}");
                        Console.WriteLine($"slides title: {slide.SlideNumber}");
                        var posY = shape.Frame.Top + shape.Frame.Height;
                        var posX = shape.Frame.CenterX;

                        if (posY < 300 && slide.SlideNumber > 1 && !String.IsNullOrEmpty(getShapeText(shape))) //body
                            body.AppendLine(getShapeText(shape).Replace(slide.Title,""));
                        //if(posY > 300 && posX < 300 && slide.SlideNumber > 1 && !String.IsNullOrEmpty(getShapeText(shape))) //verse
                        //    verse.AppendLine(getShapeText(shape));
                        //if(posY > 300 && posX > 300 && slide.SlideNumber > 1 && !String.IsNullOrEmpty(getShapeText(shape))) //song name
                        //    title.AppendLine(getShapeText(shape));
                    }
                }
                Console.WriteLine($"body:{body.ToString()}");
                Console.WriteLine($"title:{title.ToString()}");
                Console.WriteLine($"verse:{verse.ToString()}");
                Clipboard.SetText(body.ToString());
                Console.WriteLine("Copied text to Clipboard");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return long_text;
        }
        public static string getShapeText(IShape shape)
        {
            StringBuilder sb = new StringBuilder();
            if (shape is IAutoShape)
            {
                IAutoShape ashape = shape as IAutoShape;
                if (ashape.TextFrame != null)
                {
                    foreach (TextParagraph pg in ashape.TextFrame.Paragraphs)
                    {
                        sb.AppendLine(pg.Text);
                    }
                }
            }
            else if (shape is GroupShape)
            {
                GroupShape gs = shape as GroupShape;
                foreach (IShape s in gs.Shapes)
                {
                    sb.AppendLine(getShapeText(s));
                }
            }
            return sb.ToString();
        }
    }
}
