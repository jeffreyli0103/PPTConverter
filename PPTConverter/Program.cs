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
                List<string> paragraphList = new List<string>() {"verse","pre chorus", "chorus", "bridge"};
                int verse_counter = 1;
                String songName = "",paragraphName = "",verse = "";
                
                Dictionary<string,Dictionary<string,StringBuilder>> songList = new Dictionary<string, Dictionary<string, StringBuilder>>(){ };
                foreach (ISlide slide in ppt.Slides)
                {
                    if (slide.SlideNumber > 1)
                    {
                        Dictionary<string, StringBuilder> paragraph = new Dictionary<string, StringBuilder>() { };

                        paragraphName = $"{verse_counter}";
                        StringBuilder lyris = new StringBuilder();
                        foreach (IShape shape in slide.Shapes)
                        {
                            var count_paragraph = paragraphList.Where(p => getShapeText(shape).ToLower().Contains(p));
                            var posY = shape.Frame.Top + shape.Frame.Height;
                            //var posX = shape.Frame.CenterX;
                            if (posY < 140 && count_paragraph.Count() == 0 && !String.IsNullOrEmpty(getShapeText(shape).Replace("\r\n", "").Trim()))
                            {  //song name 
                                songName = getShapeText(shape).Replace("\r\n", "");
                                verse_counter = 1;
                            }
                            if (posY < 140 && count_paragraph.Count() > 0 && !String.IsNullOrEmpty(getShapeText(shape).Replace("\r\n", "").Trim()))
                            {  //paragraph name
                                verse = getShapeText(shape).Replace("\r\n", "");
                                lyris.AppendLine($";{verse}");
                            }
                            if (posY < 500  && posY > 140)  //body
                                lyris.AppendLine(getShapeText(shape).Replace(slide.Title, ""));

                        }
                        if (lyris.ToString().Trim() != "")
                        { 
                            paragraph.Add(verse_counter.ToString(), lyris);
                            if (songList.ContainsKey(songName))
                                songList[$"{songName}"].Add(paragraphName, lyris);
                            else
                                songList.Add(songName, paragraph);

                            verse_counter++;
                        }

                    }
                    
                }
                string sqlquery = "";
                foreach (var song in songList)
                {
                    sqlquery += "Insert into xxx ('";
                    foreach (var page in song.Value)
                    {
                        if (page.Value.ToString().Contains(";")) 
                            sqlquery += $"[{page.Value.ToString().Split(';')[1].Trim()}] \r\n {page.Value.ToString().Split(';')[0]}";
                        else
                            sqlquery += $"[{page.Key}] \r\n {page.Value}";
                    }
                    sqlquery += "');";
                }
                Clipboard.SetText(sqlquery);
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
