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
                            Console.WriteLine("Powerpoint 2007");
                            ProcessTextFromPowerpoint(args[1]);
                        }
                        if (fi.Extension == ".ppt")
                        {
                            Console.WriteLine("Powerpoint 1997-2003");
                            ProcessTextFromPowerpoint(args[1]);
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
            string sqlquery = "";
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
                            //if (posY < 140 && count_paragraph.Count() > 0 && !String.IsNullOrEmpty(getShapeText(shape).Replace("\r\n", "").Trim()))
                            //{  //paragraph name
                            //    verse = getShapeText(shape).Replace("\r\n", "");
                            //    lyris.AppendLine($";{verse}");
                            //}
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
                
                foreach (var song in songList)
                {
                    sqlquery += $"Insert into items(title1,author,lastmodified,songnumber,folderno,oldfolderno,cjkwordcount,cjkstrokecount,formatdata,contents)VALUES('{song.Key}','',date('now'),0,6,0,'00','000{song.Key}','<ShowSongHeadings>0</ShowSongHeadings><ShowSongHeadingsAlign>0</ShowSongHeadingsAlign><UseShadowFont>1</UseShadowFont><ShowNotations>0</ShowNotations><CapoZero>0</CapoZero><UseOutlineFont>0</UseOutlineFont><DisplayRegions>2</DisplayRegions><DisplayRegionsLayout>0</DisplayRegionsLayout><ScreenColour1>-16777056</ScreenColour1><ScreenColour2>-16777056</ScreenColour2><ScreenPatternStyle>0</ScreenPatternStyle><BackgroundPicture /><BackgroundPictureMode>2</BackgroundPictureMode><VerticalAlign>1</VerticalAlign><ScreenLeftMargin>2</ScreenLeftMargin><ScreenRightMargin>2</ScreenRightMargin><ScreenBottomMargin>0</ScreenBottomMargin><ShowItemTransition>0</ShowItemTransition><ShowSlideTransition>0</ShowSlideTransition><FontVPosition1>0</FontVPosition1><FontVPosition2>50</FontVPosition2><MediaOption>0</MediaOption><MediaVolume>50</MediaVolume><MediaBalance>-1</MediaBalance><MediaMute>0</MediaMute><MediaRepeat>0</MediaRepeat><MediaWidescreen>0</MediaWidescreen><MediaCaptureDeviceNumber>1</MediaCaptureDeviceNumber><HeadingFontFormat>1</HeadingFontFormat><HeadingFontPercentSize>100</HeadingFontPercentSize><HeadingFontBold>0</HeadingFontBold><HeadingFontItalic>0</HeadingFontItalic><HeadingFontUnderline>0</HeadingFontUnderline><HeadingFontChorusItalic>0</HeadingFontChorusItalic><FontBold1>1</FontBold1><FontItalic1>0</FontItalic1><FontUnderline1>0</FontUnderline1><FontChorusBold1>0</FontChorusBold1><FontChorusItalic1>0</FontChorusItalic1><FontChorusUnderline1>0</FontChorusUnderline1><FontBold2>0</FontBold2><FontItalic2>0</FontItalic2><FontUnderline2>0</FontUnderline2><FontChorusBold2>0</FontChorusBold2><FontChorusItalic2>0</FontChorusItalic2><FontChorusUnderline2>0</FontChorusUnderline2><FontName1>Microsoft Sans Serif</FontName1><FontName2>Microsoft Sans Serif</FontName2><FontSize1>40</FontSize1><FontSize2>40</FontSize2><FontColour1>-1</FontColour1><FontColour2>-1</FontColour2><FontRTL1>0</FontRTL1><FontRTL2>0</FontRTL2><FontAlign1>2</FontAlign1><FontAlign2>2</FontAlign2><ShowDataPanel>0</ShowDataPanel><AutoTextOverflow>2</AutoTextOverflow><UseLargestFontSize>2</UseLargestFontSize><LineBetweenRegions>2</LineBetweenRegions><WordWrapLeftAlignIndent>2</WordWrapLeftAlignIndent>','";
                    foreach (var page in song.Value)
                    {
                        //if (page.Value.ToString().Contains(";")) 
                        //    sqlquery += $"[{page.Value.ToString().Split(';')[1].Trim()}] \r\n {page.Value.ToString().Split(';')[0].Replace("'","''")}";
                        //else
                        sqlquery += $"[{page.Key}] \r\n\r\n {page.Value.Replace("'", "''")}";
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
            return sqlquery;
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
