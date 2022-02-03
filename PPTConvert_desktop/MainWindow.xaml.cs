using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.IO;
using Microsoft.Win32;
using Spire.Presentation;
using System.Data.SQLite;
using System.IO.IsolatedStorage;
using System.Data;

namespace PPTConvert_desktop
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            CachedPath("db_dir.txt", isoStore);
            CachedPath("ppt_dir.txt", isoStore);
            CachedPath("lib_dir.txt", isoStore);
        }

        public string db_dir;
        public string ppt_dir;
        public string lib_dir;
        public string format = "lower";
        IsolatedStorageFile isoStore = IsolatedStorageFile.GetStore(IsolatedStorageScope.User | IsolatedStorageScope.Assembly, null, null);


        #region Select file
        private void selectDB(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Database files|*.db";
            if(File.Exists(db_dir))
            {
                openFileDialog.InitialDirectory = Path.GetDirectoryName(db_dir);
            }
            if (openFileDialog.ShowDialog() == true) { 
                db_dir = openFileDialog.FileName;
                db_dir_tb.Text = db_dir;
                //store path in txt
                createNewPath("db_dir.txt", db_dir);
            }
        }

        private void selectPPT(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (File.Exists(ppt_dir))
            {
                openFileDialog.InitialDirectory = Path.GetDirectoryName(ppt_dir);
            }
            openFileDialog.Filter = "PowerPoint Presentations|*.ppt;*.pptx";
            if (openFileDialog.ShowDialog() == true) {
                ppt_dir = openFileDialog.FileName;
                ppt_dir_label.Content = ppt_dir;
                //store path in txt
                createNewPath("ppt_dir.txt", ppt_dir);
            }
        }


        private void selectLib(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (File.Exists(lib_dir))
            {
                openFileDialog.InitialDirectory = Path.GetDirectoryName(lib_dir);
            }
            openFileDialog.Filter = "Database files|*.db";
            if (openFileDialog.ShowDialog() == true)
            {
                lib_dir = openFileDialog.FileName;
                lib_dir_tb.Text = lib_dir;
                //store path in txt
                createNewPath("lib_dir.txt", lib_dir);
            }

        }
        #endregion

        private void createNewPath(string fileName,string filePath)
        {
            using (IsolatedStorageFileStream isoStream = new IsolatedStorageFileStream(fileName, FileMode.Create, isoStore))
            {
                using (StreamWriter writer = new StreamWriter(isoStream))
                {
                    writer.WriteLine(filePath);
                }
            }
        }

        private void importToDB(object sender, RoutedEventArgs e)
        {
            if (File.Exists(db_dir) && File.Exists(ppt_dir) && (format == "lower" || format == "upper"))
            {
                ProcessTextFromPowerpoint(ppt_dir, format, db_dir);
            }
            else consoleLog("File not found, please try again.");
        }
        public void consoleLog(string text)
        {
            console.Text += text + "\r\n";
        }

        public void ProcessTextFromPowerpoint(string filePath, string format, string db)
        {
            try
            {
                Presentation ppt = new Presentation();
                ppt.LoadFromFile(filePath);
                List<string> paragraphList = new List<string>() { "verse", "pre chorus", "chorus", "bridge" };
                int verse_counter = 1;
                String songName = "", paragraphName = "", verse = "";
                //default upper
                int AXIS_TOP_Y = 140;
                int AXIS_BOTTOM_Y = 500;
                int AXIS_BOTTOM_X = 1000;
                if (format == "lower")
                {
                    AXIS_TOP_Y = 350;
                    AXIS_BOTTOM_Y = 400;
                    AXIS_BOTTOM_X = 300;
                }

                Dictionary<string, Dictionary<string, StringBuilder>> songList = new Dictionary<string, Dictionary<string, StringBuilder>>() { };
                foreach (ISlide slide in ppt.Slides)
                {
                    if (slide.SlideNumber > 1)
                    {
                        Dictionary<string, StringBuilder> paragraph = new Dictionary<string, StringBuilder>() { };

                        paragraphName = $"{verse_counter}";
                        StringBuilder lyris = new StringBuilder();
                        foreach (IShape shape in slide.Shapes)
                        {
                            //slides size
                            var slideHeight = slide.Presentation.SlideSize.Size.Height;
                            var slideWidth = slide.Presentation.SlideSize.Size.Width;
                            //shape size
                            var frameHeight = shape.Height;
                            var frameWidth = shape.Width;
                            //shape pos
                            var posY = shape.Frame.Top + shape.Frame.Height;
                            var posX = shape.Frame.Left + shape.Frame.Width;

                            var count_paragraph = paragraphList.Where(p => getShapeText(shape).ToLower().Contains(p));
                            var content = getShapeText(shape);
                            /*
                             * Lower song name indicator
                             */
                            if
                            (
                                format == "lower" &&
                                posY > AXIS_TOP_Y && // y axis condition
                                count_paragraph.Count() == 0 && //not a paragraph indicator
                                !String.IsNullOrEmpty(content.Replace("\r\n", "").Trim()) && // not empty
                                posX < AXIS_BOTTOM_X   // x axis condition
                            )
                            {
                                songName = content.Replace("\r\n", "").Trim();//song name 
                                verse_counter = 1;
                            }
                            /*
                             * Upper song name indicator
                             */
                            else if
                            (
                                 format == "upper" &&
                                 posY < AXIS_TOP_Y && // y axis condition
                                 count_paragraph.Count() == 0 && //not a paragraph indicator
                                 !String.IsNullOrEmpty(content.Replace("\r\n", "")) // not empty
                            )
                            {
                                songName = content.Replace("\r\n", "").Trim();//song name 
                                verse_counter = 1;
                            }
                            //if (posY < 140 && count_paragraph.Count() > 0 && !String.IsNullOrEmpty(getShapeText(shape).Replace("\r\n", "").Trim()))
                            //{  //paragraph name
                            //    verse = getShapeText(shape).Replace("\r\n", "");
                            //    lyris.AppendLine($";{verse}");
                            //}

                            if (frameWidth > 300 && !String.IsNullOrEmpty(content) && count_paragraph.Count() == 0)
                                lyris.AppendLine(content);//body

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
                consoleLog("start cleaning db");
                executeSQL("Delete from items", db);
                foreach (var song in songList)
                {
                    var libRecords = selectQuery(song.Key, lib_dir);
                    var sqlquery = "";
                    if (!string.IsNullOrEmpty(libRecords))
                    {
                        sqlquery = $"Insert into items VALUES " + libRecords;
                    }
                    else { 
                        sqlquery +=
                            $"Insert into items(title1,author,lastmodified,songnumber,folderno," +
                            $"oldfolderno,cjkwordcount,cjkstrokecount,formatdata,contents)VALUES('" +
                            $"{song.Key}','',date('now'),0,6,0,'00','000{song.Key}','<ShowSongHeadings>0</ShowSongHeadings>" +
                            $"<ShowSongHeadingsAlign>0</ShowSongHeadingsAlign><UseShadowFont>1</UseShadowFont>" +
                            $"<ShowNotations>0</ShowNotations><CapoZero>0</CapoZero><UseOutlineFont>0</UseOutlineFont>" +
                            $"<DisplayRegions>2</DisplayRegions><DisplayRegionsLayout>0</DisplayRegionsLayout>" +
                            $"<ScreenColour1>-16777056</ScreenColour1><ScreenColour2>-16777056</ScreenColour2>" +
                            $"<ScreenPatternStyle>0</ScreenPatternStyle><BackgroundPicture />" +
                            $"<BackgroundPictureMode>2</BackgroundPictureMode><VerticalAlign>1</VerticalAlign>" +
                            $"<ScreenLeftMargin>2</ScreenLeftMargin><ScreenRightMargin>2</ScreenRightMargin>" +
                            $"<ScreenBottomMargin>0</ScreenBottomMargin><ShowItemTransition>0</ShowItemTransition>" +
                            $"<ShowSlideTransition>0</ShowSlideTransition><FontVPosition1>0</FontVPosition1>" +
                            $"<FontVPosition2>50</FontVPosition2><MediaOption>0</MediaOption><MediaVolume>50</MediaVolume>" +
                            $"<MediaBalance>-1</MediaBalance><MediaMute>0</MediaMute><MediaRepeat>0</MediaRepeat>" +
                            $"<MediaWidescreen>0</MediaWidescreen><MediaCaptureDeviceNumber>1</MediaCaptureDeviceNumber>" +
                            $"<HeadingFontFormat>1</HeadingFontFormat><HeadingFontPercentSize>100</HeadingFontPercentSize>" +
                            $"<HeadingFontBold>0</HeadingFontBold><HeadingFontItalic>0</HeadingFontItalic>" +
                            $"<HeadingFontUnderline>0</HeadingFontUnderline><HeadingFontChorusItalic>0</HeadingFontChorusItalic>" +
                            $"<FontBold1>1</FontBold1><FontItalic1>0</FontItalic1><FontUnderline1>0</FontUnderline1>" +
                            $"<FontChorusBold1>0</FontChorusBold1><FontChorusItalic1>0</FontChorusItalic1>" +
                            $"<FontChorusUnderline1>0</FontChorusUnderline1><FontBold2>0</FontBold2><FontItalic2>0</FontItalic2>" +
                            $"<FontUnderline2>0</FontUnderline2><FontChorusBold2>0</FontChorusBold2><FontChorusItalic2>0</FontChorusItalic2>" +
                            $"<FontChorusUnderline2>0</FontChorusUnderline2><FontName1>Microsoft Sans Serif</FontName1>" +
                            $"<FontName2>Microsoft Sans Serif</FontName2><FontSize1>40</FontSize1><FontSize2>40</FontSize2>" +
                            $"<FontColour1>-1</FontColour1><FontColour2>-1</FontColour2><FontRTL1>0</FontRTL1><FontRTL2>0</FontRTL2>" +
                            $"<FontAlign1>2</FontAlign1><FontAlign2>2</FontAlign2><ShowDataPanel>0</ShowDataPanel>" +
                            $"<AutoTextOverflow>2</AutoTextOverflow><UseLargestFontSize>2</UseLargestFontSize>" +
                            $"<LineBetweenRegions>2</LineBetweenRegions><WordWrapLeftAlignIndent>2</WordWrapLeftAlignIndent>','";
                        foreach (var page in song.Value)
                        {
                            //if (page.Value.ToString().Contains(";")) 
                            //    sqlquery += $"[{page.Value.ToString().Split(';')[1].Trim()}] \r\n {page.Value.ToString().Split(';')[0].Replace("'","''")}";
                            //else
                            sqlquery += $"[{page.Key}] \r\n{page.Value.Replace("'", "''")}";
                        }
                        sqlquery += "');";
                    }
                    executeSQL(sqlquery, db);
                    consoleLog($"Inserted song:{song.Key} into database");
                    //TODO: use song.Key to search lib
                }
                
            }
            catch (Exception ex)
            {
                consoleLog(ex.Message);
            }
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

        public void executeSQL(string sqlquery, string db)
        {
            try
            {
                FileInfo fi = new FileInfo(db);
                if (fi.Extension == ".db")
                {
                    string cs = $"URI=file:{db}";

                    var con = new SQLiteConnection(cs);
                    con.Open();
                    var cmd = new SQLiteCommand(con);

                    cmd.CommandText = sqlquery;
                    cmd.ExecuteNonQuery();
                    con.Close();
                    consoleLog("Success!");
                }
                else
                {
                    consoleLog("Invalid database path, please attach the correct path.");
                }
            }
            catch (Exception ex)
            {
                consoleLog($"Something wrong. Exception msg: {ex.Message}");
            }
        }
        private string selectQuery(string song,string db)
        {
            string data = "";
            try
            {
                FileInfo fi = new FileInfo(db);
                if (fi.Extension == ".db")
                {
                    string cs = $"URI=file:{db}";
                    var con = new SQLiteConnection(cs);
                    con.Open();
                    var sqlquery = $"SELECT * FROM items where title1 ='{song}'";
                    var cmd = new SQLiteCommand(con);
                    cmd.CommandText = sqlquery;
                    SQLiteDataReader reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            data = "(";
                            for(int i = 0; i < reader.FieldCount; i++)
                            {
                                data += $"'{reader.GetValue(i)}'";
                                if (i != reader.FieldCount - 1) {
                                    data += ",";
                                }
                            }
                            data += ");";
                        }
                    }
                    con.Close();
                }
                else
                {
                    consoleLog("Invalid database path, please attach the correct path.");
                }
            }
            catch (Exception ex)
            {
                consoleLog($"Something wrong. Exception msg: {ex.Message}");
            }
            return data;
        }

        private void upper_radio_Checked(object sender, RoutedEventArgs e)
        {
            format = "upper";
        }

        private void lower_radio_Checked(object sender, RoutedEventArgs e)
        {
            format = "lower";
        }
        private void CachedPath(string fileName, IsolatedStorageFile isoStore)
        {
            if (isoStore.FileExists(fileName))
            {
                using (IsolatedStorageFileStream isoStream = new IsolatedStorageFileStream(fileName, FileMode.Open, isoStore))
                {
                    using (StreamReader reader = new StreamReader(isoStream))
                    {
                        if (fileName.Contains("ppt_dir"))
                        {
                            ppt_dir = reader.ReadToEnd().Replace("\r\n","");
                            ppt_dir_label.Content = ppt_dir;
                        }
                        if (fileName.Contains("db_dir"))
                        {
                            db_dir = reader.ReadToEnd().Replace("\r\n", ""); 
                            db_dir_tb.Text = db_dir;
                        }
                        if (fileName.Contains("lib_dir"))
                        {
                            lib_dir = reader.ReadToEnd().Replace("\r\n", "");
                            lib_dir_tb.Text = lib_dir;
                        }
                    }
                }
            }
        }


        private void updateLib(object sender, RoutedEventArgs e)
        {
            consoleLog("feature is coming soon!");
        }

    }
}
