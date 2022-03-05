using Microsoft.VisualBasic.FileIO;
using ScriptPortal.Vegas;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ScoreSample
{
    /// <summary>
    /// 
    /// 
    /// </summary>
    public class EntryPoint
    {
        private Vegas vegas = null;

        public void FromVegas(Vegas vegas)
        {
            this.vegas = vegas;

            CreateInning("4表");

        }



        private void CreateOut(int count)
        {
            var t = FindTrack("OUT");


        }

        private void CreateInning(string text)
        {
            var t = FindTrack("INNING");
            var te = FindTemplate(t);

            var gen = GetGenerator();
            Media media = Media.CreateInstance(vegas.Project, gen);
            var ofx = media.Generator.OFXEffect;

            UpdateTextInfo(te.ActiveTake.Media, media, text);
            UpdateLocation(te.ActiveTake.Media, media);

            //var nx = ofx.FindParameterByName("Text") as OFXStringParameter;
            //RichTextBox b = new RichTextBox();
            //b.Rtf = nx.Value;
            //b.Text = text;
            //nx.Value = b.Rtf;
            ofx.AllParametersChanged();

            Timecode startTimeCode = Timecode.FromString("00:00:20;00");
            VideoEvent ve = new VideoEvent(vegas.Project, startTimeCode, Timecode.FromString("00:00:10;00"), Guid.NewGuid().ToString());
            t.Events.Add(ve);
            Take tk = new Take(media.GetVideoStreamByIndex(0));
            ve.Takes.Add(tk);

            //dst.Name = Guid.NewGuid().ToString();
            //dst.ActiveTake.Name = Guid.NewGuid().ToString();
            //dst.Length = Timecode.FromString("00:00:10;00");

            //var dst = te.Copy(t, startTimeCode);

            //OFXStringParameter x = dst.ActiveTake.Media.Generator.OFXEffect.FindParameterByName("Text") as OFXStringParameter;

            //RichTextBox b = new RichTextBox();
            //b.Rtf = x.Value;
            //b.Text = text;

            //var nx = ofx.FindParameterByName("Text") as OFXStringParameter;
            //nx.Value = b.Rtf;
            //ofx.AllParametersChanged();

            //dst.AddTake(media.GetVideoStreamByIndex(0));

        }

        public void UpdateTextInfo(Media src, Media dst, string text)
        {
            // TEXT
            var srcParam = src.Generator.OFXEffect.FindParameterByName("Text") as OFXStringParameter;
            var dstParam = dst.Generator.OFXEffect.FindParameterByName("Text") as OFXStringParameter;

            RichTextBox b = new RichTextBox();
            b.Rtf = srcParam.Value;
            FontFamily fm = b.SelectionFont.FontFamily;
            b.Text = text;
            b.SelectAll();
            b.SelectionFont = new Font(fm, b.Font.Size);

            dstParam.Value = b.Rtf;
        }

        public void UpdateLocation(Media src, Media dst)
        {
            var srcParam = src.Generator.OFXEffect.FindParameterByName("Location") as OFXDouble2DParameter;
            var dstParam = dst.Generator.OFXEffect.FindParameterByName("Location") as OFXDouble2DParameter;

            OFXDouble2D location;
            location.X = srcParam.Value.X;
            location.Y = srcParam.Value.Y;

            dstParam.Value = location;
        }


        private Track FindTrack(string name)
        {
            return vegas.Project.Tracks.First(t => t.Name == name);
        }

        private TrackEvent FindTemplate(Track t)
        {
            return t.Events.First(te => te.ActiveTake.Name == "Template");
        }

        private PlugInNode GetGenerator()
        {
            return vegas.Generators.GetChildByUniqueID("{Svfx:com.vegascreativesoftware:titlesandtext}");
        }

    }

    public class ScoreFile
    {
        public string FilePath { get; private set; }

        public ScoreFile(string filePath)
        {
            this.FilePath = filePath;
        }

        private List<ScoreEvent> events = new List<ScoreEvent>();

        public IReadOnlyList<ScoreEvent> EventList
        {
            get { return events.AsReadOnly(); }
        }

        public void Load()
        {
            bool firstRow = true;

            using (TextFieldParser p = new TextFieldParser(FilePath, Encoding.UTF8))
            {
                p.TextFieldType = FieldType.Delimited;
                p.SetDelimiters(",");
                p.TrimWhiteSpace = true;

                ScoreEvent beforeEvent = null;

                while (!p.EndOfData)
                {
                    string[] cols = p.ReadFields();
                    if (firstRow)
                    {
                        firstRow = false;
                        continue;
                    }

                    if (cols.Length < 10)
                    {
                        Debug.WriteLine("カラム数が規定未満のためスキップします 行[{0}],カラム数[{1}]", p.LineNumber, cols.Length);
                        continue;
                    }

                    var se = ScoreEvent.Create(cols);
                    events.Add(se);

                    // 直前のイベントに対して時間を設定する
                    if (beforeEvent != null)
                    {
                        var currentTime = DateTime.ParseExact(se.Time, "HH:mm:ss", DateTimeFormatInfo.InvariantInfo);
                        var beforeTime = DateTime.ParseExact(beforeEvent.Time, "HH:mm:ss", DateTimeFormatInfo.InvariantInfo);

                        var ts = currentTime - beforeTime;
                        beforeEvent.LengthSeconds = ts.TotalSeconds;
                    }
                    beforeEvent = se;
                }
            }
        }

    }


    public class ScoreEvent
    {
        private ScoreEvent()
        {
        }

        /// <summary>
        /// イベント時間 (00:00:00 表記)
        /// </summary>
        public string Time;

        /// <summary>
        /// 先攻
        /// </summary>
        public string TeamA;

        /// <summary>
        /// 後攻
        /// </summary>
        public string TeamB;

        /// <summary>
        /// イニング表示
        /// </summary>
        public string Inning;

        /// <summary>
        /// スコア表示
        /// </summary>
        public string Score;

        /// <summary>
        /// アウト
        /// </summary>
        public int OutCount;

        /// <summary>
        /// 1塁走者
        /// </summary>
        public bool RunnerFirst;

        /// <summary>
        /// 2塁走者
        /// </summary>
        public bool RunnerSecond;

        /// <summary>
        /// 3塁走者
        /// </summary>
        public bool RunnerThird;

        /// <summary>
        /// メモ
        /// </summary>
        public string Note;

        /// <summary>
        /// イベントの長さ(秒)
        /// </summary>
        public double LengthSeconds { get; set; } = 10;

        /// <summary>
        /// イベント開始時間
        /// </summary>
        public Timecode StartTime
        {
            get
            {
                return Timecode.FromString(string.Format("{0};00", this.Time));
            }
        }

        public static ScoreEvent Create(string[] cols)
        {
            var s = new ScoreEvent
            {
                Time = cols[0],
                TeamA = cols[1],
                TeamB = cols[2],
                Inning = cols[3],
                Score = cols[4],
                OutCount = int.Parse(cols[5]),
                RunnerFirst = (cols[6] == "1"),
                RunnerSecond = (cols[7] == "1"),
                RunnerThird = (cols[8] == "1"),
                Note = cols[9]
            };

            return s;
        }

    }

    public class ScoreEntry
    {


    }
}

