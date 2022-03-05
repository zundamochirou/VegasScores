using log4net;
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
        private const string SCORE_PATH = @"D:\Media\VegasPro\20220227\score.csv";

        private const string OUT_0 = @"I:\素材\BS004\アウト表示\0アウト.png";
        private const string OUT_1 = @"I:\素材\BS004\アウト表示\1アウト.png";
        private const string OUT_2 = @"I:\素材\BS004\アウト表示\2アウト.png";

        private static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        private Vegas vegas = null;

        public void FromVegas(Vegas vegas)
        {
            this.vegas = vegas;

            ScoreFile sFile = new ScoreFile(SCORE_PATH);
            sFile.Load();

            foreach (var se in sFile.EventList)
            {
                Debug.WriteLine(se.Time);
                CreateNewTrackEvents(se);
            }

            //CreateOut(1);

            //CreateInning("4表");
        }

        /// <summary>
        /// スコアの記載によって新しいトラックイベント群を作成する
        /// </summary>
        /// <param name="se"></param>
        private void CreateNewTrackEvents(ScoreEvent se)
        {
            CreateOut(se);
        }

        /// <summary>
        /// Track名 "OUT"
        /// </summary>
        /// <param name="se"></param>
        private void CreateOut(ScoreEvent se)
        {
            var take = SelectOutImageTake(se.OutCount);
            if ( take == null)
            {
                log.InfoFormat("3アウトのため非表示とします({0:H:mm:ss}})", se.Time);
                return;     
            }

            var t = FindTrack("OUT");
            VideoEvent ve = new VideoEvent(vegas.Project, se.StartTime, se.Length, string.Format("{0}アウト", se.OutCount));
            t.Events.Add(ve);
            ve.Takes.Add(take);
        }

        /// <summary>
        /// アウトカウントのTakeを取得します。
        /// </summary>
        /// <param name="outCount"></param>
        /// <returns>Takeオブジェクト。3カウントの場合はNull</returns>
        private Take SelectOutImageTake(int outCount)
        {
            string mediaPath = string.Empty;

            if (outCount == 0)
            {
                mediaPath = OUT_0;
            }
            else if (outCount == 1)
            {
                mediaPath = OUT_1;
            }
            else if (outCount == 2)
            {
                mediaPath = OUT_2;
            }
            else
            {
                return null;
            }

            var m = Media.CreateInstance(vegas.Project, mediaPath);
            return new Take(m.GetVideoStreamByIndex(0));
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
                        var ts = se.Time - beforeEvent.Time;
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
        public DateTime Time;

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
                return Timecode.FromString(string.Format("{0:HH:mm:ss};00", this.Time));
            }
        }

        public Timecode Length
        {
            get
            {
                return Timecode.FromSeconds(LengthSeconds);
            }
        }

        public static ScoreEvent Create(string[] cols)
        {
            var s = new ScoreEvent
            {
                Time = DateTime.ParseExact(cols[0], "H:mm:ss", DateTimeFormatInfo.InvariantInfo),
                TeamA = cols[1],
                TeamB = cols[2],
                Inning = cols[3],
                Score = cols[4],
                OutCount = !string.IsNullOrEmpty(cols[5]) ? int.Parse(cols[5]) : 0,
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

