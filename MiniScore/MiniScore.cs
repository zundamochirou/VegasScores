using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualBasic.FileIO;
using ScriptPortal.Vegas;

namespace MiniScore
{
    public class EntryPoint
    {
        private const string BASE_1 = @"I:\素材\BS003\メインテロップ\先攻攻撃中.png";
        private const string BASE_2 = @"I:\素材\BS003\メインテロップ\後攻攻撃中.png";

        private const string RUNNER_1_OFF = @"I:\素材\BS003\メインテロップ\ランナー\一塁OFF.png";
        private const string RUNNER_1_ON = @"I:\素材\BS003\メインテロップ\ランナー\一塁ON.png";
        private const string RUNNER_2_OFF = @"I:\素材\BS003\メインテロップ\ランナー\二塁OFF.png";
        private const string RUNNER_2_ON = @"I:\素材\BS003\メインテロップ\ランナー\二塁ON.png";
        private const string RUNNER_3_OFF = @"I:\素材\BS003\メインテロップ\ランナー\三塁OFF.png";
        private const string RUNNER_3_ON = @"I:\素材\BS003\メインテロップ\ランナー\三塁ON.png";

        private Vegas vegas = null;

        public void FromVegas(Vegas vegas)
        {
            this.vegas = vegas;

            string scoreFile = SelectFile();
            if (string.IsNullOrEmpty(scoreFile))
            {
                MessageBox.Show("ファイル未選択のため終了します");
                return;
            }

            ScoreFile sFile = new ScoreFile(scoreFile);
            sFile.Load();

            foreach (var se in sFile.EventList)
            {
                Debug.WriteLine(se.Time);
                CreateNewTrackEvents(se);
            }
        }


        /// <summary>
        /// スコアの記載によって新しいトラックイベント群を作成する
        /// </summary>
        /// <param name="se"></param>
        private void CreateNewTrackEvents(ScoreEvent se)
        {
            CreateBase(se);

            CreateTextTrack(se, "INNING", se.Inning);
            CreateTextTrack(se, "OUT", string.Format("{0}アウト", se.OutCount));
            CreateTextTrack(se, "TEAM1", se.TeamA);
            CreateTextTrack(se, "TEAM2", se.TeamB);
            CreateTextTrack(se, "SCORE1", se.ScoreA);
            CreateTextTrack(se, "SCORE2", se.ScoreB);

            CreateRunner(se);
        }

        private void CreateTextTrack(ScoreEvent se, string trackName, string text)
        {
            var track = FindTrack(trackName);
            var template = FindTemplate(track);
            var gen = GetGenerator();

            Media media = Media.CreateInstance(vegas.Project, gen);
            UpdateTextInfo(template.ActiveTake.Media, media, text);
            UpdateLocation(template.ActiveTake.Media, media);
            UpdateTextColor(template.ActiveTake.Media, media);
            UpdateScale(template.ActiveTake.Media, media);
            media.Generator.OFXEffect.AllParametersChanged();

            // MEMO は 表示時間をを 最大5秒にする
            var length = (trackName == "MEMO" && (se.LengthSeconds > 5)) ? Timecode.FromSeconds(5) : se.Length;
            VideoEvent ve = new VideoEvent(vegas.Project, se.StartTime, length, trackName);
            track.Events.Add(ve);
            ve.Takes.Add(new Take(media.GetVideoStreamByIndex(0)));

            // MEMO だけ 前後にフェード設定
            if (trackName == "MEMO")
            {
                ve.FadeIn.Length = Timecode.FromMilliseconds(500);
                ve.FadeOut.Length = Timecode.FromMilliseconds(500);
            }
        }

        /// <summary>
        /// ランナー
        /// </summary>
        /// <param name="se"></param>
        private void CreateRunner(ScoreEvent se)
        {
            CreateRunnderTrack(se, "RUNNER1", (se.OutCount < 3 && se.RunnerFirst) ? RUNNER_1_ON : RUNNER_1_OFF);
            CreateRunnderTrack(se, "RUNNER2", (se.OutCount < 3 && se.RunnerSecond) ? RUNNER_2_ON : RUNNER_2_OFF);
            CreateRunnderTrack(se, "RUNNER3", (se.OutCount < 3 && se.RunnerThird) ? RUNNER_3_ON : RUNNER_3_OFF);
        }

        private void CreateRunnderTrack(ScoreEvent se, string trackName, string mediaPath)
        {
            var t = FindTrack(trackName);
            VideoEvent ve = new VideoEvent(vegas.Project, se.StartTime, se.Length, trackName);
            t.Events.Add(ve);

            var m = Media.CreateInstance(vegas.Project, mediaPath);
            ve.Takes.Add(new Take(m.GetVideoStreamByIndex(0)));
        }

        /// <summary>
        /// Track名 "BASE"
        /// </summary>
        /// <param name="se"></param>
        private void CreateBase(ScoreEvent se)
        {
            string mediaPath = BASE_1;
            string trackName = "表";

            if (se.Inning.Contains("裏"))
            {
                mediaPath = BASE_2;
                trackName = "裏";
            }

            var take = new Take(Media.CreateInstance(vegas.Project, mediaPath).GetVideoStreamByIndex(0));

            var t = FindTrack("BASE");
            VideoEvent ve = new VideoEvent(vegas.Project, se.StartTime, se.Length, trackName);
            t.Events.Add(ve);
            ve.Takes.Add(take);
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

        public void UpdateTextColor(Media src, Media dst)
        {
            var srcParam = src.Generator.OFXEffect.FindParameterByName("TextColor") as OFXRGBAParameter;
            var dstParam = dst.Generator.OFXEffect.FindParameterByName("TextColor") as OFXRGBAParameter;

            dstParam.Value = srcParam.Value;
        }

        public void UpdateScale(Media src, Media dst)
        {
            var srcParam = src.Generator.OFXEffect.FindParameterByName("Scale") as OFXDoubleParameter;
            var dstParam = dst.Generator.OFXEffect.FindParameterByName("Scale") as OFXDoubleParameter;

            dstParam.Value = srcParam.Value;
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

        private string SelectFile()
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Title = "スコア情報CSVファイルを選択";

                var res = ofd.ShowDialog();
                if (res == DialogResult.OK)
                {
                    return ofd.FileName;
                }
            }
            return String.Empty;
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

                    if (cols.Length < 11)
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
        /// 先行スコア
        /// </summary>
        public string ScoreA;

        /// <summary>
        /// 後攻スコア
        /// </summary>
        public string ScoreB;

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
                ScoreA = cols[4],
                ScoreB = cols[5],
                OutCount = !string.IsNullOrEmpty(cols[6]) ? int.Parse(cols[6]) : 0,
                RunnerFirst = (cols[7] == "1"),
                RunnerSecond = (cols[8] == "1"),
                RunnerThird = (cols[9] == "1"),
                Note = cols[10]
            };

            return s;
        }

    }

}

