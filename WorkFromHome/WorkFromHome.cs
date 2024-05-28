using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

using Outlook = Microsoft.Office.Interop.Outlook;


namespace WorkFromHome
{
    public partial class WorkFromHome : Form
    {
        Label labelStatus;
        private bool FlagCategories;

        public WorkFromHome()
        {
            InitializeComponent();

            // バージョンアップデート時に設定を引き継ぐ
            if (Properties.Settings.Default.UpgradeRequired)
            {
                Properties.Settings.Default.Upgrade();
                Properties.Settings.Default.UpgradeRequired = false;
                Properties.Settings.Default.Save();
            }

            // ここでラベルを生成し、フォームに追加します。
            labelStatus = new Label();
            labelStatus.Location = new Point(105, 198); // 適切な位置に設定してください。
            labelStatus.Size = new Size(100, 13); // 適切なサイズに設定してください。
            labelStatus.TextAlign = ContentAlignment.MiddleCenter; // テキストを中央揃えにする
            this.Controls.Add(labelStatus);            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // カレンダーで選択された日付を取得する
            DateTime selectedDate = monthCalendar1.SelectionStart;

            string flg_1 = "申請";

            // BackgroundWorkerのインスタンスを作成
            BackgroundWorker worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true; // 進行状況の報告を有効にする
            worker.DoWork += (s, ev) => { CreateEmail(selectedDate, flg_1, worker); }; // 非同期に実行する処理
            worker.ProgressChanged += (s, ev) =>
            {
                // 進行状況が報告されたときに行う処理                
                labelStatus.Text = $"作成中... {ev.ProgressPercentage}% 完了";
            };
            worker.RunWorkerCompleted += (s, ev) =>
            {
                // 非同期処理が完了したときに行う処理
                Application.Exit();
            };

            // 非同期処理を開始する
            worker.RunWorkerAsync();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // カレンダーで選択された日付を取得する
            DateTime selectedDate = monthCalendar1.SelectionStart;

            string flg_1 = "開始";

            // BackgroundWorkerのインスタンスを作成
            BackgroundWorker worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true; // 進行状況の報告を有効にする
            worker.DoWork += (s, ev) => { CreateEmail(selectedDate, flg_1, worker); }; // 非同期に実行する処理
            worker.ProgressChanged += (s, ev) =>
            {
                // 進行状況が報告されたときに行う処理
                labelStatus.Text = $"作成中... {ev.ProgressPercentage}% 完了";
            };
            worker.RunWorkerCompleted += (s, ev) =>
            {
                // 非同期処理が完了したときに行う処理
                Application.Exit();
            };

            // 非同期処理を開始する
            worker.RunWorkerAsync();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            // カレンダーで選択された日付を取得する
            DateTime selectedDate = monthCalendar1.SelectionStart;

            string flg_1 = "終了";

            // BackgroundWorkerのインスタンスを作成
            BackgroundWorker worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true; // 進行状況の報告を有効にする
            worker.DoWork += (s, ev) => { CreateEmail(selectedDate, flg_1, worker); }; // 非同期に実行する処理
            worker.ProgressChanged += (s, ev) =>
            {
                // 進行状況が報告されたときに行う処理
                labelStatus.Text = $"作成中... {ev.ProgressPercentage}% 完了";
            };
            worker.RunWorkerCompleted += (s, ev) =>
            {
                // 非同期処理が完了したときに行う処理
                Application.Exit();
            };

            // 非同期処理を開始する
            worker.RunWorkerAsync();

        }

        private void button4_Click(object sender, EventArgs e)
        {
            Form2 form2 = new Form2(); // Form2をインスタンス化
            form2.Show(); // Form2を表示
        }


        void CreateEmail(DateTime selectedDate, string flg_1, BackgroundWorker worker)
        {
            // フラグ読み取り
            FlagCategories = Properties.Settings.Default.FlagCategories;

            // 検索する日付
            string schedule = "";
            Dictionary<string, string> subjects = new Dictionary<string, string>();
            Dictionary<string, string> templates = new Dictionary<string, string>();

            subjects.Add("申請", selectedDate.ToString("M/d(ddd)") + " TWします_" + Properties.Settings.Default.Sender);
            templates.Add("申請", "{0}\n\n" +
                                        selectedDate.ToString("M/d(ddd)") + "はテレワーク業務の了承をお願いいたします。\n" +
                                        "業務内容は以下の通りです。\n" +
                                        "よろしくお願いします。\n\n" +
                                        "{1}\n" +
                                        "{2}");

            subjects.Add("開始", "TW開始します_" + Properties.Settings.Default.Sender);
            templates.Add("開始", "{0}\n\n" +
                                        "本日の業務を開始します。\n" +
                                        "業務内容は以下の通りです。\n" +
                                        "よろしくお願いします。\n\n" +
                                        "{1}\n" +
                                        "{2}");

            subjects.Add("終了", "TW終了します_" + Properties.Settings.Default.Sender);
            templates.Add("終了", "{0}\n\n" +
                                        "本日の業務を終了します。\n" +
                                        "業務内容は以下の通りです。\n" +
                                        "お疲れ様でした。\n\n" +
                                        "{1}\n" +
                                        "{2}");

         
            // ProgressBarを更新する
            System.Threading.Thread.Sleep(50); // 一時停止（模擬的な進行状況を作成）
            worker.ReportProgress(20); // 進行状況を報告


            // Outlookアプリケーションのオブジェクトを作成
            Outlook.Application outlookApp = new Outlook.Application();

            // Outlookのカレンダーを取得
            Outlook.MAPIFolder calendarFolder = outlookApp.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);

            // カレンダーから予定を検索
            Outlook.Items calendarItems = calendarFolder.Items;
            calendarItems.IncludeRecurrences = true;
            calendarItems.Sort("[Start]");

            // ProgressBarを更新する
            System.Threading.Thread.Sleep(50); // 一時停止（模擬的な進行状況を作成）
            worker.ReportProgress(50); // 進行状況を報告


            foreach (Outlook.AppointmentItem item in calendarItems)
            {
                if (item.Sensitivity != Outlook.OlSensitivity.olPrivate)
                {
                    DateTime start = item.Start;
                    DateTime end = item.End;

                    // 表示しないスケジュールの条件
                    // 開始時刻＝終了時刻、非公開の予定、終日設定予定
                    if (start.Date == selectedDate.Date && !item.AllDayEvent && start != end)
                    {

                        // スケジュールが見つかった場合、Outlookのメールアイテムを作成
                        if (FlagCategories)
                        {
                            string classification = item.Categories;
                            if (classification == null)
                            {
                                schedule += string.Format("{0} - {1}  {2}\n", start.ToString("HH:mm"), end.ToString("HH:mm"), item.Subject);
                            }
                            else
                            {
                                schedule += string.Format("{0} - {1}  {2}_{3}\n", start.ToString("HH:mm"), end.ToString("HH:mm"), classification, item.Subject);
                            }
                        }
                        else
                        {
                            schedule += string.Format("{0} - {1}  {2}\n", start.ToString("HH:mm"), end.ToString("HH:mm"), item.Subject);
                        }
                    }
                }
            }


            // ProgressBarを更新する
            System.Threading.Thread.Sleep(50); // 一時停止（模擬的な進行状況を作成）
            worker.ReportProgress(80); // 進行状況を報告

            //本文の連結 
            string body = string.Format(templates[flg_1], Properties.Settings.Default.Recipient, schedule, Properties.Settings.Default.Sender);


            // メールを作成
            Outlook.MailItem mailItem = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);
            mailItem.To = Properties.Settings.Default.AddressTo;
            mailItem.CC = Properties.Settings.Default.AddressCc;
            mailItem.Subject = subjects[flg_1];
            mailItem.Body = body;
            mailItem.Display(false);
        }

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }

        private void WorkFromHome_Load(object sender, EventArgs e)
        {

        }
    }
}

