using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WorkFromHome
{
    public partial class ScreenSetting : Form
    {
        private bool FlagCategories;

        public ScreenSetting()
        {
            InitializeComponent();
            LoadAddresses();
        }

        private void LoadAddresses()
        {
            // アドレスを表示する
            string[] addresses_To = Properties.Settings.Default.AddressTo.Split(';');
            string[] addresses_Cc = Properties.Settings.Default.AddressCc.Split(';');
            textBox_setting_sender.Text = Properties.Settings.Default.Sender;
            textBox_setting_recipient.Text = Properties.Settings.Default.Recipient;
            textBox_setting_To.Text = string.Join(Environment.NewLine, addresses_To);
            textBox_setting_Cc.Text = string.Join(Environment.NewLine, addresses_Cc);
            // チェックボックスの初期値を設定する
            checkBox_categories.Checked = Properties.Settings.Default.FlagCategories;



            // アドレスを変更できるようにする
            textBox_setting_To.Enabled = true;
            //textBox2.Enabled = true;
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            // ここに初期化処理などを記述する
        }

        private void textBox_addressTo_TextChanged(object sender, EventArgs e)
        {

        }

        private void button_save_tab1_Click(object sender, EventArgs e)
        {
            SaveSetteing();
        }


        private void button_save_tab2_Click(object sender, EventArgs e)
        {
            SaveSetteing();
        }

        private void SaveSetteing()
        {
            // テキストボックスのテキストをセミコロンで区切った文字列に変換する
            Properties.Settings.Default.Sender = textBox_setting_sender.Text;
            Properties.Settings.Default.Recipient = textBox_setting_recipient.Text;
            Properties.Settings.Default.AddressTo = textBox_setting_To.Text.Replace(Environment.NewLine, ";");
            Properties.Settings.Default.AddressCc = textBox_setting_Cc.Text.Replace(Environment.NewLine, ";");
            Properties.Settings.Default.FlagCategories = checkBox_categories.Checked;

            Properties.Settings.Default.Save();

            MessageBox.Show("設定が変更されました。");

            this.Hide();

        }

        private void checkBox_categories_CheckedChanged(object sender, EventArgs e)
        {
            FlagCategories = checkBox_categories.Checked;
        }

        
    }

}
