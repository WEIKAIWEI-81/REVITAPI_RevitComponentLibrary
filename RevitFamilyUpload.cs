using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using System.Data.SqlClient;
using System.IO;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Security.Permissions;

namespace TRUEDREAMS
{
    public partial class Form2 : System.Windows.Forms.Form
    {
        SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
        string myfilename = null;

        // WindowsログオンAPI（ユーザー認証）
        [DllImport("advapi32.dll", SetLastError = true)]
        public static extern bool LogonUser(
           string lpszUsername,
           string lpszDomain,
           string lpszPassword,
           int dwLogonType,
           int dwLogonProvider,
           ref IntPtr phToken);

        // ハンドル解放API
        [DllImport("kernel32.dll")]
        public extern static bool CloseHandle(IntPtr hToken);

        public Form2()
        {
            InitializeComponent();

            // データベース接続設定
            builder.DataSource = "";
            builder.UserID = "";
            builder.Password = "";
            builder.InitialCatalog = "";

            // データベース接続と主要カテゴリ（comboBox1）読み込み
            SqlConnection connection = new SqlConnection(builder.ConnectionString);
            connection.Open();
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT * FROM main");
            String sql = sb.ToString();
            SqlCommand command = new SqlCommand(sql, connection);
            SqlDataReader reader = command.ExecuteReader();

            while (reader.Read())
            {
                comboBox1.Items.Add(reader.GetString(1).ToString());
            }
        }

        // [ファイル選択] ボタンクリックイベント
        private void Button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "ファイルを選択";
            dialog.InitialDirectory = ".\\";
            dialog.Filter = "rfaファイル (*.rfa)|*.rfa";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                textBox7.Text = dialog.FileName;
                myfilename = Path.GetFileName(dialog.FileName);
                dialog.Reset();
                dialog.Dispose();
            }
        }

        // [アップロード] ボタンクリックイベント
        private void Button3_Click(object sender, EventArgs e)
        {
            if (textBox7.Text != "" && comboBox1.Text != "請選擇" && comboBox2.Text != "請選擇" && comboBox3.Text != "請選擇")
            {
                // 初期値
                int whomade = 1, parameter = 1;
                if (radioButton1.Checked == true) whomade = 1;
                if (radioButton2.Checked == true) whomade = 2;
                if (radioButton3.Checked == true) whomade = 3;
                if (radioButton4.Checked == true) parameter = 1;
                if (radioButton5.Checked == true) parameter = 2;

                // データベース接続
                SqlConnection connection = new SqlConnection(builder.ConnectionString);
                connection.Open();
                StringBuilder sb = new StringBuilder();
                sb.Append("SELECT * FROM Persons WHERE fname = '" + myfilename + "'");
                SqlCommand command1 = new SqlCommand(sb.ToString(), connection);
                SqlDataReader reader = command1.ExecuteReader();

                int i = 0;
                while (reader.Read())
                {
                    i++;
                }

                if (i == 0)
                {
                    // 重複がなければ新規登録
                    reader.Close();
                    connection.Close();
                    connection.Open();

                    String sql = "INSERT INTO Persons (mname,subname,fname,rversion,whomade,parameter) values('" + comboBox1.Text.Trim() + "','" + comboBox2.Text.Trim() + "','" + myfilename + "','" + comboBox3.Text.Trim() + "','" + whomade.ToString() + "','" + parameter.ToString() + "')";
                    SqlCommand command2 = new SqlCommand(sql, connection);

                    if (command2.ExecuteNonQuery() == 1)
                    {
                        // ネットワークフォルダにファイルをコピー
                        string UserName = "dwe";
                        string MachineName = "192.168.51.132";
                        string Pw = "wra12345";
                        string category1 = comboBox1.Text;
                        string category2 = comboBox2.Text;
                        string IPath = @"\\" + MachineName + @"\family\" + category1 + @"\" + category2;
                        string sourcePath = textBox7.Text.ToString();
                        string targetPath = Path.Combine(IPath, myfilename);

                        const int LOGON32_PROVIDER_DEFAULT = 0;
                        const int LOGON32_LOGON_NEW_CREDENTIALS = 9;
                        IntPtr tokenHandle = new IntPtr(0);
                        tokenHandle = IntPtr.Zero;

                        // ユーザーとしてログイン
                        bool returnValue = LogonUser(UserName, "truedreams", Pw, LOGON32_LOGON_NEW_CREDENTIALS, LOGON32_PROVIDER_DEFAULT, ref tokenHandle);
                        WindowsIdentity w = new WindowsIdentity(tokenHandle);
                        w.Impersonate();

                        if (false == returnValue)
                        {
                            return;
                        }

                        Directory.CreateDirectory(IPath);
                        File.Copy(sourcePath, targetPath, true);
                        MessageBox.Show("ファミリファイルがアップロードされました！", "アップロード成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    // 重複ファイルエラー
                    MessageBox.Show("同じ名前のファミリがすでに存在します、新しい名前を入力してください。", "確認エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                // 入力漏れ警告
                MessageBox.Show("すべての情報を入力してください（キーワード以外の項目は必須です）。", "確認エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // comboBox1の選択変更イベント → comboBox2更新
        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            SqlConnection connection = new SqlConnection(builder.ConnectionString);
            connection.Open();
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT * FROM sub");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            SqlDataReader reader = command.ExecuteReader();

            comboBox2.Focus();
            comboBox2.Enabled = true;
            comboBox2.Items.Clear();
            comboBox2.Text = "請選擇";
            string comboboxchoiced = comboBox1.Text.ToString();

            while (reader.Read())
            {
                if (reader.GetString(0).Trim() == comboboxchoiced.Trim())
                {
                    comboBox2.Items.Add(reader.GetString(1).ToString());
                }
            }
        }

        // 初期化ボタン
        private void Button4_Click(object sender, EventArgs e)
        {
            InitializeComponent();
        }

      
    }
}
