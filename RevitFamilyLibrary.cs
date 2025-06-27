using System;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System.Data.SqlClient;
using System.IO;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Security.Principal;

namespace TRUEDREAMS
{
    // サムネイル取得オプション（フラグ列挙型）
    [Flags]
    public enum ThumbnailOptions
    {
        None = 0x00,                 // オプションなし
        BiggerSizeOk = 0x01,        // より大きなサイズが許容される
        InMemoryOnly = 0x02,        // メモリ内のみに限定
        IconOnly = 0x04,            // アイコンのみ取得
        ThumbnailOnly = 0x08,       // サムネイルのみ取得
        InCacheOnly = 0x10          // キャッシュ内のみに限定
    }

    public partial class Form1 : System.Windows.Forms.Form
    {
        private UIApplication uiapp;
        private UIDocument uidoc;
        private Document doc;
        private ExternalCommandData commandData;
        SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();

        // Windowsユーザーでログオンするための外部関数（advapi32.dll）
        [DllImport("advapi32.dll", SetLastError = true)]
        public static extern bool LogonUser(
           string lpszUsername,
           string lpszDomain,
           string lpszPassword,
           int dwLogonType,
           int dwLogonProvider,
           ref IntPtr phToken);

        // トークンのクローズ用（kernel32.dll）
        [DllImport("kernel32.dll")]
        public extern static bool CloseHandle(IntPtr hToken);

        public Form1(ExternalCommandData commandData1)
        {
            InitializeComponent();
            commandData = commandData1;
            uiapp = commandData.Application;
            uidoc = uiapp.ActiveUIDocument;
            doc = uidoc.Document;           

            // SQL接続設定
            builder.DataSource = ""; // サーバー名
            builder.UserID = "";                             // ユーザー名
            builder.Password = "";                          // パスワード
            builder.InitialCatalog = "";            // データベース名

            tabControl1.Width = 1060;
            tabControl1.Height = 600;

            // メインカテゴリデータの取得と ComboBox1 への追加
            SqlConnection connection = new SqlConnection(builder.ConnectionString);
            connection.Open();
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT * FROM main");
            String sql = sb.ToString();
            SqlCommand command = new SqlCommand(sql, connection);
            SqlDataReader reader = command.ExecuteReader();

            while (reader.Read())
            {
                comboBox1.Items.Add(reader.GetString(1).ToString()); // 第二列の文字列を追加
            }
            connection.Close();
        }

        // comboBox1選択変更時に comboBox2 の内容を更新
        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            SqlConnection connection = new SqlConnection(builder.ConnectionString);
            connection.Open();
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT * FROM sub");
            String sql = sb.ToString();
            SqlCommand command = new SqlCommand(sql, connection);
            SqlDataReader reader = command.ExecuteReader();

            comboBox2.Focus();
            comboBox2.Enabled = true;
            comboBox2.Items.Clear();
            comboBox2.Text = "請選擇"; // 「選択してください」
            string comboboxchoiced = comboBox1.Text.ToString();

            while (reader.Read())
            {                
                if (reader.GetString(0).ToString().Trim() == comboboxchoiced.Trim())
                {
                    comboBox2.Items.Add(reader.GetString(1).ToString());
                }
            }
        }

        // ボタン2（Form2表示）
        private void Button2_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2();
            f2.Visible = true;
        }

        // ボタン4（Form5表示）
        private void Button4_Click(object sender, EventArgs e)
        {
            Form5 f3 = new Form5();
            f3.Visible = true;
        }

        // Revitファミリ読み込み時のオプション処理
        private class FamilyLoadOptions : IFamilyLoadOptions
        {
            public bool OnFamilyFound(bool familyInUse, out bool overwriteParameterValues)
            {
                overwriteParameterValues = true;
                return true;
            }

            public bool OnSharedFamilyFound(Family sharedFamily, bool familyInUse, out FamilySource source, out bool overwriteParameterValues)
            {
                source = FamilySource.Family;
                overwriteParameterValues = true;
                return true;
            }
        }
       
      
               // Windows のサムネイルを取得するためのクラス
        public class WindowsThumbnailProvider
        {
            // IShellItem2 の GUID（サムネイル取得に必要）
            private const string IShellItem2Guid = "7E9FB0D3-919F-4307-AB2E-9B1860310C93";

            // ファイルパスから IShellItem を生成する WinAPI
            [DllImport("shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
            internal static extern int SHCreateItemFromParsingName(
                [MarshalAs(UnmanagedType.LPWStr)] string path,
                IntPtr pbc, // バインドコンテキスト（未使用）
                ref Guid riid,
                [MarshalAs(UnmanagedType.Interface)] out IShellItem shellItem);

            // HBITMAP を解放する WinAPI
            [DllImport("gdi32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool DeleteObject(IntPtr hObject);

            // IShellItem インターフェース定義
            [ComImport]
            [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
            [Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe")]
            internal interface IShellItem
            {
                void BindToHandler(IntPtr pbc, [MarshalAs(UnmanagedType.LPStruct)]Guid bhid, [MarshalAs(UnmanagedType.LPStruct)]Guid riid, out IntPtr ppv);
                void GetParent(out IShellItem ppsi);
                void GetDisplayName(SIGDN sigdnName, out IntPtr ppszName);
                void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
                void Compare(IShellItem psi, uint hint, out int piOrder);
            };

            // パス表示方法の列挙型
            internal enum SIGDN : uint
            {
                NORMALDISPLAY = 0,
                PARENTRELATIVEPARSING = 0x80018001,
                PARENTRELATIVEFORADDRESSBAR = 0x8001c001,
                DESKTOPABSOLUTEPARSING = 0x80028000,
                PARENTRELATIVEEDITING = 0x80031001,
                DESKTOPABSOLUTEEDITING = 0x8004c000,
                FILESYSPATH = 0x80058000,
                URL = 0x80068000
            }

            // HRESULT 結果定義
            internal enum HResult
            {
                Ok = 0x0000,
                False = 0x0001,
                InvalidArguments = unchecked((int)0x80070057),
                OutOfMemory = unchecked((int)0x8007000E),
                NoInterface = unchecked((int)0x80004002),
                Fail = unchecked((int)0x80004005),
                ElementNotFound = unchecked((int)0x80070490),
                TypeElementNotFound = unchecked((int)0x8002802B),
                NoObject = unchecked((int)0x800401E5),
                Win32ErrorCanceled = 1223,
                Canceled = unchecked((int)0x800704C7),
                ResourceInUse = unchecked((int)0x800700AA),
                AccessDenied = unchecked((int)0x80030005)
            }

            // サムネイル取得用のファクトリインターフェース
            [ComImportAttribute()]
            [GuidAttribute("bcc18b79-ba16-442f-80c4-8a59c30c463b")]
            [InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
            internal interface IShellItemImageFactory
            {
                [PreserveSig]
                HResult GetImage(
                    [In, MarshalAs(UnmanagedType.Struct)] NativeSize size,
                    [In] ThumbnailOptions flags,
                    [Out] out IntPtr phbm);
            }

            // 画像サイズ構造体
            [StructLayout(LayoutKind.Sequential)]
            internal struct NativeSize
            {
                private int width;
                private int height;
                public int Width { set { width = value; } }
                public int Height { set { height = value; } }
            };

            // RGB カラー構造体（未使用）
            [StructLayout(LayoutKind.Sequential)]
            public struct RGBQUAD
            {
                public byte rgbBlue;
                public byte rgbGreen;
                public byte rgbRed;
                public byte rgbReserved;
            }

            // ファイルからサムネイル画像を取得
            public static Bitmap GetThumbnail(string fileName, int width, int height, ThumbnailOptions options)
            {
                IntPtr hBitmap = GetHBitmap(Path.GetFullPath(fileName), width, height, options);
                try
                {
                    return GetBitmapFromHBitmap(hBitmap);
                }
                finally
                {
                    DeleteObject(hBitmap); // メモリリーク防止
                }
            }

            // HBITMAP から .NET Bitmap に変換
            public static Bitmap GetBitmapFromHBitmap(IntPtr nativeHBitmap)
            {
                Bitmap bmp = Bitmap.FromHbitmap(nativeHBitmap);
                if (Bitmap.GetPixelFormatSize(bmp.PixelFormat) < 32)
                    return bmp;
                return CreateAlphaBitmap(bmp, PixelFormat.Format32bppArgb);
            }

            // 透過情報を持つ Bitmap を作成
            public static Bitmap CreateAlphaBitmap(Bitmap srcBitmap, PixelFormat targetPixelFormat)
            {
                Bitmap result = new Bitmap(srcBitmap.Width, srcBitmap.Height, targetPixelFormat);
                System.Drawing.Rectangle bmpBounds = new System.Drawing.Rectangle(0, 0, srcBitmap.Width, srcBitmap.Height);
                BitmapData srcData = srcBitmap.LockBits(bmpBounds, ImageLockMode.ReadOnly, srcBitmap.PixelFormat);

                bool isAlplaBitmap = false;

                try
                {
                    for (int y = 0; y <= srcData.Height - 1; y++)
                    {
                        for (int x = 0; x <= srcData.Width - 1; x++)
                        {
                            System.Drawing.Color pixelColor = System.Drawing.Color.FromArgb(
                                Marshal.ReadInt32(srcData.Scan0, (srcData.Stride * y) + (4 * x)));
                            if (pixelColor.A > 0 & pixelColor.A < 255)
                            {
                                isAlplaBitmap = true;
                            }
                            result.SetPixel(x, y, pixelColor);
                        }
                    }
                }
                finally
                {
                    srcBitmap.UnlockBits(srcData);
                }

                if (isAlplaBitmap)
                {
                    return result;
                }
                else
                {
                    return srcBitmap;
                }
            }

            // 実際に Windows API 経由でサムネイル HBITMAP を取得
            private static IntPtr GetHBitmap(string fileName, int width, int height, ThumbnailOptions options)
            {
                IShellItem nativeShellItem;
                Guid shellItem2Guid = new Guid(IShellItem2Guid);
                int retCode = SHCreateItemFromParsingName(fileName, IntPtr.Zero, ref shellItem2Guid, out nativeShellItem);
                if (retCode != 0)
                    throw Marshal.GetExceptionForHR(retCode);

                NativeSize nativeSize = new NativeSize();
                nativeSize.Width = width;
                nativeSize.Height = height;

                IntPtr hBitmap;
                HResult hr = ((IShellItemImageFactory)nativeShellItem).GetImage(nativeSize, options, out hBitmap);

                Marshal.ReleaseComObject(nativeShellItem);

                if (hr == HResult.Ok) return hBitmap;

                throw Marshal.GetExceptionForHR((int)hr);
            }            
        }
        

        protected void btn_Click(Object sender, EventArgs e)
{
    // 進度條の初期設定
    progressBar1.Minimum = 0;
    progressBar1.Maximum = 3;
    progressBar1.Value = 0;
    progressBar1.Step = 1;
    statushow.Text = "     初期化中...";

    Button temp = (Button)sender;

    Transaction trans = new Transaction(doc, "ExComm");

    // ログイン用ユーザー情報
    string UserName = "";
    string MachineName = "";
    string Pw = "";

    string category1 = comboBox1.Text.Trim();
    string category2 = comboBox2.Text.Trim();
    string fid = temp.Name.Trim('b', 't', 'n');

    const int LOGON32_PROVIDER_DEFAULT = 0;
    const int LOGON32_LOGON_NEW_CREDENTIALS = 9;
    IntPtr tokenHandle = IntPtr.Zero;

    // Windowsユーザーとしてログインする
    bool returnValue = LogonUser(UserName, "", Pw, LOGON32_LOGON_NEW_CREDENTIALS, LOGON32_PROVIDER_DEFAULT, ref tokenHandle);
    WindowsIdentity w = new WindowsIdentity(tokenHandle);
    w.Impersonate();

    if (!returnValue) return;

    progressBar1.Value++;
    statushow.Text = "     リモートフォルダへ接続中...";

    SqlConnection connection = new SqlConnection(builder.ConnectionString);
    connection.Open();

    StringBuilder sb = new StringBuilder();
    sb.Append("SELECT * FROM Persons where fid='" + fid + "'");
    SqlCommand command = new SqlCommand(sb.ToString(), connection);
    SqlDataReader reader = command.ExecuteReader();
    reader.Read();

    string IPath = @"\\" + MachineName + @"\family\" + reader.GetString(1) + @"\" + reader.GetString(2);

    progressBar1.Value++;
    statushow.Text = "     データ取得中...";
    string filename = reader.GetString(3);
    string filepath = IPath + @"\" + filename;
    filepath.Replace(" ", "");

    trans.Start();
    Family fs = null;
    doc.LoadFamily(filepath, new FamilyLoadOptions(), out fs);
    temp.Enabled = false;
    temp.Text = "プロジェクトに読込済み";
    statushow.Text = "     ファミリがプロジェクトに読込まれました";
    progressBar1.Value++;
    trans.Commit();
    connection.Close();
}

private void Button3_Click(object sender, EventArgs e)
{
    statushow.Text = "検索中...お待ちください";
    progressBar1.Visible = true;
    statushow.Visible = true;
    tabControl1.TabPages.Clear();

    SqlConnection connection = new SqlConnection(builder.ConnectionString);
    connection.Open();
    StringBuilder sb = new StringBuilder();

    // 条件によって SQL を組み立てる
    if (radioButton1.Checked && radioButton5.Checked)
    {
        sb.Append("SELECT * FROM Persons where FNAME LIKE '%" + textBox1.Text + "%'");
    }
    else if (radioButton1.Checked && radioButton6.Checked)
    {
        sb.Append("SELECT * FROM Persons where FNAME LIKE '%" + textBox1.Text + "%' AND PARAMETER=1");
    }
    else if (radioButton2.Checked && radioButton6.Checked)
    {
        sb.Append("SELECT * FROM Persons where FNAME LIKE '%" + textBox1.Text + "%' AND WHOMADE=1 AND PARAMETER=1");
    }
    else if (radioButton3.Checked && radioButton6.Checked)
    {
        sb.Append("SELECT * FROM Persons where FNAME LIKE '%" + textBox1.Text + "%' AND WHOMADE=2 AND PARAMETER=1");
    }
    else if (radioButton4.Checked && radioButton6.Checked)
    {
        sb.Append("SELECT * FROM Persons where FNAME LIKE '%" + textBox1.Text + "%' AND WHOMADE=3 AND PARAMETER=1");
    }
    else if (radioButton2.Checked && radioButton5.Checked)
    {
        sb.Append("SELECT * FROM Persons where FNAME LIKE '%" + textBox1.Text + "%' AND WHOMADE=1");
    }
    else if (radioButton3.Checked && radioButton5.Checked)
    {
        sb.Append("SELECT * FROM Persons where FNAME LIKE '%" + textBox1.Text + "%' AND WHOMADE=2");
    }
    else if (radioButton4.Checked && radioButton5.Checked)
    {
        sb.Append("SELECT * FROM Persons where FNAME LIKE '%" + textBox1.Text + "%' AND WHOMADE=3");
    }

    string sql = sb.ToString();
    SqlCommand command = new SqlCommand(sql, connection);
    SqlDataReader reader = command.ExecuteReader();

    if (reader.HasRows)
    {
        if (textBox1.Text != "")
        {
            // ログイン処理
            string UserName = "dwe";
            string MachineName = "192.168.50.101";
            string Pw = "wra12345";

            const int LOGON32_PROVIDER_DEFAULT = 0;
            const int LOGON32_LOGON_NEW_CREDENTIALS = 9;
            IntPtr tokenHandle = IntPtr.Zero;

            bool returnValue = LogonUser(UserName, "truedreams", Pw, LOGON32_LOGON_NEW_CREDENTIALS, LOGON32_PROVIDER_DEFAULT, ref tokenHandle);
            WindowsIdentity w = new WindowsIdentity(tokenHandle);
            w.Impersonate();
            if (!returnValue) return;

            // データ数を数える
            int amount = 0;
            while (reader.Read()) amount++;
            reader.Close();

            SqlConnection connection2 = new SqlConnection(builder.ConnectionString);
            connection2.Open();
            SqlCommand command2 = new SqlCommand(sql, connection2);
            SqlDataReader reader2 = command2.ExecuteReader();

            PictureBox[] pic = new PictureBox[amount];
            Label[] lab = new Label[amount * 4];
            Button[] btn = new Button[amount];
            TabPage[] tab = new TabPage[amount];
            int i = 0, j = 0, k = 0;

            // 進度條初期化
            progressBar1.Minimum = 0;
            progressBar1.Maximum = amount;
            progressBar1.Value = 0;
            progressBar1.Step = 1;

            while (reader2.Read())
            {
                string IPath = @"\\" + MachineName + @"\family\" + reader2.GetString(1) + @"\" + reader2.GetString(2);
                string filename = reader2.GetString(3);
                string filepath = Path.Combine(IPath, filename);

                if (i % 10 == 0 || i == 0)
                {
                    tab[j] = new TabPage { Text = (j + 1).ToString(), BackColor = Color.White, AutoScroll = true };
                    tabControl1.TabPages.Add(tab[j]);
                    i = 0;
                }

                int[] x = { 25, 230, 435, 640, 845, 25, 230, 435, 640, 845 };
                int[] y_pic = { 20, 20, 20, 20, 20, 330, 330, 330, 330, 330 };
                int[] y_lab = { 195, 195, 195, 195, 195, 505, 505, 505, 505, 505 };
                int[] y_btn = { 280, 280, 280, 280, 280, 590, 590, 590, 590, 590 };

                // サムネイル
                pic[j * 10 + i] = new PictureBox { Height = 175, Width = 200, Location = new Point(x[i], y_pic[i]), Image = WindowsThumbnailProvider.GetThumbnail(filepath, 150, 150, ThumbnailOptions.None) };
                tab[j].Controls.Add(pic[j * 10 + i]);

                // ファミリ名表示
                lab[k] = new Label
                {
                    Text = reader2.GetString(3).TrimEnd('.', 'r', 'f', 'a'),
                    Font = new Font("微軟正黑體", 8),
                    Height = 50,
                    Width = 180,
                    Location = new Point(x[i], y_lab[i]),
                    AutoSize = false
                };
                tab[j].Controls.Add(lab[k++]);

                // 出處情報
                string[] originLabels = { "晨禎", "網路", "官方" };
                int originIndex = reader2.GetInt32(5) - 1;
                if (originIndex >= 0)
                {
                    lab[k] = new Label
                    {
                        Text = originLabels[originIndex],
                        Font = new Font("微軟正黑體", 8),
                        Height = 20,
                        Width = 50,
                        Location = new Point(x[i], y_lab[i] + 65),
                        TextAlign = ContentAlignment.MiddleCenter,
                        BackColor = Color.LightSkyBlue,
                        AutoSize = false
                    };
                    tab[j].Controls.Add(lab[k++]);
                }

                // パラメトリック情報
                string paramText = reader2.GetInt32(6) == 1 ? "可參數化" : "無參數化";
                lab[k] = new Label
                {
                    Text = paramText,
                    Font = new Font("微軟正黑體", 8),
                    Height = 20,
                    Width = 80,
                    Location = new Point(x[i] + 50, y_lab[i] + 65),
                    TextAlign = ContentAlignment.MiddleCenter,
                    BackColor = Color.Khaki,
                    AutoSize = false
                };
                tab[j].Controls.Add(lab[k++]);

                // 使用頻率
                lab[k] = new Label
                {
                    Text = reader2.GetInt32(4).ToString(),
                    Font = new Font("微軟正黑體", 8),
                    Height = 20,
                    Width = 50,
                    Location = new Point(x[i] + 130, y_lab[i] + 65),
                    TextAlign = ContentAlignment.MiddleCenter,
                    BackColor = Color.PaleGreen,
                    AutoSize = false
                };
                tab[j].Controls.Add(lab[k++]);

                // ファミリ読込ボタン
                btn[j * 10 + i] = new Button
                {
                    Text = "プロジェクトに読込",
                    Font = new Font("微軟正黑體", 10),
                    Height = 30,
                    Width = 180,
                    Name = "btn" + reader2.GetInt32(0),
                    Location = new Point(x[i], y_btn[i])
                };
                btn[j * 10 + i].Click += new EventHandler(btn_Click);
                tab[j].Controls.Add(btn[j * 10 + i]);

                i++;
                if (i == 10) j++;
                progressBar1.Value++;
            }
            connection.Close();
            connection2.Close();
            statushow.Text = "     ファミリライブラリの表示完了";
        }
        else
        {
            MessageBox.Show("一致するファミリが見つかりませんでした！", "確認", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
    else
    {
        MessageBox.Show("検索キーワードを入力してください！", "確認", MessageBoxButtons.OK, MessageBoxIcon.Error);
    }
}
       

        private void Button1_Click_2(object sender, EventArgs e)
        {
            statushow.Text = "お待ちください... ファミリーデータベースを検索中";
			progressBar1.Visible = true;
			statushow.Visible = true;
			tabControl1.TabPages.Clear();

			// SQL接続を開始
			SqlConnection connection = new SqlConnection(builder.ConnectionString);
			connection.Open();
			StringBuilder sb = new StringBuilder();

			// ラジオボタンとコンボボックスの状態に応じてSQL文を構築
			if (comboBox2.Text == "請選擇" && radioButton1.Checked && radioButton5.Checked)
			{
				sb.Append("SELECT * FROM Persons where MNAME='" + comboBox1.Text.ToString() + "'");
			}
			else if (comboBox2.Text == "請選擇" && radioButton2.Checked && radioButton5.Checked)
			{
				sb.Append("SELECT * FROM Persons where MNAME='" + comboBox1.Text.ToString() + "' AND whomade=1");
			}
			else if (comboBox2.Text == "請選擇" && radioButton3.Checked && radioButton5.Checked)
			{
				sb.Append("SELECT * FROM Persons where MNAME='" + comboBox1.Text.ToString() + "' AND whomade=2");
			}
			else if (comboBox2.Text == "請選擇" && radioButton4.Checked && radioButton5.Checked)
			{
				sb.Append("SELECT * FROM Persons where MNAME='" + comboBox1.Text.ToString() + "' AND whomade=3");
			}
			else if (comboBox2.Text == "請選擇" && radioButton2.Checked && radioButton6.Checked)
			{
				sb.Append("SELECT * FROM Persons where MNAME='" + comboBox1.Text.ToString() + "' AND whomade=1 AND parameter=1");
			}
			else if (comboBox2.Text == "請選擇" && radioButton3.Checked && radioButton6.Checked)
			{
				sb.Append("SELECT * FROM Persons where MNAME='" + comboBox1.Text.ToString() + "' AND whomade=2 AND parameter=1");
			}
			else if (comboBox2.Text == "請選擇" && radioButton4.Checked && radioButton6.Checked)
			{
				sb.Append("SELECT * FROM Persons where MNAME='" + comboBox1.Text.ToString() + "' AND whomade=3 AND parameter=1");
			}
			else if (comboBox2.Text == "請選擇" && radioButton1.Checked && radioButton6.Checked)
			{
				sb.Append("SELECT * FROM Persons where MNAME='" + comboBox1.Text.ToString() + "' AND parameter=1");
			}
			else if (radioButton1.Checked && radioButton5.Checked)
			{
				sb.Append("SELECT * FROM Persons where MNAME='" + comboBox1.Text.ToString() + "' AND SUBNAME='" + comboBox2.Text.ToString() + "'");
			}
			else if (radioButton1.Checked && radioButton6.Checked)
			{
				sb.Append("SELECT * FROM Persons where MNAME='" + comboBox1.Text.ToString() + "' AND SUBNAME='" + comboBox2.Text.ToString() + "' AND PARAMETER=1");
			}
			else if (radioButton2.Checked && radioButton6.Checked)
			{
				sb.Append("SELECT * FROM Persons where MNAME='" + comboBox1.Text.ToString() + "' AND SUBNAME='" + comboBox2.Text.ToString() + "'AND WHOMADE=1 AND PARAMETER=1");
			}
			else if (radioButton3.Checked && radioButton6.Checked)
			{
				sb.Append("SELECT * FROM Persons where MNAME='" + comboBox1.Text.ToString() + "' AND SUBNAME='" + comboBox2.Text.ToString() + "'AND WHOMADE=2 AND PARAMETER=1");
			}
			else if (radioButton4.Checked && radioButton6.Checked)
			{
				sb.Append("SELECT * FROM Persons where MNAME='" + comboBox1.Text.ToString() + "' AND SUBNAME='" + comboBox2.Text.ToString() + "'AND WHOMADE=3 AND PARAMETER=1");
			}
			else if (radioButton2.Checked && radioButton5.Checked)
			{
				sb.Append("SELECT * FROM Persons where MNAME='" + comboBox1.Text.ToString() + "' AND SUBNAME='" + comboBox2.Text.ToString() + "'AND WHOMADE=1");
			}
			else if (radioButton3.Checked && radioButton5.Checked)
			{
				sb.Append("SELECT * FROM Persons where MNAME='" + comboBox1.Text.ToString() + "' AND SUBNAME='" + comboBox2.Text.ToString() + "'AND WHOMADE=2");
			}
			else if (radioButton4.Checked && radioButton5.Checked)
			{
				sb.Append("SELECT * FROM Persons where MNAME='" + comboBox1.Text.ToString() + "' AND SUBNAME='" + comboBox2.Text.ToString() + "'AND WHOMADE=3");
			}
           
		      // クエリをSQLコマンドとして実行
                String sql = sb.ToString();
                SqlCommand command = new SqlCommand(sql, connection);
                SqlDataReader reader = command.ExecuteReader();

                // 結果が存在するか確認
                if (reader.HasRows == true)
                {
                    // 遠隔サーバーへのログイン設定
                    string UserName = "dwe";
                    string MachineName = "192.168.50.101";
                    string Pw = "wra12345";
                    string category1 = comboBox1.Text.Trim();
                    string category2 = comboBox2.Text.Trim();

                    const int LOGON32_PROVIDER_DEFAULT = 0;
                    const int LOGON32_LOGON_NEW_CREDENTIALS = 9;
                    IntPtr tokenHandle = new IntPtr(0);
                    tokenHandle = IntPtr.Zero;

                    // トークンを取得してユーザーの偽装を行う
                    bool returnValue = LogonUser(UserName, "truedreams", Pw, LOGON32_LOGON_NEW_CREDENTIALS, LOGON32_PROVIDER_DEFAULT, ref tokenHandle);

                    // ユーザー偽装の実行
                    WindowsIdentity w = new WindowsIdentity(tokenHandle);
                    w.Impersonate();
                    if (false == returnValue)
                    {
                        // ログイン失敗時の処理
                        return;
                    }

                    // 結果件数のカウント
                    int amount = 0;
                    while (reader.Read())
                    {
                        amount++;
                    }
                    reader.Close();

                    // 2つ目のDB接続を開始
                    SqlConnection connection2 = new SqlConnection(builder.ConnectionString);
                    connection2.Open();
                    SqlCommand command2 = new SqlCommand(sql, connection2);
                    SqlDataReader reader2 = command2.ExecuteReader();

                    // UI要素の初期化
                    PictureBox[] pic = new PictureBox[amount];
                    Label[] lab = new Label[amount * 4];
                    Button[] btn = new Button[amount];
                    TabPage[] tab = new TabPage[amount];
                    int i = 0;
                    int j = 0;
                    int k = 0;

                    // プログレスバーの初期設定
                    progressBar1.Minimum = 0;
                    progressBar1.Maximum = amount;
                    progressBar1.Value = 0;
                    progressBar1.Step = 1;

                    // 取得結果を処理してUIを生成
                    while (reader2.Read())
                    {
                        string IPath = @"\\" + MachineName + @"\family\" + reader2.GetString(1) + @"\" + reader2.GetString(2);
                        string filename = reader2.GetString(3);

                        // 新しいタブページの生成
                        if (i % 10 == 0 || i == 0)
                        {
                            tab[j] = new TabPage();
                            tabControl1.TabPages.Add(tab[j]);
                            tab[j].Text = (j + 1).ToString();
                            i = 0;
                        }

                        // 表示位置の配列定義
                        int[] x = { 25, 230, 435, 640, 845, 25, 230, 435, 640, 845 };
                        int[] y_pic = { 20, 20, 20, 20, 20, 330, 330, 330, 330, 330 };

                        // サムネイル画像用のPictureBox生成
                        pic[j * 10 + i] = new PictureBox();
                        pic[j * 10 + i].Height = 175;
                        pic[j * 10 + i].Width = 200;
                        pic[j * 10 + i].Location = new System.Drawing.Point(x[i], y_pic[i]);

                        // Tabに配置
                        tab[j].Controls.Add(pic[j * 10 + i]);
                        tab[j].UseVisualStyleBackColor = false;
                        tab[j].BackColor = System.Drawing.Color.White;
                        tab[j].AutoScroll = true;

                        // ファイルパスを構築
                        string filepath = Path.Combine(IPath, filename);

                        // サムネイル画像を取得して設定
                        int THUMB_SIZE = 150;
                        Bitmap thumbnail = WindowsThumbnailProvider.GetThumbnail(
                           filepath, THUMB_SIZE, THUMB_SIZE, ThumbnailOptions.None);

                        Image image = thumbnail;
                        pic[j * 10 + i].Image = image;



						/* データベースからファミリ情報を取得して、各UI要素を生成する */
						lab[k] = new Label();
						char[] mychar = { '.', 'r', 'f', 'a' };
						lab[k].Text = reader2.GetString(3).ToString().TrimEnd(mychar); // 拡張子を除去したファイル名をラベルに設定
						lab[k].Font = new Font("メイリオ", 8);
						lab[k].Height = 50;
						lab[k].Width = 180;
						lab[k].Location = new System.Drawing.Point(x[i], y_lab[i]);
						lab[k].AutoSize = false;
						tab[j].Controls.Add(lab[k]);
						k++;

						// 製作者を判定して、ラベルに表示
						if (reader2.GetInt32(5) == 1)
						{
							lab[k] = new Label();
							lab[k].Text = "晨禎"; // 社内
							lab[k].Font = new Font("メイリオ", 8);
							lab[k].TextAlign = ContentAlignment.MiddleCenter;
							lab[k].Height = 20;
							lab[k].Width = 50;
							lab[k].Location = new System.Drawing.Point(x[i], y_lab[i] + 65);
							lab[k].BackColor = System.Drawing.Color.LightSkyBlue;
							lab[k].AutoSize = false;
							tab[j].Controls.Add(lab[k]);
							k++;
						}
						else if (reader2.GetInt32(5) == 2)
						{
							lab[k] = new Label();
							lab[k].Text = "ネット"; // オンライン
							lab[k].Font = new Font("メイリオ", 8);
							lab[k].TextAlign = ContentAlignment.MiddleCenter;
							lab[k].Height = 20;
							lab[k].Width = 50;
							lab[k].Location = new System.Drawing.Point(x[i], y_lab[i] + 65);
							lab[k].BackColor = System.Drawing.Color.LightSkyBlue;
							lab[k].AutoSize = false;
							tab[j].Controls.Add(lab[k]);
							k++;
						}
						else if (reader2.GetInt32(5) == 3)
						{
							lab[k] = new Label();
							lab[k].Text = "公式"; // 公式
							lab[k].Font = new Font("メイリオ", 8);
							lab[k].TextAlign = ContentAlignment.MiddleCenter;
							lab[k].Height = 20;
							lab[k].Width = 50;
							lab[k].Location = new System.Drawing.Point(x[i], y_lab[i] + 65);
							lab[k].BackColor = System.Drawing.Color.LightSkyBlue;
							lab[k].AutoSize = false;
							tab[j].Controls.Add(lab[k]);
							k++;
						}

						// パラメータ有無の判定
						if (reader2.GetInt32(6) == 1)
						{
							lab[k] = new Label();
							lab[k].Text = "パラメータ有";
							lab[k].Font = new Font("メイリオ", 8);
							lab[k].TextAlign = ContentAlignment.MiddleCenter;
							lab[k].Height = 20;
							lab[k].Width = 80;
							lab[k].Location = new System.Drawing.Point(x[i] + 50, y_lab[i] + 65);
							lab[k].BackColor = System.Drawing.Color.Khaki;
							lab[k].AutoSize = false;
							tab[j].Controls.Add(lab[k]);
							k++;
						}
						else if (reader2.GetInt32(6) == 2)
						{
							lab[k] = new Label();
							lab[k].Text = "パラメータ無";
							lab[k].Font = new Font("メイリオ", 8);
							lab[k].TextAlign = ContentAlignment.MiddleCenter;
							lab[k].Height = 20;
							lab[k].Width = 80;
							lab[k].Location = new System.Drawing.Point(x[i] + 50, y_lab[i] + 65);
							lab[k].BackColor = System.Drawing.Color.Khaki;
							lab[k].AutoSize = false;
							tab[j].Controls.Add(lab[k]);
							k++;
						}

						// 使用頻度を表示
						lab[k] = new Label();
						lab[k].Text = reader2.GetInt32(4).ToString(); // 使用回数？
						lab[k].Font = new Font("メイリオ", 8);
						lab[k].TextAlign = ContentAlignment.MiddleCenter;
						lab[k].Height = 20;
						lab[k].Width = 50;
						lab[k].Location = new System.Drawing.Point(x[i] + 130, y_lab[i] + 65);
						lab[k].BackColor = System.Drawing.Color.PaleGreen;
						lab[k].AutoSize = false;
						tab[j].Controls.Add(lab[k]);
						k++;

						/* ダウンロードボタンを作成 */
						btn[j * 10 + i] = new Button();
						btn[j * 10 + i].Text = "プロジェクトにダウンロード";
						btn[j * 10 + i].Font = new Font("メイリオ", 10);
						btn[j * 10 + i].Height = 30;
						btn[j * 10 + i].Width = 180;
						int fid = reader2.GetInt32(0);
						btn[j * 10 + i].Name = "btn" + fid.ToString();
						btn[j * 10 + i].Location = new System.Drawing.Point(x[i], y_btn[i]);
						btn[j * 10 + i].Click += new EventHandler(btn_Click);
						tab[j].Controls.Add(btn[j * 10 + i]);


						i++;
						if (i == 10) { j++; } // 1ページに10個まで表示、それ以上は新しいタブページを作成
						progressBar1.Value++; // プログレスバーの値を1つ進める
					}

				connection.Close(); // 最初のDB接続を閉じる
				connection2.Close(); // 2つ目のDB接続を閉じる

				statushow.Text = "     ファミリーデータの読み込みが完了しました"; // 完了メッセージを表示
			}
			else
			{
				// 該当するファミリーデータが見つからなかった場合の警告メッセージ
				MessageBox.Show("条件に一致するファミリーデータが見つかりませんでした！", "再確認してください", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
        }
      
    }
}

