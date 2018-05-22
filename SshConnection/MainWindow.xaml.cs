using Renci.SshNet;
using System;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Forms;

namespace SshConnection
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            // Info
            var hostNameAddress = AppInfoConstant.hostNameAddress;
            var portNo = AppInfoConstant.portNo;
            var userName = AppInfoConstant.userName;
            var passWord = AppInfoConstant.passWord;

            ConnectionInfo info = new ConnectionInfo(hostNameAddress, portNo, userName,
                new AuthenticationMethod[]{
                    new PasswordAuthenticationMethod(userName, passWord)
                });

            // 入力チェック
            var id1Str = VaridationDelimitter(id1TextBox.Text);
            var id2Str = VaridationDelimitter(id2TextBox.Text);

            var commandStr = AppInfoConstant.command + " '" + id1Str + "' '" + id2Str + "'";

            // ssh接続
            var fileName = ConnectSsh(info, commandStr);

            // sshダウンロード
            DownloadSftp(info, fileName);

            // 終了時、最前面へメッセージボックス表示
            DialogResult msgResult = System.Windows.Forms.MessageBox.Show(fileName + "\r\nが作成出来ました。\r\n\r\nエクスプローラーを開きますか？"
                , "caption"
                , System.Windows.Forms.MessageBoxButtons.OKCancel
                , MessageBoxIcon.Information
                , MessageBoxDefaultButton.Button1
                , System.Windows.Forms.MessageBoxOptions.DefaultDesktopOnly);

            // エクスプローラ実行
            if(msgResult == System.Windows.Forms.DialogResult.OK)
            {
                this.Activate();
                System.Diagnostics.Process.Start("EXPLORER.EXE", System.Environment.CurrentDirectory);
            }
            else
            {
                this.Activate();
            }
        }

        /// <summary>
        /// ssh接続コマンド実行
        /// </summary>
        /// <param name="info"></param>
        /// <param name="commandStr"></param>
        /// <returns></returns>
        private string ConnectSsh(ConnectionInfo info, string commandStr)
        {
            string fileName;

            try
            {
                // Connection
                using (SshClient ssh = new SshClient(info))
                {
                    ssh.Connect();

                    // 接続確認
                    if (ssh.IsConnected)
                    {
                        Console.WriteLine("[OK] ssh connection succeeded!");
                    }
                    else
                    {
                        Console.WriteLine("[ERROR] ssh connection failed!");
                    }

                    // コマンドインスタンス
                    SshCommand cmd = ssh.CreateCommand(commandStr);

                    // コマンド実行
                    //Console.WriteLine("[CMD] {0}", commandStr);
                    cmd.Execute();

                    // 結果を表示する
                    //Console.WriteLine("終了コード:{0}", cmd.Result);

                    // ファイル名取得
                    string[] resultArray = cmd.Result.Split('\n');
                    fileName = resultArray[resultArray.Length - 1].Substring(5);

                    // dissconnection
                    ssh.Disconnect();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                throw ex;
            }

            return fileName;
        }

        /// <summary>
        /// sshダウンロード
        /// </summary>
        /// <param name="info"></param>
        /// <param name="fileName"></param>
        private void DownloadSftp(ConnectionInfo info, string fileName)
        {
            try
            {
                // fileDL
                using (var sftpClient = new SftpClient(info))
                {
                    sftpClient.Connect();
                    sftpClient.ChangeDirectory(AppInfoConstant.dirName);
                    using (var fs = System.IO.File.OpenWrite(fileName))
                    {
                        sftpClient.DownloadFile(fileName, fs);
                    }

                    sftpClient.Disconnect();
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                throw ex;
            }
        }

        /// <summary>
        /// varidation
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        private string VaridationDelimitter(string str)
        {
            string result = str;
            string pattern = @"\.+|\、+|\s+|\|+|\｜+|\t+";

            if (Regex.IsMatch(str, pattern))
            {
                result = Regex.Replace(str, pattern, ",");
            }

            return result;
        }
    }
}
