using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using MailKit.Net.Imap;
using MailKit.Search;
using MailKit.Security;
using MailKit;
using System.Text;

namespace GetCodeTool
{
    public partial class Form1 : Form
    {
        private Label notificationLabel;
        private System.Windows.Forms.Timer notificationTimer;

        private void InitializeLoadingPictureBox()
        {
            loadingPictureBox.Image = Properties.Resources.LoadingGif;
            loadingPictureBox.SizeMode = PictureBoxSizeMode.StretchImage;
            loadingPictureBox.Size = new Size(100, 100);
            loadingPictureBox.Visible = false;

            CenterLoadingPictureBox();
        }
        private void CenterLoadingPictureBox()
        {
            int x = (this.ClientSize.Width - loadingPictureBox.Width) / 2;
            int y = (this.ClientSize.Height - loadingPictureBox.Height) / 2;
            loadingPictureBox.Location = new Point(x, y);
        }

        private void InitializeNotificationControls()
        {
            notificationLabel = new Label();
            notificationTimer = new System.Windows.Forms.Timer();

            notificationLabel.AutoSize = true;
            notificationLabel.MinimumSize = new Size(300, 40);
            notificationLabel.MaximumSize = new Size(300, 0);
            notificationLabel.BackColor = Color.FromArgb(76, 175, 80);
            notificationLabel.ForeColor = Color.White;
            notificationLabel.TextAlign = ContentAlignment.MiddleCenter;
            notificationLabel.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            notificationLabel.Visible = false;
            notificationLabel.Text = "";
            notificationLabel.Padding = new Padding(5);
            notificationLabel.BorderStyle = BorderStyle.FixedSingle;
            notificationLabel.BringToFront();

            notificationLabel.Location = new Point(0, 0);

            this.Controls.Add(notificationLabel);

            notificationTimer.Interval = 2500;
            notificationTimer.Tick += NotificationTimer_Tick;
        }

        private void NotificationTimer_Tick(object sender, EventArgs e)
        {
            notificationTimer.Stop();
            notificationLabel.Visible = false;
        }

        [DllImport("gdi32.dll", EntryPoint = "AddFontResourceW", SetLastError = true)]
        private static extern int AddFontResource([In][MarshalAs(UnmanagedType.LPWStr)] string lpFileName);

        [DllImport("user32.dll", EntryPoint = "SendMessageTimeoutW", SetLastError = true)]
        private static extern IntPtr SendMessageTimeout(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam, int fuFlags, int uTimeout, out IntPtr lpdwResult);

        private readonly EmailService _zohoService;
        private readonly EmailService _yandexService;
        private string defaultText =
            @"Example
            zohodomain1+123xyys2@zohomail.com
            zohodomain2+bsxxyys2@zohomail.com
            zohodomain3+123xyys2@zohomail.com";

        public Form1()
        {
            InitializeComponent();
            InitializeLoadingPictureBox();
            InitializeNotificationControls();

            var emailAccounts = GetBuiltInEmailAccounts();

            _zohoService = new EmailService(emailAccounts, "imap.zoho.com", SecureSocketOptions.Auto);
            _yandexService = new EmailService(emailAccounts, "imap.yandex.com", SecureSocketOptions.SslOnConnect);
            txtEmail.Text = defaultText;
            txtEmail.ForeColor = Color.Gray;

            FormatDataGridView();
            dataGridViewOTP.CellClick += DataGridViewOTP_CellClick;
        }

        private void ShowFloatingNotification(string message, Color backgroundColor, Color textColor)
        {
            notificationLabel.Text = message;
            notificationLabel.BackColor = backgroundColor;
            notificationLabel.ForeColor = textColor;

            int margin = 10;
            notificationLabel.Location = new Point(
                this.ClientSize.Width - notificationLabel.Width - margin,
                this.ClientSize.Height - notificationLabel.Height - margin
            );

            notificationLabel.Visible = true;
            notificationLabel.BringToFront();

            notificationTimer.Stop();
            notificationTimer.Start();
        }

        private void DataGridViewOTP_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                string columnName = dataGridViewOTP.Columns[e.ColumnIndex].Name;
                if (columnName == "OTP")
                {
                    object cellValue = dataGridViewOTP.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
                    if (cellValue != null)
                    {
                        string otpToCopy = cellValue.ToString();
                        try
                        {
                            Clipboard.SetText(otpToCopy);
                            ShowFloatingNotification($"OTP '{otpToCopy}' Copied!", Color.SeaGreen, Color.White);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Failed to copy OTP: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
        }
        private void FormatDataGridView()
        {
            dataGridViewOTP.AllowUserToResizeColumns = false;
            dataGridViewOTP.AllowUserToResizeRows = false;

            if (dataGridViewOTP.Columns["No"] != null)
            {
                dataGridViewOTP.Columns["No"].HeaderText = "No";
                dataGridViewOTP.Columns["No"].Width = 30;
                dataGridViewOTP.Columns["No"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            if (dataGridViewOTP.Columns["Email"] != null)
            {
                dataGridViewOTP.Columns["Email"].HeaderText = "Email Address";
                dataGridViewOTP.Columns["Email"].Width = 200;
                dataGridViewOTP.Columns["Email"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            }

            if (dataGridViewOTP.Columns["Subject"] != null)
            {
                dataGridViewOTP.Columns["Subject"].HeaderText = "Subject";
                dataGridViewOTP.Columns["Subject"].Width = 250;
                dataGridViewOTP.Columns["Subject"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            }

            if (dataGridViewOTP.Columns["OTP"] != null)
            {
                dataGridViewOTP.Columns["OTP"].HeaderText = "OTP";
                dataGridViewOTP.Columns["OTP"].Width = 120;
                dataGridViewOTP.Columns["OTP"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridViewOTP.Columns["OTP"].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
        }

        private string NormalizeEmail(string email)
        {
            int plusIndex = email.IndexOf('+');
            if (plusIndex >= 0)
            {
                int atIndex = email.IndexOf('@');
                if (atIndex > plusIndex)
                {
                    return email.Substring(0, plusIndex) + email.Substring(atIndex);
                }
            }
            return email;
        }
        private List<EmailAccount> GetBuiltInEmailAccounts()
        {
            return new List<EmailAccount>
            {
                new EmailAccount("flyy666@zohomail.com", "QDV3unGpf3hr"),
                new EmailAccount("foxx666@zohomail.com", "5WTw8hY5GjGU"),
                new EmailAccount("funn666@zohomail.com", "hKzYbhcYzTpU"),
                new EmailAccount("gass666@zohomail.com", "jRiEwzww7Fnd"),
                new EmailAccount("drry666@zohomail.com", "6PkT1dGprz7K"),
                new EmailAccount("earr666@zohomail.com", "1Wi2WZMxJ3fx"),
                new EmailAccount("dogg666@zohomail.com", "B1kn1LjJfVCx"),
                new EmailAccount("dott666@zohomail.com", "5FxJEYGT7Rb2"),
                new EmailAccount("eatt666@zohomail.com", "S789BB0kHytG"),
                new EmailAccount("eggg666@zohomail.com", "39LckAt5GunZ"),
                new EmailAccount("endd666@zohomail.com", "BtxbqytW94uC"),
                new EmailAccount("eyee666@zohomail.com", "bZMP141Tn3M8"),
                new EmailAccount("fann666@zohomail.com", "syFEaBtH0n9W"),
                new EmailAccount("farr666@zohomail.com", "1C6YMJQ2WU0w"),
                new EmailAccount("fatt666@zohomail.com", "GwBVR6g8Wkbc"),
                new EmailAccount("feew666@zohomail.com", "xz5WtTWnvTyF"),
                new EmailAccount("fixx666@zohomail.com", "ngmwCdu5AXHh"),
                new EmailAccount("boyy666@zohomail.com", "YXB3YBhYRiyJ"),
                new EmailAccount("bug666@zohomail.com", "nRXZFwKF7Nh5"),
                new EmailAccount("buss666@zohomail.com", "9FmfEew0c6Gf"),
                new EmailAccount("cann666@zohomail.com", "7ZqcaGhTFSM8"),
                new EmailAccount("caap666@zohomail.com", "hVciRNGMB85f"),
                new EmailAccount("carr666@zohomail.com", "H4n6vrLgQwtJ"),
                new EmailAccount("catt666@zohomail.com", "MMMJMsvJTUcu"),
                new EmailAccount("coow666@zohomail.com", "si1vBrTSJQg2"),
                new EmailAccount("cutt666@zohomail.com", "GVnkmbh4ZtbF"),
                new EmailAccount("cupp666@zohomail.com", "dm35LmXRmZ4g"),
                new EmailAccount("daay666@zohomail.com", "duyLgvqgWhiT"),
                new EmailAccount("denn666@zohomail.com", "eZYwCbQ6V6rQ"),
                new EmailAccount("diig666@zohomail.com", "iuTLd8HMRs6d"),
                new EmailAccount("diip666@zohomail.com", "hZpGXws83Npq"),
                new EmailAccount("arrm666@zohomail.com", "2AENcBjP5PMv"),
                new EmailAccount("baag666@zohomail.com", "PJJ2zJQ63qQJ"),
                new EmailAccount("bar666@zohomail.com", "PTjV1fSfbT0s"),
                new EmailAccount("bat666@zohomail.com", "eF6vmJnm8pKN"),
                new EmailAccount("bay666@zohomail.com", "3tK8xLGQZ8KD"),
                new EmailAccount("bedd666@zohomail.com", "adZA2v4uVuPL"),
                new EmailAccount("beee666@zohomail.com", "1UV0KVT7Tza1"),
                new EmailAccount("bet666@zohomail.com", "j1LVTD7PjqpQ"),
                new EmailAccount("big666@zohomail.com", "53P8DD1Y0Fum"),
                new EmailAccount("bit666@zohomail.com", "Wxu9x0eacEU8"),
                new EmailAccount("boxx666@zohomail.com", "BbP3zYuKsE6i"),
                new EmailAccount("whale666@zohomail.com", "TUpw4LHDxP4t"),
                new EmailAccount("bears666@zohomail.com", "WwDRMMn11G2E"),
                new EmailAccount("pum666@zohomail.com", "TRwSCdyK3V1E"),
                new EmailAccount("dump666@zohomail.com", "3Pv3psZXC5XQ"),
                new EmailAccount("mint666@zohomail.com", "vY8HADMyAaFN"),
                new EmailAccount("flip666@zohomail.com", "xzWi0YN8M6vq"),
                new EmailAccount("hold666@zohomail.com", "FQUDjM02WPAq"),
                new EmailAccount("acee666@zohomail.com", "9CH9wenT1M6x"),
                new EmailAccount("ant666@zohomail.com", "Uckj032G28tx"),
                new EmailAccount("dust168@zohomail.com", "4kP4Ep4JZY4P"),
                new EmailAccount("zuck168@zohomail.com", "094grgncpfUa"),
                new EmailAccount("giga168@zohomail.com", "aw99HQvcdZ96"),
                new EmailAccount("dank168@zohomail.com", "8Y2cZPh0r3Wr"),
                new EmailAccount("snip168@zohomail.com", "dbxb12BhQSht"),
                new EmailAccount("hype168@zohomail.com", "DYMraksSFWbe"),
                new EmailAccount("yolo168@zohomail.com", "KANNtMVmhm3C"),
                new EmailAccount("meme168@zohomail.com", "PdrD1DRsrXGM"),
                new EmailAccount("defi168@zohomail.com", "xgJkzeBKgH0u"),
                new EmailAccount("swap168@zohomail.com", "bGsURfFi0SjY"),
                new EmailAccount("farm168@zohomail.com", "CDVcVRwJn6yV"),
                new EmailAccount("dapp168@zohomail.com", "J3KDgYt7MUAM"),
                new EmailAccount("pump168@zohomail.com", "1FAjS8jrq2Gc"),
                new EmailAccount("dump168@zohomail.com", "b8wzvMCCV1RZ"),
                new EmailAccount("hold168@zohomail.com", "nhqc4WKN6VBb"),
                new EmailAccount("pei168@zohomail.com", "bAiP8jNr4g7x"),
                new EmailAccount("lamb168@zohomail.com", "NxLnzP39vevt"),
                new EmailAccount("fomo168@zohomail.com", "mgSAQkygBTWJ"),
                new EmailAccount("rugp168@zohomail.com", "Kkz3ss3iwBqK"),
                new EmailAccount("holdd168@zohomail.com", "B0YcYi5qdXAJ"),
                new EmailAccount("mars168@zohomail.com", "FbwZURjviq70"),
                new EmailAccount("inuz168@zohomail.com", "KxQCMFjSXD6F"),
                new EmailAccount("coin168@zohomail.com", "49yduuxqCyJ4"),
                new EmailAccount("bull168@zohomail.com", "xsZsdxSZt7Qq"),
                new EmailAccount("bear168@zohomail.com", "hbeaythnwSdQ"),
                new EmailAccount("dipp168@zohomail.com", "RwhQamtkpVqz"),
                new EmailAccount("rekt168@zohomail.com", "LKLQPUrYSeRv"),
                new EmailAccount("zoom168@zohomail.com", "9svv1wCGxPF1"),
                new EmailAccount("boom168@zohomail.com", "ztbB9RCawiwP"),
                new EmailAccount("silver168@zohomail.com", "H3M1WZqFBzfL"),
                new EmailAccount("paws168@zohomail.com", "S4J7HaExa1CR"),
                new EmailAccount("chad168@zohomail.com", "FztTBcV5R8QL"),
                new EmailAccount("deep168@zohomail.com", "cYdNb4FY2CvC"),
                new EmailAccount("rich666@zohomail.com", "7peji37cby5t"),
                new EmailAccount("gain168@zohomail.com", "gBYMQsGMJRFN"),
                new EmailAccount("tend168@zohomail.com", "skk360LNiqHc"),
                new EmailAccount("wagm168@zohomail.com", "cZcZNnuNTz2U"),
                new EmailAccount("aird168@zohomail.com", "sQBeeuxE6mQh"),
                new EmailAccount("zcash168@zohomail.com", "U3ZiWcz0z1Lx"),
                new EmailAccount("story168@zohomail.com", "vmF6rDgeMftg"),
                new EmailAccount("core168@zohomail.com", "mp1Y1Aev5Ds1"),
                new EmailAccount("beam168@zohomail.com", "dnXcmsMjYQfJ"),
                new EmailAccount("amp168@zohomail.com", "mQzYwddhTTfy"),
                new EmailAccount("just168@zohomail.com", "j4s2Cm5B3tc0"),
                new EmailAccount("wemix168@zohomail.com", "YN6wbfx9tiZS"),
                new EmailAccount("safe168@zohomail.com", "0wB6LzSKXfF3"),
                new EmailAccount("blur168@zohomail.com", "hdjj9Egcg7cx"),
                new EmailAccount("astar168@zohomail.com", "aw7UJ8KE45aH"),
                new EmailAccount("holo168@zohomail.com", "kGj3q4f29D6h"),
                new EmailAccount("doge168@zohomail.com", "aK6dh3xgz0Jk"),
                new EmailAccount("shib168@zohomail.com", "vR6BGceqHKAn"),
                new EmailAccount("flok168@zohomail.com", "LV5FcfXAEFi1"),
                new EmailAccount("moon168@zohomail.com", "dm7nFimv7YWb"),
                new EmailAccount("gold168@zohomail.com", "9FxNJG2TLLWW"),
                new EmailAccount("usdt168@zohomail.com", "jWuNc5bWvF45"),
                new EmailAccount("neo168@zohomail.com", "mr5YSG5zkaTP"),
                new EmailAccount("cake168@zohomail.com", "41GZzK37J6t9"),
                new EmailAccount("chilliz168@zohomail.com", "wHFgS7xUCNRe"),
                new EmailAccount("spx168@zohomail.com", "NEg6Wev9XfgX"),
                new EmailAccount("pax168@zohomail.com", "1uQCU4K7Jnhr"),
                new EmailAccount("stacks168@zohomail.com", "jY7tismh7Zv3"),
                new EmailAccount("bonk168@zohomail.com", "tSkZmMU57EEU"),
                new EmailAccount("floki168@zohomail.com", "FyFiyiQ11PCp"),
                new EmailAccount("nexo168@zohomail.com", "jLykfQnRRRpi"),
                new EmailAccount("iota168@zohomail.com", "eMvmXg94VLJ2"),
                new EmailAccount("eos168@zohomail.com", "GE4pNSDeXmWv"),
                new EmailAccount("mantle168@zohomail.com", "qEyYJy9Zjh52"),
                new EmailAccount("lido168@zohomail.com", "snT1rgHyiceq"),
                new EmailAccount("cosmo168@zohomail.com", "GmYe1VTxDEfY"),
                new EmailAccount("binance168@zohomail.com", "vW06XcjGxG5Z"),
                new EmailAccount("sui168@zohomail.com", "ws7zvFWycdTN"),
                new EmailAccount("dai168@zohomail.com", "hi0u4syjtWu7"),
                new EmailAccount("apt168@zohomail.com", "29Fb2vhKy111"),
                new EmailAccount("okb168@zohomail.com", "C5PQLdwnqWsD"),
                new EmailAccount("kaspa168@zohomail.com", "UN7yZqQrvcAZ"),
                new EmailAccount("ondo168@zohomail.com", "sAuRKJHj0wfL"),
                new EmailAccount("raven168@zohomail.com", "euKrw2MiHdz3"),
                new EmailAccount("dxy168@zohomail.com", "wAqrYCcLuwh2"),
                new EmailAccount("alt168@zohomail.com", "ewz7smb469AM"),
                new EmailAccount("aii168@zohomail.com", "gfvcX8EyHDWn"),
                new EmailAccount("avax168@zohomail.com", "ydaUqLHEgsvE"),
                new EmailAccount("storm168@zohomail.com", "WgbBmSrm7aWP"),
                new EmailAccount("luna168@zohomail.com", "C3DyTL2gNsJ5"),
                new EmailAccount("ftx168@zohomail.com", "U6LrUKaLvgW8"),
                new EmailAccount("bybit168@zohomail.com", "x2xZykcZd9fk"),
                new EmailAccount("one168@zohomail.com", "FYEt1C72VBSD"),
                new EmailAccount("matic168@zohomail.com", "r7qHKNpsYSQG"),
                new EmailAccount("beta168@zohomail.com", "gq6xY5MJq5gf"),
                new EmailAccount("gmt168@zohomail.com", "HhFbgHrSGSw3"),
                new EmailAccount("atom168@zohomail.com", "R5pdf3Schu94"),
                new EmailAccount("btc168@zohomail.com", "DY2FszYFnU2P"),
                new EmailAccount("dot168@zohomail.com", "4iJnsDB8EBXa"),
                new EmailAccount("axs168@zohomail.com", "Cx0L0Rbnighq"),
                new EmailAccount("etc168@zohomail.com", "jgYd4yRs1c78"),
                new EmailAccount("ape168@zohomail.com", "uVXq7GGUzsvg"),
                new EmailAccount("bat168@zohomail.com", "TiaynCzTDTfg"),
                new EmailAccount("trx168@zohomail.com", "JbDRPq0XtwLC"),
                new EmailAccount("gala168@zohomail.com", "GtLUBxNGF5CW"),
                new EmailAccount("sand168@zohomail.com", "KLnP1S1i8yCX"),
                new EmailAccount("mana666@zohomail.com", "bpE6GqMf1BNG"),
                new EmailAccount("ada168@zohomail.com", "Aucu32i92DTm"),
                new EmailAccount("near168@zohomail.com", "DURbciQNj5wr"),
                new EmailAccount("bnb168@zohomail.com", "ThCQMwH903Pf"),
                new EmailAccount("trump168@zohomail.com", "e2vjYvZEzvUv"),
                new EmailAccount("ftm168@zohomail.com", "s4tBXH6N7qLm"),
                new EmailAccount("xrp168@zohomail.com", "UKth3J6unMky"),
                new EmailAccount("sol168@zohomail.com", "aqD7fXNAxvKd"),
                new EmailAccount("polkadot168@zohomail.com", "YyEy5Nh6eia6"),
                new EmailAccount("eth_creator@zohomail.com", "XNykdqmr6iP8"),
                new EmailAccount("eth168@zohomail.com", "cQttq0m73CbV"),
                new EmailAccount("labubu168@zohomail.com", "c1y4TrQrvc3n"),
                new EmailAccount("eth.05@yandex.com", "rmymzyalkgqageko"),
                new EmailAccount("eth570@yandex.com", "eiixxsokpzimbhhc"),
                new EmailAccount("eth169@yandex.com", "rytjbqbinwndajng"),
                new EmailAccount("xrp05@yandex.com", "jemkbsquucsaljjf"),
                new EmailAccount("xrp.570@yandex.com", "jvuqrvxsnhxirvps")
            };
        }


        private async void btnGetCode_Click(object sender, EventArgs e)
        {
            btnGetCode.Enabled = false;
            btnGetCode.Text = "Loading...";
            loadingPictureBox.Visible = true;

            try
            {
                var lines = txtEmail.Lines
                    .Select(line => line.Trim())
                    .Where(line => !string.IsNullOrEmpty(line) && line != defaultText)
                    .ToList();

                if (!lines.Any())
                {
                    ShowFloatingNotification("Please enter at least one email address.", Color.OrangeRed, Color.White);
                    return;
                }

                var fetchTasks = new List<Task<(string Code, string EmailBody, string MatchingEmail, string AliasEmail)>>();

                foreach (var aliasEmail in lines)
                {
                    string normalizedEmail = NormalizeEmail(aliasEmail);
                    EmailService service = normalizedEmail.IndexOf("yandex.com", StringComparison.OrdinalIgnoreCase) >= 0
                        ? _yandexService
                        : _zohoService;

                    // Start the task without awaiting it immediately
                    fetchTasks.Add(service.FetchLatestCodeAndDeleteAsync(normalizedEmail, aliasEmail, CancellationToken.None)
                                          .ContinueWith(t => (t.Result.Code, t.Result.EmailBody, t.Result.MatchingEmail, AliasEmail: aliasEmail)));
                }

                // Await all tasks to complete concurrently
                var results = await Task.WhenAll(fetchTasks);

                var otpResults = new List<OtpData>();
                int index = 1;

                foreach (var result in results)
                {
                    if (!string.IsNullOrEmpty(result.Code))
                    {
                        otpResults.Add(new OtpData
                        {
                            No = index++,
                            Email = result.AliasEmail, // Use AliasEmail from the task result
                            OTP = result.Code,
                            Subject = result.MatchingEmail,
                        });
                    }
                }

                if (otpResults.Count > 0)
                {
                    dataGridViewOTP.DataSource = otpResults;
                    FormatDataGridView();

                    StringBuilder otpListMessage = new StringBuilder();
                    foreach (var otpData in otpResults)
                    {
                        otpListMessage.AppendLine(otpData.OTP);
                    }
                    ShowFloatingNotification(otpListMessage.ToString().Trim(), Color.SeaGreen, Color.White);
                }
                else
                {
                    ShowFloatingNotification("No OTP codes found for any email.", Color.Orange, Color.White);
                    dataGridViewOTP.DataSource = null;
                }
            }
            catch (Exception ex)
            {
                ShowFloatingNotification($"Error: {ex.Message}", Color.Red, Color.White);
            }
            finally
            {
                btnGetCode.Enabled = true;
                btnGetCode.Text = "Get Code";
                loadingPictureBox.Visible = false;
            }
        }

        private void txtEmail_Enter(object sender, EventArgs e)
        {
            if (txtEmail.Text == defaultText)
            {
                txtEmail.Text = "";
                txtEmail.ForeColor = Color.Black;
            }
        }

        private void txtEmail_Leave(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtEmail.Text))
            {
                txtEmail.Text = defaultText;
                txtEmail.ForeColor = Color.Gray;
            }
        }
    }

    public class OtpData
    {
        public int No { get; set; }
        public string Email { get; set; }
        public string Subject { get; set; }
        public string OTP { get; set; }
    }

    public class EmailService
    {
        private readonly List<EmailAccount> _accounts;
        private readonly string _imapHost;
        private readonly SecureSocketOptions _socketOptions;

        public EmailService(List<EmailAccount> accounts, string imapHost, SecureSocketOptions socketOptions)
        {
            _accounts = accounts;
            _imapHost = imapHost;
            _socketOptions = socketOptions;
        }

        public async Task<(string Code, string EmailBody, string MatchingEmail, string AliasEmail)> FetchLatestCodeAndDeleteAsync(string normalizedEmail, string aliasEmail, CancellationToken cancellationToken)
        {
            var account = _accounts.FirstOrDefault(a => a.Email.Equals(normalizedEmail, StringComparison.OrdinalIgnoreCase));
            if (account == null)
                throw new Exception("No credentials found for this email.");

            using (var client = new ImapClient())
            {
                await client.ConnectAsync(_imapHost, 993, _socketOptions, cancellationToken);
                await client.AuthenticateAsync(account.Email, account.Password, cancellationToken);

                var inbox = client.Inbox;
                await inbox.OpenAsync(FolderAccess.ReadWrite, cancellationToken);

                var uids = await inbox.SearchAsync(SearchQuery.All, cancellationToken);

                string latestCode = null;
                string latestEmailBody = null;
                string latestMatchingEmail = null;

                foreach (var uid in uids.Reverse())
                {
                    var message = await inbox.GetMessageAsync(uid, cancellationToken);

                    var toHeader = message.Headers["To"] ?? "";
                    var deliveredTo = message.Headers["Delivered-To"] ?? "";
                    if (toHeader.IndexOf(aliasEmail, StringComparison.OrdinalIgnoreCase) < 0 &&
                        deliveredTo.IndexOf(aliasEmail, StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        continue;
                    }

                    latestMatchingEmail = message.Subject;

                    latestCode = ExtractSecurityCode(message.Subject);
                    if (string.IsNullOrEmpty(latestCode))
                    {
                        latestCode = ExtractSecurityCode(message.TextBody);
                    }
                    if (!string.IsNullOrEmpty(latestCode))
                    {
                        latestEmailBody = message.TextBody;
                        await inbox.AddFlagsAsync(uid, MessageFlags.Deleted, true, cancellationToken);
                        await inbox.ExpungeAsync(cancellationToken);
                        break;
                    }
                }

                await client.DisconnectAsync(true, cancellationToken);
                return (latestCode, latestEmailBody, latestMatchingEmail, aliasEmail); // Also return aliasEmail
            }
        }

        private string ExtractSecurityCode(string subject)
        {
            var match = Regex.Match(subject, "\\b\\d{5,8}\\b");
            return match.Success ? match.Value : null;
        }
    }

    public class EmailAccount
    {
        public string Email { get; }
        public string Password { get; }

        public EmailAccount(string email, string password)
        {
            Email = email;
            Password = password;
        }
    }
}
