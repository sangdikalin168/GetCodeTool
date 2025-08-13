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
            loadingPictureBox.Size = new Size(300, 300);
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

            var emailAccounts = EmailConfiguration.GetBuiltInAccounts();

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
            notificationLabel.Location = new Point(margin, margin); // Top-left corner with margin

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

                var fetchTasks = new List<Task<(string Code, string EmailBody, string MatchingEmail, string AliasEmail, string Error)>>();
                foreach (var aliasEmail in lines)
                {
                    string normalizedEmail = NormalizeEmail(aliasEmail);
                    EmailService service = normalizedEmail.IndexOf("yandex.com", StringComparison.OrdinalIgnoreCase) >= 0
                        ? _yandexService
                        : _zohoService;

                    // Start the task and handle exceptions individually
                    fetchTasks.Add(service.FetchLatestCodeAndDeleteAsync(normalizedEmail, aliasEmail, CancellationToken.None)
                        .ContinueWith<(string Code, string EmailBody, string MatchingEmail, string AliasEmail, string Error)>(t =>
                        {
                            if (t.IsFaulted)
                            {
                                string errorMessage = t.Exception?.InnerException?.Message ?? "Unknown error";
                                return (Code: null, EmailBody: null, MatchingEmail: null, AliasEmail: aliasEmail, Error: errorMessage);
                            }
                            return (t.Result.Code, t.Result.EmailBody, t.Result.MatchingEmail, AliasEmail: aliasEmail, Error: null);
                        }));
                }

                // Await all tasks to complete concurrently
                var results = await Task.WhenAll(fetchTasks);

                var otpResults = new List<OtpData>();
                int index = 1;
                StringBuilder errorMessages = new StringBuilder();

                foreach (var result in results)
                {
                    if (!string.IsNullOrEmpty(result.Error))
                    {
                        errorMessages.AppendLine($"Error for {result.AliasEmail}: {result.Error}");
                    }
                    else if (!string.IsNullOrEmpty(result.Code))
                    {
                        otpResults.Add(new OtpData
                        {
                            No = index++,
                            Email = result.AliasEmail,
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

                if (errorMessages.Length > 0)
                {
                    ShowFloatingNotification(errorMessages.ToString().Trim(), Color.Red, Color.White);
                }
            }
            catch (Exception ex)
            {
                ShowFloatingNotification($"Unexpected error: {ex.Message}", Color.Red, Color.White);
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

        private void btnCopyAll_Click(object sender, EventArgs e)
        {
            FormatDataGridView();

            if (dataGridViewOTP.DataSource == null || dataGridViewOTP.Rows.Count == 0)
            {
                ShowFloatingNotification("No OTP codes to copy.", Color.Orange, Color.White);
                return;
            }

            StringBuilder allOtps = new StringBuilder();
            string otpColumnName = "OTP";

            if (!dataGridViewOTP.Columns.Contains(otpColumnName))
            {
                ShowFloatingNotification($"The column '{otpColumnName}' was not found.", Color.OrangeRed, Color.White);
                return;
            }

            foreach (DataGridViewRow row in dataGridViewOTP.Rows)
            {
                if (!row.IsNewRow)
                {
                    DataGridViewCell otpCell = row.Cells[otpColumnName];
                    if (otpCell != null && otpCell.Value != null)
                    {
                        allOtps.AppendLine(otpCell.Value.ToString());
                    }
                }
            }

            if (allOtps.Length > 0)
            {
                string finalOtps = allOtps.ToString().TrimEnd(Environment.NewLine.ToCharArray());
                try
                {
                    Clipboard.SetText(finalOtps);
                    ShowFloatingNotification(finalOtps, Color.SeaGreen, Color.White);
                }
                catch (Exception ex)
                {
                    ShowFloatingNotification($"Failed to copy all OTPs: {ex.Message}", Color.Red, Color.White);
                }
            }
            else
            {
                ShowFloatingNotification("No OTP codes found to copy.", Color.Orange, Color.White);
            }
        }

        private void loadingPictureBox_Click(object sender, EventArgs e)
        {

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
                throw new Exception($"No credentials found for email: {normalizedEmail}");

            using (var client = new ImapClient())
            {
                try
                {
                    await client.ConnectAsync(_imapHost, 993, _socketOptions, cancellationToken);
                    await client.AuthenticateAsync(account.Email, account.Password, cancellationToken);

                    var inbox = client.Inbox;
                    await inbox.OpenAsync(FolderAccess.ReadWrite, cancellationToken);

                    // Get or create RLFARM folder WITHOUT opening it
                    IMailFolder rlfarm;
                    var personal = client.GetFolder(client.PersonalNamespaces[0]);
                    var folders = await personal.GetSubfoldersAsync(false, cancellationToken);

                    var rlfarmFolder = folders.FirstOrDefault(f =>
                        f.Name.Equals("RLFARM", StringComparison.OrdinalIgnoreCase));

                    if (rlfarmFolder == null)
                    {
                        rlfarm = await personal.CreateAsync("RLFARM", true, cancellationToken);
                    }
                    else
                    {
                        rlfarm = rlfarmFolder;
                    }

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
                            try
                            {
                                // CRITICAL FIX: Only inbox is open - no need to open destination folder
                                await inbox.CopyToAsync(uid, rlfarm, cancellationToken);
                                await inbox.AddFlagsAsync(uid, MessageFlags.Deleted, true, cancellationToken);
                                await inbox.ExpungeAsync(cancellationToken);
                            }
                            catch (Exception ex)
                            {
                                throw new Exception($"Failed to move email to RLFARM: {ex.Message}", ex);
                            }
                            break;
                        }
                    }

                    await client.DisconnectAsync(true, cancellationToken);
                    return (latestCode, latestEmailBody, latestMatchingEmail, aliasEmail);
                }
                catch (Exception ex)
                {
                    throw new Exception($"Error processing email {normalizedEmail}: {ex.Message}", ex);
                }
            }
        }

        private string ExtractSecurityCode(string text)
        {
            var match = Regex.Match(text ?? "", "\\b\\d{5,8}\\b");
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
