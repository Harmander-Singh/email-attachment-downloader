using Microsoft.Office.Interop.Outlook;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using Application = Microsoft.Office.Interop.Outlook.Application;
using Attachment = Microsoft.Office.Interop.Outlook.Attachment;
using Exception = System.Exception;
using Path = System.IO.Path;

namespace OutlookAttachmentDownloader
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private CancellationTokenSource _cancellationTokenSource;
        private int _totalAttachments = 0;
        private int _downloadedAttachments = 0;

        // Windows API for folder browser dialog
        [DllImport("shell32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern IntPtr SHBrowseForFolder([In] ref BrowseInfo lpbi);

        [DllImport("shell32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern bool SHGetPathFromIDList([In] IntPtr pidl, [In, Out] char[] pszPath);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        public struct BrowseInfo
        {
            public IntPtr hwndOwner;
            public IntPtr pidlRoot;
            public string pszDisplayName;
            public string lpszTitle;
            public uint ulFlags;
            public IntPtr lpfn;
            public IntPtr lParam;
            public int iImage;
        }

        private const uint BIF_RETURNONLYFSDIRS = 0x0001;
        private const uint BIF_NEWDIALOGSTYLE = 0x0040;


        public MainWindow()
        {
            InitializeComponent();
            InitializeUI();
        }

        private void InitializeUI()
        {
            // Set default download path to user's downloads folder
            string defaultPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads", "EmailAttachments");
            txtDownloadPath.Text = defaultPath;

            // Set default date range (last 30 days)
            dpToDate.SelectedDate = DateTime.Now;
            dpFromDate.SelectedDate = DateTime.Now.AddDays(-30);

            UpdateStatus("Ready to download attachments");
        }

        private void BrowseFolder_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var folderPath = ShowFolderBrowserDialog("Select folder to save attachments");
                if (!string.IsNullOrEmpty(folderPath))
                {
                    txtDownloadPath.Text = folderPath;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error selecting folder: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private string ShowFolderBrowserDialog(string description)
        {
            var browseInfo = new BrowseInfo
            {
                hwndOwner = new WindowInteropHelper(this).Handle,
                pidlRoot = IntPtr.Zero,
                pszDisplayName = null,
                lpszTitle = description,
                ulFlags = BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE,
                lpfn = IntPtr.Zero,
                lParam = IntPtr.Zero,
                iImage = 0
            };

            IntPtr pidl = SHBrowseForFolder(ref browseInfo);

            if (pidl == IntPtr.Zero)
                return string.Empty;

            try
            {
                char[] path = new char[260];
                if (SHGetPathFromIDList(pidl, path))
                {
                    return new string(path).TrimEnd('\0');
                }
                return string.Empty;
            }
            finally
            {
                Marshal.FreeCoTaskMem(pidl);
            }
        }

        private async void StartDownload_Click(object sender, RoutedEventArgs e)
        {
            if (!ValidateInput())
                return;

            _cancellationTokenSource = new CancellationTokenSource();
            SetUIDownloadingState(true);

            try
            {
                await DownloadAttachmentsAsync(_cancellationTokenSource.Token);
            }
            catch (OperationCanceledException)
            {
                AppendResult("Download cancelled by user.");
                UpdateStatus("Download cancelled");
            }
            catch (Exception ex)
            {
                AppendResult($"Error: {ex.Message}");
                UpdateStatus("Download failed");
                System.Windows.MessageBox.Show($"An error occurred: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                SetUIDownloadingState(false);
            }
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            _cancellationTokenSource?.Cancel();
        }

        private bool ValidateInput()
        {
            if (string.IsNullOrWhiteSpace(txtSenderEmail.Text))
            {
                System.Windows.MessageBox.Show("Please enter a sender email address.", "Validation Error",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                txtSenderEmail.Focus();
                return false;
            }

            if (string.IsNullOrWhiteSpace(txtDownloadPath.Text))
            {
                System.Windows.MessageBox.Show("Please select a download folder.", "Validation Error",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            if (!IsValidEmail(txtSenderEmail.Text))
            {
                System.Windows.MessageBox.Show("Please enter a valid email address.", "Validation Error",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                txtSenderEmail.Focus();
                return false;
            }

            return true;
        }

        private bool IsValidEmail(string email)
        {
            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return addr.Address == email;
            }
            catch
            {
                return false;
            }
        }

        private async Task DownloadAttachmentsAsync(CancellationToken cancellationToken)
        {
            UpdateStatus("Connecting to Outlook...");
            AppendResult("Starting download process...");
            AppendResult($"Sender: {txtSenderEmail.Text}");
            AppendResult($"Download folder: {txtDownloadPath.Text}");

            if (dpFromDate.SelectedDate.HasValue || dpToDate.SelectedDate.HasValue)
            {
                AppendResult($"Date range: {dpFromDate.SelectedDate?.ToShortDateString() ?? "Any"} to {dpToDate.SelectedDate?.ToShortDateString() ?? "Any"}");
            }

            AppendResult(""); // Empty line for spacing

            Application outlookApp = null;
            NameSpace nameSpace = null;
            MAPIFolder inboxFolder = null;

            try
            {
                // Connect to Outlook
                outlookApp = new Application();
                nameSpace = outlookApp.GetNamespace("MAPI");
                inboxFolder = nameSpace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

                UpdateStatus("Searching for emails...");

                // Create download directory
                Directory.CreateDirectory(txtDownloadPath.Text);

                // Search for emails from specific sender
                string filter = CreateSearchFilter();
                Items mailItems = inboxFolder.Items.Restrict(filter);

                AppendResult($"Found {mailItems.Count} emails from {txtSenderEmail.Text}");

                if (mailItems.Count == 0)
                {
                    AppendResult("No emails found matching the criteria.");
                    UpdateStatus("No emails found");
                    return;
                }

                // Count total attachments first
                _totalAttachments = 0;
                foreach (MailItem mailItem in mailItems)
                {
                    if (cancellationToken.IsCancellationRequested)
                        return;

                    if (mailItem.Attachments.Count > 0)
                    {
                        _totalAttachments += GetValidAttachments(mailItem).Count();
                    }
                }

                AppendResult($"Total attachments to download: {_totalAttachments}");
                AppendResult(""); // Empty line

                if (_totalAttachments == 0)
                {
                    AppendResult("No attachments found in the emails.");
                    UpdateStatus("No attachments found");
                    return;
                }

                _downloadedAttachments = 0;
                progressBar.Maximum = _totalAttachments;

                // Download attachments
                foreach (MailItem mailItem in mailItems)
                {
                    if (cancellationToken.IsCancellationRequested)
                        return;

                    await ProcessEmailAttachments(mailItem, cancellationToken);
                }

                AppendResult("");
                AppendResult($"Download completed! {_downloadedAttachments} attachments saved to:");
                AppendResult(txtDownloadPath.Text);
                UpdateStatus($"Download completed - {_downloadedAttachments} attachments downloaded");

                // Ask user if they want to open the download folder
                var result = System.Windows.MessageBox.Show(
                    $"Download completed successfully!\n\n{_downloadedAttachments} attachments were saved.\n\nWould you like to open the download folder?",
                    "Download Complete",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Information);

                if (result == MessageBoxResult.Yes)
                {
                    System.Diagnostics.Process.Start("explorer.exe", txtDownloadPath.Text);
                }
            }
            finally
            {
                // Clean up COM objects
                if (inboxFolder != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(inboxFolder);
                if (nameSpace != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(nameSpace);
                if (outlookApp != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(outlookApp);
            }
        }

        private string CreateSearchFilter()
        {
            string filter = $"[SenderEmailAddress] = '{txtSenderEmail.Text}' OR [SenderName] = '{txtSenderEmail.Text}'";

            if (dpFromDate.SelectedDate.HasValue)
            {
                filter += $" AND [ReceivedTime] >= '{dpFromDate.SelectedDate.Value:MM/dd/yyyy}'";
            }

            if (dpToDate.SelectedDate.HasValue)
            {
                DateTime toDate = dpToDate.SelectedDate.Value.AddDays(1); // Include the entire day
                filter += $" AND [ReceivedTime] < '{toDate:MM/dd/yyyy}'";
            }

            return filter;
        }

        private IEnumerable<Attachment> GetValidAttachments(MailItem mailItem)
        {
            var validAttachments = new List<Attachment>();

            foreach (Attachment attachment in mailItem.Attachments)
            {
                // Skip embedded images and other non-file attachments
                if (attachment.Type == OlAttachmentType.olByValue)
                {
                    if (IsValidFileType(attachment.FileName))
                    {
                        validAttachments.Add(attachment);
                    }
                }
            }

            return validAttachments;
        }

        private bool IsValidFileType(string fileName)
        {
            if (string.IsNullOrWhiteSpace(txtFileTypes.Text))
                return true; // No filter means all files are valid

            string[] allowedExtensions = txtFileTypes.Text
                .Split(',')
                .Select(ext => ext.Trim().ToLower())
                .Where(ext => !string.IsNullOrWhiteSpace(ext))
                .ToArray();

            if (allowedExtensions.Length == 0)
                return true;

            string fileExtension = Path.GetExtension(fileName).ToLower();
            return allowedExtensions.Contains(fileExtension);
        }

        private async Task ProcessEmailAttachments(MailItem mailItem, CancellationToken cancellationToken)
        {
            if (mailItem.Attachments.Count == 0)
                return;

            var validAttachments = GetValidAttachments(mailItem);
            if (!validAttachments.Any())
                return;

            string emailDate = mailItem.ReceivedTime.ToString("yyyy-MM-dd");
            string emailSubject = SanitizeFileName(mailItem.Subject);
            if (emailSubject.Length > 50)
                emailSubject = emailSubject.Substring(0, 50);

            string emailFolder = Path.Combine(txtDownloadPath.Text, $"{emailDate}_{emailSubject}");
            Directory.CreateDirectory(emailFolder);

            AppendResult($"Processing email: {mailItem.Subject} ({mailItem.ReceivedTime:yyyy-MM-dd HH:mm})");

            foreach (Attachment attachment in validAttachments)
            {
                if (cancellationToken.IsCancellationRequested)
                    return;

                try
                {
                    string fileName = SanitizeFileName(attachment.FileName);
                    string filePath = Path.Combine(emailFolder, fileName);

                    // Handle duplicate file names
                    int counter = 1;
                    string originalFilePath = filePath;
                    while (File.Exists(filePath))
                    {
                        string nameWithoutExt = Path.GetFileNameWithoutExtension(originalFilePath);
                        string extension = Path.GetExtension(originalFilePath);
                        filePath = Path.Combine(emailFolder, $"{nameWithoutExt}_{counter}{extension}");
                        counter++;
                    }

                    attachment.SaveAsFile(filePath);
                    AppendResult($"  ✓ Downloaded: {fileName} ({FormatFileSize(new FileInfo(filePath).Length)})");

                    _downloadedAttachments++;

                    // Update progress on UI thread
                    Dispatcher.Invoke(() =>
                    {
                        progressBar.Value = _downloadedAttachments;
                        lblProgress.Text = $"Downloaded {_downloadedAttachments} of {_totalAttachments} attachments";
                        lblStats.Text = $"{_downloadedAttachments} attachments downloaded";
                    });

                    // Small delay to prevent UI freezing
                    await Task.Delay(10, cancellationToken);
                }
                catch (Exception ex)
                {
                    AppendResult($"  ✗ Failed to download {attachment.FileName}: {ex.Message}");
                }
            }
        }

        private string SanitizeFileName(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                return "unnamed_file";

            // Remove invalid file name characters
            char[] invalidChars = Path.GetInvalidFileNameChars();
            string sanitized = new string(fileName.Where(c => !invalidChars.Contains(c)).ToArray());

            // Replace common problematic characters
            sanitized = sanitized.Replace(":", "_")
                                .Replace("?", "_")
                                .Replace("*", "_")
                                .Replace("\"", "_")
                                .Replace("<", "_")
                                .Replace(">", "_")
                                .Replace("|", "_");

            // Ensure the filename is not too long
            if (sanitized.Length > 200)
                sanitized = sanitized.Substring(0, 200);

            return sanitized.Trim();
        }

        private string FormatFileSize(long bytes)
        {
            string[] sizes = { "B", "KB", "MB", "GB" };
            double len = bytes;
            int order = 0;
            while (len >= 1024 && order < sizes.Length - 1)
            {
                order++;
                len = len / 1024;
            }
            return $"{len:0.##} {sizes[order]}";
        }

        private void SetUIDownloadingState(bool isDownloading)
        {
            btnDownload.IsEnabled = !isDownloading;
            btnCancel.IsEnabled = isDownloading;
            txtSenderEmail.IsEnabled = !isDownloading;
            txtDownloadPath.IsEnabled = !isDownloading;
            txtFileTypes.IsEnabled = !isDownloading;
            dpFromDate.IsEnabled = !isDownloading;
            dpToDate.IsEnabled = !isDownloading;

            if (!isDownloading)
            {
                progressBar.Value = 0;
                lblProgress.Text = _downloadedAttachments > 0 ?
                    $"Completed - {_downloadedAttachments} attachments downloaded" :
                    "Ready to start...";
            }
        }

        private void UpdateStatus(string message)
        {
            Dispatcher.Invoke(() =>
            {
                lblStatus.Text = message;
            });
        }

        private void AppendResult(string message)
        {
            Dispatcher.Invoke(() =>
            {
                txtResults.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}{Environment.NewLine}");
                txtResults.ScrollToEnd();
            });
        }
    }
}