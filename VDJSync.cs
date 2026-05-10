using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using System.Xml;
using Microsoft.Win32;

namespace VDJSync
{
    #region Program

    static class Program
    {
        private static Mutex _mutex;

        [STAThread]
        static void Main()
        {
            bool createdNew;
            _mutex = new Mutex(true, @"Global\VDJSyncApp", out createdNew);
            if (!createdNew)
            {
                MessageBox.Show("VDJ Sync is already running.", "VDJ Sync",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                string pcName = RegistryHelper.GetPCName();
                if (string.IsNullOrEmpty(pcName))
                {
                    using (var dlg = new PCNameDialog())
                    {
                        if (dlg.ShowDialog() != DialogResult.OK)
                            return;
                        RegistryHelper.SetPCName(dlg.PCName);
                    }
                }

                Application.Run(new TrayApplicationContext());
            }
            finally
            {
                _mutex.ReleaseMutex();
                _mutex.Dispose();
            }
        }
    }

    #endregion

    #region PCNameDialog

    class PCNameDialog : Form
    {
        private TextBox _txtName;
        private Label _lblError;
        public string PCName { get; private set; }

        public PCNameDialog()
        {
            Text = "VDJ Sync - First Run Setup";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(360, 160);

            var lbl = new Label
            {
                Text = "Enter a 2-3 character PC identifier\n(e.g. DJ1, AB, PC3):",
                Location = new Point(20, 20),
                Size = new Size(320, 36),
                Font = new Font("Segoe UI", 9.5f)
            };
            Controls.Add(lbl);

            _txtName = new TextBox
            {
                Location = new Point(20, 62),
                Size = new Size(120, 26),
                Font = new Font("Segoe UI", 12f),
                MaxLength = 3,
                CharacterCasing = CharacterCasing.Upper
            };
            Controls.Add(_txtName);

            _lblError = new Label
            {
                Location = new Point(150, 66),
                Size = new Size(200, 20),
                ForeColor = Color.Red,
                Font = new Font("Segoe UI", 8.5f)
            };
            Controls.Add(_lblError);

            var btnOK = new Button
            {
                Text = "OK",
                Location = new Point(160, 110),
                Size = new Size(80, 30),
                DialogResult = DialogResult.None
            };
            btnOK.Click += (s, e) =>
            {
                string val = _txtName.Text.Trim();
                if (!Regex.IsMatch(val, @"^[A-Za-z0-9]{2,3}$"))
                {
                    _lblError.Text = "Must be 2-3 letters/numbers.";
                    return;
                }
                PCName = val.ToUpper();
                DialogResult = DialogResult.OK;
                Close();
            };
            Controls.Add(btnOK);

            var btnCancel = new Button
            {
                Text = "Cancel",
                Location = new Point(250, 110),
                Size = new Size(80, 30),
                DialogResult = DialogResult.Cancel
            };
            Controls.Add(btnCancel);

            AcceptButton = btnOK;
            CancelButton = btnCancel;
        }
    }

    #endregion

    #region TrayApplicationContext

    class TrayApplicationContext : ApplicationContext
    {
        private NotifyIcon _trayIcon;
        private System.Threading.Timer _syncTimer;
        private ToolStripMenuItem _lastSyncItem;
        private ToolStripMenuItem _nextSyncItem;
        private ToolStripMenuItem _syncNowItem;
        private AppSettings _settings;
        private Logger _logger;
        private string _pcName;
        private bool _isSyncing;
        private readonly object _syncLock = new object();
        private DateTime _nextSyncTime;
        private DateTime? _lastSyncTime;
        private Icon _iconGreen, _iconBlue, _iconRed;

        public TrayApplicationContext()
        {
            _settings = AppSettings.Load();
            _pcName = RegistryHelper.GetPCName();
            _logger = new Logger();

            _iconGreen = IconHelper.CreateTrayIcon(Color.FromArgb(76, 175, 80));
            _iconBlue = IconHelper.CreateTrayIcon(Color.FromArgb(33, 150, 243));
            _iconRed = IconHelper.CreateTrayIcon(Color.FromArgb(244, 67, 54));

            var menu = new ContextMenuStrip();

            var titleItem = new ToolStripMenuItem("VDJ Sync - " + _pcName);
            titleItem.Font = new Font(titleItem.Font, FontStyle.Bold);
            titleItem.Enabled = false;
            menu.Items.Add(titleItem);
            menu.Items.Add(new ToolStripSeparator());

            _syncNowItem = new ToolStripMenuItem("Sync Now");
            _syncNowItem.Click += (s, e) => RunSyncOnBackground();
            menu.Items.Add(_syncNowItem);
            menu.Items.Add(new ToolStripSeparator());

            _lastSyncItem = new ToolStripMenuItem("Last Sync: Never");
            _lastSyncItem.Enabled = false;
            menu.Items.Add(_lastSyncItem);

            _nextSyncItem = new ToolStripMenuItem("Next Sync: ...");
            _nextSyncItem.Enabled = false;
            menu.Items.Add(_nextSyncItem);

            menu.Items.Add(new ToolStripSeparator());

            var settingsItem = new ToolStripMenuItem("Settings...");
            settingsItem.Click += (s, e) => ShowSettings();
            menu.Items.Add(settingsItem);

            var logItem = new ToolStripMenuItem("View Log...");
            logItem.Click += (s, e) =>
            {
                if (File.Exists(Logger.LogFilePath))
                    System.Diagnostics.Process.Start("notepad.exe", Logger.LogFilePath);
                else
                    MessageBox.Show("No log file yet.", "VDJ Sync");
            };
            menu.Items.Add(logItem);

            menu.Items.Add(new ToolStripSeparator());

            var quitItem = new ToolStripMenuItem("Quit");
            quitItem.Click += (s, e) => Quit();
            menu.Items.Add(quitItem);

            _trayIcon = new NotifyIcon
            {
                Icon = _iconGreen,
                Text = "VDJ Sync - " + _pcName,
                ContextMenuStrip = menu,
                Visible = true
            };

            _trayIcon.DoubleClick += (s, e) => RunSyncOnBackground();

            _nextSyncTime = DateTime.Now.AddMinutes(1);
            UpdateMenuTimestamps();
            _syncTimer = new System.Threading.Timer(OnTimerTick, null,
                TimeSpan.FromMinutes(1), TimeSpan.FromHours(1));

            if (string.IsNullOrEmpty(_settings.WebDavUrl))
            {
                _trayIcon.ShowBalloonTip(5000, "VDJ Sync",
                    "Please configure settings before first sync.", ToolTipIcon.Info);
                ShowSettings();
            }
        }

        private void OnTimerTick(object state)
        {
            _nextSyncTime = DateTime.Now.AddHours(1);
            RunSyncOnBackground();
        }

        private void RunSyncOnBackground()
        {
            lock (_syncLock)
            {
                if (_isSyncing)
                {
                    ShowBalloon("Sync already in progress...", ToolTipIcon.Info);
                    return;
                }
                _isSyncing = true;
            }

            InvokeOnUI(() =>
            {
                _trayIcon.Icon = _iconBlue;
                _syncNowItem.Enabled = false;
            });

            ThreadPool.QueueUserWorkItem(_ =>
            {
                try
                {
                    _settings = AppSettings.Load();

                    if (string.IsNullOrEmpty(_settings.WebDavUrl))
                    {
                        _logger.Warn("Sync skipped: WebDAV URL not configured.");
                        ShowBalloon("Sync skipped: configure settings first.", ToolTipIcon.Warning);
                        return;
                    }

                    var client = new WebDavClient(_settings.WebDavUrl,
                        _settings.WebDavUsername, _settings.WebDavPassword);
                    var engine = new SyncEngine(_settings, client, _logger, _pcName);
                    var result = engine.RunFullSync();

                    _lastSyncTime = DateTime.Now;
                    _settings.LastSyncTime = _lastSyncTime.Value.ToString("o");
                    _settings.Save();

                    InvokeOnUI(() => _trayIcon.Icon = result.Success ? _iconGreen : _iconRed);
                    ShowBalloon(result.Summary, result.Success ? ToolTipIcon.Info : ToolTipIcon.Warning);
                }
                catch (Exception ex)
                {
                    _logger.Error("Sync failed with unhandled exception", ex);
                    InvokeOnUI(() => _trayIcon.Icon = _iconRed);
                    ShowBalloon("Sync failed: " + ex.Message, ToolTipIcon.Error);
                }
                finally
                {
                    lock (_syncLock) _isSyncing = false;
                    InvokeOnUI(() =>
                    {
                        _syncNowItem.Enabled = true;
                        UpdateMenuTimestamps();
                    });
                }
            });
        }

        private void ShowSettings()
        {
            using (var frm = new SettingsForm())
            {
                if (frm.ShowDialog() == DialogResult.OK)
                    _settings = AppSettings.Load();
            }
        }

        private void UpdateMenuTimestamps()
        {
            if (_lastSyncTime.HasValue)
                _lastSyncItem.Text = "Last Sync: " + _lastSyncTime.Value.ToString("yyyy-MM-dd HH:mm:ss");
            else
                _lastSyncItem.Text = "Last Sync: Never";

            _nextSyncItem.Text = "Next Sync: " + _nextSyncTime.ToString("yyyy-MM-dd HH:mm:ss");
        }

        private void ShowBalloon(string message, ToolTipIcon icon)
        {
            try
            {
                _trayIcon.ShowBalloonTip(3000, "VDJ Sync", message, icon);
            }
            catch { }
        }

        private void InvokeOnUI(Action action)
        {
            try
            {
                if (_trayIcon.ContextMenuStrip != null && _trayIcon.ContextMenuStrip.InvokeRequired)
                    _trayIcon.ContextMenuStrip.Invoke(action);
                else
                    action();
            }
            catch { }
        }

        private void Quit()
        {
            if (_syncTimer != null) _syncTimer.Dispose();
            _trayIcon.Visible = false;
            _trayIcon.Dispose();
            if (_iconGreen != null) _iconGreen.Dispose();
            if (_iconBlue != null) _iconBlue.Dispose();
            if (_iconRed != null) _iconRed.Dispose();
            Application.Exit();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (_syncTimer != null) _syncTimer.Dispose();
                if (_trayIcon != null) _trayIcon.Dispose();
                if (_iconGreen != null) _iconGreen.Dispose();
                if (_iconBlue != null) _iconBlue.Dispose();
                if (_iconRed != null) _iconRed.Dispose();
            }
            base.Dispose(disposing);
        }
    }

    #endregion

    #region IconHelper

    static class IconHelper
    {
        public static Icon CreateTrayIcon(Color fillColor)
        {
            using (var bmp = new Bitmap(16, 16))
            {
                using (var g = Graphics.FromImage(bmp))
                {
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    g.Clear(Color.Transparent);

                    using (var brush = new SolidBrush(fillColor))
                        g.FillEllipse(brush, 1, 1, 13, 13);

                    using (var pen = new Pen(Color.FromArgb(60, 0, 0, 0), 1f))
                        g.DrawEllipse(pen, 1, 1, 13, 13);

                    using (var pen = new Pen(Color.White, 1.5f))
                    {
                        g.DrawArc(pen, 4, 4, 7, 7, 200, 160);
                        g.DrawArc(pen, 4, 4, 7, 7, 20, 160);
                    }
                }
                return Icon.FromHandle(bmp.GetHicon());
            }
        }
    }

    #endregion

    #region RegistryHelper

    static class RegistryHelper
    {
        private const string KeyPath = @"SOFTWARE\VDJSync";

        public static string GetPCName()
        {
            using (var key = Registry.CurrentUser.OpenSubKey(KeyPath))
                return key != null ? key.GetValue("PCName") as string : null;
        }

        public static void SetPCName(string name)
        {
            using (var key = Registry.CurrentUser.CreateSubKey(KeyPath))
                key.SetValue("PCName", name, RegistryValueKind.String);
        }
    }

    #endregion

    #region AppSettings

    class AppSettings
    {
        public string WebDavUrl { get; set; }
        public string WebDavUsername { get; set; }
        public string WebDavPassword { get; set; }
        public string TracklistFilePath { get; set; }
        public string TracklistingFolderPath { get; set; }
        public string PlaylistsFolderPath { get; set; }
        public string MP3ExtractPath { get; set; }
        public string MP3ListingPageUrl { get; set; }
        public string MP3PageUsername { get; set; }
        public string MP3PagePassword { get; set; }
        public string SparePath1 { get; set; }
        public string SparePath2 { get; set; }
        public string LastSyncTime { get; set; }

        public static string SettingsDir
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "VDJSync");
            }
        }

        public static string SettingsFilePath
        {
            get { return Path.Combine(SettingsDir, "settings.xml"); }
        }

        public AppSettings()
        {
            WebDavUrl = "";
            WebDavUsername = "";
            WebDavPassword = "";
            MP3ExtractPath = @"X:\";
            MP3ListingPageUrl = "";
            MP3PageUsername = "";
            MP3PagePassword = "";
            SparePath1 = "";
            SparePath2 = "";
            LastSyncTime = "";

            string docs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string vdj = Path.Combine(docs, "VirtualDJ");
            TracklistFilePath = Path.Combine(vdj, "History", "tracklist.txt");
            TracklistingFolderPath = Path.Combine(vdj, "Tracklisting");
            PlaylistsFolderPath = Path.Combine(vdj, "Playlists");
        }

        public static AppSettings Load()
        {
            var s = new AppSettings();
            if (!File.Exists(SettingsFilePath))
                return s;

            try
            {
                var doc = new XmlDocument();
                doc.Load(SettingsFilePath);

                s.WebDavUrl = ReadNode(doc, "/VDJSyncSettings/WebDAV/ServerUrl") ?? "";
                s.WebDavUsername = ReadNode(doc, "/VDJSyncSettings/WebDAV/Username") ?? "";

                string encPwd = ReadNode(doc, "/VDJSyncSettings/WebDAV/Password");
                if (!string.IsNullOrEmpty(encPwd))
                {
                    try
                    {
                        byte[] encrypted = Convert.FromBase64String(encPwd);
                        byte[] decrypted = ProtectedData.Unprotect(encrypted, null,
                            DataProtectionScope.CurrentUser);
                        s.WebDavPassword = Encoding.UTF8.GetString(decrypted);
                    }
                    catch { s.WebDavPassword = ""; }
                }

                s.TracklistFilePath = ReadNode(doc, "/VDJSyncSettings/Paths/TracklistFile")
                    ?? s.TracklistFilePath;
                s.TracklistingFolderPath = ReadNode(doc, "/VDJSyncSettings/Paths/TracklistingFolder")
                    ?? s.TracklistingFolderPath;
                s.PlaylistsFolderPath = ReadNode(doc, "/VDJSyncSettings/Paths/PlaylistsFolder")
                    ?? s.PlaylistsFolderPath;
                s.MP3ExtractPath = ReadNode(doc, "/VDJSyncSettings/Paths/MP3ExtractPath")
                    ?? s.MP3ExtractPath;
                s.SparePath1 = ReadNode(doc, "/VDJSyncSettings/Paths/SparePath1") ?? "";
                s.SparePath2 = ReadNode(doc, "/VDJSyncSettings/Paths/SparePath2") ?? "";
                s.MP3ListingPageUrl = ReadNode(doc, "/VDJSyncSettings/MP3/ListingPageUrl") ?? "";
                s.MP3PageUsername = ReadNode(doc, "/VDJSyncSettings/MP3/Username") ?? "";

                string encMP3Pwd = ReadNode(doc, "/VDJSyncSettings/MP3/Password");
                if (!string.IsNullOrEmpty(encMP3Pwd))
                {
                    try
                    {
                        byte[] encrypted = Convert.FromBase64String(encMP3Pwd);
                        byte[] decrypted = ProtectedData.Unprotect(encrypted, null,
                            DataProtectionScope.CurrentUser);
                        s.MP3PagePassword = Encoding.UTF8.GetString(decrypted);
                    }
                    catch { s.MP3PagePassword = ""; }
                }

                s.LastSyncTime = ReadNode(doc, "/VDJSyncSettings/LastSyncTime") ?? "";
            }
            catch { }

            return s;
        }

        public void Save()
        {
            Directory.CreateDirectory(SettingsDir);

            var doc = new XmlDocument();
            var root = doc.CreateElement("VDJSyncSettings");
            doc.AppendChild(root);

            var webdav = doc.CreateElement("WebDAV");
            root.AppendChild(webdav);
            AddElement(doc, webdav, "ServerUrl", WebDavUrl);
            AddElement(doc, webdav, "Username", WebDavUsername);

            string encPwd = "";
            if (!string.IsNullOrEmpty(WebDavPassword))
            {
                try
                {
                    byte[] plain = Encoding.UTF8.GetBytes(WebDavPassword);
                    byte[] encrypted = ProtectedData.Protect(plain, null,
                        DataProtectionScope.CurrentUser);
                    encPwd = Convert.ToBase64String(encrypted);
                }
                catch { }
            }
            AddElement(doc, webdav, "Password", encPwd);

            var paths = doc.CreateElement("Paths");
            root.AppendChild(paths);
            AddElement(doc, paths, "TracklistFile", TracklistFilePath);
            AddElement(doc, paths, "TracklistingFolder", TracklistingFolderPath);
            AddElement(doc, paths, "PlaylistsFolder", PlaylistsFolderPath);
            AddElement(doc, paths, "MP3ExtractPath", MP3ExtractPath);
            AddElement(doc, paths, "SparePath1", SparePath1);
            AddElement(doc, paths, "SparePath2", SparePath2);

            var mp3 = doc.CreateElement("MP3");
            root.AppendChild(mp3);
            AddElement(doc, mp3, "ListingPageUrl", MP3ListingPageUrl);
            AddElement(doc, mp3, "Username", MP3PageUsername);

            string encMP3Pwd = "";
            if (!string.IsNullOrEmpty(MP3PagePassword))
            {
                try
                {
                    byte[] plain = Encoding.UTF8.GetBytes(MP3PagePassword);
                    byte[] encrypted = ProtectedData.Protect(plain, null,
                        DataProtectionScope.CurrentUser);
                    encMP3Pwd = Convert.ToBase64String(encrypted);
                }
                catch { }
            }
            AddElement(doc, mp3, "Password", encMP3Pwd);

            AddElement(doc, root, "LastSyncTime", LastSyncTime);

            var ws = new XmlWriterSettings { Indent = true, Encoding = Encoding.UTF8 };
            using (var writer = XmlWriter.Create(SettingsFilePath, ws))
                doc.Save(writer);
        }

        private static string ReadNode(XmlDocument doc, string xpath)
        {
            var node = doc.SelectSingleNode(xpath);
            return node != null ? node.InnerText : null;
        }

        private static void AddElement(XmlDocument doc, XmlElement parent, string name, string value)
        {
            var el = doc.CreateElement(name);
            el.InnerText = value ?? "";
            parent.AppendChild(el);
        }
    }

    #endregion

    #region Logger

    class Logger
    {
        private readonly object _lock = new object();
        private const int MaxLogSizeBytes = 1048576;

        public static string LogDir { get { return AppSettings.SettingsDir; } }
        public static string LogFilePath { get { return Path.Combine(LogDir, "sync.log"); } }

        public Logger()
        {
            Directory.CreateDirectory(LogDir);
        }

        public void Info(string message) { Write("INFO ", message); }
        public void Warn(string message) { Write("WARN ", message); }

        public void Error(string message, Exception ex = null)
        {
            if (ex != null)
                Write("ERROR", message + " | " + ex.GetType().Name + ": " + ex.Message);
            else
                Write("ERROR", message);
        }

        private void Write(string level, string message)
        {
            lock (_lock)
            {
                try
                {
                    string line = string.Format("[{0:yyyy-MM-dd HH:mm:ss}] {1} {2}\r\n",
                        DateTime.Now, level, message);
                    File.AppendAllText(LogFilePath, line, Encoding.UTF8);
                }
                catch { }
            }
        }

        public void RotateIfNeeded()
        {
            lock (_lock)
            {
                try
                {
                    if (!File.Exists(LogFilePath)) return;
                    var fi = new FileInfo(LogFilePath);
                    if (fi.Length <= MaxLogSizeBytes) return;

                    string all = File.ReadAllText(LogFilePath, Encoding.UTF8);
                    int keep = Math.Min(all.Length, 512000);
                    string trimmed = all.Substring(all.Length - keep);
                    int firstNewline = trimmed.IndexOf('\n');
                    if (firstNewline >= 0)
                        trimmed = trimmed.Substring(firstNewline + 1);
                    File.WriteAllText(LogFilePath, "[Log rotated]\r\n" + trimmed, Encoding.UTF8);
                }
                catch { }
            }
        }
    }

    #endregion

    #region WebDavItem

    class WebDavItem
    {
        public string Href { get; set; }
        public string Name { get; set; }
        public bool IsCollection { get; set; }
        public long ContentLength { get; set; }
        public DateTime LastModified { get; set; }
    }

    #endregion

    #region WebDavClient

    class WebDavClient : IDisposable
    {
        private readonly HttpClient _client;
        private readonly string _baseUrl;

        public WebDavClient(string baseUrl, string username, string password)
        {
            _baseUrl = baseUrl.TrimEnd('/');
            var handler = new HttpClientHandler();
            if (!string.IsNullOrEmpty(username))
            {
                handler.Credentials = new NetworkCredential(username, password);
                handler.PreAuthenticate = true;
            }
            handler.ServerCertificateCustomValidationCallback = (msg, cert, chain, errors) => true;
            _client = new HttpClient(handler) { Timeout = TimeSpan.FromMinutes(10) };
        }

        public List<WebDavItem> ListDirectory(string remotePath)
        {
            string url = BuildUrl(remotePath);
            var request = new HttpRequestMessage(new HttpMethod("PROPFIND"), url);
            request.Headers.Add("Depth", "1");
            request.Content = new StringContent(
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                "<D:propfind xmlns:D=\"DAV:\"><D:allprop/></D:propfind>",
                Encoding.UTF8, "application/xml");

            var response = _client.SendAsync(request).GetAwaiter().GetResult();

            if ((int)response.StatusCode != 207)
            {
                response.EnsureSuccessStatusCode();
                return new List<WebDavItem>();
            }

            string xml = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
            var items = ParsePropfindResponse(xml);

            if (items.Count > 0)
                items.RemoveAt(0);

            return items;
        }

        public void UploadFile(string localPath, string remotePath)
        {
            string url = BuildUrl(remotePath);
            byte[] data = File.ReadAllBytes(localPath);
            var content = new ByteArrayContent(data);
            content.Headers.ContentType =
                new System.Net.Http.Headers.MediaTypeHeaderValue("application/octet-stream");

            var response = _client.PutAsync(url, content).GetAwaiter().GetResult();
            if (response.StatusCode != HttpStatusCode.OK &&
                response.StatusCode != HttpStatusCode.Created &&
                response.StatusCode != HttpStatusCode.NoContent)
            {
                throw new Exception("Upload failed: " + (int)response.StatusCode + " " +
                    response.ReasonPhrase + " for " + remotePath);
            }
        }

        public void DownloadFile(string remotePath, string localPath)
        {
            string url = BuildUrl(remotePath);
            var response = _client.GetAsync(url).GetAwaiter().GetResult();
            response.EnsureSuccessStatusCode();

            string dir = Path.GetDirectoryName(localPath);
            if (!string.IsNullOrEmpty(dir))
                Directory.CreateDirectory(dir);

            byte[] data = response.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult();
            File.WriteAllBytes(localPath, data);
        }

        public void EnsureDirectory(string remotePath)
        {
            string url = BuildUrl(remotePath);
            var request = new HttpRequestMessage(new HttpMethod("MKCOL"), url);
            var response = _client.SendAsync(request).GetAwaiter().GetResult();

            if (response.StatusCode != HttpStatusCode.Created &&
                (int)response.StatusCode != 405 &&
                response.StatusCode != HttpStatusCode.OK &&
                (int)response.StatusCode != 301)
            {
                // 405 = already exists, 301 = some servers redirect existing dirs
            }
        }

        public byte[] DownloadBytes(string fullUrl, string username = null, string password = null)
        {
            using (var handler = new HttpClientHandler())
            {
                if (!string.IsNullOrEmpty(username))
                {
                    handler.Credentials = new NetworkCredential(username, password);
                    handler.PreAuthenticate = true;
                }
                using (var client = new HttpClient(handler) { Timeout = TimeSpan.FromMinutes(30) })
                {
                    var response = client.GetAsync(fullUrl).GetAwaiter().GetResult();
                    response.EnsureSuccessStatusCode();
                    return response.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult();
                }
            }
        }

        public string DownloadString(string fullUrl, string username = null, string password = null)
        {
            using (var handler = new HttpClientHandler())
            {
                if (!string.IsNullOrEmpty(username))
                {
                    handler.Credentials = new NetworkCredential(username, password);
                    handler.PreAuthenticate = true;
                }
                using (var client = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(30) })
                {
                    return client.GetStringAsync(fullUrl).GetAwaiter().GetResult();
                }
            }
        }

        private string BuildUrl(string remotePath)
        {
            if (string.IsNullOrEmpty(remotePath))
                return _baseUrl + "/";
            return _baseUrl + "/" + remotePath.TrimStart('/');
        }

        private List<WebDavItem> ParsePropfindResponse(string xml)
        {
            var items = new List<WebDavItem>();
            var doc = new XmlDocument();
            doc.LoadXml(xml);

            var nsMgr = new XmlNamespaceManager(doc.NameTable);
            nsMgr.AddNamespace("D", "DAV:");

            var responses = doc.SelectNodes("//D:response", nsMgr);
            if (responses == null || responses.Count == 0)
                responses = doc.GetElementsByTagName("response");

            foreach (XmlNode resp in responses)
            {
                var item = new WebDavItem();

                var hrefNode = FindChild(resp, "href", nsMgr);
                item.Href = hrefNode != null ? hrefNode.InnerText : "";

                string name = Uri.UnescapeDataString(item.Href.TrimEnd('/'));
                int lastSlash = name.LastIndexOf('/');
                item.Name = lastSlash >= 0 ? name.Substring(lastSlash + 1) : name;

                var propstat = FindChild(resp, "propstat", nsMgr);
                if (propstat != null)
                {
                    var prop = FindChild(propstat, "prop", nsMgr);
                    if (prop != null)
                    {
                        var resType = FindChild(prop, "resourcetype", nsMgr);
                        if (resType != null)
                        {
                            var collection = FindChild(resType, "collection", nsMgr);
                            item.IsCollection = collection != null;
                        }

                        var contentLen = FindChild(prop, "getcontentlength", nsMgr);
                        if (contentLen != null)
                        {
                            long len;
                            if (long.TryParse(contentLen.InnerText, out len))
                                item.ContentLength = len;
                        }

                        var lastMod = FindChild(prop, "getlastmodified", nsMgr);
                        if (lastMod != null)
                        {
                            DateTime dt;
                            if (DateTime.TryParse(lastMod.InnerText, out dt))
                                item.LastModified = dt;
                        }
                    }
                }

                items.Add(item);
            }

            return items;
        }

        private XmlNode FindChild(XmlNode parent, string localName, XmlNamespaceManager nsMgr)
        {
            var node = parent.SelectSingleNode("D:" + localName, nsMgr);
            if (node != null) return node;

            foreach (XmlNode child in parent.ChildNodes)
            {
                if (child.LocalName.Equals(localName, StringComparison.OrdinalIgnoreCase))
                    return child;
            }
            return null;
        }

        public void Dispose()
        {
            if (_client != null) _client.Dispose();
        }
    }

    #endregion

    #region SyncResult

    class SyncResult
    {
        public bool Success { get; set; }
        public int ErrorCount { get; set; }
        public string Summary { get; set; }
    }

    #endregion

    #region MP3ZipTracker

    static class MP3ZipTracker
    {
        private const string RegistryKeyPath = @"SOFTWARE\VDJSync\CompletedZips";

        public static HashSet<string> GetCompletedZips()
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            using (var key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath))
            {
                if (key == null) return set;
                foreach (string name in key.GetValueNames())
                    set.Add(name);
            }
            return set;
        }

        public static void MarkCompleted(string zipFileName)
        {
            using (var key = Registry.CurrentUser.CreateSubKey(RegistryKeyPath))
            {
                key.SetValue(zipFileName,
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                    RegistryValueKind.String);
            }
        }

        public static List<string> ParseListingPage(string content, string baseUrl)
        {
            var urls = new List<string>();

            foreach (string rawLine in content.Split(new[] { '\r', '\n' },
                StringSplitOptions.RemoveEmptyEntries))
            {
                string line = rawLine.Trim();
                if (string.IsNullOrEmpty(line)) continue;

                if (line.EndsWith(".zip", StringComparison.OrdinalIgnoreCase) &&
                    (line.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                     line.StartsWith("https://", StringComparison.OrdinalIgnoreCase)))
                {
                    urls.Add(line);
                    continue;
                }

                var match = Regex.Match(line, @"href\s*=\s*[""']([^""']*\.zip)[""']",
                    RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    string href = match.Groups[1].Value;
                    if (href.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                        href.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
                    {
                        urls.Add(href);
                    }
                    else
                    {
                        string bUrl = baseUrl.TrimEnd('/');
                        int lastSlash = bUrl.LastIndexOf('/');
                        if (lastSlash > 8)
                            bUrl = bUrl.Substring(0, lastSlash);
                        urls.Add(bUrl + "/" + href.TrimStart('/'));
                    }
                }
            }

            return urls;
        }
    }

    #endregion

    #region SyncEngine

    class SyncEngine
    {
        private readonly AppSettings _settings;
        private readonly WebDavClient _client;
        private readonly Logger _log;
        private readonly string _pcName;
        private int _errorCount;
        private int _filesUploaded;
        private int _filesDownloaded;

        public SyncEngine(AppSettings settings, WebDavClient client, Logger log, string pcName)
        {
            _settings = settings;
            _client = client;
            _log = log;
            _pcName = pcName;
        }

        public SyncResult RunFullSync()
        {
            _errorCount = 0;
            _filesUploaded = 0;
            _filesDownloaded = 0;

            _log.RotateIfNeeded();
            _log.Info("========== Sync started ==========");

            RunTask("Task 1: Upload tracklist archive", Task1_UploadTracklist);
            RunTask("Task 2: Upload Tracklisting folder", Task2_UploadTracklistingFolder);
            RunTask("Task 3: Download Tracklisting from others", Task3_DownloadTracklisting);
            RunTask("Task 4: Upload Playlists folder", Task4_UploadPlaylistsFolder);
            RunTask("Task 5: Download Playlists from others", Task5_DownloadPlaylists);
            RunTask("Task 6: Download MP3 zips", Task6_DownloadMP3Zips);
            RunTask("Task 7: Upload log", Task7_UploadLog);

            string summary = string.Format("Sync complete. {0} uploaded, {1} downloaded, {2} error(s).",
                _filesUploaded, _filesDownloaded, _errorCount);
            _log.Info("========== " + summary + " ==========");

            return new SyncResult
            {
                Success = _errorCount == 0,
                ErrorCount = _errorCount,
                Summary = summary
            };
        }

        private void RunTask(string taskName, Action action)
        {
            try
            {
                _log.Info(taskName + " - starting...");
                action();
                _log.Info(taskName + " - done.");
            }
            catch (Exception ex)
            {
                _errorCount++;
                _log.Error(taskName + " - FAILED", ex);
            }
        }

        private void Task1_UploadTracklist()
        {
            if (!File.Exists(_settings.TracklistFilePath))
            {
                _log.Warn("Tracklist file not found: " + _settings.TracklistFilePath);
                return;
            }

            _client.EnsureDirectory("Archive");

            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss");
            string remoteName = string.Format("{0} Tracklist log {1}.txt", _pcName, timestamp);
            _client.UploadFile(_settings.TracklistFilePath, "Archive/" + remoteName);
            _filesUploaded++;
            _log.Info("  Uploaded tracklist archive: " + remoteName);
        }

        private void Task2_UploadTracklistingFolder()
        {
            if (!Directory.Exists(_settings.TracklistingFolderPath))
            {
                _log.Warn("Tracklisting folder not found: " + _settings.TracklistingFolderPath);
                return;
            }

            _client.EnsureDirectory("Tracklisting");
            _client.EnsureDirectory("Tracklisting/" + _pcName);

            UploadFolderRecursive(_settings.TracklistingFolderPath,
                "Tracklisting/" + _pcName, "zOthers");
        }

        private void Task3_DownloadTracklisting()
        {
            List<WebDavItem> items;
            try
            {
                items = _client.ListDirectory("Tracklisting");
            }
            catch
            {
                _log.Warn("Tracklisting folder not found on server, skipping download.");
                return;
            }

            string localZOthers = Path.Combine(_settings.TracklistingFolderPath, "zOthers");
            Directory.CreateDirectory(localZOthers);

            foreach (var item in items)
            {
                if (!item.IsCollection) continue;
                if (item.Name.Equals(_pcName, StringComparison.OrdinalIgnoreCase)) continue;

                string localDir = Path.Combine(localZOthers, item.Name);
                DownloadFolderRecursive("Tracklisting/" + item.Name, localDir);
            }
        }

        private void Task4_UploadPlaylistsFolder()
        {
            if (!Directory.Exists(_settings.PlaylistsFolderPath))
            {
                _log.Warn("Playlists folder not found: " + _settings.PlaylistsFolderPath);
                return;
            }

            _client.EnsureDirectory("Playlists");
            _client.EnsureDirectory("Playlists/" + _pcName);

            UploadFolderRecursive(_settings.PlaylistsFolderPath,
                "Playlists/" + _pcName, "zOthers");
        }

        private void Task5_DownloadPlaylists()
        {
            List<WebDavItem> items;
            try
            {
                items = _client.ListDirectory("Playlists");
            }
            catch
            {
                _log.Warn("Playlists folder not found on server, skipping download.");
                return;
            }

            string localZOthers = Path.Combine(_settings.PlaylistsFolderPath, "zOthers");
            Directory.CreateDirectory(localZOthers);

            foreach (var item in items)
            {
                if (!item.IsCollection) continue;
                if (item.Name.Equals(_pcName, StringComparison.OrdinalIgnoreCase)) continue;

                string localDir = Path.Combine(localZOthers, item.Name);
                DownloadFolderRecursive("Playlists/" + item.Name, localDir);
            }
        }

        private void Task6_DownloadMP3Zips()
        {
            if (string.IsNullOrEmpty(_settings.MP3ListingPageUrl))
            {
                _log.Info("  MP3 listing page URL not configured, skipping.");
                return;
            }

            if (!Directory.Exists(_settings.MP3ExtractPath))
            {
                _log.Warn("  MP3 extract path not available: " + _settings.MP3ExtractPath);
                return;
            }

            string mp3User = _settings.MP3PageUsername;
            string mp3Pass = _settings.MP3PagePassword;

            string html = _client.DownloadString(_settings.MP3ListingPageUrl, mp3User, mp3Pass);
            var zipUrls = MP3ZipTracker.ParseListingPage(html, _settings.MP3ListingPageUrl);
            var completed = MP3ZipTracker.GetCompletedZips();

            _log.Info(string.Format("  Found {0} zip(s) on server, {1} already completed.",
                zipUrls.Count, completed.Count));

            foreach (string zipUrl in zipUrls)
            {
                string zipName = Path.GetFileName(new Uri(zipUrl).LocalPath);
                if (completed.Contains(zipName))
                    continue;

                try
                {
                    _log.Info("  Downloading: " + zipName);
                    string tempPath = Path.Combine(Path.GetTempPath(), "VDJSync_" + zipName);

                    byte[] data = _client.DownloadBytes(zipUrl, mp3User, mp3Pass);
                    File.WriteAllBytes(tempPath, data);

                    _log.Info("  Extracting to: " + _settings.MP3ExtractPath);
                    ZipFile.ExtractToDirectory(tempPath, _settings.MP3ExtractPath);

                    File.Delete(tempPath);
                    MP3ZipTracker.MarkCompleted(zipName);
                    _filesDownloaded++;
                    _log.Info("  Completed: " + zipName);
                }
                catch (Exception ex)
                {
                    _errorCount++;
                    _log.Error("  Failed to process zip: " + zipName, ex);
                }
            }
        }

        private void Task7_UploadLog()
        {
            if (!File.Exists(Logger.LogFilePath))
                return;

            _client.EnsureDirectory("Logs");
            _client.UploadFile(Logger.LogFilePath, "Logs/" + _pcName + "_sync.log");
            _log.Info("  Log uploaded to server.");
        }

        private void UploadFolderRecursive(string localDir, string remoteDir, string excludeFolder)
        {
            _client.EnsureDirectory(remoteDir);

            foreach (string file in Directory.GetFiles(localDir))
            {
                string fileName = Path.GetFileName(file);
                try
                {
                    _client.UploadFile(file, remoteDir + "/" + fileName);
                    _filesUploaded++;
                }
                catch (Exception ex)
                {
                    _errorCount++;
                    _log.Error("  Failed to upload: " + fileName, ex);
                }
            }

            foreach (string subDir in Directory.GetDirectories(localDir))
            {
                string dirName = new DirectoryInfo(subDir).Name;
                if (!string.IsNullOrEmpty(excludeFolder) &&
                    dirName.Equals(excludeFolder, StringComparison.OrdinalIgnoreCase))
                    continue;

                UploadFolderRecursive(subDir, remoteDir + "/" + dirName, null);
            }
        }

        private void DownloadFolderRecursive(string remoteDir, string localDir)
        {
            Directory.CreateDirectory(localDir);

            List<WebDavItem> items;
            try
            {
                items = _client.ListDirectory(remoteDir);
            }
            catch (Exception ex)
            {
                _log.Error("  Failed to list: " + remoteDir, ex);
                _errorCount++;
                return;
            }

            foreach (var item in items)
            {
                if (item.IsCollection)
                {
                    DownloadFolderRecursive(remoteDir + "/" + item.Name,
                        Path.Combine(localDir, item.Name));
                }
                else
                {
                    try
                    {
                        string localPath = Path.Combine(localDir, item.Name);
                        _client.DownloadFile(remoteDir + "/" + item.Name, localPath);
                        _filesDownloaded++;
                    }
                    catch (Exception ex)
                    {
                        _errorCount++;
                        _log.Error("  Failed to download: " + item.Name, ex);
                    }
                }
            }
        }
    }

    #endregion

    #region SettingsForm

    class SettingsForm : Form
    {
        private TextBox _txtServerUrl, _txtUsername, _txtPassword;
        private TextBox _txtTracklistFile, _txtTracklistingFolder, _txtPlaylistsFolder;
        private TextBox _txtMP3ExtractPath, _txtSpare1, _txtSpare2;
        private TextBox _txtMP3ListingUrl, _txtMP3Username, _txtMP3Password;

        public SettingsForm()
        {
            Text = "VDJ Sync - Settings";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(580, 580);
            Font = new Font("Segoe UI", 9f);

            int y = 10;

            // WebDAV group
            var grpWebDav = new GroupBox
            {
                Text = "WebDAV Connection",
                Location = new Point(10, y),
                Size = new Size(555, 110)
            };
            Controls.Add(grpWebDav);

            AddLabel(grpWebDav, "Server URL:", 15, 25);
            _txtServerUrl = AddTextBox(grpWebDav, 110, 22, 430);

            AddLabel(grpWebDav, "Username:", 15, 55);
            _txtUsername = AddTextBox(grpWebDav, 110, 52, 170);

            AddLabel(grpWebDav, "Password:", 300, 55);
            _txtPassword = AddTextBox(grpWebDav, 380, 52, 160);
            _txtPassword.UseSystemPasswordChar = true;

            var btnTest = new Button
            {
                Text = "Test",
                Location = new Point(110, 78),
                Size = new Size(70, 25)
            };
            btnTest.Click += (s, e) => TestConnection();
            grpWebDav.Controls.Add(btnTest);

            y += 120;

            // Paths group
            var grpPaths = new GroupBox
            {
                Text = "Local Paths",
                Location = new Point(10, y),
                Size = new Size(555, 230)
            };
            Controls.Add(grpPaths);

            int py = 25;
            AddLabel(grpPaths, "Tracklist File:", 15, py);
            _txtTracklistFile = AddTextBox(grpPaths, 130, py - 3, 370);
            AddBrowseButton(grpPaths, _txtTracklistFile, true, 510, py - 4);

            py += 32;
            AddLabel(grpPaths, "Tracklisting:", 15, py);
            _txtTracklistingFolder = AddTextBox(grpPaths, 130, py - 3, 370);
            AddBrowseButton(grpPaths, _txtTracklistingFolder, false, 510, py - 4);

            py += 32;
            AddLabel(grpPaths, "Playlists:", 15, py);
            _txtPlaylistsFolder = AddTextBox(grpPaths, 130, py - 3, 370);
            AddBrowseButton(grpPaths, _txtPlaylistsFolder, false, 510, py - 4);

            py += 32;
            AddLabel(grpPaths, "MP3 Extract:", 15, py);
            _txtMP3ExtractPath = AddTextBox(grpPaths, 130, py - 3, 370);
            AddBrowseButton(grpPaths, _txtMP3ExtractPath, false, 510, py - 4);

            py += 32;
            AddLabel(grpPaths, "Spare Path 1:", 15, py);
            _txtSpare1 = AddTextBox(grpPaths, 130, py - 3, 370);
            AddBrowseButton(grpPaths, _txtSpare1, false, 510, py - 4);

            py += 32;
            AddLabel(grpPaths, "Spare Path 2:", 15, py);
            _txtSpare2 = AddTextBox(grpPaths, 130, py - 3, 370);
            AddBrowseButton(grpPaths, _txtSpare2, false, 510, py - 4);

            y += 240;

            // MP3 group
            var grpMP3 = new GroupBox
            {
                Text = "MP3 Downloads",
                Location = new Point(10, y),
                Size = new Size(555, 95)
            };
            Controls.Add(grpMP3);

            AddLabel(grpMP3, "Listing URL:", 15, 25);
            _txtMP3ListingUrl = AddTextBox(grpMP3, 110, 22, 430);

            AddLabel(grpMP3, "Username:", 15, 58);
            _txtMP3Username = AddTextBox(grpMP3, 110, 55, 170);

            AddLabel(grpMP3, "Password:", 300, 58);
            _txtMP3Password = AddTextBox(grpMP3, 380, 55, 160);
            _txtMP3Password.UseSystemPasswordChar = true;

            y += 105;

            // Buttons
            var btnSave = new Button
            {
                Text = "Save",
                Location = new Point(390, y),
                Size = new Size(80, 32)
            };
            btnSave.Click += (s, e) => SaveSettings();
            Controls.Add(btnSave);

            var btnCancel = new Button
            {
                Text = "Cancel",
                Location = new Point(480, y),
                Size = new Size(80, 32),
                DialogResult = DialogResult.Cancel
            };
            Controls.Add(btnCancel);

            AcceptButton = btnSave;
            CancelButton = btnCancel;

            LoadSettings();
        }

        private void LoadSettings()
        {
            var s = AppSettings.Load();
            _txtServerUrl.Text = s.WebDavUrl;
            _txtUsername.Text = s.WebDavUsername;
            _txtPassword.Text = s.WebDavPassword;
            _txtTracklistFile.Text = s.TracklistFilePath;
            _txtTracklistingFolder.Text = s.TracklistingFolderPath;
            _txtPlaylistsFolder.Text = s.PlaylistsFolderPath;
            _txtMP3ExtractPath.Text = s.MP3ExtractPath;
            _txtSpare1.Text = s.SparePath1;
            _txtSpare2.Text = s.SparePath2;
            _txtMP3ListingUrl.Text = s.MP3ListingPageUrl;
            _txtMP3Username.Text = s.MP3PageUsername;
            _txtMP3Password.Text = s.MP3PagePassword;
        }

        private void SaveSettings()
        {
            string url = _txtServerUrl.Text.Trim();
            if (!string.IsNullOrEmpty(url) &&
                !url.StartsWith("http://", StringComparison.OrdinalIgnoreCase) &&
                !url.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show("Server URL must start with http:// or https://",
                    "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var s = new AppSettings
            {
                WebDavUrl = url,
                WebDavUsername = _txtUsername.Text.Trim(),
                WebDavPassword = _txtPassword.Text,
                TracklistFilePath = _txtTracklistFile.Text.Trim(),
                TracklistingFolderPath = _txtTracklistingFolder.Text.Trim(),
                PlaylistsFolderPath = _txtPlaylistsFolder.Text.Trim(),
                MP3ExtractPath = _txtMP3ExtractPath.Text.Trim(),
                SparePath1 = _txtSpare1.Text.Trim(),
                SparePath2 = _txtSpare2.Text.Trim(),
                MP3ListingPageUrl = _txtMP3ListingUrl.Text.Trim(),
                MP3PageUsername = _txtMP3Username.Text.Trim(),
                MP3PagePassword = _txtMP3Password.Text
            };
            s.Save();

            DialogResult = DialogResult.OK;
            Close();
        }

        private void TestConnection()
        {
            string url = _txtServerUrl.Text.Trim();
            if (string.IsNullOrEmpty(url))
            {
                MessageBox.Show("Enter a server URL first.", "Test",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Cursor = Cursors.WaitCursor;
            try
            {
                using (var client = new WebDavClient(url, _txtUsername.Text.Trim(), _txtPassword.Text))
                {
                    client.ListDirectory("");
                }
                MessageBox.Show("Connection successful!", "Test",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Connection failed:\n" + ex.Message, "Test",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private Label AddLabel(Control parent, string text, int x, int y)
        {
            var lbl = new Label
            {
                Text = text,
                Location = new Point(x, y),
                AutoSize = true
            };
            parent.Controls.Add(lbl);
            return lbl;
        }

        private TextBox AddTextBox(Control parent, int x, int y, int width)
        {
            var txt = new TextBox
            {
                Location = new Point(x, y),
                Size = new Size(width, 23)
            };
            parent.Controls.Add(txt);
            return txt;
        }

        private Button AddBrowseButton(Control parent, TextBox target, bool isFile, int x, int y)
        {
            var btn = new Button
            {
                Text = "...",
                Location = new Point(x, y),
                Size = new Size(30, 25)
            };
            btn.Click += (s, e) =>
            {
                if (isFile)
                {
                    using (var dlg = new OpenFileDialog())
                    {
                        dlg.Filter = "Text Files|*.txt|All Files|*.*";
                        if (!string.IsNullOrEmpty(target.Text))
                            dlg.InitialDirectory = Path.GetDirectoryName(target.Text);
                        if (dlg.ShowDialog() == DialogResult.OK)
                            target.Text = dlg.FileName;
                    }
                }
                else
                {
                    using (var dlg = new FolderBrowserDialog())
                    {
                        if (!string.IsNullOrEmpty(target.Text))
                            dlg.SelectedPath = target.Text;
                        if (dlg.ShowDialog() == DialogResult.OK)
                            target.Text = dlg.SelectedPath;
                    }
                }
            };
            parent.Controls.Add(btn);
            return btn;
        }
    }

    #endregion
}
