using System;
using System.Diagnostics;
using System.IO;
using System.Text.Json;
using System.Windows.Forms;
using Microsoft.Web.WebView2.WinForms;
using Microsoft.Web.WebView2.Core;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Linq;

namespace KaggleApp
{
    public class Main : Form
    {
        private WebView2 webView;
        private PythonServer pyServer;
        private DataManager dataManager;
        private MenuStrip menuStrip;
        private StatusStrip statusStrip;
        private ToolStripStatusLabel statusLabel;

        [DllImport("kernel32.dll")]
        private static extern bool AllocConsole();

        public Main()
        {
            this.Text = "Kaggle App - Comment Analyzer";
            this.Width = 1000;
            this.Height = 700;

            // Allocate console
            AllocConsole();
            Console.WriteLine("Console attached. Python output will appear here.");

            // Initialize data manager
            string dataFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data");
            dataManager = new DataManager(dataFolder);

            // Setup UI
            SetupMenu();
            SetupStatusBar();

            // Initialize WebView2
            webView = new WebView2 { Dock = DockStyle.Fill };
            this.Controls.Add(webView);

            InitializeAsync();

            // Start Python server
            string scriptsFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Scripts");
            string scriptPath = Path.Combine(scriptsFolder, "algo.py");
            pyServer = new PythonServer("python", scriptPath);

            UpdateStatus("Ready");
        }

        private void SetupMenu()
        {
            menuStrip = new MenuStrip();

            // File Menu
            var fileMenu = new ToolStripMenuItem("File");
            fileMenu.DropDownItems.Add("Load Comments from File", null, LoadCommentsFromFile);
            fileMenu.DropDownItems.Add("Load Session", null, LoadSessionDialog);
            fileMenu.DropDownItems.Add(new ToolStripSeparator());
            fileMenu.DropDownItems.Add("Exit", null, (s, e) => this.Close());

            // Export Menu
            var exportMenu = new ToolStripMenuItem("Export");
            exportMenu.DropDownItems.Add("Export All Results", null, (s, e) => ExportToCsv(null));
            exportMenu.DropDownItems.Add("Export Toxic Only", null, (s, e) => ExportToCsv("Toxic"));
            exportMenu.DropDownItems.Add("Export Clean Only", null, (s, e) => ExportToCsv("Clean"));
            exportMenu.DropDownItems.Add(new ToolStripSeparator());
            exportMenu.DropDownItems.Add("Export Current Session", null, ExportCurrentSession);

            // Sessions Menu
            var sessionsMenu = new ToolStripMenuItem("Sessions");
            sessionsMenu.DropDownItems.Add("View All Sessions", null, ViewAllSessions);
            sessionsMenu.DropDownItems.Add("Delete Session", null, DeleteSessionDialog);
            sessionsMenu.DropDownItems.Add("Clear All Sessions", null, ClearAllSessions);

            // Statistics Menu
            var statsMenu = new ToolStripMenuItem("Statistics");
            statsMenu.DropDownItems.Add("Show Current Stats", null, ShowCurrentStats);
            statsMenu.DropDownItems.Add("Show All-Time Stats", null, ShowAllTimeStats);

            menuStrip.Items.Add(fileMenu);
            menuStrip.Items.Add(exportMenu);
            menuStrip.Items.Add(sessionsMenu);
            menuStrip.Items.Add(statsMenu);

            this.MainMenuStrip = menuStrip;
            this.Controls.Add(menuStrip);
        }

        private void SetupStatusBar()
        {
            statusStrip = new StatusStrip();
            statusLabel = new ToolStripStatusLabel("Ready");
            statusStrip.Items.Add(statusLabel);
            this.Controls.Add(statusStrip);
        }

        private void UpdateStatus(string message)
        {
            if (statusLabel.Owner.InvokeRequired)
            {
                statusLabel.Owner.Invoke(new Action(() => statusLabel.Text = message));
            }
            else
            {
                statusLabel.Text = message;
            }
        }

        private async void InitializeAsync()
        {
            await webView.EnsureCoreWebView2Async(null);
            webView.CoreWebView2.Settings.IsWebMessageEnabled = true;

            string htmlPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Web", "index.html");
            string htmlUri = new Uri(htmlPath).AbsoluteUri;
            webView.CoreWebView2.Navigate(htmlUri);

            webView.CoreWebView2.WebMessageReceived += WebMessageReceived;
        }

        private void WebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs args)
        {
            try
            {
                var msg = args.TryGetWebMessageAsString();
                var doc = JsonDocument.Parse(msg);

                string action = doc.RootElement.TryGetProperty("action", out var actionProp)
                    ? actionProp.GetString()
                    : "analyze";

                switch (action)
                {
                    case "analyze":
                        HandleAnalyze(doc);
                        break;

                    case "filter":
                        HandleFilter(doc);
                        break;

                    case "export":
                        string filterType = doc.RootElement.TryGetProperty("filterType", out var ft)
                            ? ft.GetString()
                            : null;
                        ExportToCsv(filterType);
                        break;

                    case "getStats":
                        SendStatistics();
                        break;
                    case "loadFile":
                        this.Invoke(new Action(() => LoadCommentsFromFile(null, null)));
                        break;

                    default:
                        HandleAnalyze(doc);
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error in WebMessageReceived: " + ex);
                UpdateStatus("Error processing request");
            }
        }

        private void HandleAnalyze(JsonDocument doc)
        {
            var comments = new List<string>();
            foreach (var c in doc.RootElement.GetProperty("comments").EnumerateArray())
                comments.Add(c.GetString());

            UpdateStatus($"Analyzing {comments.Count} comments...");

            // Send to Python
            var resultsJson = pyServer.SendRequest(comments);

            // Parse and store results
            var results = ParseResults(comments, resultsJson);

            // Save session
            string sessionId = dataManager.SaveSession(results);
            UpdateStatus($"Analysis complete. Session saved: {sessionId}");

            // Send results back to JS
            webView.CoreWebView2.PostWebMessageAsString(resultsJson);

            // Update statistics
            SendStatistics();
        }

        private void HandleFilter(JsonDocument doc)
        {
            string filterType = doc.RootElement.GetProperty("filterType").GetString();
            var currentSession = dataManager.GetCurrentSession();

            if (currentSession != null)
            {
                var filtered = currentSession.Results
                    .Where(r => string.IsNullOrEmpty(filterType) || r.Status == filterType)
                    .ToList();

                var response = new
                {
                    type = "filtered",
                    results = filtered
                };

                webView.CoreWebView2.PostWebMessageAsString(JsonSerializer.Serialize(response));
                UpdateStatus($"Filtered: {filtered.Count} results");
            }
        }

        private List<CommentResult> ParseResults(List<string> comments, string resultsJson)
        {
            var results = new List<CommentResult>();
            var doc = JsonDocument.Parse(resultsJson);
            var resultsArray = doc.RootElement.GetProperty("results").EnumerateArray();

            int i = 0;
            foreach (var result in resultsArray)
            {
                results.Add(new CommentResult
                {
                    Comment = i < comments.Count ? comments[i] : "",
                    Probability = result.GetProperty("probability").GetDouble(),
                    Status = result.GetProperty("status").GetString(),
                    Rule = result.GetProperty("rule").GetString(),
                    Timestamp = DateTime.Now
                });
                i++;
            }

            return results;
        }

        private void SendStatistics()
        {
            var currentSession = dataManager.GetCurrentSession();
            if (currentSession != null)
            {
                var stats = CalculateStats(currentSession.Results);
                var response = new
                {
                    type = "stats",
                    data = stats
                };

                webView.CoreWebView2.PostWebMessageAsString(JsonSerializer.Serialize(response));
            }
        }

        private object CalculateStats(List<CommentResult> results)
        {
            if (results.Count == 0)
            {
                return new
                {
                    total = 0,
                    toxic = 0,
                    clean = 0,
                    avgProbability = 0.0,
                    topRule = "N/A"
                };
            }

            return new
            {
                total = results.Count,
                toxic = results.Count(r => r.Status == "Toxic"),
                clean = results.Count(r => r.Status == "Clean"),
                avgProbability = results.Average(r => r.Probability),
                topRule = results.GroupBy(r => r.Rule)
                                 .OrderByDescending(g => g.Count())
                                 .FirstOrDefault()?.Key ?? "N/A"
            };
        }

        private void ExportToCsv(string filterStatus)
        {
            var currentSession = dataManager.GetCurrentSession();
            if (currentSession == null || currentSession.Results.Count == 0)
            {
                MessageBox.Show("No results to export.", "Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var saveDialog = new SaveFileDialog
            {
                Filter = "CSV files (*.csv)|*.csv",
                FileName = $"results_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
            };

            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                var filtered = string.IsNullOrEmpty(filterStatus)
                    ? currentSession.Results
                    : currentSession.Results.Where(r => r.Status == filterStatus).ToList();

                using (var writer = new StreamWriter(saveDialog.FileName))
                {
                    writer.WriteLine("Timestamp,Comment,Probability,Status,Rule");
                    foreach (var result in filtered)
                    {
                        writer.WriteLine($"\"{result.Timestamp:yyyy-MM-dd HH:mm:ss}\",\"{EscapeCsv(result.Comment)}\",{result.Probability},\"{result.Status}\",\"{result.Rule}\"");
                    }
                }
                MessageBox.Show($"Exported {filtered.Count} results to {saveDialog.FileName}", "Export Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                UpdateStatus($"Exported {filtered.Count} results");
            }
        }

        private void ExportCurrentSession(object sender, EventArgs e)
        {
            var currentSession = dataManager.GetCurrentSession();
            if (currentSession == null)
            {
                MessageBox.Show("No active session to export.", "Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var saveDialog = new SaveFileDialog
            {
                Filter = "JSON files (*.json)|*.json",
                FileName = $"session_{currentSession.SessionId}.json"
            };

            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                string json = JsonSerializer.Serialize(currentSession, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(saveDialog.FileName, json);
                MessageBox.Show($"Session exported successfully.", "Export Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void LoadCommentsFromFile(object sender, EventArgs e)
        {
            var openDialog = new OpenFileDialog
            {
                Filter = "Text files (*.txt)|*.txt|CSV files (*.csv)|*.csv|All files (*.*)|*.*"
            };

            if (openDialog.ShowDialog() == DialogResult.OK)
            {
                var comments = File.ReadAllLines(openDialog.FileName)
                    .Where(line => !string.IsNullOrWhiteSpace(line))
                    .ToList();

                if (comments.Count == 0)
                {
                    MessageBox.Show("No comments found in file.", "Load Comments", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                UpdateStatus($"Processing {comments.Count} comments from file...");

                // Send to Python
                var resultsJson = pyServer.SendRequest(comments);

                // Parse and store
                var results = ParseResults(comments, resultsJson);
                string sessionId = dataManager.SaveSession(results);

                // Send to JS
                webView.CoreWebView2.PostWebMessageAsString(resultsJson);

                UpdateStatus($"Loaded and analyzed {comments.Count} comments. Session: {sessionId}");
                SendStatistics();
            }
        }

        private void LoadSessionDialog(object sender, EventArgs e)
        {
            var sessions = dataManager.GetAllSessions();
            if (sessions.Count == 0)
            {
                MessageBox.Show("No saved sessions found.", "Load Session", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var form = new Form
            {
                Text = "Load Session",
                Width = 500,
                Height = 400,
                StartPosition = FormStartPosition.CenterParent
            };

            var listBox = new ListBox
            {
                Dock = DockStyle.Fill,
                DisplayMember = "Display"
            };

            var sessionList = sessions.Select(s => new
            {
                Session = s,
                Display = $"{s.SessionId} - {s.Timestamp:yyyy-MM-dd HH:mm:ss} ({s.Results.Count} results)"
            }).ToList();

            listBox.DataSource = sessionList;

            var btnPanel = new Panel { Dock = DockStyle.Bottom, Height = 40 };
            var btnLoad = new Button { Text = "Load", DialogResult = DialogResult.OK, Location = new System.Drawing.Point(10, 5) };
            var btnCancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Location = new System.Drawing.Point(100, 5) };

            btnPanel.Controls.Add(btnLoad);
            btnPanel.Controls.Add(btnCancel);
            form.Controls.Add(listBox);
            form.Controls.Add(btnPanel);

            if (form.ShowDialog() == DialogResult.OK && listBox.SelectedItem != null)
            {
                var selected = ((dynamic)listBox.SelectedItem).Session as AnalysisSession;
                dataManager.SetCurrentSession(selected.SessionId);

                // Send to UI
                var resultsJson = JsonSerializer.Serialize(new
                {
                    comments = selected.Results.Select(r => r.Comment).ToList(),
                    results = selected.Results
                });

                webView.CoreWebView2.PostWebMessageAsString(resultsJson);
                UpdateStatus($"Loaded session: {selected.SessionId}");
                SendStatistics();
            }
        }

        private void ViewAllSessions(object sender, EventArgs e)
        {
            var sessions = dataManager.GetAllSessions();
            if (sessions.Count == 0)
            {
                MessageBox.Show("No saved sessions found.", "Sessions", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var form = new Form
            {
                Text = "All Sessions",
                Width = 700,
                Height = 500,
                StartPosition = FormStartPosition.CenterParent
            };

            var dataGrid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoGenerateColumns = false,
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect
            };

            dataGrid.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "SessionId", HeaderText = "Session ID", Width = 200 });
            dataGrid.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Timestamp", HeaderText = "Date/Time", Width = 150 });
            dataGrid.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "ResultCount", HeaderText = "Results", Width = 80 });
            dataGrid.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "ToxicCount", HeaderText = "Toxic", Width = 80 });
            dataGrid.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "CleanCount", HeaderText = "Clean", Width = 80 });

            var sessionData = sessions.Select(s => new
            {
                s.SessionId,
                s.Timestamp,
                ResultCount = s.Results.Count,
                ToxicCount = s.Results.Count(r => r.Status == "Toxic"),
                CleanCount = s.Results.Count(r => r.Status == "Clean")
            }).ToList();

            dataGrid.DataSource = sessionData;
            form.Controls.Add(dataGrid);
            form.ShowDialog();
        }

        private void DeleteSessionDialog(object sender, EventArgs e)
        {
            var sessions = dataManager.GetAllSessions();
            if (sessions.Count == 0)
            {
                MessageBox.Show("No sessions to delete.", "Delete Session", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var form = new Form
            {
                Text = "Delete Session",
                Width = 500,
                Height = 400,
                StartPosition = FormStartPosition.CenterParent
            };

            var listBox = new ListBox
            {
                Dock = DockStyle.Fill,
                DisplayMember = "Display"
            };

            var sessionList = sessions.Select(s => new
            {
                Session = s,
                Display = $"{s.SessionId} - {s.Timestamp:yyyy-MM-dd HH:mm:ss}"
            }).ToList();

            listBox.DataSource = sessionList;

            var btnPanel = new Panel { Dock = DockStyle.Bottom, Height = 40 };
            var btnDelete = new Button { Text = "Delete", DialogResult = DialogResult.OK, Location = new System.Drawing.Point(10, 5) };
            var btnCancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Location = new System.Drawing.Point(100, 5) };

            btnPanel.Controls.Add(btnDelete);
            btnPanel.Controls.Add(btnCancel);
            form.Controls.Add(listBox);
            form.Controls.Add(btnPanel);

            if (form.ShowDialog() == DialogResult.OK && listBox.SelectedItem != null)
            {
                var selected = ((dynamic)listBox.SelectedItem).Session as AnalysisSession;
                if (MessageBox.Show($"Are you sure you want to delete session {selected.SessionId}?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    dataManager.DeleteSession(selected.SessionId);
                    MessageBox.Show("Session deleted.", "Delete Session", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    UpdateStatus($"Deleted session: {selected.SessionId}");
                }
            }
        }

        private void ClearAllSessions(object sender, EventArgs e)
        {
            if (MessageBox.Show("Are you sure you want to delete ALL sessions? This cannot be undone.", "Confirm Clear", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                dataManager.ClearAllSessions();
                MessageBox.Show("All sessions cleared.", "Clear Sessions", MessageBoxButtons.OK, MessageBoxIcon.Information);
                UpdateStatus("All sessions cleared");
            }
        }

        private void ShowCurrentStats(object sender, EventArgs e)
        {
            var currentSession = dataManager.GetCurrentSession();
            if (currentSession == null || currentSession.Results.Count == 0)
            {
                MessageBox.Show("No current session data available.", "Statistics", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var stats = CalculateStats(currentSession.Results);
            var statsText = $@"Current Session Statistics
Session ID: {currentSession.SessionId}
Timestamp: {currentSession.Timestamp:yyyy-MM-dd HH:mm:ss}

Total Comments: {((dynamic)stats).total}
Toxic: {((dynamic)stats).toxic}
Clean: {((dynamic)stats).clean}
Average Probability: {((dynamic)stats).avgProbability:F4}
Top Rule: {((dynamic)stats).topRule}";

            MessageBox.Show(statsText, "Current Session Statistics", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ShowAllTimeStats(object sender, EventArgs e)
        {
            var allSessions = dataManager.GetAllSessions();
            if (allSessions.Count == 0)
            {
                MessageBox.Show("No session data available.", "Statistics", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var allResults = allSessions.SelectMany(s => s.Results).ToList();
            var stats = CalculateStats(allResults);

            var statsText = $@"All-Time Statistics
Total Sessions: {allSessions.Count}
Total Comments Analyzed: {((dynamic)stats).total}
Toxic: {((dynamic)stats).toxic}
Clean: {((dynamic)stats).clean}
Average Probability: {((dynamic)stats).avgProbability:F4}
Top Rule: {((dynamic)stats).topRule}

First Session: {allSessions.Min(s => s.Timestamp):yyyy-MM-dd HH:mm:ss}
Last Session: {allSessions.Max(s => s.Timestamp):yyyy-MM-dd HH:mm:ss}";

            MessageBox.Show(statsText, "All-Time Statistics", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private string EscapeCsv(string value)
        {
            return value?.Replace("\"", "\"\"") ?? "";
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            pyServer?.Stop();
            base.OnFormClosing(e);
        }
    }

    // ------------------------------
    // Data Models
    // ------------------------------
    public class CommentResult
    {
        public string Comment { get; set; }
        public double Probability { get; set; }
        public string Status { get; set; }
        public string Rule { get; set; }
        public DateTime Timestamp { get; set; }
    }

    public class AnalysisSession
    {
        public string SessionId { get; set; }
        public DateTime Timestamp { get; set; }
        public List<CommentResult> Results { get; set; }
    }

    // ------------------------------
    // Data Manager - CRUD Operations
    // ------------------------------
    public class DataManager
    {
        private string dataFolder;
        private string sessionsFile;
        private Dictionary<string, AnalysisSession> sessions;
        private string currentSessionId;

        public DataManager(string dataFolder)
        {
            this.dataFolder = dataFolder;
            this.sessionsFile = Path.Combine(dataFolder, "sessions.json");

            Directory.CreateDirectory(dataFolder);
            LoadSessions();
        }

        private void LoadSessions()
        {
            if (File.Exists(sessionsFile))
            {
                try
                {
                    string json = File.ReadAllText(sessionsFile);
                    sessions = JsonSerializer.Deserialize<Dictionary<string, AnalysisSession>>(json)
                        ?? new Dictionary<string, AnalysisSession>();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error loading sessions: {ex.Message}");
                    sessions = new Dictionary<string, AnalysisSession>();
                }
            }
            else
            {
                sessions = new Dictionary<string, AnalysisSession>();
            }
        }

        private void SaveSessions()
        {
            try
            {
                string json = JsonSerializer.Serialize(sessions, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(sessionsFile, json);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving sessions: {ex.Message}");
            }
        }

        // Create
        public string SaveSession(List<CommentResult> results)
        {
            string sessionId = $"session_{DateTime.Now:yyyyMMdd_HHmmss}_{Guid.NewGuid().ToString().Substring(0, 8)}";

            var session = new AnalysisSession
            {
                SessionId = sessionId,
                Timestamp = DateTime.Now,
                Results = results
            };

            sessions[sessionId] = session;
            currentSessionId = sessionId;
            SaveSessions();

            return sessionId;
        }

        // Read
        public AnalysisSession GetSession(string sessionId)
        {
            return sessions.ContainsKey(sessionId) ? sessions[sessionId] : null;
        }

        public AnalysisSession GetCurrentSession()
        {
            return !string.IsNullOrEmpty(currentSessionId) ? GetSession(currentSessionId) : null;
        }

        public List<AnalysisSession> GetAllSessions()
        {
            return sessions.Values.OrderByDescending(s => s.Timestamp).ToList();
        }

        // Update
        public void UpdateSession(string sessionId, List<CommentResult> results)
        {
            if (sessions.ContainsKey(sessionId))
            {
                sessions[sessionId].Results = results;
                SaveSessions();
            }
        }

        public void SetCurrentSession(string sessionId)
        {
            if (sessions.ContainsKey(sessionId))
            {
                currentSessionId = sessionId;
            }
        }

        // Delete
        public void DeleteSession(string sessionId)
        {
            if (sessions.ContainsKey(sessionId))
            {
                sessions.Remove(sessionId);
                if (currentSessionId == sessionId)
                    currentSessionId = null;
                SaveSessions();
            }
        }

        public void ClearAllSessions()
        {
            sessions.Clear();
            currentSessionId = null;
            SaveSessions();
        }

        // Statistics
        public int GetTotalSessions()
        {
            return sessions.Count;
        }

        public int GetTotalComments()
        {
            return sessions.Values.Sum(s => s.Results.Count);
        }
    }

    // ------------------------------
    // Python persistent process helper
    // ------------------------------
    public class PythonServer : IDisposable
    {
        private Process pythonProcess;
        private StreamWriter pythonStdin;
        private StreamReader pythonStdout;

        public PythonServer(string pythonExePath, string scriptPath)
        {
            var psi = new ProcessStartInfo
            {
                FileName = pythonExePath,
                Arguments = $"\"{scriptPath}\"",
                UseShellExecute = false,
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            pythonProcess = Process.Start(psi);
            pythonStdin = pythonProcess.StandardInput;
            pythonStdout = pythonProcess.StandardOutput;

            // Wait for Python server ready message
            string readyMsg = pythonStdout.ReadLine();
            Console.WriteLine(readyMsg);

            // Capture stderr asynchronously
            pythonProcess.ErrorDataReceived += (s, e) =>
            {
                if (!string.IsNullOrEmpty(e.Data))
                    Console.WriteLine("--- Python stderr ---\n" + e.Data);
            };
            pythonProcess.BeginErrorReadLine();
        }

        public string SendRequest(List<string> comments)
        {
            try
            {
                string jsonInput = JsonSerializer.Serialize(new { comments });
                pythonStdin.WriteLine(jsonInput);
                pythonStdin.Flush();

                string jsonResponse = pythonStdout.ReadLine();
                return jsonResponse;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error sending request to Python: " + ex);
                return JsonSerializer.Serialize(new
                {
                    comments,
                    results = comments.ConvertAll(c => new
                    {
                        probability = 0.0,
                        status = "Error",
                        rule = ""
                    })
                });
            }
        }

        public void Stop()
        {
            try
            {
                if (!pythonProcess.HasExited)
                    pythonProcess.Kill(true);
            }
            catch { }

            pythonStdin?.Dispose();
            pythonStdout?.Dispose();
            pythonProcess?.Dispose();
        }

        public void Dispose() => Stop();
    }
}