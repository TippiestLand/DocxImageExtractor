using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;

namespace DocxImageExtractor
{
    public partial class Form1 : Form
    {
        private static readonly Color BgDark = Color.FromArgb(24, 26, 32);
        private static readonly Color BgPanel = Color.FromArgb(32, 35, 43);
        private static readonly Color BgDrop = Color.FromArgb(38, 42, 52);
        private static readonly Color BgDropHover = Color.FromArgb(46, 52, 66);
        private static readonly Color BorderSoft = Color.FromArgb(58, 63, 76);
        private static readonly Color TextMain = Color.FromArgb(230, 233, 239);
        private static readonly Color TextDim = Color.FromArgb(140, 146, 160);
        private static readonly Color Danger = Color.FromArgb(240, 90, 90);

        private static readonly Color Accent = Color.FromArgb(139, 92, 246);  // #8B5CF6
        private static readonly Color AccentHover = Color.FromArgb(160, 120, 255);
        private static readonly Color AccentDown = Color.FromArgb(110, 72, 210);

        private static readonly Color Ghost = Color.FromArgb(38, 42, 52);
        private static readonly Color GhostHover = Color.FromArgb(52, 57, 70);
        private static readonly Color GhostDown = Color.FromArgb(30, 34, 42);

        private static readonly HashSet<string> AllowedExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".emf", ".wmf", ".svg"
        };

        private TextBox txtOutputFolder = null!;
        private RoundedButton btnChooseFolder = null!;
        private RoundedButton btnClearLog = null!;
        private RoundedButton btnSaveLog = null!;
        private Label lblFolderError = null!;
        private Panel dropPanel = null!;
        private Label lblDropHint = null!;
        private ListBox lstLog = null!;

        private bool isBusy = false;
        private bool isDragHover = false;

        public Form1()
        {
            InitializeComponent();
            BuildUi();
        }

        private void BuildUi()
        {
            this.Text = "Docx Image Extractor";
            this.BackColor = BgDark;
            this.ForeColor = TextMain;
            this.Font = new Font("Segoe UI", 10);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.DoubleBuffered = true;

            var topPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 130,
                Padding = new Padding(20, 18, 20, 8),
                BackColor = BgDark
            };

            txtOutputFolder = new TextBox
            {
                Dock = DockStyle.Top,
                Height = 38,
                Font = new Font("Segoe UI", 10),
                BackColor = BgPanel,
                ForeColor = TextMain,
                BorderStyle = BorderStyle.FixedSingle,
                ReadOnly = true
            };

            var gap1 = new Panel { Dock = DockStyle.Top, Height = 8, BackColor = BgDark };

            btnChooseFolder = new RoundedButton
            {
                Text = "Выбрать папку",
                Dock = DockStyle.Top,
                Height = 40,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                BackColor = Accent,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                CornerRadius = 10,
                HoverBackColor = AccentHover,
                DownBackColor = AccentDown
            };
            btnChooseFolder.FlatAppearance.BorderSize = 0;
            btnChooseFolder.Click += BtnChooseFolder_Click;

            lblFolderError = new Label
            {
                Text = "Выберите папку отправки",
                ForeColor = Danger,
                AutoSize = false,
                Visible = false,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                Dock = DockStyle.Top,
                Height = 22,
                TextAlign = ContentAlignment.MiddleLeft,
                BackColor = Color.Transparent
            };

            topPanel.Controls.Add(lblFolderError);
            topPanel.Controls.Add(btnChooseFolder);
            topPanel.Controls.Add(gap1);
            topPanel.Controls.Add(txtOutputFolder);

            var centerWrap = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(20, 6, 20, 6),
                BackColor = BgDark
            };

            dropPanel = new Panel
            {
                Dock = DockStyle.Fill,
                AllowDrop = true,
                BackColor = BgDrop,
                Padding = new Padding(2)
            };
            dropPanel.Paint += DropPanel_Paint;
            dropPanel.DragEnter += DropPanel_DragEnter;
            dropPanel.DragLeave += DropPanel_DragLeave;
            dropPanel.DragDrop += DropPanel_DragDrop;
            dropPanel.Resize += (s, e) => dropPanel.Invalidate();

            lblDropHint = new Label
            {
                Text = "Перетащите сюда .docx файл",
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 18, FontStyle.Bold),
                ForeColor = TextMain,
                Dock = DockStyle.Fill,
                BackColor = Color.Transparent
            };
            dropPanel.Controls.Add(lblDropHint);

            centerWrap.Controls.Add(dropPanel);

            var bottomWrap = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 240,
                Padding = new Padding(20, 6, 20, 18),
                BackColor = BgDark
            };

            var headerRow = new Panel
            {
                Dock = DockStyle.Top,
                Height = 34,
                BackColor = BgDark
            };

            var logHeader = new Label
            {
                Text = "Журнал",
                Dock = DockStyle.Left,
                Width = 100,
                ForeColor = TextDim,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft,
                BackColor = Color.Transparent
            };

            btnSaveLog = new RoundedButton
            {
                Text = "Скачать журнал (TXT)",
                Dock = DockStyle.Right,
                Width = 210,
                Height = 30,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                BackColor = Ghost,
                ForeColor = TextMain,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                CornerRadius = 0,
                HoverBackColor = GhostHover,
                DownBackColor = GhostDown
            };
            btnSaveLog.FlatAppearance.BorderSize = 0;
            btnSaveLog.Click += BtnSaveLog_Click;

            var gapBtns = new Panel { Dock = DockStyle.Right, Width = 8, BackColor = BgDark };

            btnClearLog = new RoundedButton
            {
                Text = "Очистить журнал",
                Dock = DockStyle.Right,
                Width = 180,
                Height = 30,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                BackColor = Ghost,
                ForeColor = TextMain,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                CornerRadius = 0,
                HoverBackColor = GhostHover,
                DownBackColor = GhostDown
            };
            btnClearLog.FlatAppearance.BorderSize = 0;
            btnClearLog.Click += BtnClearLog_Click;

            headerRow.Controls.Add(logHeader);
            headerRow.Controls.Add(gapBtns);
            headerRow.Controls.Add(btnClearLog);
            headerRow.Controls.Add(btnSaveLog);

            lstLog = new ListBox
            {
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 9.5f),
                HorizontalScrollbar = true,
                BackColor = BgPanel,
                ForeColor = TextMain,
                BorderStyle = BorderStyle.FixedSingle
            };

            bottomWrap.Controls.Add(lstLog);
            bottomWrap.Controls.Add(headerRow);

            this.Controls.Add(centerWrap);
            this.Controls.Add(bottomWrap);
            this.Controls.Add(topPanel);
        }

        private void DropPanel_Paint(object sender, PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            var rect = new Rectangle(0, 0, dropPanel.Width - 1, dropPanel.Height - 1);
            int radius = 16;

            Color back = isDragHover ? BgDropHover : BgDrop;
            using (var path = RoundedRect(rect, radius))
            using (var brush = new SolidBrush(back))
            using (var pen = new Pen(BorderSoft, 1.2f) { DashStyle = DashStyle.Dash })
            {
                g.FillPath(brush, path);
                g.DrawPath(pen, path);
            }
        }

        private static GraphicsPath RoundedRect(Rectangle bounds, int radius)
        {
            var path = new GraphicsPath();
            if (radius <= 0)
            {
                path.AddRectangle(bounds);
                return path;
            }
            int d = radius * 2;
            path.AddArc(bounds.X, bounds.Y, d, d, 180, 90);
            path.AddArc(bounds.Right - d, bounds.Y, d, d, 270, 90);
            path.AddArc(bounds.Right - d, bounds.Bottom - d, d, d, 0, 90);
            path.AddArc(bounds.X, bounds.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }

        private void BtnChooseFolder_Click(object sender, EventArgs e)
        {
            using (var dlg = new FolderBrowserDialog())
            {
                dlg.Description = "Выберите папку для сохранения картинок";
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    txtOutputFolder.Text = dlg.SelectedPath;
                    lblFolderError.Visible = false;
                }
            }
        }

        private void BtnClearLog_Click(object sender, EventArgs e)
        {
            lstLog.Items.Clear();
        }

        private void BtnSaveLog_Click(object sender, EventArgs e)
        {
            try
            {
                string desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                string fileName = $"journal_{DateTime.Now:yyyyMMdd_HHmmss}.txt";
                string path = Path.Combine(desktop, fileName);

                var lines = lstLog.Items.Cast<object>().Select(o => o?.ToString() ?? string.Empty);
                File.WriteAllLines(path, lines);

                Log($"Журнал сохранён: {path}");
            }
            catch (Exception ex)
            {
                Log($"ОШИБКА сохранения журнала: {ex.Message}");
            }
        }

        private void DropPanel_DragEnter(object sender, DragEventArgs e)
        {
            if (isBusy)
            {
                e.Effect = DragDropEffects.None;
                return;
            }

            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Any(f => f.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)))
                {
                    e.Effect = DragDropEffects.Copy;
                    isDragHover = true;
                    dropPanel.Invalidate();
                    return;
                }
            }
            e.Effect = DragDropEffects.None;
        }

        private void DropPanel_DragLeave(object sender, EventArgs e)
        {
            isDragHover = false;
            dropPanel.Invalidate();
        }

        private void DropPanel_DragDrop(object sender, DragEventArgs e)
        {
            isDragHover = false;
            dropPanel.Invalidate();
            if (isBusy) return;

            if (string.IsNullOrWhiteSpace(txtOutputFolder.Text))
            {
                lblFolderError.Visible = true;
                return;
            }
            lblFolderError.Visible = false;

            var files = (string[])e.Data.GetData(DataFormats.FileDrop);
            var docxFiles = files.Where(f => f.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)).ToArray();

            if (docxFiles.Length == 0)
            {
                Log("Нет .docx файлов для обработки.");
                return;
            }

            var docx = docxFiles[0];
            if (docxFiles.Length > 1)
                Log($"Будет обработан только первый файл: {Path.GetFileName(docx)}");

            try
            {
                isBusy = true;
                dropPanel.AllowDrop = false;
                lblDropHint.Text = "Обработка...";
                lblDropHint.ForeColor = Accent;
                Cursor = Cursors.WaitCursor;
                Application.DoEvents();

                ProcessDocx(docx, txtOutputFolder.Text);
            }
            catch (Exception ex)
            {
                Log($"ОШИБКА [{Path.GetFileName(docx)}]: {ex.Message}");
            }
            finally
            {
                isBusy = false;
                dropPanel.AllowDrop = true;
                Cursor = Cursors.Default;
                if (lblDropHint.Text != "Картинок нет")
                {
                    lblDropHint.Text = "Перетащите сюда .docx файл";
                    lblDropHint.ForeColor = TextMain;
                }
            }
        }

        private void ProcessDocx(string docxPath, string outputFolder)
        {
            string baseName = Path.GetFileNameWithoutExtension(docxPath);
            Log($"--- Обработка: {Path.GetFileName(docxPath)} ---");

            using (var zip = ZipFile.OpenRead(docxPath))
            {
                var orderedMedia = GetMediaInDocumentOrder(zip);

                if (orderedMedia.Count == 0)
                {
                    Log("Картинок нет.");
                    lblDropHint.Text = "Картинок нет";
                    lblDropHint.ForeColor = Danger;
                    return;
                }

                int index = 1;
                foreach (var mediaEntryName in orderedMedia)
                {
                    var entry = zip.GetEntry(mediaEntryName);
                    if (entry == null) continue;

                    string ext = Path.GetExtension(entry.Name);
                    if (!AllowedExtensions.Contains(ext)) continue;

                    string newName = $"{baseName}_{index}{ext}";
                    string outPath = Path.Combine(outputFolder, newName);

                    if (File.Exists(outPath))
                        File.Delete(outPath);

                    entry.ExtractToFile(outPath);

                    Log($"  [{index}] {entry.Name} -> {newName}");
                    index++;
                }

                if (index == 1)
                {
                    Log("Картинок нет.");
                    lblDropHint.Text = "Картинок нет";
                    lblDropHint.ForeColor = Danger;
                }
                else
                {
                    Log($"Готово. Извлечено картинок: {index - 1}");
                    lblDropHint.Text = $"Готово — {index - 1} шт.";
                    lblDropHint.ForeColor = Accent;
                }
            }
        }

        private List<string> GetMediaInDocumentOrder(ZipArchive zip)
        {
            var result = new List<string>();

            var docEntry = zip.GetEntry("word/document.xml");
            var relsEntry = zip.GetEntry("word/_rels/document.xml.rels");
            if (docEntry == null || relsEntry == null)
                return result;

            var relMap = new Dictionary<string, string>();
            using (var relsStream = relsEntry.Open())
            {
                var relsDoc = XDocument.Load(relsStream);
                XNamespace rns = "http://schemas.openxmlformats.org/package/2006/relationships";
                foreach (var rel in relsDoc.Root.Elements(rns + "Relationship"))
                {
                    string id = rel.Attribute("Id")?.Value;
                    string target = rel.Attribute("Target")?.Value;
                    if (id != null && target != null)
                        relMap[id] = target;
                }
            }

            using (var docStream = docEntry.Open())
            {
                var doc = XDocument.Load(docStream);

                XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
                XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

                var blips = doc.Descendants(a + "blip")
                               .Select(b => b.Attribute(r + "embed")?.Value)
                               .Where(id => !string.IsNullOrEmpty(id));

                var added = new HashSet<string>();
                foreach (var rid in blips)
                {
                    if (relMap.TryGetValue(rid, out var target))
                    {
                        string normalized = target.Replace('\\', '/');
                        if (normalized.StartsWith("media/"))
                            normalized = "word/" + normalized;
                        else if (!normalized.StartsWith("word/"))
                            normalized = "word/" + normalized.TrimStart('/');

                        if (added.Add(normalized))
                            result.Add(normalized);
                    }
                }
            }

            return result;
        }

        private void Log(string message)
        {
            lstLog.Items.Add(message);
            lstLog.TopIndex = lstLog.Items.Count - 1;
        }
    }

    public class RoundedButton : Button
    {
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        [Browsable(false)]
        public int CornerRadius { get; set; } = 10;

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        [Browsable(false)]
        public Color HoverBackColor { get; set; } = Color.Gray;

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        [Browsable(false)]
        public Color DownBackColor { get; set; } = Color.DimGray;

        private bool _hover, _down;

        public RoundedButton()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.UserPaint |
                     ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.ResizeRedraw, true);
        }

        protected override void OnMouseEnter(EventArgs e)
        {
            _hover = true; Invalidate(); base.OnMouseEnter(e);
        }
        protected override void OnMouseLeave(EventArgs e)
        {
            _hover = false; _down = false; Invalidate(); base.OnMouseLeave(e);
        }
        protected override void OnMouseDown(MouseEventArgs mevent)
        {
            _down = true; Invalidate(); base.OnMouseDown(mevent);
        }
        protected override void OnMouseUp(MouseEventArgs mevent)
        {
            _down = false; Invalidate(); base.OnMouseUp(mevent);
        }

        protected override void OnPaint(PaintEventArgs pevent)
        {
            var g = pevent.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            Color back = _down ? DownBackColor : (_hover ? HoverBackColor : BackColor);

            var rect = new Rectangle(0, 0, Width - 1, Height - 1);

            if (CornerRadius <= 0)
            {
                using (var brush = new SolidBrush(back))
                {
                    g.FillRectangle(brush, rect);
                }
            }
            else
            {
                using (var path = RoundedRect(rect, CornerRadius))
                using (var brush = new SolidBrush(back))
                {
                    g.FillPath(brush, path);
                }
            }

            TextRenderer.DrawText(
                g,
                Text,
                Font,
                rect,
                ForeColor,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        }

        private static GraphicsPath RoundedRect(Rectangle bounds, int radius)
        {
            var path = new GraphicsPath();
            if (radius <= 0)
            {
                path.AddRectangle(bounds);
                return path;
            }
            int d = radius * 2;
            path.AddArc(bounds.X, bounds.Y, d, d, 180, 90);
            path.AddArc(bounds.Right - d, bounds.Y, d, d, 270, 90);
            path.AddArc(bounds.Right - d, bounds.Bottom - d, d, d, 0, 90);
            path.AddArc(bounds.X, bounds.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }
    }
}