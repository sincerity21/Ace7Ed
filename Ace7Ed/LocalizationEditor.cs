using Ace7Ed.Properties;
using Ace7Ed.Prompt;
using Ace7LocalizationFormat.Formats;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using CMN = Ace7LocalizationFormat.Formats.CmnFile;
using DAT = Ace7LocalizationFormat.Formats.DatFile;
using Ace7Ed.Interact;
using System.ComponentModel.Design.Serialization;
namespace Ace7Ed
{
    internal interface IUndoAction
    {
        void Undo(LocalizationEditor editor);
    }

    internal sealed class TextEditUndoAction : IUndoAction
    {
        private readonly int _datIndex;
        private readonly int _stringNumber;
        private readonly string _oldText;

        public TextEditUndoAction(int datIndex, int stringNumber, string oldText)
        {
            _datIndex = datIndex;
            _stringNumber = stringNumber;
            _oldText = oldText ?? "";
        }

        public void Undo(LocalizationEditor editor) => editor.UndoTextEdit(_datIndex, _stringNumber, _oldText);
    }

    public partial class LocalizationEditor : Form
    {
        private (CMN?, List<DAT>?) _modifiedLocalization = (null, null);

        private string _directory { get; set; } = "";

        private (int, List<int>?) _copyVariableStrings { get; set; }
        private List<string> _copyStrings = new List<string>();

        /// <summary>Node that was right-clicked for context menu; used so Delete works even if TreeView clears selection when menu opens.</summary>
        private TreeNode? _cmnNodeForContextMenu;

        private int _selectedRowIndex = -1;
        private int _selectedColumnIndex = -1;

        /// <summary>Index of the selected language DAT in the loaded list (0-based), or -1 if none.</summary>
        private int _selectedDatIndex
        {
            get
            {
                if (DatLanguageComboBox.SelectedItem is not DAT dat || _modifiedLocalization.Item2 == null)
                    return -1;
                int index = _modifiedLocalization.Item2.IndexOf(dat);
                return index >= 0 ? index : -1;
            }
        }

        private bool _savedChanges = true;
        private bool _isClosingDeferred;

        private readonly Stack<IUndoAction> _undoStack = new Stack<IUndoAction>();

        private System.Windows.Forms.Timer? _closeTreeTimer;
        private Stopwatch? _closeTreeStopwatch;
        private Stopwatch? _closeTotalStopwatch;
        private long _closeLogReleaseMs;
        private long _closeLogComboMs;
        private long _closeLogGridRowsMs;
        private long _closeLogGridColsMs;

        public bool SavedChanges
        {
            get
            {
                return _savedChanges;
            }
            set
            {
                _savedChanges = value;
                if (_savedChanges)
                {
                    Text = Text.Substring(0, Text.Length - 1); // Remove the "*"
                }
                else
                {
                    Text += "*";
                }
            }
        }

        public LocalizationEditor()
        {
            InitializeComponent();
            ToggleDarkTheme();
            MSMainUndo.Enabled = false;
            MSMainSave.Enabled = false;

            if (Clipboard.GetText() != null)
            {
                _copyStrings.Add(Clipboard.GetText());
            }

            MSOptionImportLocalization.Enabled = false;
            MSOptionImportLocalization.Visible = false;
        }

        private void LocalizationEditor_FormClosing(object? sender, FormClosingEventArgs e)
        {
            // Only ask for confirmation on the first close attempt; skip when deferred Close() runs
            if (!_isClosingDeferred && !SavedChanges)
            {
                var result = MessageBox.Show(
                    "You have unsaved changes. Are you sure you want to close?",
                    "Unsaved changes",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);
                if (result != DialogResult.Yes)
                {
                    e.Cancel = true;
                    return;
                }
            }

            if (!_isClosingDeferred)
            {
                _isClosingDeferred = true;
                e.Cancel = true;
                Hide();
                BeginInvoke(StartDeferredClose);
            }
            else
            {
                ClearHeavyDataBeforeClose();
            }
        }

        private void StartDeferredClose()
        {
            _closeTotalStopwatch = Stopwatch.StartNew();
            var sw = Stopwatch.StartNew();

            _modifiedLocalization = (null, null);
            ClearUndoStack();
            _closeLogReleaseMs = sw.ElapsedMilliseconds;

            DatsDataGridView.SuspendLayout();
            CmnTreeView.BeginUpdate();
            try
            {
                sw.Restart();
                DatLanguageComboBox.Items.Clear();
                _closeLogComboMs = sw.ElapsedMilliseconds;

                sw.Restart();
                DatsDataGridView.Rows.Clear();
                DatsDataGridView.Columns.Clear();
                _closeLogGridRowsMs = sw.ElapsedMilliseconds;
                _closeLogGridColsMs = 0; // combined with rows for log
            }
            finally
            {
                DatsDataGridView.ResumeLayout(false);
            }

            if (CmnTreeView.Nodes.Count == 0)
            {
                CmnTreeView.EndUpdate();
                FinishCloseAndLog(treeMs: 0);
                return;
            }

            _closeTreeStopwatch = Stopwatch.StartNew();
            _closeTreeTimer = new System.Windows.Forms.Timer { Interval = TreeClearTimerIntervalMs };
            _closeTreeTimer.Tick += CloseTreeTimer_Tick;
            _closeTreeTimer.Start();
        }

        private void CloseTreeTimer_Tick(object? sender, EventArgs e)
        {
            ClearTreeNodesOneBatch(CmnTreeView.Nodes, TreeClearBatchSize);
            if (CmnTreeView.Nodes.Count != 0)
                return;

            _closeTreeTimer!.Stop();
            _closeTreeTimer.Dispose();
            _closeTreeTimer = null;
            _closeTreeStopwatch!.Stop();
            long treeMs = _closeTreeStopwatch.ElapsedMilliseconds;
            _closeTreeStopwatch = null;
            CmnTreeView.EndUpdate();
            FinishCloseAndLog(treeMs);
        }

        private void FinishCloseAndLog(long treeMs)
        {
            _closeTotalStopwatch!.Stop();
            long totalMs = _closeTotalStopwatch.ElapsedMilliseconds;
            _closeTotalStopwatch = null;
            WriteCloseTimingsLog(_closeLogReleaseMs, _closeLogComboMs, treeMs, _closeLogGridRowsMs, _closeLogGridColsMs, totalMs);
            MSMainSave.Enabled = false;
            Close();
        }

        private void ClearHeavyDataBeforeClose()
        {
            var totalSw = Stopwatch.StartNew();
            long releaseMs = 0, comboMs = 0, treeMs = 0, gridRowsMs = 0, gridColsMs = 0;

            var sw = Stopwatch.StartNew();
            _modifiedLocalization = (null, null);
            ClearUndoStack();
            releaseMs = sw.ElapsedMilliseconds;

            DatsDataGridView.SuspendLayout();
            CmnTreeView.BeginUpdate();
            try
            {
                sw.Restart();
                DatLanguageComboBox.Items.Clear();
                comboMs = sw.ElapsedMilliseconds;

                sw.Restart();
                ClearTreeNodesBottomUp(CmnTreeView.Nodes);
                treeMs = sw.ElapsedMilliseconds;

                sw.Restart();
                DatsDataGridView.Rows.Clear();
                gridRowsMs = sw.ElapsedMilliseconds;

                sw.Restart();
                DatsDataGridView.Columns.Clear();
                gridColsMs = sw.ElapsedMilliseconds;
            }
            finally
            {
                CmnTreeView.EndUpdate();
                DatsDataGridView.ResumeLayout(false);
            }

            totalSw.Stop();
            WriteCloseTimingsLog(releaseMs, comboMs, treeMs, gridRowsMs, gridColsMs, totalSw.ElapsedMilliseconds);

            MSMainSave.Enabled = false;
        }

        private const int TreeClearBatchSize = 80;
        private const int TreeClearTimerIntervalMs = 50;

        /// <summary>Removes up to maxCount nodes in bottom-up order; used by the close timer.</summary>
        private static void ClearTreeNodesOneBatch(TreeNodeCollection nodes, int maxCount)
        {
            int removed = 0;
            ClearTreeNodesOneBatchCore(nodes, maxCount, ref removed);
        }

        private static void ClearTreeNodesOneBatchCore(TreeNodeCollection nodes, int maxCount, ref int removed)
        {
            for (int i = nodes.Count - 1; i >= 0 && removed < maxCount; i--)
            {
                ClearTreeNodesOneBatchCore(nodes[i].Nodes, maxCount, ref removed);
                if (removed >= maxCount) return;
                nodes.RemoveAt(i);
                removed++;
            }
        }

        /// <summary>Clears tree nodes from the leaves up; yields every batch so the UI stays responsive.</summary>
        private static void ClearTreeNodesBottomUp(TreeNodeCollection nodes, ref int removedCount)
        {
            for (int i = nodes.Count - 1; i >= 0; i--)
            {
                ClearTreeNodesBottomUp(nodes[i].Nodes, ref removedCount);
                nodes.RemoveAt(i);
                removedCount++;
                if (removedCount % TreeClearBatchSize == 0)
                    Application.DoEvents();
            }
        }

        private static void ClearTreeNodesBottomUp(TreeNodeCollection nodes)
        {
            int removed = 0;
            ClearTreeNodesBottomUp(nodes, ref removed);
        }

        private static void WriteCloseTimingsLog(long releaseMs, long comboMs, long treeMs, long gridRowsMs, long gridColsMs, long totalMs)
        {
            try
            {
                string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Ace7Ed");
                Directory.CreateDirectory(dir);
                string path = Path.Combine(dir, "close_timings.log");
                string line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} | Release: {releaseMs} ms, Combo: {comboMs} ms, Tree: {treeMs} ms, GridRows: {gridRowsMs} ms, GridCols: {gridColsMs} ms, Total: {totalMs} ms{Environment.NewLine}";
                File.AppendAllText(path, line);
            }
            catch { /* avoid affecting close if log fails */ }
        }

        private void ToggleDarkTheme()
        {
            BackColor = Theme.ControlColor;
            ForeColor = Theme.ControlTextColor;

            LocalizationEditorMenuStrip.Renderer = new Theme.MenuStripRenderer();

            Theme.SetDarkThemeToolStripMenuItem(MenuStripMain);
            Theme.SetDarkThemeToolStripMenuItem(MSMainOpenFolder);
            Theme.SetDarkThemeToolStripMenuItem(MSMainOpenSingleLanguage);
            Theme.SetDarkThemeToolStripMenuItem(MSMainUndo);
            Theme.SetDarkThemeToolStripMenuItem(MSMainSave);

            Theme.SetDarkThemeToolStripMenuItem(MenuStripOptions);
            Theme.SetDarkThemeToolStripMenuItem(MSOptionImportLocalization);
            Theme.SetDarkThemeToolStripMenuItem(MSOptionBatchCopyLanguage);
            Theme.SetDarkThemeToolStripMenuItem(MSOptionsToggleDarkTheme);
            Theme.SetDarkThemeToolStripMenuItem(MSOptionAddAddon);
            Theme.SetDarkThemeToolStripMenuItem(MSOptionExport);
            Theme.SetDarkThemeToolStripMenuItem(MSOptionImport);

            Theme.SetDarkThemeComboBox(DatLanguageComboBox);
            Theme.SetDarkThemeComboBox(SearchModeComboBox);
            Theme.SetDarkThemeTextBox(SearchTextBox);
            Theme.SetDarkThemeLabel(SearchLabel);

            Theme.SetDarkThemeTreeView(CmnTreeView);

            Theme.SetDarkThemeDataGridView(DatsDataGridView);
        }

        private void CopyVariables()
        {
            List<int> copyVariableNumbers = new List<int>();

            // Copy variables from selected cells
            foreach (DataGridViewCell selectedCell in DatsDataGridView.SelectedCells)
            {
                int variableStringNumber = (int)DatsDataGridView.Rows[selectedCell.RowIndex].Cells[0].Value; // Get the variable string number from the first column
                copyVariableNumbers.Add(variableStringNumber); // Add the variable string number to the list of variables to copy
            }
            _copyVariableStrings = (_selectedDatIndex, copyVariableNumbers);
        }

        private void PasteVariables()
        {
            if (_copyVariableStrings.Item2 == null || _modifiedLocalization.Item2 == null)
                return;
            foreach (int copyVariableNumber in _copyVariableStrings.Item2)
            {
                _modifiedLocalization.Item2[_selectedDatIndex].Strings[copyVariableNumber] = _modifiedLocalization.Item2[_copyVariableStrings.Item1].Strings[copyVariableNumber];
            }
            LoadDatsDataGridView();
        }

        private void CopyStrings()
        {
            List<string> copyStrings = new List<string>();

            foreach (DataGridViewCell selectedCell in DatsDataGridView.SelectedCells)
            {
                string datString = selectedCell.Value?.ToString() ?? "";
                copyStrings.Add(datString);
            }

            _copyStrings = copyStrings;

            if (_copyStrings.Count > 0)
                Clipboard.SetText(_copyStrings[0]);
        }

        private void PasteStrings()
        {
            if (_modifiedLocalization.Item2 == null)
                return;
            var dats = _modifiedLocalization.Item2;
            int index = 0;
            foreach (DataGridViewCell selectedCell in DatsDataGridView.SelectedCells)
            {
                int variableStringNumber = (int)DatsDataGridView.Rows[selectedCell.RowIndex].Cells[0].Value;
                if (_copyStrings.Count == 1)
                {
                    DatsDataGridView.Rows[selectedCell.RowIndex].Cells[2].Value = _copyStrings[0];
                    dats[_selectedDatIndex].Strings[variableStringNumber] = _copyStrings[0];
                }
                else if (index < _copyStrings.Count)
                {
                    DatsDataGridView.Rows[selectedCell.RowIndex].Cells[2].Value = _copyStrings[index];
                    dats[_selectedDatIndex].Strings[variableStringNumber] = _copyStrings[index];
                }
                index++;
            }
        }

        #region Loading

        public (CMN, List<DAT>) LoadLocalization(string[] files)
        {
            CMN? modifiedCmn = null;
            List<DAT> modifiedDats = new List<DAT>();

            foreach (string filePath in files)
            {
                if (Path.GetExtension(filePath) == ".dat")
                {
                    if (Path.GetFileNameWithoutExtension(filePath).Equals("Cmn", StringComparison.OrdinalIgnoreCase))
                    {
                        modifiedCmn = new CMN(filePath);
                    }
                    else if (AceLocalizationConstants.DatLetters.Keys.Contains(Path.GetFileNameWithoutExtension(filePath)[0]))
                    {
                        modifiedDats.Add(new DAT(filePath, Path.GetFileNameWithoutExtension(filePath)[0]));
                    }
                }
            }

            if (modifiedCmn == null || modifiedDats.Count != 13)
            {
                MessageBox.Show("Missing Dats", "Error");
                throw new Exception("Missing Dats");
            }

            return (modifiedCmn, modifiedDats);
        }

        /// <summary>
        /// Load Cmn.dat and one or more language DATs (e.g. for editing selected languages).
        /// </summary>
        public (CMN, List<DAT>) LoadLocalizationSingle(string cmnPath, string[] datPaths)
        {
            if (!File.Exists(cmnPath))
            {
                MessageBox.Show("Cmn.dat file not found.", "Error");
                throw new FileNotFoundException("Cmn.dat not found.", cmnPath);
            }
            if (datPaths == null || datPaths.Length == 0)
            {
                MessageBox.Show("Select at least one language DAT (A.dat through M.dat).", "Error");
                throw new ArgumentException("At least one language DAT must be selected.", nameof(datPaths));
            }

            string cmnName = Path.GetFileNameWithoutExtension(cmnPath);
            if (!cmnName.Equals("Cmn", StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show("First file must be Cmn.dat.", "Error");
                throw new ArgumentException("First file must be Cmn.dat.", nameof(cmnPath));
            }

            List<DAT> modifiedDats = new List<DAT>();
            foreach (string datPath in datPaths)
            {
                if (string.IsNullOrEmpty(datPath) || !File.Exists(datPath))
                {
                    MessageBox.Show($"Language DAT file not found: {datPath}", "Error");
                    throw new FileNotFoundException("Language DAT not found.", datPath);
                }
                string datName = Path.GetFileNameWithoutExtension(datPath);
                if (datName.Length != 1 || !AceLocalizationConstants.DatLetters.Keys.Contains(datName[0]))
                {
                    MessageBox.Show($"Invalid language DAT: {Path.GetFileName(datPath)}. Must be A.dat through M.dat.", "Error");
                    throw new ArgumentException($"Invalid language DAT: {datPath}", nameof(datPaths));
                }
                modifiedDats.Add(new DAT(datPath, datName[0]));
            }

            CMN modifiedCmn = new CMN(cmnPath);
            return (modifiedCmn, modifiedDats);
        }

        private void LoadLocalizationForUI(string folder)
        {
            var cmn = _modifiedLocalization.Item1;
            var dats = _modifiedLocalization.Item2;
            if (cmn == null || dats == null)
                return;
            foreach (var dat in dats)
            {
                if (dat.Strings.Count < cmn.MaxStringNumber)
                {
                    dat.Strings.AddRange(Enumerable.Repeat("\0", cmn.MaxStringNumber + 1 - dat.Strings.Count));
                }
            }

            _directory = folder;

            LoadDatLanguageComboBox();
            LoadCmnTreeView();
        }

        private void LoadDatLanguageComboBox()
        {
            if (_modifiedLocalization.Item2 == null)
                return;
            DatLanguageComboBox.BeginUpdate();

            DatLanguageComboBox.Items.Clear();

            _modifiedLocalization.Item2.ForEach(dat => DatLanguageComboBox.Items.Add(dat));

            DatLanguageComboBox.EndUpdate();
        }

        private void LoadCmnTreeView()
        {
            if (_modifiedLocalization.Item1 == null)
                return;
            // Ensure the whole tree has the correct connected hierarchy (e.g. AircraftTest_Name_ with _f15c and _f15t under it)
            _modifiedLocalization.Item1.EnsureAllIntermediateParents();

            CmnTreeView.BeginUpdate();

            CmnTreeView.Nodes.Clear();

            foreach (var node in _modifiedLocalization.Item1.Root.Values)
            {
                CmnTreeView.Nodes.Add(GetTreeNodeFromCmn(node));
            }

            CmnTreeView.EndUpdate();
        }

        private void LoadDatsDataGridView()
        {
            DatsDataGridView.Columns.Clear();
            DatsDataGridView.Columns.Add("designNumber", "Number");
            DatsDataGridView.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            DatsDataGridView.Columns[0].ReadOnly = true;
            DatsDataGridView.Columns.Add("designID", "ID");
            DatsDataGridView.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            DatsDataGridView.Columns[1].ReadOnly = true;
            DatsDataGridView.Columns.Add("designText", "Text");
            DatsDataGridView.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            DatsDataGridView.Columns[2].ReadOnly = true;

            string searchTerm = SearchTextBox?.Text?.Trim() ?? "";
            int searchModeIndex = SearchModeComboBox?.SelectedIndex ?? 0;
            bool filterBySearch = searchTerm.Length > 0;

            if (DatLanguageComboBox.SelectedItem != null && _modifiedLocalization.Item1 != null && _modifiedLocalization.Item2 != null)
            {
                DatsDataGridView.Rows.Clear();
                DAT dat = _modifiedLocalization.Item2[_selectedDatIndex];

                {
                    // Load strings of the selected CMN
                    if (CmnTreeView.SelectedNode is TreeNode treeNode)
                    {
                        AddCmnNodeToDataGridView(dat, (CmnString)treeNode.Tag, filterBySearch ? searchModeIndex : (int?)null, filterBySearch ? searchTerm : null);
                    }
                    // Load all strings
                    else if (CmnTreeView.SelectedNode == null)
                    {
                        foreach (var node in _modifiedLocalization.Item1.Root.Values)
                        {
                            AddCmnNodeToDataGridView(dat, node, filterBySearch ? searchModeIndex : (int?)null, filterBySearch ? searchTerm : null);
                        }
                    }
                }
            }

            DatsDataGridView.Sort(DatsDataGridView.Columns[0], ListSortDirection.Ascending);

            DatsDataGridView.ClearSelection();
        }

        /// <summary>
        /// Returns true if the row (number, id, text) matches the current search.
        /// searchMode: 0 = Number (exact), 1 = ID (keyword), 2 = Text (keyword).
        /// </summary>
        private static bool RowMatchesSearch(int stringNumber, string id, string text, int searchMode, string searchTerm)
        {
            if (string.IsNullOrEmpty(searchTerm))
                return true;
            switch (searchMode)
            {
                case 0: // Number - exact match
                    return int.TryParse(searchTerm, out int n) && n == stringNumber;
                case 1: // ID - keyword (contains, case-insensitive)
                    return id.Contains(searchTerm, StringComparison.OrdinalIgnoreCase);
                case 2: // Text - keyword (contains, case-insensitive)
                    return (text ?? "").Contains(searchTerm, StringComparison.OrdinalIgnoreCase);
                default:
                    return true;
            }
        }

        private void SearchTextBox_KeyDown(object? sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                LoadDatsDataGridView();
            }
        }

        private void SearchTextBox_TextChanged(object? sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(SearchTextBox?.Text))
                LoadDatsDataGridView();
        }

        private void SearchModeComboBox_SelectedIndexChanged(object? sender, EventArgs e)
        {
            LoadDatsDataGridView();
        }

        #endregion

        #region Undo
        private void ClearUndoStack()
        {
            _undoStack.Clear();
            UpdateUndoMenuItemEnabled();
        }

        private void UpdateUndoMenuItemEnabled()
        {
            if (MSMainUndo != null)
                MSMainUndo.Enabled = _undoStack.Count > 0;
        }

        internal void UndoTextEdit(int datIndex, int stringNumber, string oldText)
        {
            if (_modifiedLocalization.Item2 == null || datIndex < 0 || datIndex >= _modifiedLocalization.Item2.Count)
                return;
            var dats = _modifiedLocalization.Item2;
            if (stringNumber < 0 || stringNumber >= dats[datIndex].Strings.Count)
                return;
            dats[datIndex].Strings[stringNumber] = oldText;
            if (datIndex == _selectedDatIndex)
            {
                foreach (DataGridViewRow row in DatsDataGridView.Rows)
                {
                    if (row.Cells[0].Value != null && Convert.ToInt32(row.Cells[0].Value) == stringNumber)
                    {
                        row.Cells[2].Value = oldText;
                        break;
                    }
                }
            }
        }

        private void PerformUndo()
        {
            if (_undoStack.Count == 0)
                return;
            IUndoAction action = _undoStack.Pop();
            action.Undo(this);
            if (_undoStack.Count == 0)
                SavedChanges = true;
            else
                SavedChanges = false;
            UpdateUndoMenuItemEnabled();
        }
        #endregion

        #region Saving
        private void SaveChanges()
        {
            if (_modifiedLocalization.Item1 != null)
            {
                string cmnFilePath = _directory + "\\Cmn.dat";
                if (File.Exists(cmnFilePath))
                    File.Copy(cmnFilePath, cmnFilePath + ".bak", true);

                _modifiedLocalization.Item1.Write(cmnFilePath);
            }

            if (_modifiedLocalization.Item2 != null)
            {
                foreach (var dat in _modifiedLocalization.Item2)
                {
                    string datFilePath = _directory + "\\" + dat.Letter + ".dat";

                    if (File.Exists(datFilePath))
                        File.Copy(datFilePath, datFilePath + ".bak", true);

                    dat.Write(datFilePath, dat.Letter);
                }
            }
            SavedChanges = true;
            ClearUndoStack();
        }

        #endregion

        #region Manage Controls

        public TreeNode GetTreeNodeFromCmn(CmnString parent)
        {
            var root = new TreeNode();
            root.Name = parent.Name;
            root.Text = parent.Name;
            root.Tag = parent;

            List<TreeNode> nodesToVisit = [root];

            int index = 0;
            while (index < nodesToVisit.Count)
            {
                TreeNode currentCmnNode = nodesToVisit[index];
                CmnString cmnString = (CmnString)currentCmnNode.Tag;
                foreach (var child in cmnString.Childrens.Values)
                {
                    TreeNode subNode = new TreeNode();
                    subNode.Name = child.Name;
                    subNode.Text = child.Name;
                    subNode.Tag = child;
                    currentCmnNode.Nodes.Add(subNode);
                    nodesToVisit.Add(subNode);
                }
                index++;
            }

            return root;
        }

        /// <summary>
        /// Finds a TreeNode under root whose CmnString tag has the given full name.
        /// </summary>
        private static TreeNode? FindTreeNodeByCmnName(TreeNode root, string cmnFullName)
        {
            if (root.Tag is CmnString cmn && cmn.Name == cmnFullName)
                return root;
            foreach (TreeNode child in root.Nodes)
            {
                var found = FindTreeNodeByCmnName(child, cmnFullName);
                if (found != null)
                    return found;
            }
            return null;
        }

        public void ImportLocalization((CMN, List<DAT>) localization)
        {
            var cmn = _modifiedLocalization.Item1;
            var dats = _modifiedLocalization.Item2;
            if (cmn == null || dats == null)
                return;
            List<CmnString> nodesToVisit = new List<CmnString>(localization.Item1.Root.Values);

            int index = 0;
            while (index < nodesToVisit.Count)
            {
                CmnString cmnString = nodesToVisit[index];
                foreach (var child in cmnString.Childrens.Values)
                {
                    bool isVariableAdded = cmn.AddVariable(child.Name, cmn.Root, out int variableStringNumber);

                    if (child.StringNumber != -1)
                    {
                        foreach (var dat in localization.Item2)
                        {
                            if (child.StringNumber < dat.Strings.Count)
                            {
                                if (isVariableAdded)
                                {
                                    dats[dat.Letter - 65].Strings.Add(localization.Item2[dat.Letter - 65].Strings[child.StringNumber]);
                                }
                                else
                                {
                                    dats[dat.Letter - 65].Strings[variableStringNumber] = localization.Item2[dat.Letter - 65].Strings[child.StringNumber];
                                }
                            }
                            else
                            {
                                dats[dat.Letter - 65].Strings.Add("\0");
                            }
                        }
                    }
                    nodesToVisit.Add(child);
                }
                index++;
            }

            Debug.WriteLine(cmn.MaxStringNumber);
        }

        private void AddCmnNodeToDataGridView(DAT dat, CmnString parent, int? searchMode = null, string? searchTerm = null)
        {
            if (parent.StringNumber != -1)
            {
                string text = dat.Strings[parent.StringNumber];
                if (searchMode == null || searchTerm == null || RowMatchesSearch(parent.StringNumber, parent.Name, text, searchMode.Value, searchTerm))
                    DatsDataGridView.Rows.Add(parent.StringNumber, parent.Name, text);
            }

            foreach (var child in parent.Childrens.Values)
            {
                AddCmnNodeToDataGridView(dat, child, searchMode, searchTerm);
            }
        }

        private void LoadLocalization_DragDrop(object? sender, DragEventArgs e)
        {
            string[]? folderPaths = e.Data?.GetData(DataFormats.FileDrop, false) as string[];
            if (folderPaths == null)
                return;

            foreach (string folderPath in folderPaths)
            {
                if (string.IsNullOrEmpty(folderPath) || !Directory.Exists(folderPath))
                    continue;
                string[] files = Directory.GetFiles(folderPath);

                _modifiedLocalization = LoadLocalization(files);
                ClearUndoStack();

                LoadLocalizationForUI(folderPath);

                MSMainSave.Enabled = true;
                MSOptionImportLocalization.Enabled = true;
                MSOptionBatchCopyLanguage.Enabled = _modifiedLocalization.Item2?.Count > 1;
                MSOptionAddAddon.Enabled = true;
                MSOptionExport.Enabled = true;
                MSOptionImport.Enabled = true;
            }

        }

        private void LoadLocalization_DragEnter(object? sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.None;
            if (e.Data?.GetDataPresent(DataFormats.FileDrop) == true)
            {
                var paths = e.Data.GetData(DataFormats.FileDrop) as string[];
                if (paths != null && paths.Length > 0 && Directory.Exists(paths[0]))
                    e.Effect = DragDropEffects.All;
            }
        }

        #endregion

        #region Menu Strip Controls

        private void MSOptionImportLocalization_Click(object? sender, EventArgs e)
        {
            using FolderBrowserDialog folderBrowser = new FolderBrowserDialog()
            {
                Description = "Select DATs directory",
                RootFolder = Environment.SpecialFolder.MyComputer,
            };

            if (folderBrowser.ShowDialog() == DialogResult.OK)
            {
                string[] files = Directory.GetFiles(folderBrowser.SelectedPath);

                var localization = LoadLocalization(files);

                ImportLocalization(localization);

                LoadCmnTreeView();
                LoadDatsDataGridView();
            }
        }

        private void MSOptionBatchCopyLanguage_Click(object? sender, EventArgs e)
        {
            if (_modifiedLocalization.Item1 == null || _modifiedLocalization.Item2 == null)
                return;
            var localization = (_modifiedLocalization.Item1, _modifiedLocalization.Item2);
            using (var batchCopyLanguage = new BatchCopyLanguage(localization))
            {
                batchCopyLanguage.ShowDialog();

                if (batchCopyLanguage.DialogResult == DialogResult.OK && batchCopyLanguage.SelectedCopyLanguageIndex != -1 && batchCopyLanguage.SelectedPasteLanguages.Count != 0)
                {
                    var dats = _modifiedLocalization.Item2;
                    foreach (var pasteLanguageLetter in batchCopyLanguage.SelectedPasteLanguages)
                    {
                        int pasteIndex = dats.FindIndex(d => d.Letter == pasteLanguageLetter);
                        if (pasteIndex < 0)
                            continue;
                        for (int i = batchCopyLanguage.StartNumber; i < batchCopyLanguage.EndNumber; i++)
                        {
                            if (dats[pasteIndex].Strings[i] == "\0" || batchCopyLanguage.OverwriteExistingString)
                            {
                                dats[pasteIndex].Strings[i] = dats[batchCopyLanguage.SelectedCopyLanguageIndex].Strings[i];
                            }
                        }
                    }
                }

                batchCopyLanguage.Dispose();
            }
            LoadDatsDataGridView();
        }

        private void MSOptionsToggleDarkTheme_Click(object? sender, EventArgs e)
        {
            Configurations.Default.DarkTheme = !Configurations.Default.DarkTheme;
            Configurations.Default.Save();
            ToggleDarkTheme();
        }

        private void MSOptionAddAddon_Click(object? sender, EventArgs e)
        {
            if (_modifiedLocalization.Item1 == null || _modifiedLocalization.Item2 == null)
            {
                MessageBox.Show("Open a localization folder first.", "Add a new plane");
                return;
            }

            string plane;
            using (var input = new Input("Enter plane string (e.g. f15t)", "") { StartPosition = FormStartPosition.CenterScreen })
            {
                input.ShowDialog();
                if (input.DialogResult != DialogResult.OK) return;
                plane = (input.InputText ?? "").Trim();
            }
            if (string.IsNullOrEmpty(plane))
            {
                MessageBox.Show("Plane string cannot be empty.", "Add a new plane");
                return;
            }

            int skinCount;
            using (var input = new Input("Enter number of skins (minimum 6)", "") { StartPosition = FormStartPosition.CenterScreen })
            {
                input.ShowDialog();
                if (input.DialogResult != DialogResult.OK) return;
                if (!int.TryParse(input.InputText?.Trim(), out skinCount) || skinCount < 6)
                {
                    MessageBox.Show("Please enter a number of skins (minimum 6).", "Add a new plane");
                    return;
                }
            }

            var cmn = _modifiedLocalization.Item1;
            var dats = _modifiedLocalization.Item2;
            var keysToAdd = new List<string>();

            keysToAdd.Add($"Aircraft_Name_{plane}");
            keysToAdd.Add($"AircraftShort_Name_{plane}");
            keysToAdd.Add($"Aircraft_Description_{plane}");
            keysToAdd.Add($"AircraftDataviewer_Description_{plane}");

            for (int i = 0; i < 6 && i < skinCount; i++)
            {
                keysToAdd.Add($"AircraftSkin_Description_{plane}_{i:D2}");
            }
            for (int i = 6; i < skinCount; i++)
            {
                int dlcNum = i - 5;
                keysToAdd.Add($"AircraftSkintype_Name_{plane}_DLC{dlcNum}");
                keysToAdd.Add($"AircraftSkin_Description_{plane}_{i:D2}");
            }

            foreach (string key in keysToAdd)
            {
                if (cmn.CheckVariableExist(key))
                    continue; // already exists, skip
                cmn.AddVariable(key, cmn.Root, out int _);
                foreach (var dat in dats)
                    dat.Strings.Add("\0");
            }

            LoadCmnTreeView();
            LoadDatsDataGridView();
            SavedChanges = false;
        }

        private static void CollectMatchingKeys(List<(string FullName, int StringNumber)> result, CmnString node, string filter)
        {
            if (node.StringNumber != -1 && node.Name.Contains(filter, StringComparison.OrdinalIgnoreCase))
                result.Add((node.Name, node.StringNumber));
            foreach (var child in node.Childrens.Values)
                CollectMatchingKeys(result, child, filter);
        }

        private static string EscapeCsvField(string? value)
        {
            if (value == null) return "";
            if (value.Contains('"') || value.Contains(',') || value.Contains('\n') || value.Contains('\r'))
            {
                return "\"" + value.Replace("\"", "\"\"") + "\"";
            }
            return value;
        }

        /// <summary>
        /// Parses a single line of CSV (no newlines inside fields). Use ParseCsvRows for full content with quoted newlines.
        /// </summary>
        private static List<string> ParseCsvLine(string line)
        {
            var fields = new List<string>();
            int i = 0;
            while (i < line.Length)
            {
                if (line[i] == '"')
                {
                    var sb = new System.Text.StringBuilder();
                    i++;
                    while (i < line.Length)
                    {
                        if (line[i] == '"')
                        {
                            i++;
                            if (i < line.Length && line[i] == '"') { sb.Append('"'); i++; }
                            else break;
                        }
                        else { sb.Append(line[i]); i++; }
                    }
                    fields.Add(sb.ToString());
                }
                else
                {
                    int start = i;
                    while (i < line.Length && line[i] != ',') i++;
                    fields.Add(line.Substring(start, i - start));
                    if (i < line.Length) i++;
                }
            }
            return fields;
        }

        /// <summary>
        /// Parses full CSV content and yields one row at a time. Handles RFC 4180: quoted fields may contain commas and newlines; "" is an escaped quote.
        /// </summary>
        private static IEnumerable<List<string>> ParseCsvRows(string csvContent)
        {
            var row = new List<string>();
            var currentField = new System.Text.StringBuilder();
            bool inQuotes = false;
            int i = 0;
            while (i < csvContent.Length)
            {
                char c = csvContent[i];
                if (inQuotes)
                {
                    if (c == '"')
                    {
                        i++;
                        if (i < csvContent.Length && csvContent[i] == '"')
                        {
                            currentField.Append('"');
                            i++;
                        }
                        else
                        {
                            inQuotes = false;
                        }
                    }
                    else
                    {
                        currentField.Append(c);
                        i++;
                    }
                }
                else
                {
                    if (c == '"')
                    {
                        inQuotes = true;
                        i++;
                    }
                    else if (c == ',')
                    {
                        row.Add(currentField.ToString());
                        currentField.Clear();
                        i++;
                    }
                    else if (c == '\n' || c == '\r')
                    {
                        row.Add(currentField.ToString());
                        currentField.Clear();
                        if (c == '\r' && i + 1 < csvContent.Length && csvContent[i + 1] == '\n')
                            i++;
                        i++;
                        yield return row;
                        row = new List<string>();
                    }
                    else
                    {
                        currentField.Append(c);
                        i++;
                    }
                }
            }
            row.Add(currentField.ToString());
            yield return row;
        }

        private void MSOptionExport_Click(object? sender, EventArgs e)
        {
            if (_modifiedLocalization.Item1 == null || _modifiedLocalization.Item2 == null)
            {
                MessageBox.Show("Open a localization folder first.", "Export");
                return;
            }

            var cmn = _modifiedLocalization.Item1;
            int startNumber;
            int endNumber;
            using (var rangeDialog = new Interact.ExportRangeDialog(cmn) { StartPosition = FormStartPosition.CenterParent })
            {
                if (rangeDialog.ShowDialog(this) != DialogResult.OK) return;
                startNumber = rangeDialog.StartNumber;
                endNumber = rangeDialog.EndNumber;
            }

            using var saveDialog = new SaveFileDialog
            {
                Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
                DefaultExt = "csv",
                Title = "Export localization to CSV"
            };
            if (saveDialog.ShowDialog() != DialogResult.OK) return;

            var dats = _modifiedLocalization.Item2;
            var allKeys = new List<(string FullName, int StringNumber)>();
            foreach (var rootChild in cmn.Root.Values)
                CollectMatchingKeys(allKeys, rootChild, "");

            var keysInRange = allKeys.Where(k => k.StringNumber >= startNumber && k.StringNumber < endNumber).ToList();

            try
            {
                using var writer = new System.IO.StreamWriter(saveDialog.FileName, false, System.Text.Encoding.UTF8);
                writer.WriteLine(string.Join(",", CsvExportConstants.ColumnHeaders.Select(EscapeCsvField)));
                foreach (var (fullName, stringNumber) in keysInRange)
                {
                    var row = new List<string> { EscapeCsvField(fullName) };
                    for (int i = 0; i < CsvExportConstants.LanguageColumnCount; i++)
                    {
                        string val = (i < dats.Count && stringNumber < dats[i].Strings.Count)
                            ? (dats[i].Strings[stringNumber] ?? "\0")
                            : "\0";
                        row.Add(EscapeCsvField(val));
                    }
                    writer.WriteLine(string.Join(",", row));
                }
                MessageBox.Show($"Exported {keysInRange.Count} string(s) to {saveDialog.FileName}.", "Export");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to export: {ex.Message}", "Export");
            }
        }

        private void MSOptionImport_Click(object? sender, EventArgs e)
        {
            if (_modifiedLocalization.Item1 == null || _modifiedLocalization.Item2 == null)
            {
                MessageBox.Show("Open a localization folder first.", "Import");
                return;
            }

            using var openDialog = new OpenFileDialog
            {
                Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
                DefaultExt = "csv",
                Title = "Import localization from CSV"
            };
            if (openDialog.ShowDialog() != DialogResult.OK) return;

            bool overwriteExisting;
            using (var optionsDialog = new Interact.ImportOptionsDialog { StartPosition = FormStartPosition.CenterParent })
            {
                if (optionsDialog.ShowDialog(this) != DialogResult.OK) return;
                overwriteExisting = optionsDialog.OverwriteExistingStrings;
            }

            var cmn = _modifiedLocalization.Item1;
            var dats = _modifiedLocalization.Item2;
            int added = 0, updated = 0;

            try
            {
                string csvContent = System.IO.File.ReadAllText(openDialog.FileName, System.Text.Encoding.UTF8);
                var rows = ParseCsvRows(csvContent).ToList();
                if (rows.Count < 2)
                {
                    MessageBox.Show("CSV file has no data rows.", "Import");
                    return;
                }
                var headerFields = rows[0];
                const int expectedColumns = 1 + CsvExportConstants.LanguageColumnCount; // Variable + 13 languages
                for (int rowIndex = 1; rowIndex < rows.Count; rowIndex++)
                {
                    var fields = rows[rowIndex];
                    if (fields.Count == 0) continue;
                    // Normalize to expected columns to avoid misalignment from malformed/multi-line rows
                    while (fields.Count < expectedColumns)
                        fields.Add(string.Empty);
                    if (fields.Count > expectedColumns)
                        fields = fields.Take(expectedColumns).ToList();
                    string key = fields[0].Trim();
                    if (string.IsNullOrEmpty(key)) continue;

                    var values = new List<string>();
                    for (int i = 1; i <= CsvExportConstants.LanguageColumnCount && i < fields.Count; i++)
                        values.Add(fields[i] ?? "\0");
                    while (values.Count < CsvExportConstants.LanguageColumnCount)
                        values.Add("\0");

                    if (!cmn.CheckVariableExist(key))
                    {
                        try
                        {
                            cmn.AddVariable(key, cmn.Root, out int variableStringNumber);
                            for (int i = 0; i < dats.Count && i < values.Count; i++)
                            {
                                while (dats[i].Strings.Count <= variableStringNumber)
                                    dats[i].Strings.Add("\0");
                                dats[i].Strings[variableStringNumber] = values[i];
                            }
                            added++;
                        }
                        catch (ArgumentException)
                        {
                            // Duplicate key (e.g. same variable in CSV twice) — treat as update
                            int stringNumber = cmn.GetVariableStringNumber(key);
                            if (stringNumber >= 0)
                            {
                                for (int i = 0; i < dats.Count && i < values.Count; i++)
                                {
                                    if (stringNumber < dats[i].Strings.Count)
                                    {
                                        bool isEmpty = dats[i].Strings[stringNumber] == "\0" || string.IsNullOrEmpty(dats[i].Strings[stringNumber]);
                                        if (overwriteExisting || isEmpty)
                                            dats[i].Strings[stringNumber] = values[i];
                                    }
                                }
                                updated++;
                            }
                        }
                    }
                    else
                    {
                        int stringNumber = cmn.GetVariableStringNumber(key);
                        if (stringNumber >= 0)
                        {
                            for (int i = 0; i < dats.Count && i < values.Count; i++)
                            {
                                if (stringNumber < dats[i].Strings.Count)
                                {
                                    bool isEmpty = dats[i].Strings[stringNumber] == "\0" || string.IsNullOrEmpty(dats[i].Strings[stringNumber]);
                                    if (overwriteExisting || isEmpty)
                                        dats[i].Strings[stringNumber] = values[i];
                                }
                            }
                            updated++;
                        }
                    }
                }

                LoadCmnTreeView();
                LoadDatsDataGridView();
                SavedChanges = false;
                int imported = added + updated;
                MessageBox.Show(imported == 0 ? "0 strings imported." : $"{imported} string(s) imported.", "Import");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to import: {ex.Message}", "Import");
            }
        }

        private void MSMainOpenFolder_Click(object? sender, EventArgs e)
        {
            using FolderBrowserDialog folderBrowser = new FolderBrowserDialog()
            {
                Description = "Select DATs directory",
                RootFolder = Environment.SpecialFolder.MyComputer,
            };

            if (folderBrowser.ShowDialog() == DialogResult.OK)
            {
                string[] files = Directory.GetFiles(folderBrowser.SelectedPath);

                _modifiedLocalization = LoadLocalization(files);
                ClearUndoStack();

                LoadLocalizationForUI(folderBrowser.SelectedPath);

                MSMainSave.Enabled = true;
                MSOptionImportLocalization.Enabled = true;
                MSOptionBatchCopyLanguage.Enabled = _modifiedLocalization.Item2?.Count > 1;
                MSOptionAddAddon.Enabled = true;
                MSOptionExport.Enabled = true;
                MSOptionImport.Enabled = true;
            }
        }

        private void MSMainOpenSingleLanguage_Click(object? sender, EventArgs e)
        {
            using OpenFileDialog cmnDialog = new OpenFileDialog()
            {
                Title = "Select Cmn.dat",
                Filter = "Cmn.dat|Cmn.dat|DAT files (*.dat)|*.dat|All files (*.*)|*.*",
                FileName = "Cmn.dat",
            };

            if (cmnDialog.ShowDialog() != DialogResult.OK)
                return;

            using OpenFileDialog datDialog = new OpenFileDialog()
            {
                Title = "Select language DAT(s) (Ctrl+click for multiple) — A.dat through M.dat",
                Filter = "DAT files (*.dat)|*.dat|All files (*.*)|*.*",
                Multiselect = true,
            };

            if (datDialog.ShowDialog() != DialogResult.OK)
                return;

            try
            {
                _modifiedLocalization = LoadLocalizationSingle(cmnDialog.FileName, datDialog.FileNames);
                ClearUndoStack();

                string folder = Path.GetDirectoryName(cmnDialog.FileName) ?? "";
                LoadLocalizationForUI(folder);

                MSMainSave.Enabled = true;
                MSOptionImportLocalization.Enabled = true;
                MSOptionBatchCopyLanguage.Enabled = _modifiedLocalization.Item2?.Count > 1;
                MSOptionAddAddon.Enabled = true;
                MSOptionExport.Enabled = true;
                MSOptionImport.Enabled = true;
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                MessageBox.Show($"Failed to load: {ex.Message}", "Open single language");
            }
        }

        private void MSMainUndo_Click(object? sender, EventArgs e)
        {
            PerformUndo();
        }

        private void MSMainSave_Click(object? sender, EventArgs e)
        {
            SaveChanges();
        }

        #endregion

        #region Tree View Controls

        private void CmnTreeView_NodeMouseClick(object? sender, TreeNodeMouseClickEventArgs e)
        {
#if DEBUG
            CmnString cmnString = (CmnString)e.Node.Tag;
            Debug.WriteLine("\"" + cmnString.Name + "\"" + " StringNumber : " + cmnString.StringNumber);
#endif
            if (e.Button == MouseButtons.Right)
            {
                CmnTreeView.SelectedNode = e.Node;
                _cmnNodeForContextMenu = e.Node;

                if (CmnTreeView.SelectedNode is TreeNode cmnTreeNode)
                {
                    ContextMenuStrip contextMenu = new ContextMenuStrip();
                    contextMenu.Closed += (_, _) => _cmnNodeForContextMenu = null;

                    ToolStripMenuItem newMenuItem = Utils.CreateToolStripMenuItem("New Variable", "NewVariable", new EventHandler(NewVariableMenuItem_Click), _modifiedLocalization.Item2 == null ? false : true);
                    contextMenu.Items.Add(newMenuItem);

                    ToolStripMenuItem deleteVariableMenuItem = Utils.CreateToolStripMenuItem("Delete Variable", "DeleteVariable", new EventHandler(DeleteVariableMenuItem_Click), _modifiedLocalization.Item2 == null ? false : true);
                    contextMenu.Items.Add(deleteVariableMenuItem);

                    contextMenu.Show(CmnTreeView, CmnTreeView.PointToClient(Cursor.Position));
                }
            }
        }

        private void CmnTreeView_AfterSelect(object? sender, TreeViewEventArgs e)
        {
            LoadDatsDataGridView();
        }

        #endregion

        #region Data Grid View Controls

        private void DatsDataGridView_SelectionChanged(object? sender, EventArgs e)
        {
            // Limit user to multiselect in columns and in only one
            switch (DatsDataGridView.SelectedCells.Count)
            {
                case 0:
                    _selectedRowIndex = -1;
                    _selectedColumnIndex = -1;
                    return;
                case 1:
                    _selectedRowIndex = DatsDataGridView.SelectedCells[0].RowIndex;
                    _selectedColumnIndex = DatsDataGridView.SelectedCells[0].ColumnIndex;
                    return;
            }

            foreach (DataGridViewCell cell in DatsDataGridView.SelectedCells)
            {
                if (cell.ColumnIndex == _selectedColumnIndex)
                {
                    if (cell.RowIndex != _selectedRowIndex)
                    {
                        _selectedRowIndex = -1;
                    }
                }
                else
                {
                    cell.Selected = false;
                }
            }
        }

        private void DatsDataGridView_CellMouseClick(object? sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right && DatsDataGridView.SelectedCells.Count > 0)
            {
                ContextMenuStrip contextMenu = new ContextMenuStrip();

                if (e.ColumnIndex != 2)
                {
                    ToolStripMenuItem copyVariablesMenuItem = Utils.CreateToolStripMenuItem("Copy Variables", "CopyVariables", new EventHandler(CopyVariableStringMenuItem_Click), DatsDataGridView.SelectedCells.Count == 0 ? false : true);
                    contextMenu.Items.Add(copyVariablesMenuItem);

                    ToolStripMenuItem pasteVariablesMenuItem = Utils.CreateToolStripMenuItem("Paste Variables", "PasteVariables", new EventHandler(PasteVariableStringMenuItem_Click), _copyVariableStrings.Item2 == null ? false : true);
                    contextMenu.Items.Add(pasteVariablesMenuItem);

                    ToolStripMenuItem copyPasteToLanguagesMenuItem = Utils.CreateToolStripMenuItem("Copy Paste to languages", "CopyPasteToLanguages", new EventHandler(CopyPasteToLanguagesMenuItem_Click), DatsDataGridView.SelectedCells.Count == 0 ? false : true);
                    contextMenu.Items.Add(copyPasteToLanguagesMenuItem);
                }
                else
                {
                    ToolStripMenuItem copyStringsMenuItem = Utils.CreateToolStripMenuItem("Copy Strings", "CopyStrings", new EventHandler(CopyStringMenuItem_Click));
                    contextMenu.Items.Add(copyStringsMenuItem);

                    ToolStripMenuItem pasteStringsMenuItem = Utils.CreateToolStripMenuItem("Paste Strings", "PasteStrings", new EventHandler(PasteStringMenuItem_Click));
                    contextMenu.Items.Add(pasteStringsMenuItem);
                }


                ToolStripMenuItem selectAllMenuItem = Utils.CreateToolStripMenuItem("Select All", "SelectAll", new EventHandler(SelectAllMenuItem_Click));
                contextMenu.Items.Add(selectAllMenuItem);

                ToolStripMenuItem deSelectAllMenuItem = Utils.CreateToolStripMenuItem("Deselect All", "DeselectAll", new EventHandler(DeselectAllMenuItem_Click), DatsDataGridView.SelectedCells.Count == 0 ? false : true);
                contextMenu.Items.Add(deSelectAllMenuItem);

                contextMenu.Show(DatsDataGridView, DatsDataGridView.PointToClient(Cursor.Position));
            }
        }

        private void DatsDataGridView_CellDoubleClick(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex != 2 || e.RowIndex == -1 || _modifiedLocalization.Item2 == null)
                return;
            var dats = _modifiedLocalization.Item2;
            string datText = DatsDataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString() ?? "";
            int stringNumber = Convert.ToInt32(DatsDataGridView.Rows[e.RowIndex].Cells[0].Value);

            using (var datStringEditor = new DatStringEditor(datText) { StartPosition = FormStartPosition.CenterScreen })
            {
                datStringEditor.ShowDialog();
                if (datStringEditor.DialogResult == DialogResult.OK)
                {
                    _undoStack.Push(new TextEditUndoAction(_selectedDatIndex, stringNumber, datText));
                    UpdateUndoMenuItemEnabled();
                    SavedChanges = false;
                    DatsDataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = datStringEditor.DatText;
                    dats[_selectedDatIndex].Strings[stringNumber] = datStringEditor.DatText;
                }
                datStringEditor.Dispose();
            }
        }

        private void DatsDataGridView_MouseDown(object? sender, MouseEventArgs e)
        {
            DataGridView.HitTestInfo hitTest = DatsDataGridView.HitTest(e.X, e.Y);

            if (hitTest.Type == DataGridViewHitTestType.None && DatsDataGridView.SelectedCells.Count > 0)
            {
                DatsDataGridView.ClearSelection();
            }
        }

        #endregion

        #region Combo Box Controls
        private void DatLanguageComboBox_SelectedValueChanged(object? sender, EventArgs e)
        {
            LoadDatsDataGridView();
        }

        #endregion

        #region Menu Items

        #region CmnTreeView Menu Items

        private void NewVariableMenuItem_Click(object? sender, EventArgs e)
        {
            if (_modifiedLocalization.Item1 == null || _modifiedLocalization.Item2 == null)
                return;
            TreeNode? treeNode = CmnTreeView.SelectedNode;
            if (treeNode == null)
                return;
            int treeNodeIndex = treeNode.Index;
            TreeNode? treeNodeParent = treeNode.Parent;

            using (var input = new Input("Enter a name for the new node", treeNode.Text) { StartPosition = FormStartPosition.CenterScreen })
            {
                input.ShowDialog();

                if (input.DialogResult == DialogResult.OK)
                {
                    string suffix = (input.InputText ?? "").Trim();
                    if (string.IsNullOrEmpty(suffix))
                        return;
                    string fullName = treeNode.Text + suffix;
                    CmnString selectedCmnNode = (CmnString)treeNode.Tag;

                    // Find the deepest existing node under the selected one that is a prefix of fullName,
                    // so the new variable is placed in the correct alphabetical position in the tree.
                    CmnString addUnder = selectedCmnNode;
                    while (true)
                    {
                        CmnString? bestChild = null;
                        foreach (var child in addUnder.Childrens.Values)
                        {
                            if (fullName.StartsWith(child.Name, StringComparison.Ordinal) && child.Name.Length < fullName.Length)
                            {
                                if (bestChild == null || child.Name.Length > bestChild.Name.Length)
                                    bestChild = child;
                            }
                        }
                        if (bestChild == null)
                            break;
                        addUnder = bestChild;
                    }

                    if (_modifiedLocalization.Item1.AddVariable(fullName, addUnder.Childrens, out int variableStringNumber))
                    {
                        foreach (var dat in _modifiedLocalization.Item2)
                        {
                            dat.Strings.Add("\0");
                        }

                        // If the new node is a prefix of existing siblings, move those siblings under it
                        _modifiedLocalization.Item1.MoveSiblingsUnderNewNode(addUnder.Childrens, fullName);

                        // Ensure the whole tree has the correct connected hierarchy (e.g. AircraftTest_Name_ with _f15c and _f15t under it)
                        _modifiedLocalization.Item1.EnsureAllIntermediateParents();

                        // Remove the treeNode because we need to update it
                        treeNode.Remove();

                        // Re-insert the node after updating it
                        TreeNode reInsertTreeNode = GetTreeNodeFromCmn(selectedCmnNode); // Create the new treeNode and its childrens

                        // Re-insert the tree node in the right place
                        if (treeNodeParent == null)
                        {
                            CmnTreeView.Nodes.Insert(treeNodeIndex, reInsertTreeNode);
                        }
                        else
                        {
                            treeNodeParent.Nodes.Insert(treeNodeIndex, reInsertTreeNode);
                        }

                        TreeNode? newNode = FindTreeNodeByCmnName(reInsertTreeNode, fullName);
                        if (newNode != null)
                        {
                            // Expand path from root to new node so it's visible
                            for (TreeNode? n = newNode.Parent; n != null; n = n.Parent)
                                n.Expand();
                            CmnTreeView.SelectedNode = newNode;
                        }
                        else
                        {
                            reInsertTreeNode.Expand();
                            CmnTreeView.SelectedNode = reInsertTreeNode;
                        }

                        SavedChanges = false;
                    }
                    else
                    {
                        // Show Message Box
                    }
                }

                input.Dispose();
            }
        }

        private void RenameVariableMenuItem_Click(object? sender, EventArgs e)
        {
            TreeNode treeNode = CmnTreeView.SelectedNode;
            int treeNodeIndex = treeNode.Index;

            TreeNode treeNodeParent = treeNode.Parent;

            using (var input = new Input("Enter a new name for the node", treeNodeParent.Text) { StartPosition = FormStartPosition.CenterScreen })
            {
                input.ShowDialog();

                if (input.DialogResult == DialogResult.OK)
                {
                    CmnString selectedCmnNode = (CmnString)treeNode.Tag;
                }

                input.Dispose();
            }
        }

        private void DeleteVariableMenuItem_Click(object? sender, EventArgs e)
        {
            if (_modifiedLocalization.Item1 == null || _modifiedLocalization.Item2 == null)
                return;
            // Use node captured at context menu open; TreeView selection is often cleared when the menu is shown
            TreeNode? treeNode = _cmnNodeForContextMenu ?? CmnTreeView.SelectedNode;
            if (treeNode == null)
                return;

            if (treeNode.Tag is not CmnString selectedCmnNode)
                return;

            string variableName = selectedCmnNode.Name;

            var result = MessageBox.Show(
                $"Delete variable \"{variableName}\"? This will remove it and all its child variables from the CMN and all language DATs.",
                "Delete Variable",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);
            if (result != DialogResult.Yes)
                return;

            List<int> variableStringNumbers = _modifiedLocalization.Item1.DeleteVariable(selectedCmnNode);
            var dats = _modifiedLocalization.Item2;

            foreach (var variableStringNumber in variableStringNumbers)
            {
                if (variableStringNumber >= 0)
                {
                    foreach (var dat in dats)
                    {
                        if (variableStringNumber < dat.Strings.Count)
                            dat.Strings.RemoveAt(variableStringNumber);
                    }
                }
            }

            LoadCmnTreeView();
            LoadDatsDataGridView();
            SavedChanges = false;
        }

        #endregion

        #region DatsGridView Menu Items

        #region Variable Column (0 and 1)

        private void CopyVariableStringMenuItem_Click(object? sender, EventArgs e)
        {
            CopyVariables();
        }

        private void PasteVariableStringMenuItem_Click(object? sender, EventArgs e)
        {
            PasteVariables();
        }

        private void CopyPasteToLanguagesMenuItem_Click(object? sender, EventArgs e)
        {
            if (DatLanguageComboBox.SelectedItem is not DAT selectedDat || _modifiedLocalization.Item2 == null)
                return;
            var dats = _modifiedLocalization.Item2;
            using (var copyPasteLanguagesSelector = new CopyPasteLanguagesSelector(dats, selectedDat) { StartPosition = FormStartPosition.CenterScreen })
            {
                copyPasteLanguagesSelector.ShowDialog();

                foreach (DataGridViewCell selectedCell in DatsDataGridView.SelectedCells)
                {
                    int variableStringNumber = (int)DatsDataGridView.Rows[selectedCell.RowIndex].Cells[0].Value;
                    if (variableStringNumber != -1)
                    {
                        foreach (char datLetter in copyPasteLanguagesSelector.SelectedLanguages)
                        {
                            int targetIndex = dats.FindIndex(d => d.Letter == datLetter);
                            if (targetIndex >= 0)
                                dats[targetIndex].Strings[variableStringNumber] = dats[_selectedDatIndex].Strings[variableStringNumber];
                        }
                    }
                }

                copyPasteLanguagesSelector.Dispose();
            }
        }

        #endregion

        #region String Column (2)

        private void CopyStringMenuItem_Click(object? sender, EventArgs e)
        {
            CopyStrings();
        }

        private void PasteStringMenuItem_Click(object? sender, EventArgs e)
        {
            PasteStrings();
        }

        #endregion

        private void SelectAllMenuItem_Click(object? sender, EventArgs e)
        {
            for (int i = 0; i < DatsDataGridView.Rows.Count; i++)
            {
                DatsDataGridView.Rows[i].Selected = true;
            }
        }

        private void DeselectAllMenuItem_Click(object? sender, EventArgs e)
        {
            DatsDataGridView.ClearSelection();
        }

        #endregion

        #endregion

        private void DatsDataGridView_KeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.Z)
            {
                PerformUndo();
                e.Handled = true;
                return;
            }
            if (e.Control && e.KeyCode == Keys.C)
            {
                if (DatsDataGridView.SelectedCells.Count > 0 && DatsDataGridView.SelectedCells[0].ColumnIndex != 2)
                {
                    CopyVariables();
                }
                else if (DatsDataGridView.SelectedCells.Count > 0)
                {
                    CopyStrings();
                }

            }
            else if (e.Control && e.KeyCode == Keys.V)
            {
                if (DatsDataGridView.SelectedCells.Count > 0 && DatsDataGridView.SelectedCells[0].ColumnIndex != 2)
                {
                    PasteVariables();
                }
                else if (DatsDataGridView.SelectedCells.Count > 0)
                {
                    PasteStrings();
                }
            }
        }

        private void LocalizationEditor_KeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.Z)
            {
                PerformUndo();
                e.Handled = true;
            }
            else if (e.Control && e.KeyCode == Keys.S && SavedChanges == false)
            {
                SaveChanges();
                SavedChanges = true;
            }
        }
    }
}
