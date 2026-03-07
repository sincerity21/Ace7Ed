using Ace7Ed.Properties;
using Ace7Ed.Prompt;
using Ace7LocalizationFormat.Formats;
using System.ComponentModel;
using System.Windows.Forms;
using static Ace7LocalizationFormat.Formats.CMN;
using Ace7Ed.Interact;
using System.Diagnostics;
using System.ComponentModel.Design.Serialization;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Ace7Ed
{
    public partial class LocalizationEditor : Form
    {
        private (CMN?, List<DAT>?) _modifiedLocalization = (null, null);

        private string _directory { get; set; } = "";

        private (int, List<int>?) _copyVariableStrings { get; set; }
        private List<string> _copyStrings = new List<string>();

        private int _selectedRowIndex = -1;
        private int _selectedColumnIndex = -1;

        private int _selectedDatIndex
        {
            get
            {
                if (DatLanguageComboBox.SelectedItem != null)
                {
                    DAT dat = (DAT)DatLanguageComboBox.SelectedItem;
                    return dat.Letter - 65;
                }
                return -1;

            }
        }

        private bool _savedChanges = true;
        private bool _isClosingDeferred;
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

            if (Clipboard.GetText() != null)
            {
                _copyStrings.Add(Clipboard.GetText());
            }

            MSOptionImportLocalization.Enabled = false;
            MSOptionImportLocalization.Visible = false;
        }

        private void LocalizationEditor_FormClosing(object? sender, FormClosingEventArgs e)
        {
            if (!_isClosingDeferred)
            {
                _isClosingDeferred = true;
                e.Cancel = true;
                Hide();
                BeginInvoke(() =>
                {
                    ClearHeavyDataBeforeClose();
                    Close();
                });
            }
            else
            {
                ClearHeavyDataBeforeClose();
            }
        }

        private void ClearHeavyDataBeforeClose()
        {
            CmnTreeView.Nodes.Clear();
            DatsDataGridView.Rows.Clear();
            DatsDataGridView.Columns.Clear();
            DatLanguageComboBox.Items.Clear();
            _modifiedLocalization = (null, null);
        }

        private void ToggleDarkTheme()
        {
            BackColor = Theme.ControlColor;
            ForeColor = Theme.ControlTextColor;

            LocalizationEditorMenuStrip.Renderer = new Theme.MenuStripRenderer();

            Theme.SetDarkThemeToolStripMenuItem(MenuStripMain);
            Theme.SetDarkThemeToolStripMenuItem(MSMainOpenFolder);
            Theme.SetDarkThemeToolStripMenuItem(MSMainSave);

            Theme.SetDarkThemeToolStripMenuItem(MenuStripOptions);
            Theme.SetDarkThemeToolStripMenuItem(MSOptionImportLocalization);
            Theme.SetDarkThemeToolStripMenuItem(MSOptionBatchCopyLanguage);
            Theme.SetDarkThemeToolStripMenuItem(MSOptionsToggleDarkTheme);
            Theme.SetDarkThemeToolStripMenuItem(MSOptionAddAddon);
            Theme.SetDarkThemeToolStripMenuItem(MSOptionCopyAddon);
            Theme.SetDarkThemeToolStripMenuItem(MSOptionPasteAddon);

            Theme.SetDarkThemeComboBox(DatLanguageComboBox);

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
            CmnTreeView.BeginUpdate();

            CmnTreeView.Nodes.Clear();

            foreach (var node in _modifiedLocalization.Item1.Root.Childrens)
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

            if (DatLanguageComboBox.SelectedItem != null && _modifiedLocalization.Item1 != null && _modifiedLocalization.Item2 != null)
            {
                DatsDataGridView.Rows.Clear();
                DAT dat = _modifiedLocalization.Item2[_selectedDatIndex];

                {
                    // Load strings of the selected CMN
                    if (CmnTreeView.SelectedNode is TreeNode treeNode)
                    {
                        AddCmnNodeToDataGridView(dat, (CmnString)treeNode.Tag);
                    }
                    // Load all strings
                    else if (CmnTreeView.SelectedNode == null)
                    {
                        foreach (var node in _modifiedLocalization.Item1.Root.Childrens)
                        {
                            AddCmnNodeToDataGridView(dat, node);
                        }
                    }
                }
            }

            DatsDataGridView.Sort(DatsDataGridView.Columns[0], ListSortDirection.Ascending);

            DatsDataGridView.ClearSelection();
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
        }

        #endregion

        #region Manage Controls

        public TreeNode GetTreeNodeFromCmn(CmnString parent)
        {
            var root = new TreeNode();
            root.Name = parent.Name;
            root.Text = parent.FullName;
            root.Tag = parent;

            List<TreeNode> nodesToVisit = [root];

            int index = 0;
            while (index < nodesToVisit.Count)
            {
                TreeNode currentCmnNode = nodesToVisit[index];
                CmnString cmnString = (CmnString)currentCmnNode.Tag;
                foreach (var child in cmnString.Childrens)
                {
                    TreeNode subNode = new TreeNode();
                    subNode.Name = child.Name;
                    subNode.Text = child.FullName;
                    subNode.Tag = child;
                    currentCmnNode.Nodes.Add(subNode);
                    nodesToVisit.Add(subNode);
                }
                index++;
            }

            return root;
        }

        public void ImportLocalization((CMN, List<DAT>) localization)
        {
            var cmn = _modifiedLocalization.Item1;
            var dats = _modifiedLocalization.Item2;
            if (cmn == null || dats == null)
                return;
            List<CmnString> nodesToVisit = [localization.Item1.Root];

            int index = 0;
            while (index < nodesToVisit.Count)
            {
                CmnString cmnString = nodesToVisit[index];
                foreach (var child in cmnString.Childrens)
                {
                    bool isVariableAdded = cmn.AddVariable(child.FullName, cmn.Root, out int variableStringNumber, child.StringNumber == -1);

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

        private void AddCmnNodeToDataGridView(DAT dat, CmnString parent)
        {
            if (parent.StringNumber != -1)
            {
                string text = dat.Strings[parent.StringNumber];
                DatsDataGridView.Rows.Add(parent.StringNumber, parent.FullName, text);
            }

            foreach (var children in parent.Childrens)
            {
                AddCmnNodeToDataGridView(dat, children);
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

                LoadLocalizationForUI(folderPath);

                MSOptionImportLocalization.Enabled = true;
                MSOptionBatchCopyLanguage.Enabled = true;
                MSOptionAddAddon.Enabled = true;
                MSOptionCopyAddon.Enabled = true;
                MSOptionPasteAddon.Enabled = true;
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
                        for (int i = batchCopyLanguage.StartNumber; i < batchCopyLanguage.EndNumber; i++)
                        {
                            if (dats[pasteLanguageLetter - 65].Strings[i] == "\0" || batchCopyLanguage.OverwriteExistingString)
                            {
                                dats[pasteLanguageLetter - 65].Strings[i] = dats[batchCopyLanguage.SelectedCopyLanguageIndex].Strings[i];
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
                MessageBox.Show("Open a localization folder first.", "Add add-on");
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
                MessageBox.Show("Plane string cannot be empty.", "Add add-on");
                return;
            }

            int skinCount;
            using (var input = new Input("Enter number of skins (minimum 6)", "") { StartPosition = FormStartPosition.CenterScreen })
            {
                input.ShowDialog();
                if (input.DialogResult != DialogResult.OK) return;
                if (!int.TryParse(input.InputText?.Trim(), out skinCount) || skinCount < 6)
                {
                    MessageBox.Show("Please enter a number of skins (minimum 6).", "Add add-on");
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
                keysToAdd.Add($"AircraftSkin_Name_{plane}_{i:D2}");
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
                if (!cmn.CheckVariableExist(key))
                    continue; // already exists, skip
                cmn.AddVariable(key, cmn.Root, out int _);
                foreach (var dat in dats)
                    dat.Strings.Add("\0");
            }

            LoadCmnTreeView();
            LoadDatsDataGridView();
            SavedChanges = false;
        }

        private static void CollectMatchingKeys(List<(string FullName, int StringNumber)> result, CMN.CmnString node, string filter)
        {
            if (node.StringNumber != -1 && node.FullName.Contains(filter, StringComparison.OrdinalIgnoreCase))
                result.Add((node.FullName, node.StringNumber));
            foreach (var child in node.Childrens)
                CollectMatchingKeys(result, child, filter);
        }

        private void MSOptionCopyAddon_Click(object? sender, EventArgs e)
        {
            if (_modifiedLocalization.Item1 == null || _modifiedLocalization.Item2 == null)
            {
                MessageBox.Show("Open a localization folder first.", "Copy add-on");
                return;
            }

            string filter;
            using (var input = new Input("Enter substring to copy (e.g. f15t)", "") { StartPosition = FormStartPosition.CenterScreen })
            {
                input.ShowDialog();
                if (input.DialogResult != DialogResult.OK) return;
                filter = (input.InputText ?? "").Trim();
            }
            if (string.IsNullOrEmpty(filter))
            {
                MessageBox.Show("Substring cannot be empty.", "Copy add-on");
                return;
            }

            var cmn = _modifiedLocalization.Item1;
            var dats = _modifiedLocalization.Item2;
            var matching = new List<(string FullName, int StringNumber)>();
            foreach (var rootChild in cmn.Root.Childrens)
                CollectMatchingKeys(matching, rootChild, filter);

            if (matching.Count == 0)
            {
                MessageBox.Show($"No strings found containing \"{filter}\".", "Copy add-on");
                return;
            }

            var entries = new List<JObject>();
            foreach (var (fullName, stringNumber) in matching)
            {
                var values = new JArray();
                foreach (var dat in dats)
                {
                    string val = stringNumber < dat.Strings.Count ? dat.Strings[stringNumber] : "\0";
                    values.Add(val ?? "\0");
                }
                entries.Add(new JObject
                {
                    ["key"] = fullName,
                    ["values"] = values
                });
            }
            var json = new JObject { ["entries"] = new JArray(entries) };
            try
            {
                Clipboard.SetText(json.ToString());
                MessageBox.Show($"Copied {matching.Count} string(s) to clipboard.", "Copy add-on");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to copy to clipboard: {ex.Message}", "Copy add-on");
            }
        }

        private void MSOptionPasteAddon_Click(object? sender, EventArgs e)
        {
            if (_modifiedLocalization.Item1 == null || _modifiedLocalization.Item2 == null)
            {
                MessageBox.Show("Open a localization folder first.", "Paste add-on");
                return;
            }

            string clipboardText;
            try
            {
                clipboardText = Clipboard.GetText();
            }
            catch
            {
                MessageBox.Show("Could not read clipboard.", "Paste add-on");
                return;
            }
            if (string.IsNullOrWhiteSpace(clipboardText))
            {
                MessageBox.Show("Clipboard is empty.", "Paste add-on");
                return;
            }

            JObject? json;
            try
            {
                json = JObject.Parse(clipboardText);
            }
            catch
            {
                MessageBox.Show("Clipboard does not contain valid add-on data (expected JSON from Copy add-on).", "Paste add-on");
                return;
            }

            var entriesToken = json["entries"] as JArray;
            if (entriesToken == null || entriesToken.Count == 0)
            {
                MessageBox.Show("No entries found in add-on data.", "Paste add-on");
                return;
            }

            var cmn = _modifiedLocalization.Item1;
            var dats = _modifiedLocalization.Item2;
            const int langCount = 13;
            int added = 0, updated = 0;

            foreach (var entry in entriesToken.Cast<JObject>())
            {
                var keyToken = entry["key"];
                var valuesToken = entry["values"] as JArray;
                if (keyToken == null || string.IsNullOrEmpty(keyToken.ToString()) || valuesToken == null)
                    continue;

                string key = keyToken.ToString();
                var values = new List<string>();
                for (int i = 0; i < langCount && i < valuesToken.Count; i++)
                    values.Add(valuesToken[i]?.ToString() ?? "\0");
                while (values.Count < langCount)
                    values.Add("\0");

                if (cmn.CheckVariableExist(key))
                {
                    cmn.AddVariable(key, cmn.Root, out int _);
                    for (int i = 0; i < dats.Count && i < values.Count; i++)
                        dats[i].Strings.Add(values[i]);
                    added++;
                }
                else
                {
                    int stringNumber = cmn.GetVariableStringNumber(key);
                    if (stringNumber >= 0)
                    {
                        for (int i = 0; i < dats.Count && i < values.Count; i++)
                        {
                            if (stringNumber < dats[i].Strings.Count)
                                dats[i].Strings[stringNumber] = values[i];
                        }
                        updated++;
                    }
                }
            }

            LoadCmnTreeView();
            LoadDatsDataGridView();
            SavedChanges = false;
            MessageBox.Show($"Paste complete: {added} new string(s), {updated} updated.", "Paste add-on");
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

                LoadLocalizationForUI(folderBrowser.SelectedPath);

                MSOptionImportLocalization.Enabled = true;
                MSOptionBatchCopyLanguage.Enabled = true;
                MSOptionAddAddon.Enabled = true;
                MSOptionCopyAddon.Enabled = true;
                MSOptionPasteAddon.Enabled = true;
            }
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
            Debug.WriteLine("\"" + cmnString.FullName + "\"" + " Index : " + cmnString.Index);
#endif
            if (e.Button == MouseButtons.Right)
            {
                CmnTreeView.SelectedNode = e.Node;

                if (CmnTreeView.SelectedNode is TreeNode cmnTreeNode)
                {
                    ContextMenuStrip contextMenu = new ContextMenuStrip();

                    ToolStripMenuItem newMenuItem = Utils.CreateToolStripMenuItem("New", "New", new EventHandler(NewVariableMenuItem_Click), _modifiedLocalization.Item2 == null ? false : true);
                    contextMenu.Items.Add(newMenuItem);

                    //ToolStripMenuItem renameMenuItem = Utils.CreateToolStripMenuItem("Rename", "Rename", new EventHandler(RenameMenuItem_Click), _modifiedLocalization.Item2 == null ? false : true);
                    //contextMenu.Items.Add(renameMenuItem);

                    //ToolStripMenuItem deleteMenuItem = Utils.CreateToolStripMenuItem("Delete", "Delete", new EventHandler(DeleteMenuItem_Click), _modifiedLocalization.Item2 == null ? false : true);
                    //contextMenu.Items.Add(deleteMenuItem);

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
                    CmnString selectedCmnNode = (CmnString)treeNode.Tag;
                    if (_modifiedLocalization.Item1.AddVariable(input.InputText, selectedCmnNode, out int variableStringNumber))
                    {
                        foreach (var dat in _modifiedLocalization.Item2)
                        {
                            dat.Strings.Add("\0");
                        }

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

                        CmnTreeView.SelectedNode = reInsertTreeNode;
                        CmnTreeView.SelectedNode.Expand();

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
            TreeNode? treeNode = CmnTreeView.SelectedNode;
            if (treeNode == null)
                return;

            CmnString selectedCmnNode = (CmnString)treeNode.Tag;

            List<int> variableStringNumbers = _modifiedLocalization.Item1.DeleteVariable(selectedCmnNode);
            var dats = _modifiedLocalization.Item2;

            foreach (var variableStringNumber in variableStringNumbers)
            {
                if (variableStringNumber != -1)
                {
                    foreach (var dat in dats)
                    {
                        dat.Strings.RemoveAt(variableStringNumber);
                    }
                }
            }
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
                            dats[datLetter - 65].Strings[variableStringNumber] = dats[_selectedDatIndex].Strings[variableStringNumber];
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
            if (e.Control && e.KeyCode == Keys.S && SavedChanges == false)
            {
                SaveChanges();
                SavedChanges = true;
            }
        }
    }
}
