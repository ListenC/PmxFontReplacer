using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using PEPlugin;

namespace PmxFontReplacer
{
    public class PmxFontReplacerPlugin : PEPluginClass
    {
        private readonly string DefaultFontName = "宋体";
        private const float DefaultFontSize = 9f;
        private readonly FontStyle DefaultFontStyle = FontStyle.Regular;

        private Font targetFont;
        private List<string> targetSourceFonts = new List<string>();

        private string configFile;
        private int runCount = 0;

        private bool enabledFontReplace;         // 当前已生效状态
        private bool tempEnabledFontReplace;     // 临时 Checkbox 状态

        private bool autoSize;                   //自动尺寸

        // 跟踪已挂事件的对象，便于在窗体关闭时清理
        private HashSet<ToolStripDropDownItem> hookedDropDownItems = new HashSet<ToolStripDropDownItem>();
        private HashSet<ContextMenuStrip> hookedContextMenus = new HashSet<ContextMenuStrip>();
        private HashSet<ToolStrip> hookedToolStrips = new HashSet<ToolStrip>();

        public PmxFontReplacerPlugin()
        {
            m_option = new PEPluginOption(true, true, "PE字体替换");

            string dllDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            configFile = Path.Combine(dllDir, "PmxFontReplacer.config");

            LoadConfig();

            RefreshAllOpenForms();

            Application.Idle += (s, e) =>
            {
                if (enabledFontReplace)
                {
                    foreach (Form f in Application.OpenForms)
                    {
                        if (!f.Tag?.Equals("FontReplaced") ?? true)
                            ReplaceFontsInForm(f);
                    }
                }
            };

            AppDomain.CurrentDomain.AssemblyLoad += (sender, e) =>
            {
                try { ReplaceFontsInAssembly(e.LoadedAssembly); } catch { }
            };
        }

        public override void Run(IPERunArgs app)
        {
            runCount++;
            if (runCount == 1)
            {
                RefreshAllOpenForms();
                return;
            }
            ShowFontSettings();
        }

        private void LoadConfig()
        {
            try
            {
                if (File.Exists(configFile))
                {
                    var lines = File.ReadAllLines(configFile);
                    if (lines.Length > 0)
                    {
                        var fontLine = lines[0];
                        if (fontLine.StartsWith("TargetFont:"))
                        {
                            var parts = fontLine.Substring("TargetFont:".Length).Split('|');
                            if (parts.Length >= 3)
                            {
                                string fname = parts[0];
                                float fsize = float.Parse(parts[1]);
                                FontStyle fstyle = (FontStyle)Enum.Parse(typeof(FontStyle), parts[2]);
                                targetFont = new Font(fname, fsize, fstyle);
                            }
                        }

                        targetSourceFonts = lines
                            .Where(l => l.StartsWith("SourceFont:"))
                            .Select(l => l.Substring("SourceFont:".Length))
                            .ToList();

                        // 读取 Checkbox 状态
                        var enableLine = lines.FirstOrDefault(l => l.StartsWith("EnableFontReplace:"));
                        if (enableLine != null)
                        {
                            bool.TryParse(enableLine.Substring("EnableFontReplace:".Length), out enabledFontReplace);
                            tempEnabledFontReplace = enabledFontReplace;
                        }

                        // 读取 AutoSize 状态
                        var sizeLine = lines.FirstOrDefault(l => l.StartsWith("AutoSize:"));
                        if (sizeLine != null)
                        {
                            bool.TryParse(sizeLine.Substring("AutoSize:".Length), out autoSize);
                        }

                        if (targetFont != null) return;
                    }
                }
            }
            catch { }

            targetFont = new Font(DefaultFontName, DefaultFontSize, DefaultFontStyle);
            targetSourceFonts = new List<string> { "MS UI Gothic", "Yu Gothic UI" };
            enabledFontReplace = tempEnabledFontReplace = true;
            autoSize = true;
            SaveConfig();
        }

        private void SaveConfig()
        {
            try
            {
                var lines = new List<string>
                {
                    $"TargetFont:{targetFont.FontFamily.Name}|{targetFont.Size}|{targetFont.Style}",
                    $"EnableFontReplace:{tempEnabledFontReplace}",
                    $"AutoSize:{autoSize}"
                };
                foreach (var sf in targetSourceFonts)
                    lines.Add("SourceFont:" + sf);
                File.WriteAllLines(configFile, lines);
            }
            catch { }
        }

        private void RefreshAllOpenForms()
        {
            foreach (Form f in Application.OpenForms)
            {
                f.Tag = null;
                ReplaceFontsInForm(f);
            }
        }

        private void ReplaceFontsInForm(Form form)
        {
            if (form.Tag?.Equals("FontReplaced") ?? false) return;

            form.Tag = "FontReplaced";
            ReplaceFontsInControlRecursive(form);

            // 绑定窗体级别事件（先清除防重复）
            form.ControlAdded -= Form_ControlAdded;
            form.ControlAdded += Form_ControlAdded;

            form.Shown -= Form_Shown;
            form.Shown += Form_Shown;

            // 订阅 FormClosed，用于清理该窗体注册的事件（释放 HashSet 与事件）
            form.FormClosed -= Form_FormClosed;
            form.FormClosed += Form_FormClosed;
        }

        private void Form_Shown(object sender, EventArgs e)
        {
            if (sender is Form f)
                ReplaceFontsInControlRecursive(f);
        }

        private void Form_ControlAdded(object sender, ControlEventArgs e)
        {
            ReplaceFontsInControlRecursive(e.Control);
        }

        // 当窗体关闭时，解除该窗体上注册的事件并清理对应 HashSet 项
        private void Form_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (sender is Form f)
            {
                try { UnhookEventsForForm(f); } catch { }
            }
        }

        // 遍历并解除该窗体所有我们曾经添加的监听（ToolStrip, ContextMenuStrip, DropDownItem, Control 的 ContextMenuStripChanged 等）
        private void UnhookEventsForForm(Form form)
        {
            if (form == null) return;

            // 移除窗体级别绑定
            try
            {
                form.ControlAdded -= Form_ControlAdded;
                form.Shown -= Form_Shown;
                form.FormClosed -= Form_FormClosed;
            }
            catch { }

            // 遍历所有控件（包含自身）
            foreach (Control ctrl in GetAllControls(form))
            {
                try
                {
                    // 解除 ContextMenuStripChanged 事件（如果之前添加过）
                    try { ctrl.ContextMenuStripChanged -= Ctrl_ContextMenuStripChanged; } catch { }

                    // 清理 ContextMenuStrip 及其项的监听
                    if (ctrl.ContextMenuStrip != null)
                    {
                        var cms = ctrl.ContextMenuStrip;
                        if (hookedContextMenus.Contains(cms))
                        {
                            try { cms.Opening -= Cms_Opening; } catch { }
                            hookedContextMenus.Remove(cms);
                        }

                        // 递归解除 cms 内部项的 DropDownOpening 绑定
                        foreach (ToolStripItem item in cms.Items)
                            UnhookToolStripItemRecursive(item);
                    }

                    // 如果控件本身是 ToolStrip（包括 MenuStrip、StatusStrip 等继承自 ToolStrip），解除 ItemAdded 并解除子项监听
                    if (ctrl is ToolStrip ts)
                    {
                        if (hookedToolStrips.Contains(ts))
                        {
                            try { ts.ItemAdded -= Ts_ItemAdded; } catch { }
                            hookedToolStrips.Remove(ts);
                        }

                        foreach (ToolStripItem item in ts.Items)
                            UnhookToolStripItemRecursive(item);
                    }
                }
                catch { }
            }
        }

        // 递归解除单个 ToolStripItem 的 DropDownOpening 监听，并从 HashSet 移除
        private void UnhookToolStripItemRecursive(ToolStripItem item)
        {
            if (item == null) return;
            try
            {
                if (item is ToolStripDropDownItem drop)
                {
                    if (hookedDropDownItems.Contains(drop))
                    {
                        try { drop.DropDownOpening -= Drop_DropDownOpening; } catch { }
                        hookedDropDownItems.Remove(drop);
                    }

                    foreach (ToolStripItem sub in drop.DropDownItems)
                        UnhookToolStripItemRecursive(sub);
                }
            }
            catch { }
        }

        // 辅助：遍历控件树（包含 parent 自身）
        private IEnumerable<Control> GetAllControls(Control parent)
        {
            var stack = new Stack<Control>();
            stack.Push(parent);
            while (stack.Count > 0)
            {
                var c = stack.Pop();
                yield return c;
                foreach (Control child in c.Controls)
                    stack.Push(child);
            }
        }

        private void ReplaceFontsInControlRecursive(Control ctrl)
        {
            if (ctrl == null || !enabledFontReplace) return;

            try
            {
                if (ctrl.Font != null &&
                    (targetSourceFonts.Count == 0 || targetSourceFonts.Contains(ctrl.Font.FontFamily.Name)))
                {
                    ctrl.Font = new Font(targetFont.FontFamily, autoSize != true ? targetFont.Size : ctrl.Font.Size, targetFont.Style, ctrl.Font.Unit);
                }
            }
            catch { }

            // 如果控件有 ContextMenuStrip，递归替换并监听 Opening（用于动态生成或在 Opening 时添加子项的情况）
            try
            {
                if (ctrl.ContextMenuStrip != null)
                    ReplaceFontsInContextMenu(ctrl.ContextMenuStrip);

                // 监听 ContextMenuStripChanged（控件运行时可能赋值新的 ContextMenuStrip）
                ctrl.ContextMenuStripChanged -= Ctrl_ContextMenuStripChanged;
                ctrl.ContextMenuStripChanged += Ctrl_ContextMenuStripChanged;
            }
            catch { }

            // ToolStrip 递归替换并监听 ItemAdded（动态添加项）
            try
            {
                if (ctrl is ToolStrip ts)
                {
                    ReplaceFontsInToolStrip(ts);

                    if (!hookedToolStrips.Contains(ts))
                    {
                        ts.ItemAdded += Ts_ItemAdded;
                        hookedToolStrips.Add(ts);
                    }
                }
            }
            catch { }

            // 对子控件递归
            foreach (Control child in ctrl.Controls)
            {
                try { ReplaceFontsInControlRecursive(child); } catch { }
            }
        }

        // 处理控件的 ContextMenuStrip 被修改
        private void Ctrl_ContextMenuStripChanged(object sender, EventArgs e)
        {
            if (sender is Control c && c.ContextMenuStrip != null)
            {
                try { ReplaceFontsInContextMenu(c.ContextMenuStrip); } catch { }
            }
        }

        // ToolStrip 的 ItemAdded 事件处理（针对动态添加项）
        private void Ts_ItemAdded(object sender, ToolStripItemEventArgs e)
        {
            try { ReplaceToolStripItemFont(e.Item); } catch { }
        }

        // ===== 递归替换 ContextMenuStrip & ToolStripItem 方法（永久替换，带事件监听） =====
        private void ReplaceFontsInContextMenu(ContextMenuStrip cms)
        {
            if (cms == null || !enabledFontReplace) return;

            foreach (ToolStripItem item in cms.Items)
            {
                ReplaceToolStripItemFont(item);
            }

            // 监听 Opening，防止在 Opening 时动态创建/修改菜单项导致漏刷
            if (!hookedContextMenus.Contains(cms))
            {
                cms.Opening += Cms_Opening;
                hookedContextMenus.Add(cms);
            }
        }

        private void Cms_Opening(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (sender is ContextMenuStrip cms)
            {
                try { ReplaceFontsInContextMenu(cms); } catch { }
            }
        }

        private void ReplaceToolStripItemFont(ToolStripItem item)
        {
            if (item == null || !enabledFontReplace) return;

            try
            {
                if (item.Font != null &&
                    (targetSourceFonts.Count == 0 || targetSourceFonts.Contains(item.Font.FontFamily.Name)))
                {
                    // 保持原 item 大小（如果 autoSize），否则使用 targetFont.Size
                    float newSize = autoSize ? item.Font.Size : targetFont.Size;
                    item.Font = new Font(targetFont.FontFamily, newSize, targetFont.Style, item.Font.Unit);
                }
            }
            catch { }

            // 如果是带下拉的项，递归处理子项并监听 DropDownOpening（动态生成子项情形）
            if (item is ToolStripDropDownItem drop)
            {
                try
                {
                    foreach (ToolStripItem subItem in drop.DropDownItems)
                        ReplaceToolStripItemFont(subItem);

                    if (!hookedDropDownItems.Contains(drop))
                    {
                        drop.DropDownOpening += Drop_DropDownOpening;
                        hookedDropDownItems.Add(drop);
                    }
                }
                catch { }
            }
        }

        private void Drop_DropDownOpening(object sender, EventArgs e)
        {
            if (sender is ToolStripDropDownItem drop)
            {
                try
                {
                    foreach (ToolStripItem subItem in drop.DropDownItems)
                        ReplaceToolStripItemFont(subItem);
                }
                catch { }
            }
        }

        private void ReplaceFontsInToolStrip(ToolStrip ts)
        {
            if (ts == null || !enabledFontReplace) return;

            foreach (ToolStripItem item in ts.Items)
            {
                ReplaceToolStripItemFont(item);
            }

            // ItemAdded 事件在上面统一注册（Ts_ItemAdded）
        }

        // ====== 对 Assembly 静态 Font 字段/属性 的替换（永久） ======
        private void ReplaceFontsInAssembly(Assembly asm)
        {
            if (asm == null || !enabledFontReplace) return;
            Type[] types = null;
            try { types = asm.GetTypes(); }
            catch (ReflectionTypeLoadException ex) { types = ex.Types; }
            catch { return; }

            if (types == null) return;

            foreach (var t in types)
            {
                if (t == null) continue;

                var fields = t.GetFields(BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
                foreach (var f in fields)
                {
                    try
                    {
                        if (f.FieldType == typeof(Font))
                        {
                            var val = f.GetValue(null) as Font;
                            if (val != null &&
                                (targetSourceFonts.Count == 0 || targetSourceFonts.Contains(val.FontFamily.Name)))
                                f.SetValue(null, new Font(targetFont.FontFamily, autoSize != true ? targetFont.Size : val.Size, targetFont.Style, val.Unit));
                        }
                    }
                    catch { }
                }

                var props = t.GetProperties(BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
                foreach (var p in props)
                {
                    try
                    {
                        if (p.PropertyType == typeof(Font) && p.CanRead && p.CanWrite && p.GetIndexParameters().Length == 0)
                        {
                            var val = p.GetValue(null, null) as Font;
                            if (val != null &&
                                (targetSourceFonts.Count == 0 || targetSourceFonts.Contains(val.FontFamily.Name)))
                                p.SetValue(null, new Font(targetFont.FontFamily, autoSize != true ? targetFont.Size : val.Size, targetFont.Style, val.Unit), null);
                        }
                    }
                    catch { }
                }
            }
        }

        // ====== 设置界面（Unchanged） ======
        private void ShowFontSettings()
        {
            using (Form settingsForm = new Form())
            {
                settingsForm.Text = "字体设置";
                settingsForm.Size = new Size(420, 420);
                settingsForm.StartPosition = FormStartPosition.CenterParent;
                settingsForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                settingsForm.MaximizeBox = false;
                settingsForm.MinimizeBox = false;

                Label lblFont = new Label { Text = $"目标字体: {targetFont.Name}, {targetFont.Size}pt, {targetFont.Style}", Location = new Point(20, 20), Size = new Size(240, 30) };
                CheckBox cblFontSize = new CheckBox { Text = "自动大小", Location = new Point(260, 20), Checked = autoSize, AutoSize = true };
                Button btnChangeFont = new Button { Text = "选择目标字体", Location = new Point(20, 50), Size = new Size(120, 30) };

                ListBox lbSourceFonts = new ListBox { Location = new Point(20, 90), Size = new Size(200, 200), SelectionMode = SelectionMode.MultiExtended };
                lbSourceFonts.Items.AddRange(targetSourceFonts.ToArray());

                Button btnAdd = new Button { Text = "添加", Location = new Point(230, 90), Size = new Size(120, 30) };
                Button btnRemove = new Button { Text = "删除", Location = new Point(230, 130), Size = new Size(120, 30) };
                Button btnClear = new Button { Text = "全清", Location = new Point(230, 170), Size = new Size(120, 30) };
                Button btnInitialize = new Button { Text = "初始化配置", Location = new Point(230, 210), Size = new Size(120, 30) };

                CheckBox cbEnable = new CheckBox { Text = "启用字体替换", Location = new Point(240, 255), Checked = tempEnabledFontReplace, AutoSize = true };

                Label lblNote = new Label { Text = "列表为空时则替换全部字体", Location = new Point(20, 300), Size = new Size(360, 30), ForeColor = Color.Blue };

                Button btnOK = new Button { Text = "确定并刷新", Location = new Point(20, 330), Size = new Size(150, 30) };
                Button btnCancel = new Button { Text = "取消", Location = new Point(200, 330), Size = new Size(150, 30) };

                // ==== 控件事件 ====
                cbEnable.CheckedChanged += (s, e) =>
                {
                    tempEnabledFontReplace = cbEnable.Checked;
                    // 不立即刷新，不弹提示
                };

                lbSourceFonts.KeyDown += (s, e) =>
                {
                    if (e.KeyCode == Keys.Delete || e.KeyCode == Keys.Back)
                    {
                        var toRemove = lbSourceFonts.SelectedItems.Cast<string>().ToList();
                        foreach (var rm in toRemove) targetSourceFonts.Remove(rm);
                        lbSourceFonts.Items.Clear();
                        lbSourceFonts.Items.AddRange(targetSourceFonts.ToArray());
                    }
                };

                btnChangeFont.Click += (s, e) =>
                {
                    using (FontDialog dlg = new FontDialog())
                    {
                        dlg.Font = targetFont;
                        dlg.ShowEffects = true;
                        if (dlg.ShowDialog() == DialogResult.OK)
                        {
                            targetFont = dlg.Font;
                            lblFont.Text = $"目标字体: {targetFont.Name}, {targetFont.Size}pt, {targetFont.Style}";
                        }
                    }
                };

                btnAdd.Click += (s, e) =>
                {
                    using (Form selectForm = new Form())
                    {
                        selectForm.Text = "添加源字体";
                        selectForm.Size = new Size(300, 400);
                        selectForm.StartPosition = FormStartPosition.CenterParent;
                        ListBox lbAllFonts = new ListBox { Location = new Point(10, 10), Size = new Size(260, 300), SelectionMode = SelectionMode.MultiExtended };
                        Button btnAddFont = new Button { Text = "添加", Location = new Point(10, 320), Size = new Size(100, 30) };
                        Button btnReturn = new Button { Text = "返回", Location = new Point(170, 320), Size = new Size(100, 30) };

                        foreach (FontFamily ff in FontFamily.Families)
                            lbAllFonts.Items.Add(ff.Name);

                        btnAddFont.Click += (s2, e2) =>
                        {
                            foreach (var item in lbAllFonts.SelectedItems)
                            {
                                if (!targetSourceFonts.Contains(item.ToString()))
                                    targetSourceFonts.Add(item.ToString());
                            }
                            lbSourceFonts.Items.Clear();
                            lbSourceFonts.Items.AddRange(targetSourceFonts.ToArray());
                        };
                        lbAllFonts.DoubleClick += (s2, e2) =>
                        {
                            foreach (var item in lbAllFonts.SelectedItems)
                            {
                                if (!targetSourceFonts.Contains(item.ToString()))
                                    targetSourceFonts.Add(item.ToString());
                            }
                            lbSourceFonts.Items.Clear();
                            lbSourceFonts.Items.AddRange(targetSourceFonts.ToArray());
                        };
                        btnReturn.Click += (s2, e2) => selectForm.Close();
                        selectForm.Controls.Add(lbAllFonts);
                        selectForm.Controls.Add(btnAddFont);
                        selectForm.Controls.Add(btnReturn);
                        selectForm.ShowDialog();
                    }
                };

                btnRemove.Click += (s, e) =>
                {
                    var toRemove = lbSourceFonts.SelectedItems.Cast<string>().ToList();
                    foreach (var rm in toRemove) targetSourceFonts.Remove(rm);
                    lbSourceFonts.Items.Clear();
                    lbSourceFonts.Items.AddRange(targetSourceFonts.ToArray());
                };

                btnClear.Click += (s, e) =>
                {
                    targetSourceFonts.Clear();
                    lbSourceFonts.Items.Clear();
                };

                btnInitialize.Click += (s, e) =>
                {
                    try { File.Delete(configFile); } catch { }
                    MessageBox.Show("配置已删除，请重启软件生效。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    settingsForm.Close();
                };

                // ==== 确定并刷新逻辑 ====
                btnOK.Click += (s, e) =>
                {
                    bool needRestartPrompt = enabledFontReplace != tempEnabledFontReplace;
                    enabledFontReplace = tempEnabledFontReplace;
                    autoSize = cblFontSize.Checked;

                    // 从配置文件读取旧的目标字体
                    string oldFontName = DefaultFontName;
                    if (File.Exists(configFile))
                    {
                        var lines = File.ReadAllLines(configFile);
                        if (lines.Length > 0 && lines[0].StartsWith("TargetFont:"))
                        {
                            var parts = lines[0].Substring("TargetFont:".Length).Split('|');
                            if (parts.Length >= 3)
                                oldFontName = parts[0]; // 取旧字体
                        }
                    }

                    SaveConfig();

                    if (enabledFontReplace)
                    {
                        var tempList = new List<string>(targetSourceFonts);
                        if (tempList.Count > 0)
                            tempList.Add(oldFontName);

                        foreach (Form f in Application.OpenForms)
                        {
                            f.Tag = null;
                            ReplaceFontsInFormWithTempSourceFonts(f, tempList);
                        }

                        foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
                            ReplaceFontsInAssemblyWithTempSourceFonts(asm, tempList);
                    }

                    if (needRestartPrompt)
                        MessageBox.Show("启用状态修改已保存，请重启软件生效。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    settingsForm.Close();
                };

                btnCancel.Click += (s, e) => settingsForm.Close();

                settingsForm.Controls.Add(lblFont);
                settingsForm.Controls.Add(cblFontSize);
                settingsForm.Controls.Add(btnChangeFont);
                settingsForm.Controls.Add(lbSourceFonts);
                settingsForm.Controls.Add(btnAdd);
                settingsForm.Controls.Add(btnRemove);
                settingsForm.Controls.Add(btnClear);
                settingsForm.Controls.Add(btnInitialize);
                settingsForm.Controls.Add(cbEnable);
                settingsForm.Controls.Add(lblNote);
                settingsForm.Controls.Add(btnOK);
                settingsForm.Controls.Add(btnCancel);

                settingsForm.ShowDialog();
            }
        }

        // ==== 临时刷新方法（WithTempList：一次性替换，不挂事件） ====
        private void ReplaceFontsInFormWithTempSourceFonts(Form form, List<string> tempList)
        {
            if (form == null) return;
            form.Tag = null;
            ReplaceFontsInControlRecursiveWithTempList(form, tempList);
        }

        private void ReplaceFontsInControlRecursiveWithTempList(Control ctrl, List<string> tempList)
        {
            if (ctrl == null || !enabledFontReplace) return;

            try
            {
                if (ctrl.Font != null &&
                    (tempList.Count == 0 || tempList.Contains(ctrl.Font.FontFamily.Name)))
                {
                    ctrl.Font = new Font(targetFont.FontFamily, autoSize != true ? targetFont.Size : ctrl.Font.Size, targetFont.Style, ctrl.Font.Unit);
                }
            }
            catch { }

            // ToolStrip 替换（递归，但不挂事件）
            if (ctrl is ToolStrip ts)
            {
                ReplaceFontsInToolStripWithTempList(ts, tempList);
            }

            // ContextMenuStrip 替换（递归，但不挂事件）
            if (ctrl.ContextMenuStrip != null)
            {
                ReplaceFontsInContextMenuWithTempList(ctrl.ContextMenuStrip, tempList);
            }

            foreach (Control child in ctrl.Controls)
            {
                try { ReplaceFontsInControlRecursiveWithTempList(child, tempList); } catch { }
            }
        }

        private void ReplaceFontsInToolStripWithTempList(ToolStrip ts, List<string> tempList)
        {
            if (ts == null || !enabledFontReplace) return;
            foreach (ToolStripItem item in ts.Items)
            {
                ReplaceToolStripItemFontWithTempList(item, tempList);
            }
        }

        private void ReplaceFontsInContextMenuWithTempList(ContextMenuStrip cms, List<string> tempList)
        {
            if (cms == null || !enabledFontReplace) return;
            foreach (ToolStripItem item in cms.Items)
            {
                ReplaceToolStripItemFontWithTempList(item, tempList);
            }
        }

        private void ReplaceToolStripItemFontWithTempList(ToolStripItem item, List<string> tempList)
        {
            if (item == null || !enabledFontReplace) return;

            try
            {
                if (item.Font != null &&
                    (tempList.Count == 0 || tempList.Contains(item.Font.FontFamily.Name)))
                {
                    float newSize = autoSize ? item.Font.Size : targetFont.Size;
                    item.Font = new Font(targetFont.FontFamily, newSize, targetFont.Style, item.Font.Unit);
                }
            }
            catch { }

            if (item is ToolStripDropDownItem drop)
            {
                foreach (ToolStripItem subItem in drop.DropDownItems)
                    ReplaceToolStripItemFontWithTempList(subItem, tempList);
            }
        }

        private void ReplaceFontsInAssemblyWithTempSourceFonts(Assembly asm, List<string> tempList)
        {
            if (asm == null || !enabledFontReplace) return;
            Type[] types = null;
            try { types = asm.GetTypes(); }
            catch (ReflectionTypeLoadException ex) { types = ex.Types; }
            catch { return; }

            if (types == null) return;

            foreach (var t in types)
            {
                if (t == null) continue;

                var fields = t.GetFields(BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
                foreach (var f in fields)
                {
                    try
                    {
                        if (f.FieldType == typeof(Font))
                        {
                            var val = f.GetValue(null) as Font;
                            if (val != null &&
                                (tempList.Count == 0 || tempList.Contains(val.FontFamily.Name)))
                                f.SetValue(null, new Font(targetFont.FontFamily, autoSize != true ? targetFont.Size : val.Size, targetFont.Style, val.Unit));
                        }
                    }
                    catch { }
                }

                var props = t.GetProperties(BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
                foreach (var p in props)
                {
                    try
                    {
                        if (p.PropertyType == typeof(Font) && p.CanRead && p.CanWrite && p.GetIndexParameters().Length == 0)
                        {
                            var val = p.GetValue(null, null) as Font;
                            if (val != null &&
                                (tempList.Count == 0 || tempList.Contains(val.FontFamily.Name)))
                                p.SetValue(null, new Font(targetFont.FontFamily, autoSize != true ? targetFont.Size : val.Size, targetFont.Style, val.Unit), null);
                        }
                    }
                    catch { }
                }
            }
        }
    }
}
