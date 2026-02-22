// =======================
// MyPluginControl.cs
// =======================
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Metadata.Query;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using XrmToolBox.Extensibility;

// Fix for the 'Label' ambiguity error
using WinFormLabel = System.Windows.Forms.Label;

namespace SecurityRoleUpdaterTool
{
    public partial class MyPluginControl : PluginControlBase
    {
        // ===== UI (3 columns) =====
        private CheckedListBox listEntities;
        private TextBox txtEntityFilter;

        private CheckedListBox listRoles;
        private TextBox txtRoleFilter;

        private ComboBox cmbBusinessUnits;
        private CheckBox chkAllBusinessUnits;

        private DataGridView gridPrivileges;

        private Button btnLoadData;
        private Button btnCancelLoad;
        private Button btnLoadPrivileges;
        private Button btnUpdate;

        private WinFormLabel lblStatus;

        // ===== Caches =====
        private List<EntityItem> _entitiesCache = new List<EntityItem>();
        private List<RoleItem> _rolesCache = new List<RoleItem>();
        private List<BusinessUnitItem> _busCache = new List<BusinessUnitItem>();

        // entityLogicalName -> (PrivilegeTypeName -> PrivilegeId)
        private readonly Dictionary<string, Dictionary<string, Guid>> _entityPrivCache =
            new Dictionary<string, Dictionary<string, Guid>>(StringComparer.OrdinalIgnoreCase);

        private CancellationTokenSource _loadCts;

        // ===== CRM palette (closer to D365 UI) =====
        private readonly Color _crmYellow = Color.FromArgb(255, 185, 0);   // #FFB900
        private readonly Color _crmGreen = Color.FromArgb(16, 124, 16);    // #107C10

        // ===== Painting =====
        private readonly Pen _gridCirclePen = new Pen(Color.FromArgb(110, 110, 110), 1);
        private readonly Pen _redPen = new Pen(Color.Red, 2);
        private readonly Pen _ghostPen = new Pen(Color.FromArgb(150, 150, 150), 1)
        { DashStyle = System.Drawing.Drawing2D.DashStyle.Dash };

        private readonly SolidBrush _yellowBrush;
        private readonly SolidBrush _greenBrush;

        private readonly string[] _privCols = { "Create", "Read", "Write", "Delete", "Append", "AppendTo", "Assign", "Share" };

        // ===== Watermarks (NET Framework) =====
        private const string EntityWatermark = "Filter entities (type to search)...";
        private const string RoleWatermark = "Filter roles (type to search)...";

        // ===== Depth states =====
        // "NoChange" keep as ghost, "SetNone" keep as red X, and the CRM-style pieces below.
        private const string S_NoChange = "NoChange";
        private const string S_SetNone = "SetNone";
        private const string S_User = "User";
        private const string S_BU = "BU";
        private const string S_Deep = "Deep"; // Parent:Child BU
        private const string S_Org = "Org";

        public MyPluginControl()
        {
            // Dynamic UI; no InitializeComponent
            _yellowBrush = new SolidBrush(_crmYellow);
            _greenBrush = new SolidBrush(_crmGreen);

            SetupCustomUI();
        }

        // =======================
        // UI (3 columns + legend)
        // =======================
        private void SetupCustomUI()
        {
            Controls.Clear();

            var mainContainer = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                SplitterDistance = 560
            };

            var table = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 1
            };
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 360)); // Entities
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 420)); // Roles
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));  // Depth grid
            table.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            var panelEntities = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                Padding = new Padding(12),
                AutoScroll = true,
                WrapContents = false
            };

            var panelRoles = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                Padding = new Padding(12),
                AutoScroll = true,
                WrapContents = false
            };

            var panelDepth = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12)
            };

            lblStatus = new WinFormLabel
            {
                Text = "Idle",
                AutoSize = true,
                ForeColor = Color.DimGray,
                Margin = new Padding(0, 0, 0, 8)
            };

            btnLoadData = new Button
            {
                Text = "Refresh Entities / Business Units / Roles",
                Width = 330,
                Height = 42,
                BackColor = Color.WhiteSmoke,
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };
            btnLoadData.Click += (s, e) => ExecuteMethod(LoadInitialDataFast);

            btnCancelLoad = new Button
            {
                Text = "Cancel Loading",
                Width = 330,
                Height = 32,
                BackColor = Color.Gainsboro,
                Enabled = false
            };
            btnCancelLoad.Click += (s, e) => CancelLoading();

            // Entities
            txtEntityFilter = new TextBox { Width = 330 };
            AddWatermark(txtEntityFilter, EntityWatermark);
            txtEntityFilter.TextChanged += (s, e) => ApplyEntityFilter();

            listEntities = new CheckedListBox
            {
                Width = 330,
                Height = 430,
                CheckOnClick = true
            };

            panelEntities.Controls.Add(new WinFormLabel
            {
                Text = "ENTITIES",
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                AutoSize = true
            });
            panelEntities.Controls.Add(lblStatus);
            panelEntities.Controls.Add(btnLoadData);
            panelEntities.Controls.Add(btnCancelLoad);
            panelEntities.Controls.Add(new WinFormLabel { Text = "Filter:", AutoSize = true, Margin = new Padding(0, 10, 0, 2) });
            panelEntities.Controls.Add(txtEntityFilter);
            panelEntities.Controls.Add(new WinFormLabel { Text = "Select Entities:", AutoSize = true, Margin = new Padding(0, 10, 0, 2) });
            panelEntities.Controls.Add(listEntities);

            // Roles + BU filter
            chkAllBusinessUnits = new CheckBox
            {
                Text = "All Business Units",
                AutoSize = true,
                Checked = true
            };

            cmbBusinessUnits = new ComboBox
            {
                Width = 390,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Enabled = false
            };

            chkAllBusinessUnits.CheckedChanged += (s, e) =>
            {
                cmbBusinessUnits.Enabled = !chkAllBusinessUnits.Checked;
                ExecuteMethod(LoadRolesForSelectedBU);
            };

            cmbBusinessUnits.SelectedIndexChanged += (s, e) =>
            {
                if (!chkAllBusinessUnits.Checked)
                    ExecuteMethod(LoadRolesForSelectedBU);
            };

            txtRoleFilter = new TextBox { Width = 390 };
            AddWatermark(txtRoleFilter, RoleWatermark);
            txtRoleFilter.TextChanged += (s, e) => ApplyRoleFilter();

            listRoles = new CheckedListBox
            {
                Width = 390,
                Height = 330,
                CheckOnClick = true
            };

            btnLoadPrivileges = new Button
            {
                Text = "Load Privileges for Selected Entities",
                Width = 390,
                Height = 40,
                BackColor = Color.WhiteSmoke,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Margin = new Padding(0, 8, 0, 0)
            };
            btnLoadPrivileges.Click += (s, e) => ExecuteMethod(LoadPrivilegesForSelectedEntities);

            btnUpdate = new Button
            {
                Text = "APPLY CHANGES (Entities × Roles)",
                Width = 390,
                Height = 58,
                BackColor = Color.FromArgb(192, 0, 0),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                Margin = new Padding(0, 12, 0, 0)
            };
            btnUpdate.Click += (s, e) => ExecuteMethod(UpdateRolesLogic_MultiEntity);

            panelRoles.Controls.Add(new WinFormLabel
            {
                Text = "SECURITY ROLES",
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                AutoSize = true
            });
            panelRoles.Controls.Add(new WinFormLabel { Text = "Business Unit:", AutoSize = true, Margin = new Padding(0, 8, 0, 2) });
            panelRoles.Controls.Add(chkAllBusinessUnits);
            panelRoles.Controls.Add(cmbBusinessUnits);
            panelRoles.Controls.Add(new WinFormLabel { Text = "Filter:", AutoSize = true, Margin = new Padding(0, 10, 0, 2) });
            panelRoles.Controls.Add(txtRoleFilter);
            panelRoles.Controls.Add(new WinFormLabel { Text = "Select Roles:", AutoSize = true, Margin = new Padding(0, 10, 0, 2) });
            panelRoles.Controls.Add(listRoles);
            panelRoles.Controls.Add(btnLoadPrivileges);
            panelRoles.Controls.Add(btnUpdate);

            // Depth grid (CRM-style drawing)
            gridPrivileges = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                RowHeadersVisible = false,
                BackgroundColor = Color.White,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                RowTemplate = { Height = 70 },
                EditMode = DataGridViewEditMode.EditProgrammatically,
                MultiSelect = false,
                SelectionMode = DataGridViewSelectionMode.CellSelect
            };

            EnableDoubleBuffering(gridPrivileges);
            gridPrivileges.CellClick += GridPrivileges_CellClick;
            gridPrivileges.CellPainting += GridPrivileges_CellPainting;

            BuildGridSkeleton();

            var depthHeader = new WinFormLabel
            {
                Text = "DEPTH CONFIGURATION",
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                AutoSize = true,
                Padding = new Padding(0, 0, 0, 10),
                Dock = DockStyle.Top
            };

            panelDepth.Controls.Add(gridPrivileges);
            panelDepth.Controls.Add(depthHeader);
            gridPrivileges.BringToFront();

            table.Controls.Add(panelEntities, 0, 0);
            table.Controls.Add(panelRoles, 1, 0);
            table.Controls.Add(panelDepth, 2, 0);

            mainContainer.Panel1.Controls.Add(table);
            mainContainer.Panel2.Controls.Add(BuildLegendPanel_CrmStyle());

            Controls.Add(mainContainer);
        }

        // CRM-like legend: same icons/palette as grid
        private Control BuildLegendPanel_CrmStyle()
        {
            var panel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10),
                BackColor = Color.White,
                WrapContents = true,
                AutoScroll = true
            };

            panel.Controls.Add(new WinFormLabel
            {
                Text = "LEGEND:",
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                AutoSize = true,
                Margin = new Padding(0, 6, 10, 0)
            });

            panel.Controls.Add(NewLegendIcon(S_User));
            panel.Controls.Add(NewLegendText("User"));

            panel.Controls.Add(NewLegendIcon(S_BU));
            panel.Controls.Add(NewLegendText("Business Unit"));

            panel.Controls.Add(NewLegendIcon(S_Deep));
            panel.Controls.Add(NewLegendText("Parent: Child Business Units"));

            panel.Controls.Add(NewLegendIcon(S_Org));
            panel.Controls.Add(NewLegendText("Organization"));

            // Keep NONE unchanged (your requirement)
            panel.Controls.Add(NewLegendIcon(S_SetNone));
            panel.Controls.Add(NewLegendText("SET TO NONE (Remove)"));

            // Keep ghost
            panel.Controls.Add(NewLegendIcon(S_NoChange));
            panel.Controls.Add(NewLegendText("NO CHANGE (Skip)"));

            return panel;
        }

        private Control NewLegendText(string text)
        {
            return new WinFormLabel
            {
                Text = text,
                AutoSize = true,
                Margin = new Padding(6, 8, 18, 0)
            };
        }

        private Control NewLegendIcon(string state)
        {
            var c = new DepthIconControl
            {
                Width = 18,
                Height = 18,
                Margin = new Padding(12, 6, 0, 0),
                State = state,
                Yellow = _crmYellow,
                Green = _crmGreen
            };
            return c;
        }

        // =========================
        // Watermark (NET Framework)
        // =========================
        private void AddWatermark(TextBox tb, string watermark)
        {
            tb.ForeColor = Color.DimGray;
            tb.Text = watermark;

            tb.GotFocus += (s, e) =>
            {
                if (tb.Text == watermark)
                {
                    tb.Text = "";
                    tb.ForeColor = SystemColors.WindowText;
                }
            };

            tb.LostFocus += (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(tb.Text))
                {
                    tb.Text = watermark;
                    tb.ForeColor = Color.DimGray;
                }
            };
        }

        private string GetFilterText(TextBox tb, string watermark)
        {
            var t = (tb.Text ?? "").Trim();
            return string.Equals(t, watermark, StringComparison.Ordinal) ? "" : t;
        }

        private void EnableDoubleBuffering(DataGridView dgv)
        {
            typeof(DataGridView).InvokeMember(
                "DoubleBuffered",
                BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.SetProperty,
                null, dgv, new object[] { true }
            );
        }

        private void SetStatus(string text)
        {
            if (lblStatus.InvokeRequired) { lblStatus.BeginInvoke(new Action(() => lblStatus.Text = text)); return; }
            lblStatus.Text = text;
        }

        private void SetLoadingUi(bool loading)
        {
            if (InvokeRequired) { BeginInvoke(new Action(() => SetLoadingUi(loading))); return; }

            btnLoadData.Enabled = !loading;
            btnCancelLoad.Enabled = loading;

            listEntities.Enabled = !loading;
            listRoles.Enabled = !loading;

            txtEntityFilter.Enabled = !loading;
            txtRoleFilter.Enabled = !loading;

            chkAllBusinessUnits.Enabled = !loading;
            cmbBusinessUnits.Enabled = !loading && !chkAllBusinessUnits.Checked;

            btnLoadPrivileges.Enabled = !loading;
            btnUpdate.Enabled = !loading;
        }

        private void CancelLoading()
        {
            try { _loadCts?.Cancel(); } catch { }
        }

        // =========================
        // Grid build / paint / click
        // =========================
        private void BuildGridSkeleton()
        {
            gridPrivileges.SuspendLayout();
            gridPrivileges.Columns.Clear();
            gridPrivileges.Rows.Clear();

            foreach (var c in _privCols)
                gridPrivileges.Columns.Add(c, c);

            gridPrivileges.Rows.Add();
            ResetGridToGhost();

            gridPrivileges.ResumeLayout();
        }

        private void ResetGridToGhost()
        {
            if (gridPrivileges.Rows.Count == 0) return;
            foreach (DataGridViewColumn col in gridPrivileges.Columns)
                gridPrivileges.Rows[0].Cells[col.Index].Value = S_NoChange;
            gridPrivileges.Invalidate();
        }

        // CRM-style geometry requirements:
        // - User: quarter, filled from bottom area (use bottom-right quarter like CRM)
        // - Business Unit: half circle horizontal (bottom half filled)
        // - Parent-Child BU: 3/4 filled, empty part UP (top-right empty)
        // - Org: full circle green
        // - None unchanged (red X)
        private void GridPrivileges_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;

            e.PaintBackground(e.CellBounds, true);

            string val = e.Value?.ToString() ?? S_NoChange;

            var rect = new Rectangle(
                e.CellBounds.X + (e.CellBounds.Width / 2) - 15,
                e.CellBounds.Y + (e.CellBounds.Height / 2) - 15,
                30, 30);

            e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

            if (val == S_NoChange)
            {
                e.Graphics.DrawEllipse(_ghostPen, rect);
                e.Handled = true;
                return;
            }

            if (val == S_SetNone)
            {
                e.Graphics.DrawEllipse(_redPen, rect);
                e.Graphics.DrawLine(_redPen, rect.Left + 8, rect.Top + 8, rect.Right - 8, rect.Bottom - 8);
                e.Graphics.DrawLine(_redPen, rect.Left + 8, rect.Bottom - 8, rect.Right - 8, rect.Top + 8);
                e.Handled = true;
                return;
            }

            if (val == S_Org)
            {
                e.Graphics.FillEllipse(_greenBrush, rect);
                e.Graphics.DrawEllipse(_gridCirclePen, rect);
                e.Handled = true;
                return;
            }

            if (val == S_User)
            {
                // USER: quarter filled from the bottom (matches CRM)
                // bottom is centered at 90°, so fill 45°..135°
                e.Graphics.FillPie(_yellowBrush, rect, 45, 90);
                e.Graphics.DrawEllipse(_gridCirclePen, rect);
                e.Handled = true;
                return;
            }

            if (val == S_BU)
            {
                // BU: bottom half filled (horizontal)
                e.Graphics.FillPie(_yellowBrush, rect, 0, 180);
                e.Graphics.DrawEllipse(_gridCirclePen, rect);
                e.Handled = true;
                return;
            }

            if (val == S_Deep)
            {
                // Parent:Child BU: top quarter empty (empty part UP), GREEN
                e.Graphics.FillPie(_greenBrush, rect, 315, 270);
                e.Graphics.DrawEllipse(_gridCirclePen, rect);
                e.Handled = true;
                return;
            }

            e.Graphics.DrawEllipse(_gridCirclePen, rect);
            e.Handled = true;
        }

        private void GridPrivileges_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;

            var cell = gridPrivileges.Rows[e.RowIndex].Cells[e.ColumnIndex];
            string current = cell.Value?.ToString() ?? S_NoChange;

            // Cycle: NoChange -> User -> BU -> Deep -> Org -> SetNone -> NoChange
            if (current == S_NoChange) cell.Value = S_User;
            else if (current == S_User) cell.Value = S_BU;
            else if (current == S_BU) cell.Value = S_Deep;
            else if (current == S_Deep) cell.Value = S_Org;
            else if (current == S_Org) cell.Value = S_SetNone;
            else cell.Value = S_NoChange;

            gridPrivileges.InvalidateCell(cell);
        }

        // =========================
        // FAST LOAD: Entities + BUs + Roles
        // =========================
        private void LoadInitialDataFast()
        {
            _loadCts?.Cancel();
            _loadCts = new CancellationTokenSource();
            var ct = _loadCts.Token;

            SetLoadingUi(true);
            SetStatus("Loading entities, business units, roles...");

            WorkAsync(new WorkAsyncInfo
            {
                Message = "Loading Entities / Business Units / Roles...",
                Work = (worker, args) =>
                {
                    ct.ThrowIfCancellationRequested();

                    var entities = RetrieveEntitiesFast(ct);
                    ct.ThrowIfCancellationRequested();

                    var bus = RetrieveBusinessUnits(ct);
                    ct.ThrowIfCancellationRequested();

                    var roles = RetrieveRolesPaged_ByBU(ct, null);

                    args.Result = new LoadResult
                    {
                        Entities = entities,
                        BusinessUnits = bus,
                        Roles = roles
                    };
                },
                PostWorkCallBack = (args) =>
                {
                    SetLoadingUi(false);

                    if (ct.IsCancellationRequested) { SetStatus("Cancelled."); return; }
                    if (args.Error != null) { SetStatus("Load failed."); MessageBox.Show(args.Error.Message, "Error"); return; }

                    var res = (LoadResult)args.Result;

                    _entitiesCache = res.Entities ?? new List<EntityItem>();
                    _busCache = res.BusinessUnits ?? new List<BusinessUnitItem>();
                    _rolesCache = res.Roles ?? new List<RoleItem>();

                    _entityPrivCache.Clear();
                    ResetGridToGhost();

                    BindEntities(_entitiesCache);
                    BindBusinessUnits(_busCache);

                    chkAllBusinessUnits.Checked = true;
                    cmbBusinessUnits.Enabled = false;

                    BindRoles(_rolesCache);

                    SetStatus($"Loaded: {_entitiesCache.Count} entities, {_busCache.Count} BUs, {_rolesCache.Count} roles.");
                }
            });
        }

        private List<EntityItem> RetrieveEntitiesFast(CancellationToken ct)
        {
            var props = new MetadataPropertiesExpression
            {
                AllProperties = false,
                PropertyNames = { "LogicalName", "DisplayName", "IsCustomizable", "IsIntersect", "IsActivity" }
            };

            var filter = new MetadataFilterExpression(LogicalOperator.And);
            filter.Conditions.Add(new MetadataConditionExpression("IsIntersect", MetadataConditionOperator.Equals, false));
            filter.Conditions.Add(new MetadataConditionExpression("IsActivity", MetadataConditionOperator.Equals, false));
            filter.Conditions.Add(new MetadataConditionExpression("IsCustomizable", MetadataConditionOperator.Equals, true));

            var query = new EntityQueryExpression
            {
                Properties = props,
                Criteria = filter
            };

            var req = new RetrieveMetadataChangesRequest { Query = query };
            var resp = (RetrieveMetadataChangesResponse)Service.Execute(req);

            var list = new List<EntityItem>(resp.EntityMetadata.Count);

            foreach (var em in resp.EntityMetadata)
            {
                ct.ThrowIfCancellationRequested();

                var dn = em.DisplayName?.UserLocalizedLabel?.Label;
                list.Add(new EntityItem
                {
                    LogicalName = em.LogicalName,
                    DisplayName = string.IsNullOrWhiteSpace(dn) ? em.LogicalName : dn
                });
            }

            return list.OrderBy(x => x.DisplayName, StringComparer.CurrentCultureIgnoreCase).ToList();
        }

        private List<BusinessUnitItem> RetrieveBusinessUnits(CancellationToken ct)
        {
            var items = new List<BusinessUnitItem>(64);

            int page = 1;
            string cookie = null;

            while (true)
            {
                ct.ThrowIfCancellationRequested();

                var qe = new QueryExpression("businessunit")
                {
                    ColumnSet = new ColumnSet("businessunitid", "name"),
                    PageInfo = new PagingInfo { PageNumber = page, Count = 5000, PagingCookie = cookie }
                };

                var res = Service.RetrieveMultiple(qe);

                foreach (var e in res.Entities)
                {
                    items.Add(new BusinessUnitItem
                    {
                        Id = e.Id,
                        Name = e.GetAttributeValue<string>("name") ?? e.Id.ToString()
                    });
                }

                if (!res.MoreRecords) break;
                cookie = res.PagingCookie;
                page++;
            }

            return items.OrderBy(x => x.Name, StringComparer.CurrentCultureIgnoreCase).ToList();
        }

        private List<RoleItem> RetrieveRolesPaged_ByBU(CancellationToken ct, Guid? businessUnitId)
        {
            var roles = new List<RoleItem>(512);
            int page = 1;
            string cookie = null;

            while (true)
            {
                ct.ThrowIfCancellationRequested();

                var qe = new QueryExpression("role")
                {
                    ColumnSet = new ColumnSet("roleid", "name", "businessunitid"),
                    Criteria = new FilterExpression(LogicalOperator.And),
                    PageInfo = new PagingInfo { PageNumber = page, Count = 5000, PagingCookie = cookie }
                };

                qe.Criteria.AddCondition("parentroleid", ConditionOperator.Null);

                if (businessUnitId.HasValue)
                    qe.Criteria.AddCondition("businessunitid", ConditionOperator.Equal, businessUnitId.Value);

                var res = Service.RetrieveMultiple(qe);

                foreach (var e in res.Entities)
                {
                    roles.Add(new RoleItem
                    {
                        Id = e.Id,
                        Name = e.GetAttributeValue<string>("name") ?? e.Id.ToString()
                    });
                }

                if (!res.MoreRecords) break;
                cookie = res.PagingCookie;
                page++;
            }

            return roles.OrderBy(r => r.Name, StringComparer.CurrentCultureIgnoreCase).ToList();
        }

        // =========================
        // Bind + Filters
        // =========================
        private void BindEntities(List<EntityItem> entities)
        {
            listEntities.BeginUpdate();
            listEntities.Items.Clear();
            foreach (var e in entities) listEntities.Items.Add(e, false);
            listEntities.EndUpdate();
            ApplyEntityFilter();
        }

        private void ApplyEntityFilter()
        {
            var f = GetFilterText(txtEntityFilter, EntityWatermark);

            var filtered = string.IsNullOrWhiteSpace(f)
                ? _entitiesCache
                : _entitiesCache.Where(x =>
                        x.DisplayName.IndexOf(f, StringComparison.CurrentCultureIgnoreCase) >= 0 ||
                        x.LogicalName.IndexOf(f, StringComparison.CurrentCultureIgnoreCase) >= 0)
                    .ToList();

            var checkedNames = new HashSet<string>(
                listEntities.CheckedItems.Cast<EntityItem>().Select(x => x.LogicalName),
                StringComparer.OrdinalIgnoreCase);

            listEntities.BeginUpdate();
            listEntities.Items.Clear();
            foreach (var e in filtered) listEntities.Items.Add(e, checkedNames.Contains(e.LogicalName));
            listEntities.EndUpdate();
        }

        private void BindBusinessUnits(List<BusinessUnitItem> bus)
        {
            cmbBusinessUnits.BeginUpdate();
            cmbBusinessUnits.Items.Clear();
            foreach (var b in bus) cmbBusinessUnits.Items.Add(b);
            cmbBusinessUnits.EndUpdate();

            if (cmbBusinessUnits.Items.Count > 0)
                cmbBusinessUnits.SelectedIndex = 0;
        }

        private void BindRoles(List<RoleItem> roles)
        {
            listRoles.BeginUpdate();
            listRoles.Items.Clear();
            foreach (var r in roles) listRoles.Items.Add(r, false);
            listRoles.EndUpdate();
            ApplyRoleFilter();
        }

        private void ApplyRoleFilter()
        {
            var f = GetFilterText(txtRoleFilter, RoleWatermark);

            var filtered = string.IsNullOrWhiteSpace(f)
                ? _rolesCache
                : _rolesCache.Where(x => x.Name.IndexOf(f, StringComparison.CurrentCultureIgnoreCase) >= 0).ToList();

            var checkedIds = new HashSet<Guid>(listRoles.CheckedItems.Cast<RoleItem>().Select(r => r.Id));

            listRoles.BeginUpdate();
            listRoles.Items.Clear();
            foreach (var r in filtered) listRoles.Items.Add(r, checkedIds.Contains(r.Id));
            listRoles.EndUpdate();
        }

        // =========================
        // Load roles by BU
        // =========================
        private void LoadRolesForSelectedBU()
        {
            WorkAsync(new WorkAsyncInfo
            {
                Message = "Loading roles...",
                Work = (w, a) =>
                {
                    Guid? buId = null;
                    if (!chkAllBusinessUnits.Checked && cmbBusinessUnits.SelectedItem is BusinessUnitItem bu)
                        buId = bu.Id;

                    a.Result = RetrieveRolesPaged_ByBU(CancellationToken.None, buId);
                },
                PostWorkCallBack = (a) =>
                {
                    if (a.Error != null) { MessageBox.Show(a.Error.Message, "Error"); return; }
                    _rolesCache = (List<RoleItem>)a.Result;
                    BindRoles(_rolesCache);

                    SetStatus(chkAllBusinessUnits.Checked
                        ? $"Roles loaded (all BUs): {_rolesCache.Count}"
                        : $"Roles loaded (BU): {_rolesCache.Count}");
                }
            });
        }

        // =========================
        // Privileges for multiple entities (cached)
        // =========================
        private void LoadPrivilegesForSelectedEntities()
        {
            var entities = listEntities.CheckedItems.Cast<EntityItem>().ToList();
            if (entities.Count == 0)
            {
                MessageBox.Show("Select at least one entity.");
                return;
            }

            WorkAsync(new WorkAsyncInfo
            {
                Message = "Loading privileges for selected entities...",
                Work = (w, a) =>
                {
                    foreach (var ent in entities)
                    {
                        if (_entityPrivCache.ContainsKey(ent.LogicalName))
                            continue;

                        var resp = (RetrieveEntityResponse)Service.Execute(new RetrieveEntityRequest
                        {
                            LogicalName = ent.LogicalName,
                            EntityFilters = EntityFilters.Privileges
                        });

                        var map = new Dictionary<string, Guid>(StringComparer.OrdinalIgnoreCase);
                        foreach (var p in resp.EntityMetadata.Privileges)
                        {
                            var key = p.PrivilegeType.ToString();
                            if (!map.ContainsKey(key)) map[key] = p.PrivilegeId;
                        }

                        _entityPrivCache[ent.LogicalName] = map;
                    }
                },
                PostWorkCallBack = (a) =>
                {
                    if (a.Error != null) { MessageBox.Show(a.Error.Message, "Error"); return; }

                    ResetGridToGhost();
                    SetStatus($"Privileges cached for {entities.Count} entities.");
                    MessageBox.Show("Privileges loaded and cached for selected entities.");
                }
            });
        }

        // =========================
        // Apply to (Entities × Roles) with ExecuteMultiple batching
        // =========================
        private void UpdateRolesLogic_MultiEntity()
        {
            var roles = listRoles.CheckedItems.Cast<RoleItem>().ToList();
            if (roles.Count == 0)
            {
                MessageBox.Show("Select at least one role.");
                return;
            }

            var entities = listEntities.CheckedItems.Cast<EntityItem>().ToList();
            if (entities.Count == 0)
            {
                MessageBox.Show("Select at least one entity.");
                return;
            }

            // Build template from grid
            var template = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (DataGridViewColumn col in gridPrivileges.Columns)
            {
                var val = gridPrivileges.Rows[0].Cells[col.Index].Value?.ToString() ?? S_NoChange;
                template[col.Name] = val;
            }

            if (!template.Values.Any(v => v != S_NoChange))
            {
                MessageBox.Show("No changes selected (all are NO CHANGE).");
                return;
            }

            SetStatus("Updating (batched) ...");

            WorkAsync(new WorkAsyncInfo
            {
                Message = "Applying changes (ExecuteMultiple)...",
                Work = (w, a) =>
                {
                    foreach (var ent in entities)
                    {
                        if (!_entityPrivCache.ContainsKey(ent.LogicalName))
                            throw new InvalidOperationException("Privileges not loaded for entity: " + ent.LogicalName);
                    }

                    var batch = NewBatch();

                    foreach (var role in roles)
                    {
                        foreach (var ent in entities)
                        {
                            var privMap = _entityPrivCache[ent.LogicalName];
                            var addList = new List<RolePrivilege>(8);

                            foreach (var kv in template)
                            {
                                var privType = kv.Key;
                                var state = kv.Value;

                                if (state == S_NoChange) continue;
                                if (!privMap.ContainsKey(privType)) continue;

                                var privId = privMap[privType];

                                if (state == S_SetNone)
                                {
                                    batch.Requests.Add(new RemovePrivilegeRoleRequest
                                    {
                                        RoleId = role.Id,
                                        PrivilegeId = privId
                                    });
                                }
                                else
                                {
                                    addList.Add(new RolePrivilege
                                    {
                                        PrivilegeId = privId,
                                        Depth = GetDepth(state)
                                    });
                                }
                            }

                            if (addList.Count > 0)
                            {
                                batch.Requests.Add(new AddPrivilegesRoleRequest
                                {
                                    RoleId = role.Id,
                                    Privileges = addList.ToArray()
                                });
                            }

                            if (batch.Requests.Count >= 200)
                            {
                                Service.Execute(batch);
                                batch = NewBatch();
                            }
                        }
                    }

                    if (batch.Requests.Count > 0)
                        Service.Execute(batch);
                },
                PostWorkCallBack = (a) =>
                {
                    if (a.Error != null)
                    {
                        SetStatus("Update failed.");
                        MessageBox.Show(a.Error.Message, "Error");
                        return;
                    }

                    SetStatus("Update complete.");
                    MessageBox.Show("Update completed.");
                }
            });
        }

        private ExecuteMultipleRequest NewBatch()
        {
            return new ExecuteMultipleRequest
            {
                Settings = new ExecuteMultipleSettings
                {
                    ContinueOnError = true,
                    ReturnResponses = false
                },
                Requests = new OrganizationRequestCollection()
            };
        }

        private PrivilegeDepth GetDepth(string val)
        {
            if (val == S_User) return PrivilegeDepth.Basic;
            if (val == S_BU) return PrivilegeDepth.Local;
            if (val == S_Deep) return PrivilegeDepth.Deep;
            return PrivilegeDepth.Global;
        }

        // =========================
        // DTOs
        // =========================
        private sealed class LoadResult
        {
            public List<EntityItem> Entities { get; set; }
            public List<RoleItem> Roles { get; set; }
            public List<BusinessUnitItem> BusinessUnits { get; set; }
        }

        public sealed class EntityItem
        {
            public string LogicalName { get; set; }
            public string DisplayName { get; set; }
            public override string ToString() => DisplayName + "  (" + LogicalName + ")";
        }

        public sealed class RoleItem
        {
            public Guid Id { get; set; }
            public string Name { get; set; }
            public override string ToString() => Name;
        }

        public sealed class BusinessUnitItem
        {
            public Guid Id { get; set; }
            public string Name { get; set; }
            public override string ToString() => Name;
        }

        // ==================================
        // Legend icon control (CRM-style)
        // ==================================
        private sealed class DepthIconControl : Control
        {
            public string State { get; set; } = "NoChange";
            public Color Yellow { get; set; } = Color.FromArgb(255, 185, 0);
            public Color Green { get; set; } = Color.FromArgb(16, 124, 16);

            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);
                e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

                var rect = new Rectangle(0, 0, Width - 1, Height - 1);

                using (var border = new Pen(Color.FromArgb(110, 110, 110), 1))
                using (var ghost = new Pen(Color.FromArgb(150, 150, 150), 1) { DashStyle = System.Drawing.Drawing2D.DashStyle.Dash })
                using (var red = new Pen(Color.Red, 2))
                using (var yb = new SolidBrush(Yellow))
                using (var gb = new SolidBrush(Green))
                {
                    if (State == "NoChange")
                    {
                        e.Graphics.DrawEllipse(ghost, rect);
                        return;
                    }

                    if (State == "SetNone")
                    {
                        e.Graphics.DrawEllipse(red, rect);
                        e.Graphics.DrawLine(red, rect.Left + 4, rect.Top + 4, rect.Right - 4, rect.Bottom - 4);
                        e.Graphics.DrawLine(red, rect.Left + 4, rect.Bottom - 4, rect.Right - 4, rect.Top + 4);
                        return;
                    }

                    if (State == "Org")
                    {
                        e.Graphics.FillEllipse(gb, rect);
                        e.Graphics.DrawEllipse(border, rect);
                        return;
                    }

                    if (State == "User")
                    {
                        // quarter from bottom (45..135)
                        e.Graphics.FillPie(yb, rect, 45, 90);
                        e.Graphics.DrawEllipse(border, rect);
                        return;
                    }

                    if (State == "BU")
                    {
                        // bottom half
                        e.Graphics.FillPie(yb, rect, 0, 180);
                        e.Graphics.DrawEllipse(border, rect);
                        return;
                    }

                    if (State == "Deep")
                    {
                        // Parent:Child BU: top quarter empty (empty part UP), GREEN
                        e.Graphics.FillPie(gb, rect, 315, 270);
                        e.Graphics.DrawEllipse(border, rect);
                        return;
                    }

                    e.Graphics.DrawEllipse(border, rect);
                }
            }
        }
    }
}