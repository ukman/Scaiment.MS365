import * as React from "react";
import {
  Stack,
  Text,
  SearchBox,
  DefaultButton,
  PrimaryButton,
  Dropdown,
  IDropdownOption,
  Pivot,
  PivotItem,
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  MessageBar,
  getTheme,
  IColumn,
} from "@fluentui/react";
import {
  PieChart,
  Pie,
  ResponsiveContainer,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  Tooltip as ReTooltip,
  Cell,
} from "recharts";
import { RequisitionService } from "../../services/RequisitionService";
import { Requisition } from "../../util/data/DBSchema";

export interface DashboardProps {
  title: string;
}

// ---------- fake data ----------
type RequestStatus =
  | "New"
  | "Awaiting Approval"
  | "Approved"
  | "Ordered"
  | "Shipped"
  | "Delivered"
  | "Partially Delivered"
  | "Cancelled";

type Risk = "Low" | "Medium" | "High";

interface RequestRow {
  id: string;
  name: string;
  project: string;
  supplier: string;
  neededBy: string; // ISO
  createdAt: string; // ISO
  status: RequestStatus;
  items: number;
  total: number; // GBP
  risk: Risk;
  owner: string;
  lastUpdate: string; // ISO
}

const names = [
    "Монтаж пола",
    "Ремонт фасада",
    "Малярка",
    "Электрика на 1 этаж",
    "Сантехника в душевую",
    "Косметика",
    "Благоустроство двора",
    "Двери и окна",
    "Входная группа",
    "Подготовка",
  ];
  
const projects = [
  "Stadium Zarechny",
  "Riverside Tower",
  "Westfield Mall",
  "Tech Park A",
  "Residential Block C",
];
const suppliers = ["Hanson Concrete", "Hilti", "Travis Perkins", "B&Q Trade", "SIG plc"];
const owners = ["Roman", "Alex", "Maria", "Sam", "Chen"];
const statuses: RequestStatus[] = [
  "New",
  "Awaiting Approval",
  "Approved",
  "Ordered",
  "Shipped",
  "Delivered",
  "Partially Delivered",
  "Cancelled",
];

const toISO = (d: Date) => d.toISOString();
const addDays = (base: Date, days: number) => new Date(base.getTime() + days * 86400000);
const currency = (n: number) => new Intl.NumberFormat(undefined, { style: "currency", currency: "GBP" }).format(n);
const formatDate = (iso: string) => new Date(iso).toLocaleDateString();

function deriveRisk(status: RequestStatus, neededByISO: string): Risk {
  const today = new Date();
  const needed = new Date(neededByISO);
  const daysLeft = Math.ceil((needed.getTime() - today.getTime()) / 86400000);
  if (status === "Delivered" || status === "Cancelled") return "Low";
  if (daysLeft <= 0) return "High";
  if (daysLeft <= 5) return "Medium";
  return "Low";
}

function genFakeRows(count = 60): Requisition[] {
  const now = new Date();
  return Array.from({ length: count }).map((_, i) => {
    const created = addDays(now, -Math.floor(Math.random() * 30));
    const needed = addDays(created, Math.floor(Math.random() * 25) + 1);
    const status = statuses[Math.floor(Math.random() * statuses.length)];
    const row: Requisition = {
      id: 1000 + i,
      name: names[Math.floor(Math.random() * names.length)],
      projectName: projects[Math.floor(Math.random() * projects.length)],
      // supplier: suppliers[Math.floor(Math.random() * suppliers.length)],
      dueDate: needed,
      createdAt: created,
      createdBy:1,
      projectId:1
      // status,
      // items: Math.floor(Math.random() * 12) + 1,
      // total: Math.floor(Math.random() * 9000) + 500,
      // risk: "Low",
      // createdBy: owners[Math.floor(Math.random() * owners.length)],
      // lastUpdate: toISO(addDays(created, Math.floor(Math.random() * 15))),
    } as Requisition;
    // row.risk = deriveRisk(row.status, row.neededBy);
    return row;
  });
}

// ---------- theming + helpers ----------
const theme = getTheme();
const colors = {
  primary: theme.palette.themePrimary,
  neutral: theme.palette.neutralSecondary,
  success: "#107C10",
  warning: "#C19C00",
  danger: "#A4262C",
  info: "#005B70",
  teal: "#038387",
  brass: "#986F0B",
};

const statusColorMap: Record<RequestStatus, string> = {
  New: colors.info,
  "Awaiting Approval": colors.primary,
  Approved: colors.success,
  Ordered: colors.brass,
  Shipped: colors.teal,
  Delivered: colors.success,
  "Partially Delivered": colors.warning,
  Cancelled: colors.danger,
};

const riskColorMap: Record<Risk, string> = {
  Low: colors.success,
  Medium: colors.warning,
  High: colors.danger,
};

const hexToRgb = (hex: string) => {
  if(!hex) {
    return '0,0,0'; 
  }
  const s = hex.replace('#','');
  if (s.length !== 6) return '0,0,0';
  const r = parseInt(s.slice(0,2), 16);
  const g = parseInt(s.slice(2,4), 16);
  const b = parseInt(s.slice(4,6), 16);
  return `${r},${g},${b}`;
};
const rgba = (hex: string, a: number) => `rgba(${hexToRgb(hex)},${a})`;

const Chip: React.FC<{ text: string; color: string }> = ({ text, color }) => (
  <span
    style={{
      display: "inline-block",
      padding: "2px 8px",
      borderRadius: 999,
      fontSize: 12,
      lineHeight: 1.6,
      color,
      background: rgba(color, 0.12),
      border: `1px solid ${rgba(color, 0.28)}`,
    }}
  >
    {text}
  </span>
);

// ---------- component ----------
const tile: React.CSSProperties = {
  background: "#fff",
  borderRadius: 12,
  padding: 14,
  boxShadow: "0 1px 3px rgba(0,0,0,0.08)",
};

const kpiNumber: React.CSSProperties = { fontSize: 28, fontWeight: 700, color: colors.primary };

const Dashboard: React.FC<DashboardProps> = ({ title }) => {
  const [rows, setRows] = React.useState<Requisition[]>(() => genFakeRows());
  const [query, setQuery] = React.useState("");
  const [tab, setTab] = React.useState("my");
  const [project, setProject] = React.useState<string | undefined>();
  const [status, setStatus] = React.useState<RequestStatus | undefined>();
  const [owner, setOwner] = React.useState<string | undefined>();

  const filtered = React.useMemo(() => {
    return rows.filter((_r) => { return true;
        /*
      const matchesQuery =
        !query ||
        r.id.toLowerCase().includes(query.toLowerCase()) ||
        r.project.toLowerCase().includes(query.toLowerCase()) ||
        r.supplier.toLowerCase().includes(query.toLowerCase());
      const matchesProject = !project || r.project === project;
      const matchesStatus = !status || r.status === status;
      const matchesOwner = !owner || r.owner === owner;

      const now = new Date();
      const needed = new Date(r.neededBy);
      const isOverdue = needed < now && r.status !== "Delivered" && r.status !== "Cancelled";
      const isAtRisk = r.risk !== "Low";
      const tabOk =
        tab === "all" ||
        (tab === "my" && r.owner === "Roman") ||
        (tab === "overdue" && isOverdue) ||
        (tab === "atrisk" && isAtRisk) ||
        (tab === "approvals" && r.status === "Awaiting Approval") ||
        (tab === "upcoming" && needed >= now && needed <= addDays(now, 7));

      return matchesQuery && matchesProject && matchesStatus && matchesOwner && tabOk;
      */
    });
  }, [rows, query, project, status, owner, tab]);

  // KPIs
  const kpi = React.useMemo(() => {
    const open = rows.filter((_r) => true /*!["Delivered", "Cancelled"].includes(r.status) */).length;
    const approvals = rows; //.filter((r) => r.status === "Awaiting Approval").length;
    const overdue = rows; //.filter((r) => new Date(r.neededBy) < new Date() && !["Delivered", "Cancelled"].includes(r.status)).length;
    const thisWeek = rows; /*.filter((r) => {
      const d = new Date(r.neededBy);
      const now = new Date();
      return d >= now && d <= addDays(now, 7);
    }).length;*/
    const atRisk = rows; //.filter((r) => r.risk !== "Low").length;
    const pendingValue = rows;
    /*
      .filter((r) => !["Delivered", "Cancelled"].includes(r.status))
      .reduce((s, r) => s + r.total, 0);*/
    return { open, approvals, overdue, thisWeek, atRisk, pendingValue };
  }, [rows]);

  // Charts data
  const statusCounts = React.useMemo(() => {
    const map: Record<string, number> = {};
    // rows.forEach((r) => (map[r.status] = (map[r.status] || 0) + 1));
    return Object.entries(map).map(([name, value]) => ({ name, value }));
  }, [rows]);
/*
  const weeklyDeliveries = React.useMemo(() => {
    const m = new Map<string, number>();
    rows.forEach((r) => {
      const d = r.dueDate;// new Date(r.neededBy);
      if(!d) {
        return;
      }
      const monday = new Date(d);
      const diffToMon = (d.getDay() + 6) % 7;
      monday.setDate(d.getDate() - diffToMon);
      monday.setHours(0, 0, 0, 0);
      const key = monday.toISOString().slice(0, 10);
      m.set(key, (m.get(key) || 0) + 1);
    });
    return Array.from(m.entries())
      .map(([week, count]) => ({ week, count }))
      .sort((a, b) => (a.week < b.week ? -1 : 1));
  }, [rows]);*/

  // Dropdown options (v8)
  const projectOpts: IDropdownOption[] = projects.map((p) => ({ key: p, text: p }));
  const statusOpts: IDropdownOption[] = statuses.map((s) => ({ key: s, text: s }));
  const ownerOpts: IDropdownOption[] = owners.map((o) => ({ key: o, text: o }));

  // Columns for DetailsList
  const columns : IColumn[] = [
    { key: "id", name: "Request", minWidth: 100, maxWidth:100, onRender: (i: RequestRow) => (
      <Stack tokens={{ childrenGap: 4 }}>
        <Text variant="mediumPlus"><b>{i.id}</b></Text>
        <Text variant="xSmall">{i.items} items</Text>
      </Stack>
    )},
    { key: "proj", name: "Name / Project", minWidth: 150, onRender: (i: RequestRow) => (
      <Stack tokens={{ childrenGap: 4 }}>
        <Text>{i.name}</Text>
        <Text variant="xSmall">{i.project}</Text>
      </Stack>
    )},
    { key: "needed", name: "Needed by", minWidth: 100, onRender: (i: RequestRow) => (
      <Stack tokens={{ childrenGap: 4 }}>
        <Text>{formatDate(i.neededBy)}</Text>
        <Text variant="xSmall">Created {formatDate(i.createdAt)}</Text>
      </Stack>
    )},
    { key: "status", name: "Status", minWidth: 100, onRender: (i: RequestRow) => (
      <Stack horizontal tokens={{ childrenGap: 8 }}>
        <Chip text={i.status} color={statusColorMap[i.status]} />
        {/* <Chip text={i.risk} color={riskColorMap[i.risk]} /> */}
      </Stack>
    )},
    // { key: "total", name: "Total", minWidth: 100, onRender: (i: RequestRow) => <Text>{currency(i.total)}</Text> },
    { key: "owner", name: "Owner / Updated", minWidth: 120, onRender: (i: RequestRow) => (
      <Stack>
        <Text>{i.owner}</Text>
        <Text variant="xSmall">{formatDate(i.lastUpdate)}</Text>
      </Stack>
    )},
    /*
    { key: "actions", name: "Actions", minWidth: 180, onRender: (i: RequestRow) => (
      <Stack horizontal tokens={{ childrenGap: 8 }}>
        {i.status === "Awaiting Approval" && (
          <PrimaryButton text="Approve" onClick={() => alert(`Approved ${i.id}`)} />
        )}
        {i.status === "Approved" && (
          <DefaultButton text="Create PO" onClick={() => alert(`Create PO for ${i.id}`)} />
        )}
        <DefaultButton text="Received" onClick={() => alert(`Received ${i.id}`)} />
      </Stack>
    )},
    */ 
  ];

  const exportCSV = () => {
    /*
    const header = ["id","project","supplier","neededBy","createdAt","status","items","total","risk","owner","lastUpdate"]; 
    const lines = [header.join(","), ...filtered.map(r => [r.id, r.project, r.supplier, r.neededBy, r.createdAt, r.status, r.items, r.total, r.risk, r.owner, r.lastUpdate].join(","))];
    const blob = new Blob([lines.join("\n")], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `requests-export-${Date.now()}.csv`;
    a.click();
    URL.revokeObjectURL(url);
    */
  };

  const resetFilters = () => { setQuery(""); setProject(undefined); setStatus(undefined); setOwner(undefined); };

  // chart palette from statuses
  const pieColors = statusCounts.map(s => statusColorMap[s.name as RequestStatus] || colors.neutral);

  const loadData = async() => {
    Excel.run(async (ctx) => {
        try {
            const requisitionService = await RequisitionService.create(ctx);
            const rows = await requisitionService.findAll();
            setRows(rows);
        } finally {
            await ctx.sync();
        }
    });
  };
  React.useEffect(() => {
    loadData();
  }, []);  

  return (
    <Stack tokens={{ childrenGap: 12 }} styles={{ root: { padding: 12 } }}>
      {/* header */}
      <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
        <Stack>
          <Text variant="xLarge"><b>{title || "Procurement Dashboard"}</b></Text>
          <Text variant="small">Control & visibility over material requests across projects</Text>
        </Stack>
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <DefaultButton text="Refresh demo data" onClick={() => setRows(genFakeRows())} />
          <PrimaryButton text="Export CSV" onClick={exportCSV} />
        </Stack>
      </Stack>

      {/* KPIs */}

      {/* <Stack horizontal tokens={{ childrenGap: 12 }} wrap>
        <Stack style={{ ...tile, borderLeft: `4px solid ${colors.primary}` }}>
          <Text variant="small">Open requests</Text>
          <span style={kpiNumber}>{kpi.open}</span>
        </Stack>
        <Stack style={{ ...tile, borderLeft: `4px solid ${colors.primary}` }}>
          <Text variant="small">Awaiting approvals</Text>
          <span style={kpiNumber}>{kpi.approvals}</span>
        </Stack>
        <Stack style={{ ...tile, borderLeft: `4px solid ${colors.danger}` }}>
          <Text variant="small">Overdue</Text>
          <span style={{ ...kpiNumber, color: colors.danger }}>{kpi.overdue}</span>
        </Stack>
        <Stack style={{ ...tile, borderLeft: `4px solid ${colors.primary}` }}>
          <Text variant="small">Due in 7 days</Text>
          <span style={kpiNumber}>{kpi.thisWeek}</span>
        </Stack>
        <Stack style={{ ...tile, borderLeft: `4px solid ${colors.warning}` }}>
          <Text variant="small">At risk</Text>
          <span style={{ ...kpiNumber, color: colors.warning }}>{kpi.atRisk}</span>
        </Stack>
        <Stack style={{ ...tile, borderLeft: `4px solid ${colors.primary}` }}>
          <Text variant="small">Pending value</Text>
          <span style={kpiNumber}>{currency(kpi.pendingValue)}</span>
        </Stack>
      </Stack> */}

      {/* tabs */}
      <Pivot selectedKey={tab} onLinkClick={(i) => i && setTab(i.props.itemKey!)}>
        <PivotItem headerText="Мои задачи" itemKey="my" />
        <PivotItem headerText="Все заявки" itemKey="all" />
        <PivotItem headerText="Просроченные" itemKey="overdue" />
        <PivotItem headerText="Под риском" itemKey="atrisk" />
        <PivotItem headerText="Ближайшие поставки" itemKey="upcoming" />
        <PivotItem headerText="На согласовании" itemKey="approvals" />
      </Pivot>

      {/* filters */}
      <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center" wrap>
        <SearchBox placeholder="Поиск по номеру, проекту, поставщику" value={query} onChange={(_e, v) => setQuery(v || '')} styles={{ root: { width: 280 } }} />
        <Dropdown placeholder="Проект" options={projectOpts} selectedKey={project} onChange={(_e, o) => setProject(o?.key as string)} styles={{ root: { width: 220 } }} />
        <Dropdown placeholder="Статус" options={statusOpts} selectedKey={status} onChange={(_e, o) => setStatus(o?.key as RequestStatus)} styles={{ root: { width: 220 } }} />
        <Dropdown placeholder="Ответственный" options={ownerOpts} selectedKey={owner} onChange={(_e, o) => setOwner(o?.key as string)} styles={{ root: { width: 220 } }} />
        <DefaultButton text="Сбросить" onClick={resetFilters} />
      </Stack>

      {/* table + charts */}
      <Stack horizontal tokens={{ childrenGap: 12 }} wrap>
        <Stack grow styles={{ root: { minWidth: 500 } }}>
          <DetailsList
            items={filtered}
            columns={columns as any}
            selectionMode={SelectionMode.none}
            layoutMode={DetailsListLayoutMode.justified}
          />
        </Stack>
        <Stack styles={{ root: { minWidth: 360, maxWidth: 420 } }} tokens={{ childrenGap: 12 }}>
          <Stack style={tile}>
            <Text variant="medium"><b>Status breakdown</b></Text>
            <div style={{ width: "100%", height: 220 }}>
              <ResponsiveContainer>
                <PieChart>
                  <Pie dataKey="value" data={statusCounts} nameKey="name" outerRadius={80} label>
                    {statusCounts.map((_entry, idx) => (
                      <Cell key={`cell-${idx}`} fill={pieColors[idx]} />
                    ))}
                  </Pie>
                  <ReTooltip />
                </PieChart>
              </ResponsiveContainer>
            </div>
          </Stack>
          {/* <Stack style={tile}>
            <Text variant="medium"><b>Deliveries by week</b></Text>
            <div style={{ width: "100%", height: 220 }}>
              <ResponsiveContainer>
                <BarChart data={weeklyDeliveries}>
                  <XAxis dataKey="week" />
                  <YAxis allowDecimals={false} />
                  <Bar dataKey="count" fill={colors.primary} />
                  <ReTooltip />
                </BarChart>
              </ResponsiveContainer>
            </div>
          </Stack> */}
        </Stack>
      </Stack>

      <MessageBar>Demo data only. Replace with API/Office.js sources.</MessageBar>
    </Stack>
  );
};

export default Dashboard;