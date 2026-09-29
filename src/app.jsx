import { useMemo, useState } from 'react';

const PROJECT_ACTIONS = [
  'Reports and Status',
  'Model vs. MTO vs. PID Check',
  'Equipment Status',
  'ISO Planning and Management',
  'Support Planning and Management',
];

const MODULE_NAMES = [
  'M0C', 'M1C', 'M2C', 'M3C', 'M4C', 'M5C',
  'M1P', 'M2P', 'M3P', 'M4P', 'M5P',
  'M1S', 'M2S', 'M3S', 'M4S', 'M5S',
];

const E3D_REPORT_NAMES = [
  'MTO Report',
  'Valve Report',
  'Pipe Branch Report',
  'Primary Support Report',
  'ATTA Report',
  'Elbow and Bend Report',
  'Equipment Orientation and Position Report',
  'Special Item Report',
  'Nozzle Report',
];

/*
  Automatic count source.
  Real report records must be added inside the relevant module and report array.
  Every array is currently empty because no real reports have been supplied yet.
*/
const REPORT_RECORDS = Object.fromEntries(
  MODULE_NAMES.map((moduleName) => [
    moduleName,
    Object.fromEntries(E3D_REPORT_NAMES.map((reportName) => [reportName, []])),
  ]),
);

const getReportCount = (moduleName, reportName) =>
  REPORT_RECORDS[moduleName]?.[reportName]?.length ?? 0;

const INTEGRATED_TOOL_GROUPS = [
  {
    title: 'E3D Reports',
    items: E3D_REPORT_NAMES,
  },
  {
    title: 'Modelling Status',
    items: [
      '2 Inch and Below',
      '3 Inch and 4 Inch',
      '6 Inch and Above',
      'Fluid Code',
      'Material and Specification',
    ],
  },
  {
    title: 'Gate Status',
    items: [
      'Line Status',
      'Equipment Status',
      'Escape Route Status',
      'Safety Equipment Status',
      'Material Handling Status',
      'Support Status',
    ],
  },
  {
    title: 'ISO Utility',
    items: [
      'Data Consistency',
      'Gusset Report',
      'ISO Break Check',
      'Specification Mismatch',
      'Pipe Aid Check',
      'Pipe Insulation Check',
      'Vendor Document Check',
      'Weld Gap Check',
      'Support Tag Check',
      'Clash Report',
    ],
  },
];

export default function App() {
  const [page, setPage] = useState('dashboard');
  const [blankTitle, setBlankTitle] = useState('');
  const [returnPage, setReturnPage] = useState('dashboard');

  const navigate = (nextPage) => {
    setPage(nextPage);
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  const openBlank = (title, previousPage) => {
    setBlankTitle(title);
    setReturnPage(previousPage);
    navigate('blank');
  };

  if (page === 'reports') {
    return (
      <ReportsPage
        onBack={() => navigate('dashboard')}
        onOpenReport={(title) => openBlank(title, 'reports')}
      />
    );
  }

  if (page === 'blank') {
    return <BlankPage title={blankTitle} onBack={() => navigate(returnPage)} />;
  }

  return (
    <Dashboard
      onHammerReports={() => navigate('reports')}
      onOpenBlank={(title) => openBlank(title, 'dashboard')}
    />
  );
}

function Dashboard({ onHammerReports, onOpenBlank }) {
  const handleAction = (projectCode, projectName, action) => {
    if (projectCode === '4193' && action === 'Reports and Status') {
      onHammerReports();
      return;
    }
    onOpenBlank(`${projectCode} · ${projectName} · ${action}`);
  };

  return (
    <main>
      <section className="hero">
        <div className="hero-overlay" />
        <div className="hero-copy">
          <p>FPSO PROJECT PORTAL</p>
          <h1>Reports Generate Tools</h1>
        </div>
      </section>

      <section className="dashboard-section">
        <div className="project-grid">
          <ProjectCard
            code="4193"
            name="HAMMER HEAD"
            onAction={(action) => handleAction('4193', 'HAMMER HEAD', action)}
          />
          <ProjectCard
            code="4173"
            name="Gato do Mato"
            onAction={(action) => handleAction('4173', 'Gato do Mato', action)}
          />
        </div>
      </section>
    </main>
  );
}

function ProjectCard({ code, name, onAction }) {
  return (
    <article className="dashboard-card">
      <header className="project-header">
        <div className="project-code">{code}</div>
        <div>
          <p>PROJECT DASHBOARD</p>
          <h2>{name}</h2>
        </div>
      </header>
      <div className="tool-list">
        {PROJECT_ACTIONS.map((action) => (
          <button key={action} className="tool-button" onClick={() => onAction(action)}>
            <span>{action}</span>
            <span className="circle-arrow">→</span>
          </button>
        ))}
      </div>
    </article>
  );
}

function ReportsPage({ onBack, onOpenReport }) {
  const [selectedModule, setSelectedModule] = useState(null);
  const [searchText, setSearchText] = useState('');

  const visibleModules = useMemo(() => {
    const query = searchText.trim().toLowerCase();
    if (!query) return MODULE_NAMES;
    return MODULE_NAMES.filter((moduleName) => moduleName.toLowerCase().includes(query));
  }, [searchText]);

  const selectOrCloseModule = (moduleName) => {
    setSelectedModule((current) => (current === moduleName ? null : moduleName));
  };

  return (
    <PageHeader title="Reports and Status" subtitle="Project 4193 · HAMMER HEAD" onBack={onBack}>
      <section className="reports-layout">
        <aside className="module-sidebar" aria-label="Project modules">
          <div className="module-sidebar-heading">
            <p>PROJECT MODULES</p>
            <h2>Module List</h2>
          </div>

          <label className="module-search">
            <span aria-hidden="true">⌕</span>
            <input
              type="search"
              value={searchText}
              onChange={(event) => setSearchText(event.target.value)}
              placeholder="Search module name"
            />
          </label>

          <div className="module-nav">
            {visibleModules.map((moduleName) => (
              <button
                key={moduleName}
                type="button"
                className={`module-nav-button ${selectedModule === moduleName ? 'is-selected' : ''}`}
                onClick={() => selectOrCloseModule(moduleName)}
                aria-pressed={selectedModule === moduleName}
              >
                {moduleName}
              </button>
            ))}
          </div>
        </aside>

        <section className="report-workspace" aria-live="polite">
          {selectedModule ? (
            <>
              <header className="workspace-header">
                <div>
                  <p>SELECTED MODULE</p>
                  <h2>{selectedModule}</h2>
                  <span>Report counts are calculated from real records in this module.</span>
                </div>
                <button type="button" className="close-module" onClick={() => setSelectedModule(null)}>
                  Close
                </button>
              </header>

              <div className="report-list">
                {E3D_REPORT_NAMES.map((reportName) => (
                  <button
                    key={reportName}
                    className="report-row"
                    onClick={() => onOpenReport(`${selectedModule} · ${reportName}`)}
                  >
                    <span>{reportName}</span>
                    <strong className="count-badge">
                      {getReportCount(selectedModule, reportName)}
                    </strong>
                  </button>
                ))}
              </div>
            </>
          ) : (
            <div className="workspace-empty">
              <div className="empty-icon">RGT</div>
              <p>MODULE SELECTION</p>
              <h2>Select a module</h2>
              <span>Module names remain fixed on the left. Select a module to open its reports on the right.</span>
            </div>
          )}
        </section>
      </section>

      <IntegratedToolSection onOpenReport={onOpenReport} />
    </PageHeader>
  );
}

function IntegratedToolSection({ onOpenReport }) {
  return (
    <section className="integrated-tool-section">
      <header>
        <p>INTEGRATED TOOL</p>
        <h2>Overall Project and Module-Wise Status</h2>
      </header>

      <div className="integrated-grid">
        {INTEGRATED_TOOL_GROUPS.map((group, groupIndex) => (
          <article className="integrated-card" key={group.title}>
            <h3>{group.title}</h3>
            <div>
              {group.items.map((item, itemIndex) => (
                <button
                  key={item}
                  className={`integrated-button color-${(groupIndex + itemIndex) % 5}`}
                  onClick={() => onOpenReport(item)}
                >
                  {item}
                </button>
              ))}
            </div>
          </article>
        ))}
      </div>
    </section>
  );
}

function BlankPage({ title, onBack }) {
  return (
    <PageHeader title={title} subtitle="Content Placeholder" onBack={onBack}>
      <section className="blank-page">
        <div>
          <div className="rgt-icon">RGT</div>
          <p>CONTENT PLACEHOLDER</p>
          <h2>{title}</h2>
          <span>No project data or report has been added to this page yet.</span>
          <button onClick={onBack}>← Return to Previous Page</button>
        </div>
      </section>
    </PageHeader>
  );
}

function PageHeader({ title, subtitle, onBack, children }) {
  return (
    <main className="internal-page">
      <header className="internal-header">
        <div className="internal-header-inner">
          <button className="back-button" onClick={onBack}>Back</button>
          <div>
            <p>REPORTS GENERATE TOOLS</p>
            <h1>{title}</h1>
            <span>{subtitle}</span>
          </div>
        </div>
      </header>
      <div className="page-content">{children}</div>
    </main>
  );
}
