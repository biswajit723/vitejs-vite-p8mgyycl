import { useState } from 'react';

const PROJECT_TOOLS = [
  'Reports and Status',
  'Model vs. MTO vs. PID Check',
  'Equipment Status',
  'ISO Planning and Management',
  'Support Planning and Management',
];

const REPORT_CATEGORIES = [
  {
    title: 'E3D Reports',
    items: [
      'MTO Report',
      'Valve Report',
      'Pipe Branch Report',
      'Primary Support Report',
      'ATTA Report',
      'Elbow and Bend Report',
      'Equipment Orientation and Position Report',
      'Special Item Report',
      'Nozzle Report',
    ],
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
        onOpen={(title) => openBlank(title, 'reports')}
      />
    );
  }

  if (page === 'blank') {
    return <BlankPage title={blankTitle} onBack={() => navigate(returnPage)} />;
  }

  return (
    <Dashboard
      onReports={() => navigate('reports')}
      onOpen={(title) => openBlank(title, 'dashboard')}
    />
  );
}

function Dashboard({ onReports, onOpen }) {
  const handleTool = (title) => {
    if (title === 'Reports and Status') {
      onReports();
      return;
    }
    onOpen(title);
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
        <article className="dashboard-card">
          <header className="project-header">
            <div className="project-code">4193</div>
            <div>
              <p>PROJECT DASHBOARD</p>
              <h2>HAMMER HEAD</h2>
            </div>
          </header>

          <div className="tool-list">
            {PROJECT_TOOLS.map((title) => (
              <button key={title} className="tool-button" onClick={() => handleTool(title)}>
                <span>{title}</span>
                <span className="circle-arrow">→</span>
              </button>
            ))}
          </div>
        </article>
      </section>
    </main>
  );
}

function ReportsPage({ onBack, onOpen }) {
  return (
    <PageHeader title="Reports and Status" subtitle="Project 4193 · HAMMER HEAD" onBack={onBack}>
      <section className="report-intro">
        <div>
          <p>INTEGRATED ENGINEERING REPORTS</p>
          <h2>Project Reports and Status</h2>
          <span>Select a report or status item to open its configured page.</span>
        </div>
        <strong>4193</strong>
      </section>

      <section className="category-grid">
        {REPORT_CATEGORIES.map((category) => (
          <article className="category-card" key={category.title}>
            <header><h3>{category.title}</h3></header>
            <div className="category-buttons">
              {category.items.map((item) => (
                <button key={item} onClick={() => onOpen(item)}>
                  <span>{item}</span><span>→</span>
                </button>
              ))}
            </div>
          </article>
        ))}
      </section>
    </PageHeader>
  );
}

function BlankPage({ title, onBack }) {
  return (
    <PageHeader title={title} subtitle="Project 4193 · HAMMER HEAD" onBack={onBack}>
      <section className="blank-page">
        <div>
          <div className="rgt-icon">RGT</div>
          <p>CONTENT PLACEHOLDER</p>
          <h2>{title}</h2>
          <span>This page is reserved for content that will be added in a future phase.</span>
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
          <button className="back-button" onClick={onBack}>←</button>
          <div><p>REPORTS GENERATE TOOLS</p><h1>{title}</h1><span>{subtitle}</span></div>
        </div>
      </header>
      <div className="page-content">{children}</div>
    </main>
  );
}
