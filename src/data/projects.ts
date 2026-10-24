export type ProjectTechGroup = {
  label: string;
  items: string[];
};

export type ProjectChallenge = {
  title: string;
  challenges?: string[];
  learnings: string[];
};

export type Project = {
  slug: string;
  title: string;
  company: string;
  category?: string;
  description: string;
  role?: string;
  image: string;
  stack: string[];
  liveUrl?: string;
  repoUrl?: string;
  overview: string[];
  achievements: string[];
  challenges: ProjectChallenge[];
  contributions: { title: string; items: string[] }[];
  technicalDetails: string[];
  technologies: ProjectTechGroup[];
  futurePlans: string[];
};

export const featuredProjects: Project[] = [
  {
    slug: 'destinique-travel-platform',
    title: 'Destinique Travel Platform',
    company: 'Destinique Travel LLC',
    category: 'Travel / booking',
    description:
      'Customer-facing travel and property platform covering booking flows, promotions, and property experiences. Contributed across Angular frontend and PHP/Laravel backend APIs used by visitors and the operations team.',
    role: 'Tech Lead — Business Applications & Automation',
    image: 'img/screenshots/destinique.png',
    stack: ['Angular', 'PHP', 'REST API', 'MySQL', 'JavaScript', 'Bootstrap', 'Cloudflare'],
    liveUrl: 'https://destinique.com/',
    overview: [
      'Destinique needed a reliable customer-facing platform for travel and property experiences—booking, promotions, and property discovery that could work alongside day-to-day operations tools.',
      'I contributed across the Angular frontend and PHP/Laravel backend APIs, helping connect visitor journeys with the data and workflows the business team depends on.',
    ],
    achievements: [
      'Supported booking, promotions, and property-related customer flows on a live production platform.',
      'Worked across frontend and backend so features stayed consistent from UI to API to database.',
      'Helped improve reliability around integrations, performance, and production issues as usage grew.',
      'Collaborated with business stakeholders to turn requirements into practical, shippable releases.',
    ],
    challenges: [
      {
        title: 'Connecting customer UX with operational reality',
        challenges: [
          'Customer-facing features had to stay aligned with booking, property, and promotion rules used by the ops team.',
          'Changes in one flow often affected APIs, data, and dashboards used internally.',
        ],
        learnings: [
          'Early alignment on business rules reduces rework between frontend, API, and operations tools.',
          'Clear ownership of data contracts makes cross-team delivery faster.',
        ],
      },
      {
        title: 'Production stability under real traffic',
        challenges: [
          'Live travel and property traffic means bugs and slow queries show up quickly.',
        ],
        learnings: [
          'Query tuning, careful releases, and fast triage matter as much as new features.',
        ],
      },
    ],
    contributions: [
      {
        title: 'Frontend & API collaboration',
        items: [
          'Contributed to Angular customer experiences for booking and property-related flows.',
          'Helped wire UI behavior to PHP/Laravel REST APIs and backend data models.',
        ],
      },
      {
        title: 'Backend & reliability',
        items: [
          'Worked on API endpoints, database queries, and production troubleshooting.',
          'Supported Cloudflare/VPS-related delivery concerns when shipping updates.',
        ],
      },
    ],
    technicalDetails: [
      'Stack centered on Angular for customer UI and PHP/Laravel APIs for business logic.',
      'MySQL as the primary data store for booking, property, and related operational data.',
      'Cloudflare and VPS hosting in the delivery path for production traffic.',
    ],
    technologies: [
      { label: 'Frontend', items: ['Angular', 'JavaScript', 'Bootstrap'] },
      { label: 'Backend', items: ['PHP', 'REST API'] },
      { label: 'Database', items: ['MySQL'] },
      { label: 'Others', items: ['Cloudflare', 'VPS'] },
    ],
    futurePlans: [
      'Deeper automation around promotions and booking operations',
      'Richer reporting and monitoring tied to customer and ops activity',
      'Continued performance and reliability improvements as traffic grows',
    ],
  },
  {
    slug: 'destinique-business-dashboard',
    title: 'Destinique Business Dashboard',
    company: 'Destinique Travel LLC',
    category: 'Business automation / reporting',
    description:
      'Business automation dashboard for ads, promotions, visitor tracking, monitoring, and operational workflows—helping the team turn day-to-day activity into clearer, actionable visibility.',
    role: 'Tech Lead — Business Applications & Automation',
    image: 'img/screenshots/destinique_dashboard_ad_generation.png',
    stack: ['PHP', 'MySQL', 'JavaScript', 'REST API', 'Bootstrap', 'Cloudflare'],
    liveUrl: 'https://quote.destinique.com/destin/dashboard/',
    overview: [
      'Operations needed a practical way to manage ads, promotions, visitor tracking, monitoring, and related workflows without relying on scattered spreadsheets and manual checks.',
      'I built and maintained a business dashboard that turns operational activity into clearer visibility—APIs, screens, and reporting pieces the team can use every day.',
    ],
    achievements: [
      'Delivered dashboard workflows for ads, promotions, visitor tracking, and monitoring.',
      'Reduced manual follow-up by centralizing operational information in one place.',
      'Connected dashboard actions to backend APIs and MySQL-backed operational data.',
      'Improved day-to-day decision-making with clearer reporting and tracking views.',
    ],
    challenges: [
      {
        title: 'Turning messy operational data into useful views',
        challenges: [
          'Activity came from multiple flows—ads, promotions, visitors, monitoring—with different shapes and update rates.',
        ],
        learnings: [
          'Good dashboards start with clear questions, then shape the data model around those answers.',
          'Small, reliable reporting beats complex screens people avoid.',
        ],
      },
      {
        title: 'Keeping automation trustworthy',
        challenges: [
          'If tracking or monitoring is wrong, teams stop trusting the tool.',
        ],
        learnings: [
          'Validation, clear status, and fast bug fixes are part of the product—not afterthoughts.',
        ],
      },
    ],
    contributions: [
      {
        title: 'Business automation',
        items: [
          'Built dashboard features for ad generation, promotions, visitor tracking, and monitoring.',
          'Supported operational workflows that help the team act on day-to-day activity.',
        ],
      },
      {
        title: 'Backend & data',
        items: [
          'Implemented PHP/REST API pieces and MySQL queries behind dashboard screens.',
          'Helped with production fixes, performance, and practical reporting improvements.',
        ],
      },
    ],
    technicalDetails: [
      'PHP + JavaScript/Bootstrap UI for operational screens.',
      'REST APIs and MySQL for workflow data and reporting.',
      'Deployed behind Cloudflare with VPS-oriented production support.',
    ],
    technologies: [
      { label: 'Frontend', items: ['JavaScript', 'Bootstrap'] },
      { label: 'Backend', items: ['PHP', 'REST API'] },
      { label: 'Database', items: ['MySQL'] },
      { label: 'Others', items: ['Cloudflare'] },
    ],
    futurePlans: [
      'Richer monitoring and alerting around operational activity',
      'More automated reporting for promotions and visitor trends',
      'Tighter links between dashboard actions and customer-facing outcomes',
    ],
  },
  {
    slug: 'paimart-wholesale-ecommerce',
    title: 'PaiMart — Wholesale E-commerce',
    company: 'PaiMart / JS Group',
    category: 'E-commerce / ERP',
    description:
      'Wholesale e-commerce and business platform spanning products, orders, inventory, payments, customers, ERP/POS functions, reporting, and APIs—built to support real sales and distribution operations.',
    role: 'Deputy Manager',
    image: 'img/screenshots/paimart.png',
    stack: ['Laravel', 'PHP', 'MySQL', 'JavaScript', 'REST API'],
    liveUrl: 'https://paimart.painest.com/',
    overview: [
      'PaiMart (previously Paikaree.com.bd), backed by JS Enterprise, needed a complete wholesale e-commerce and business platform—not only a storefront, but the operations behind sales, distribution, inventory, and back-office work.',
      'Over 4+ years I helped design, build, and maintain core platform pieces across e-commerce, ERP/POS-style workflows, APIs, reporting, and day-to-day business operations.',
    ],
    achievements: [
      'Supported a full business platform covering products, orders, inventory, payments, customers, and reporting.',
      'Delivered ERP/POS-related workflows used in real sales and distribution operations.',
      'Exposed APIs for storefront and mobile integrations.',
      'Improved production reliability through releases, query work, and ongoing support as volume grew.',
    ],
    challenges: [
      {
        title: 'Platform breadth across sales and operations',
        challenges: [
          'Wholesale, retail, field sales, inventory, payments, and back-office workflows all had to stay consistent.',
          'Business rules changed as operations scaled.',
        ],
        learnings: [
          'Understanding the business process first prevents building the wrong feature well.',
          'Shared data models and clear module boundaries make large platforms maintainable.',
        ],
      },
      {
        title: 'Growth and production pressure',
        challenges: [
          'Catalog and order volume increased pain around queries, integrations, and releases.',
        ],
        learnings: [
          'Performance work and operational ownership are continuous—not a one-time cleanup.',
        ],
      },
    ],
    contributions: [
      {
        title: 'Engineering ownership',
        items: [
          'Owned major Laravel platform areas across e-commerce, ERP/POS, inventory, and reporting.',
          'Designed tables, APIs, and operational screens used by sales and warehouse teams.',
        ],
      },
      {
        title: 'Delivery & support',
        items: [
          'Handled production issues, releases, and integration work with storefront/mobile teams.',
          'Helped translate business requirements into practical software the company could run daily.',
        ],
      },
    ],
    technicalDetails: [
      'Laravel/PHP monolith-style business platform with MySQL as the system of record.',
      'REST APIs for storefront and mobile consumers.',
      'JavaScript-enhanced operational screens for internal business workflows.',
    ],
    technologies: [
      { label: 'Frontend', items: ['JavaScript'] },
      { label: 'Backend', items: ['Laravel', 'PHP', 'REST API'] },
      { label: 'Database', items: ['MySQL'] },
      { label: 'Others', items: ['Git', 'VPS'] },
    ],
    futurePlans: [
      'Stronger automation around distribution and inventory operations',
      'Richer BI and reporting for sales and stock decisions',
      'Continued API and performance improvements for growth',
    ],
  },
  {
    slug: 'js-erp-sales-billing-inventory',
    title: 'JS ERP — Sales, Billing & Inventory',
    company: 'JS Group',
    category: 'ERP',
    description:
      'ERP-style application for sales, billing, and inventory workflows—helping teams manage core commercial operations in one place with reporting and backend reliability in mind.',
    role: 'Technical Manager — JS Enterprise',
    image: 'img/portfolio/js-erp.png',
    stack: ['Laravel', 'PHP', 'MySQL', 'JavaScript'],
    liveUrl: 'https://js-erp.paitechno.com',
    overview: [
      'JS ERP was built to bring sales, billing, and inventory workflows into one practical system for commercial operations.',
      'I contributed backend engineering around Laravel/MySQL modules that support day-to-day business processes and reporting needs.',
    ],
    achievements: [
      'Supported ERP modules covering sales, billing, and inventory operations.',
      'Helped teams reduce reliance on disconnected spreadsheets for core commercial work.',
      'Improved backend reliability around transactional and reporting workflows.',
    ],
    challenges: [
      {
        title: 'Unifying commercial workflows',
        challenges: [
          'Sales, billing, and inventory each had their own edge cases and reporting expectations.',
        ],
        learnings: [
          'ERP value comes from consistency—shared masters, clear statuses, and trustworthy reports.',
        ],
      },
    ],
    contributions: [
      {
        title: 'Backend engineering',
        items: [
          'Built and maintained Laravel modules for commercial operations.',
          'Worked on MySQL data models, reporting queries, and production fixes.',
        ],
      },
    ],
    technicalDetails: [
      'Laravel/PHP backend with MySQL for transactional and reporting data.',
      'JavaScript used where operational screens needed richer interaction.',
    ],
    technologies: [
      { label: 'Frontend', items: ['JavaScript'] },
      { label: 'Backend', items: ['Laravel', 'PHP'] },
      { label: 'Database', items: ['MySQL'] },
    ],
    futurePlans: [
      'Deeper reporting and analytics for sales and inventory',
      'More automation around billing and stock movement',
    ],
  },
  {
    slug: 'ecommerce-operations-back-office',
    title: 'E-commerce Operations — Back Office',
    company: 'PaiMart / JS Group',
    category: 'Operations',
    description:
      'Back-office operations portal for e-commerce day-to-day work—supporting inventory, order handling, and business workflows tied to the storefront.',
    role: 'Deputy Manager',
    image: 'img/portfolio/paikaree.png',
    stack: ['Laravel', 'PHP', 'MySQL', 'JavaScript'],
    liveUrl: 'https://operation.paikaree.com.bd/',
    overview: [
      'Storefront success depends on a practical back office. This portal supported inventory, order handling, and related business workflows behind the e-commerce experience.',
      'I contributed backend features that helped operations teams run day-to-day work without living in spreadsheets.',
    ],
    achievements: [
      'Supported back-office workflows for inventory and order operations.',
      'Connected operational screens to Laravel/MySQL business data.',
      'Helped improve speed and clarity for teams handling daily e-commerce work.',
    ],
    challenges: [
      {
        title: 'Ops tools must match real warehouse and order reality',
        challenges: [
          'Processes change quickly during busy sales periods.',
        ],
        learnings: [
          'Sit with the users, watch the process, then automate the painful steps first.',
        ],
      },
    ],
    contributions: [
      {
        title: 'Operations software',
        items: [
          'Built and maintained Laravel back-office features for e-commerce operations.',
          'Supported inventory/order workflows tied to the live storefront.',
        ],
      },
    ],
    technicalDetails: [
      'Laravel/PHP operations portal on MySQL.',
      'JavaScript enhancements for day-to-day operational screens.',
    ],
    technologies: [
      { label: 'Frontend', items: ['JavaScript'] },
      { label: 'Backend', items: ['Laravel', 'PHP'] },
      { label: 'Database', items: ['MySQL'] },
    ],
    futurePlans: [
      'More automation around order and inventory exceptions',
      'Better operational reporting for back-office teams',
    ],
  },
  {
    slug: 'ngte-welfare-trust',
    title: 'NGTE Welfare Trust',
    company: 'ADN Telecom PLC',
    category: 'Government ICT',
    description:
      'Government ICT welfare platform covering application processing, tracking, information management, document generation, and administrative workflows where reliability and data accuracy matter.',
    role: 'Assistant Manager — Engineering',
    image: 'img/portfolio/terbb.png',
    stack: ['Laravel', 'PHP', 'MySQL', 'JavaScript'],
    liveUrl: 'https://apps.ngte-welfaretrust.gov.bd/',
    overview: [
      'The NGTE Welfare Trust platform supported government ICT workflows for applications, tracking, information management, document generation, and administration.',
      'While at ADN Telecom PLC, I contributed to Laravel-based development where process reliability and data accuracy were non-negotiable.',
    ],
    achievements: [
      'Supported application processing and tracking workflows for a government ICT platform.',
      'Contributed to information management and document generation features.',
      'Helped deliver administrative flows with attention to reliability and accuracy.',
    ],
    challenges: [
      {
        title: 'Process-heavy government workflows',
        challenges: [
          'Multiple statuses, documents, and administrative checks had to stay consistent.',
          'Mistakes in data or process are costly in welfare/application systems.',
        ],
        learnings: [
          'Clear process mapping and careful validation matter more than flashy UI.',
          'Auditability and accuracy are features, not extras.',
        ],
      },
    ],
    contributions: [
      {
        title: 'Application development',
        items: [
          'Built Laravel backend features for applications, tracking, and admin workflows.',
          'Supported schema work, testing, deploy help, and production issue fixes.',
        ],
      },
    ],
    technicalDetails: [
      'Laravel/PHP application stack with MySQL.',
      'JavaScript/Bootstrap for administrative interfaces.',
    ],
    technologies: [
      { label: 'Frontend', items: ['JavaScript', 'Bootstrap'] },
      { label: 'Backend', items: ['Laravel', 'PHP'] },
      { label: 'Database', items: ['MySQL'] },
    ],
    futurePlans: [
      'Stronger reporting and tracking visibility for administrators',
      'Continued hardening around document and application accuracy',
    ],
  },
];

export function getProjectBySlug(slug: string) {
  return featuredProjects.find((p) => p.slug === slug);
}
