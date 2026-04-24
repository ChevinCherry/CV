import {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  LevelFormat, BorderStyle, ExternalHyperlink, TabStopType, TabStopPosition
} from "docx"
import fs from "fs";
import path from "path";

const OUT_DIR = path.resolve(__dirname, "./out");

const ACCENT = "1F4E79";
const LIGHT_GRAY = "555555";
const PAGE_WIDTH_DXA = 12240;
const MARGIN = 864; // 0.6 inch
const CONTENT_WIDTH = PAGE_WIDTH_DXA - MARGIN * 2;

const heading = (text) => {
  return new Paragraph({
    spacing: { before: 160, after: 60 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: ACCENT, space: 1 } },
    children: [new TextRun({ text, bold: true, size: 22, color: ACCENT, font: "Arial" })]
  });
}

const jobHeader = (company, role, period) => {
  return new Paragraph({
    spacing: { before: 120, after: 40 },
    tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_WIDTH }],
    children: [
      new TextRun({ text: company, bold: true, size: 20, font: "Arial" }),
      new TextRun({ text: "  |  ", size: 20, font: "Arial", color: LIGHT_GRAY }),
      new TextRun({ text: role, size: 20, font: "Arial", italics: true }),
      new TextRun({ text: "\t", size: 20, font: "Arial" }),
      new TextRun({ text: period, size: 18, font: "Arial", color: LIGHT_GRAY }),
    ]
  });
}

const bullet = (text, bold_prefix = null) => {
  const children = [];
  if (bold_prefix) {
    children.push(new TextRun({ text: bold_prefix + " ", bold: true, size: 18, font: "Arial" }));
    children.push(new TextRun({ text, size: 18, font: "Arial" }));
  } else {
    children.push(new TextRun({ text, size: 18, font: "Arial" }));
  }
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 30, after: 30 },
    children
  });
}

const subtext = (text) => {
  return new Paragraph({
    spacing: { before: 0, after: 40 },
    children: [new TextRun({ text, size: 18, font: "Arial", color: LIGHT_GRAY, italics: true })]
  });
}

const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 360, hanging: 200 } } }
        }]
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN }
      }
    },
    children: [
      // NAME
      new Paragraph({
        spacing: { before: 0, after: 40 },
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "Kevin", bold: true, size: 40, font: "Arial", color: ACCENT })]
      }),
      // CONTACT
      new Paragraph({
        spacing: { before: 0, after: 80 },
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({ text: "Toronto, ON  •  ", size: 18, font: "Arial", color: LIGHT_GRAY }),
          new TextRun({ text: "github.com/kevin  •  linkedin.com/in/kevin  •  kevin@email.com", size: 18, font: "Arial", color: LIGHT_GRAY }),
        ]
      }),
      // SUMMARY
      new Paragraph({
        spacing: { before: 0, after: 160 },
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "Full-stack generalist with 6 years of experience owning complex systems end-to-end, from low-latency infrastructure serving 5 million transactions a day to leading features independently from proof-of-concept through production.", size: 18, font: "Arial", italics: true, color: LIGHT_GRAY })]
      }),

      // SKILLS
      heading("Technical Skills"),
      new Paragraph({
        spacing: { before: 60, after: 60 },
        children: [
          new TextRun({ text: "Languages: ", bold: true, size: 18, font: "Arial" }),
          new TextRun({ text: "TypeScript, JavaScript, Python, C++  ", size: 18, font: "Arial" }),
          new TextRun({ text: "  |  Backend: ", bold: true, size: 18, font: "Arial" }),
          new TextRun({ text: "Node.js, Express, REST APIs, WebSockets, OAuth/SSO  ", size: 18, font: "Arial" }),
        ]
      }),
      new Paragraph({
        spacing: { before: 0, after: 60 },
        children: [
          new TextRun({ text: "Frontend: ", bold: true, size: 18, font: "Arial" }),
          new TextRun({ text: "React, MUI, Storybook  ", size: 18, font: "Arial" }),
          new TextRun({ text: "  |  Infra & Tools: ", bold: true, size: 18, font: "Arial" }),
          new TextRun({ text: "GCP, Firebase, PostgreSQL, MySQL, MongoDB, Jenkins, GitLab CI/CD, Playwright, Jest, Jira, Electron  ", size: 18, font: "Arial" }),
        ]
      }),

      // EXPERIENCE
      heading("Experience"),

      // Fortran
      jobHeader("Fortran Traffic Systems", "Full-Stack Developer", "Apr 2023 – Mar 2026"),
      subtext("Aria (15,000+ intersections, ~5M vehicles/day) · Fenyx (universal traffic controller interface)"),
      bullet("Led end-to-end development of an unprecedented configuration transfer system that handled 100,000+ objects in seconds across devices from different manufacturers, automating a process that previously took traffic engineers hours or days manually."),
      bullet("Designed and built a caching layer that reduced network calls to hardware controllers by 100–1000x while simultaneously decreasing update latency, enabling more frequent polling across thousands of live devices."),
      bullet("Solely led Microsoft SSO integration from proof-of-concept to production: designed the OAuth token linking system, handled edge cases across auth flows, coordinated with client IT departments, and authored end-user setup guides."),
      bullet("Built an async data-sync system to reliably synchronize state between intermittently offline devices and a central database, ensuring consistency at scale across unreliable connections."),
      bullet("Collaborated cross-functionally with UI/UX designers, product managers, QA, and hardware engineers across multiple concurrent product tracks to ship features used globally in production."),
      bullet("Drove UI modernization (TypeScript, React, ESLint/Prettier) and built a full component library from Figma specs with Storybook coverage; hardened security by removing SQL injection vulnerabilities and gating all UI calls behind auth tokens."),
      bullet("Built an internal tool to ingest and map real-time GPS data from buses against GTFS municipal route data, including an algorithm to automatically match buses to their closest route without requiring city-provided route assignments; used to measure the effectiveness of a software-only Transit Signal Priority system and deliver performance data to clients."),
      bullet("Mentored an intern developer, providing technical guidance and code reviews throughout their placement."),
      bullet("Delivered a demo-ready MVP for external app iframe integration in one week for the ITS Canada international conference showcase."),

      // MyTechne
      jobHeader("MyTechne", "Software Engineer", "Jun 2020 – Nov 2023"),
      bullet("Built a graph-based survey routing engine for a Slack bot with custom routing rules, multi-format answers, and structured data types, enabling downstream analytics and OpenAI-powered manager insights."),
      bullet("Designed a custom analytics dashboard with team/manager/org-level aggregation; reached first revenue after alpha testing with 4 organizations."),
      bullet("Deployed a cost-effective serverless backend on GCP Cloud Functions, keeping infra lean during the early development cycle."),

      // Vention
      jobHeader("Vention", "Full-Stack Developer (Co-op)", "May 2022 – Aug 2022"),
      bullet("Built a custom reactive component renderer in JavaScript/jQuery, modeled after React's reconciliation pattern, that automatically re-rendered UI based on state variable dependencies. The team later adopted it across other components."),
      bullet("Wrote a customer-facing Python API for direct motor control, collaborating with the hardware team to expose low-level device commands as a clean, documented library."),

      // EDUCATION
      heading("Education"),
      jobHeader("McGill University", "B.Sc. Computer Science, First Class Honours  ·  3.9 / 4.0 GPA", "Sep 2018 – Dec 2022"),
      bullet("Relevant coursework: Algorithm Design & Data Structures, Databases, Software Design, AI, NLP, ML Application Design, Computer Vision, Computer Graphics, Robotics"),
      bullet("Awarded Bourse d'excellence en sciences de l'informatique scholarship ($1,000)"),

      new Paragraph({ spacing: { before: 80, after: 40 }, children: [new TextRun({ text: "Research Publication", bold: true, size: 18, font: "Arial" })] }),
      new Paragraph({
        spacing: { before: 0, after: 40 },
        children: [
          new TextRun({ text: "SUMMIT: Scaffolding OSS Issue Discussion Through Summarization", bold: true, size: 18, font: "Arial" }),
          new TextRun({ text: "  |  ", size: 18, font: "Arial", color: LIGHT_GRAY }),
          new TextRun({ text: "🏆 Best Paper Award", size: 18, font: "Arial", bold: true, color: ACCENT }),
          new TextRun({ text: ", CSCW 2023  (top 1% of all submissions)", size: 18, font: "Arial", color: LIGHT_GRAY }),
        ]
      }),
      bullet("Built a Chrome extension integrating with GitHub to classify and summarize issue tracker threads using a custom NLP model, helping contributors parse large discussion threads faster."),
      new Paragraph({
        spacing: { before: 20, after: 0 },
        children: [
          new TextRun({ text: "arxiv.org/abs/2308.02780", size: 17, font: "Arial", color: "1155CC" })
        ]
      }),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(`${OUT_DIR}/KevinCherryResume_${(new Date()).toLocaleDateString("en-GB", { year: "2-digit", month:"2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", second: "2-digit"})}.docx`, buf);
  console.log("Done");
});