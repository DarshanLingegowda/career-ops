# generate-cvs-windows.ps1
# Run from your career-ops folder on Windows.
# Creates all 5 CV HTML files and generates PDFs using Chrome.

$ErrorActionPreference = 'Stop'
$repoRoot = $PSScriptRoot

$chromePaths = @(
    'C:\Program Files\Google\Chrome\Application\chrome.exe',
    'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe',
    "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
)
$chrome = $chromePaths | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $chrome) { Write-Error "Chrome not found."; exit 1 }
Write-Host "Chrome: $chrome"

$htmlDir = Join-Path $repoRoot 'cv-html'
$outDir  = Join-Path $repoRoot 'output'
New-Item -ItemType Directory -Force -Path $htmlDir | Out-Null
New-Item -ItemType Directory -Force -Path $outDir  | Out-Null

# --- KPMG ---
$htmlPath_kpmg = Join-Path $htmlDir "cv-darshan-lingegowda-kpmg.html"
@"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Darshan Linge Gowda — CV</title>
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  html { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  body { font-family: 'DM Sans', sans-serif; font-size: 11px; line-height: 1.5; color: #1a1a2e; background: #ffffff; }
  .page { width: 100%; max-width: 210mm; margin: 0 auto; padding: 2px 0; }
  .header { margin-bottom: 20px; }
  .header h1 { font-family: 'Space Grotesk', sans-serif; font-size: 28px; font-weight: 700; color: #1a1a2e; letter-spacing: -0.02em; margin-bottom: 6px; line-height: 1.1; }
  .header-gradient { height: 2px; background: linear-gradient(to right, hsl(187,74%,32%), hsl(270,70%,45%)); border-radius: 1px; margin-bottom: 10px; }
  .contact-row { display: flex; flex-wrap: wrap; gap: 8px 14px; font-family: 'DM Sans', sans-serif; font-size: 10.5px; color: #555; }
  .contact-row a { color: #555; text-decoration: none; }
  .contact-row .separator { color: #ccc; }
  .section { margin-bottom: 18px; }
  .section-title { font-family: 'Space Grotesk', sans-serif; font-size: 12px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.06em; color: hsl(187,74%,32%); border-bottom: 1.5px solid #e2e2e2; padding-bottom: 4px; margin-bottom: 10px; }
  .summary-text { font-size: 11px; line-height: 1.7; color: #2f2f2f; }
  a { white-space: nowrap; }
  .competencies-grid { display: flex; flex-wrap: wrap; gap: 8px; }
  .competency-tag { font-family: 'DM Sans', sans-serif; font-size: 10px; font-weight: 500; color: hsl(187,74%,28%); background: hsl(187,40%,95%); padding: 4px 10px; border-radius: 3px; border: 1px solid hsl(187,40%,88%); }
  .job { margin-bottom: 14px; }
  .job-header { display: flex; justify-content: space-between; align-items: baseline; gap: 12px; margin-bottom: 4px; }
  .job-company { font-family: 'Space Grotesk', sans-serif; font-size: 12.5px; font-weight: 600; color: hsl(270,70%,45%); }
  .job-period { font-size: 10.5px; color: #777; white-space: nowrap; }
  .job-role { font-size: 11px; font-weight: 600; color: #333; margin-bottom: 6px; }
  .job ul { padding-left: 18px; margin-top: 6px; }
  .job li { font-size: 10.5px; line-height: 1.6; color: #333; margin-bottom: 4px; }
  .job li strong { font-weight: 600; }
  .project { margin-bottom: 12px; }
  .project-title { font-family: 'Space Grotesk', sans-serif; font-size: 11.5px; font-weight: 600; color: hsl(270,70%,45%); }
  .project-badge { font-size: 9px; font-weight: 500; color: hsl(187,74%,32%); background: hsl(187,40%,95%); padding: 1px 6px; border-radius: 2px; margin-left: 6px; }
  .project-desc { font-size: 10.5px; color: #444; margin-top: 3px; line-height: 1.55; }
  .project-tech { font-size: 9.5px; color: #888; margin-top: 3px; }
  .edu-item { margin-bottom: 8px; }
  .edu-header { display: flex; justify-content: space-between; align-items: baseline; gap: 12px; }
  .edu-title { font-weight: 600; font-size: 11px; color: #333; }
  .edu-org { color: hsl(270,70%,45%); font-weight: 500; }
  .edu-year { font-size: 10px; color: #777; white-space: nowrap; }
  .edu-desc { font-size: 10px; color: #666; margin-top: 2px; line-height: 1.5; }
  .cert-item { display: flex; justify-content: space-between; align-items: baseline; gap: 12px; margin-bottom: 6px; }
  .cert-title { font-size: 10.5px; font-weight: 500; color: #333; }
  .cert-org { color: hsl(270,70%,45%); }
  .cert-year { font-size: 10px; color: #777; white-space: nowrap; }
  .skills-grid { display: flex; flex-wrap: wrap; gap: 6px 14px; }
  .skill-category { font-weight: 600; color: #333; font-size: 10.5px; }
  .skill-item { font-size: 10.5px; color: #444; }
  .avoid-break, .job, .project, .edu-item, .cert-item { break-inside: avoid; page-break-inside: avoid; }
</style>
<link rel="preconnect" href="https://fonts.googleapis.com"><link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@300..700&family=DM+Sans:ital,opsz,wght@0,9..40,100..1000;1,9..40,100..1000&display=swap" rel="stylesheet">
</head>
<body>
<div class="page">

  <div class="header avoid-break">
    <h1>Darshan Linge Gowda</h1>
    <div class="header-gradient"></div>
    <div class="contact-row">
      <span>darshanl1711@gmail.com</span>
      <span class="separator">|</span>
      <a href="https://www.linkedin.com/in/darshan-lingegowda-ai/">linkedin.com/in/darshan-lingegowda-ai</a>
      <span class="separator">|</span>
      <a href="https://github.com/DarshanLingegowda">github.com/DarshanLingegowda</a>
      <span class="separator">|</span>
      <span>Bangalore, India · Open to relocation, Zurich Switzerland</span>
    </div>
  </div>

  <div class="section avoid-break">
    <div class="section-title">Professional Summary</div>
    <div class="summary-text">AI/ML Engineer with production experience delivering RAG implementations, AI agent systems, and LLM-powered applications using Gemini, OpenAI, and LangChain. Proven track record: 80% accuracy improvement in ML-driven finance automation for BMW-sector clients, 70% faster invoice workflows via conversational multi-agent AI. Three years of professional experience in Germany (Munich) and M.Sc. from TU Chemnitz — a combination that positions me to contribute directly to KPMG Switzerland's AI practice with both the technical depth and the European professional context clients expect. Experienced across GCP/Vertex AI, prompt engineering, model evaluation, and MLOps best practices in production environments.</div>
  </div>

  <div class="section">
    <div class="section-title">Core Competencies</div>
    <div class="competencies-grid">
      <span class="competency-tag">RAG Pipeline Design</span>
      <span class="competency-tag">LLM Agents &amp; Orchestration</span>
      <span class="competency-tag">Gemini 2.5 / OpenAI API / Mistral</span>
      <span class="competency-tag">Prompt Engineering &amp; Model Evaluation</span>
      <span class="competency-tag">Python &amp; FastAPI</span>
      <span class="competency-tag">MLOps Best Practices</span>
      <span class="competency-tag">GCP / Vertex AI / Cloud Run</span>
      <span class="competency-tag">Multi-Agent Systems (Google ADK)</span>
    </div>
  </div>

  <div class="section">
    <div class="section-title">Work Experience</div>

    <div class="job">
      <div class="job-header">
        <span class="job-company">PICA GmbH</span>
        <span class="job-period">2018 – 2021</span>
      </div>
      <div class="job-role">Software Developer · Munich, Germany</div>
      <ul>
        <li>Delivered <strong>80% accuracy improvement</strong> in ML-driven SAP invoice automation for BMW service centres across South Germany and Austria — applying model evaluation, validation logic, and automated data quality checks in a regulated production environment.</li>
        <li>Built backend data systems for healthcare, mobility, and industrial domains; designed normalisation rules and anomaly detection pipelines for structured and semi-structured datasets.</li>
        <li>Strengthened observability via structured error handling and logging, reducing recurring production failures — directly applicable to LLMOps monitoring practices.</li>
        <li>Collaborated cross-functionally with business stakeholders across German-language enterprise clients to translate requirements into production ML solutions.</li>
      </ul>
    </div>

    <div class="job">
      <div class="job-header">
        <span class="job-company">AAM IT GmbH / Vorwerk Switzerland</span>
        <span class="job-period">2018</span>
      </div>
      <div class="job-role">Software Developer (Contract) · Munich, Germany</div>
      <ul>
        <li>Delivered backend migration and API development for the Thermomix product platform (Vorwerk, Switzerland) — experience directly relevant to Swiss enterprise client delivery.</li>
        <li>Built Python scripts for real-time data mining and ETL optimisation against live, frequently-updated datasets; integrated SOAP web services via Jenkins and Bash automation.</li>
        <li>Provided cross-timezone technical support for teams across Asia, ensuring pipeline stability and rapid incident resolution.</li>
      </ul>
    </div>

    <div class="job">
      <div class="job-header">
        <span class="job-company">ADREMES GmbH</span>
        <span class="job-period">2017</span>
      </div>
      <div class="job-role">Software Developer (Trainee) · Hamburg, Germany</div>
      <ul>
        <li>Backend validation and session-handling workflows; dashboard enhancements for cloud ad-analytics platform.</li>
      </ul>
    </div>
  </div>

  <div class="section avoid-break">
    <div class="section-title">Selected AI/ML Projects</div>

    <div class="project">
      <span class="project-title">AI Invoice Agent — Conversational Multi-Agent System</span>
      <span class="project-badge">Gen AI Academy APAC · 2025</span>
      <div class="project-desc">Built production-grade conversational AI agent automating invoice creation and real-time financial analytics using multi-agent orchestration and Gemini 2.5 Flash. Resolved LLM reliability issues (model/API version mismatch) in production. <strong>Impact: ~70% faster invoice workflows.</strong></div>
      <div class="project-tech">Gemini 2.5 Flash · LangChain · Google ADK · FastAPI · Docker · Cloud Run · Cloud Build CI/CD</div>
    </div>

    <div class="project">
      <span class="project-title">Multi-Agent Productivity Assistant — Orchestrated AI System</span>
      <span class="project-badge">Gen AI Academy APAC · 2025</span>
      <div class="project-desc">Architected an AI agent system where an Orchestrator Agent routes intent to four specialist sub-agents (Calendar, Task, Notes, Email) — applying multi-agent orchestration, real-time tool-use (Google Calendar MCP, Gmail MCP), and stateful memory (AlloyDB AI).</div>
      <div class="project-tech">Gemini 2.0 Flash · Google ADK · LangGraph · Node.js · Cloud Run · AlloyDB AI</div>
    </div>

    <div class="project">
      <span class="project-title">TravelAI — RAG Concierge</span>
      <span class="project-badge">Kaggle Capstone · 2025</span>
      <div class="project-desc">End-to-end RAG pipeline: heterogeneous data ingestion, embedding-based retrieval indexes, semantic search, LLM grounding, and controlled retrieval experiments measuring relevance and consistency. Directly demonstrates RAG implementation skills.</div>
      <div class="project-tech">RAG · Pinecone-style Vector DB · Semantic Search · LLM Grounding · Python</div>
    </div>

    <div class="project">
      <span class="project-title">AI SRE Copilot — Multimodal LLM Observability</span>
      <span class="project-badge">DEVPOST Hackathon · 2025</span>
      <div class="project-desc">Multimodal AI copilot for automated root-cause analysis integrating Datadog metrics/logs with Vertex AI (Gemini). FastAPI service on Cloud Run — demonstrates LLMOps observability and production monitoring.</div>
      <div class="project-tech">Vertex AI · Gemini · Datadog · FastAPI · Cloud Run · MLOps</div>
    </div>
  </div>

  <div class="section avoid-break">
    <div class="section-title">Education</div>
    <div class="edu-item">
      <div class="edu-header">
        <span class="edu-title">M.Sc. Information &amp; Communication Systems · <span class="edu-org">TU Chemnitz, Germany</span></span>
        <span class="edu-year">2017 – 2021</span>
      </div>
      <div class="edu-desc">Thesis: IEEE-published — large-scale geospatial data processing adopted by a Galileo Map Service Provider for production use.</div>
    </div>
    <div class="edu-item">
      <div class="edu-header">
        <span class="edu-title">B.E. Electronics &amp; Communication Engineering · <span class="edu-org">Visvesvaraya Technological University, India</span></span>
        <span class="edu-year">2013 – 2017</span>
      </div>
    </div>
  </div>

  <div class="section avoid-break">
    <div class="section-title">Certifications</div>
    <div class="cert-item">
      <span class="cert-title"><span class="cert-org">DataCamp</span> — Associate AI Engineer for Developers (LLMOps, OpenAI API, LangChain, Pinecone, Semantic Search, Embeddings)</span>
      <span class="cert-year">2025</span>
    </div>
    <div class="cert-item">
      <span class="cert-title"><span class="cert-org">DataCamp</span> — Developing AI Applications · OpenAI Fundamentals</span>
      <span class="cert-year">2025</span>
    </div>
    <div class="cert-item">
      <span class="cert-title"><span class="cert-org">MIT-IDSS</span> — Data Science &amp; Machine Learning · <span class="cert-org">HarvardX</span> CS50 Python · <span class="cert-org">Kaggle</span> Generative AI &amp; AI Agents Intensive</span>
      <span class="cert-year">2024–2025</span>
    </div>
  </div>

  <div class="section avoid-break">
    <div class="section-title">Skills</div>
    <div class="skills-grid">
      <span class="skill-item"><span class="skill-category">LLMs &amp; GenAI:</span> Gemini 2.5, OpenAI API, Mistral (via API), LangChain, LangGraph, RAG, prompt engineering, fine-tuning, LLMOps</span>
      <span class="skill-item"><span class="skill-category">Agentic AI:</span> Multi-agent orchestration, Google ADK, tool-use, stateful agent design, ReAct pattern</span>
      <span class="skill-item"><span class="skill-category">Cloud / MLOps:</span> GCP, Vertex AI, Cloud Run, Cloud Build CI/CD, Docker, Datadog, MLflow</span>
      <span class="skill-item"><span class="skill-category">Languages:</span> Python, TypeScript/Node.js, FastAPI, SQL, C#, Bash</span>
      <span class="skill-item"><span class="skill-category">Languages spoken:</span> English (fluent), German (professional — 3 years work experience in Germany)</span>
    </div>
  </div>

</div>
</body>
</html>
"@
 | Out-File -FilePath $htmlPath_kpmg -Encoding utf8NoBOM
$pdfPath_kpmg = Join-Path $outDir "cv-darshan-lingegowda-kpmg-2026-04-13.pdf"
Write-Host "Generating cv-darshan-lingegowda-kpmg-2026-04-13.pdf..."
& $chrome --headless=new --no-sandbox --disable-gpu --paper-width=8.27 --paper-height=11.69 `
    --print-to-pdf="$pdfPath_kpmg" --print-to-pdf-no-header `
    "$htmlPath_kpmg" 2>$null
if (Test-Path $pdfPath_kpmg) { Write-Host "  OK: cv-darshan-lingegowda-kpmg-2026-04-13.pdf" } else { Write-Warning "  FAILED: cv-darshan-lingegowda-kpmg-2026-04-13.pdf" }

# --- EVOLUTIONIQ ---
$htmlPath_evolutioniq = Join-Path $htmlDir "cv-darshan-lingegowda-evolutioniq.html"
@"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Darshan Linge Gowda — CV</title>
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  html { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  body { font-family: 'DM Sans', sans-serif; font-size: 11px; line-height: 1.5; color: #1a1a2e; background: #ffffff; }
  .page { width: 100%; max-width: 8.5in; margin: 0 auto; padding: 2px 0; }
  .header { margin-bottom: 20px; }
  .header h1 { font-family: 'Space Grotesk', sans-serif; font-size: 28px; font-weight: 700; color: #1a1a2e; letter-spacing: -0.02em; margin-bottom: 6px; line-height: 1.1; }
  .header-gradient { height: 2px; background: linear-gradient(to right, hsl(187,74%,32%), hsl(270,70%,45%)); border-radius: 1px; margin-bottom: 10px; }
  .contact-row { display: flex; flex-wrap: wrap; gap: 8px 14px; font-family: 'DM Sans', sans-serif; font-size: 10.5px; color: #555; }
  .contact-row a { color: #555; text-decoration: none; }
  .contact-row .separator { color: #ccc; }
  .section { margin-bottom: 18px; }
  .section-title { font-family: 'Space Grotesk', sans-serif; font-size: 12px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.06em; color: hsl(187,74%,32%); border-bottom: 1.5px solid #e2e2e2; padding-bottom: 4px; margin-bottom: 10px; }
  .summary-text { font-size: 11px; line-height: 1.7; color: #2f2f2f; }
  a { white-space: nowrap; }
  .competencies-grid { display: flex; flex-wrap: wrap; gap: 8px; }
  .competency-tag { font-family: 'DM Sans', sans-serif; font-size: 10px; font-weight: 500; color: hsl(187,74%,28%); background: hsl(187,40%,95%); padding: 4px 10px; border-radius: 3px; border: 1px solid hsl(187,40%,88%); }
  .job { margin-bottom: 14px; }
  .job-header { display: flex; justify-content: space-between; align-items: baseline; gap: 12px; margin-bottom: 4px; }
  .job-company { font-family: 'Space Grotesk', sans-serif; font-size: 12.5px; font-weight: 600; color: hsl(270,70%,45%); }
  .job-period { font-size: 10.5px; color: #777; white-space: nowrap; }
  .job-role { font-size: 11px; font-weight: 600; color: #333; margin-bottom: 6px; }
  .job ul { padding-left: 18px; margin-top: 6px; }
  .job li { font-size: 10.5px; line-height: 1.6; color: #333; margin-bottom: 4px; }
  .job li strong { font-weight: 600; }
  .project { margin-bottom: 12px; }
  .project-title { font-family: 'Space Grotesk', sans-serif; font-size: 11.5px; font-weight: 600; color: hsl(270,70%,45%); }
  .project-badge { font-size: 9px; font-weight: 500; color: hsl(187,74%,32%); background: hsl(187,40%,95%); padding: 1px 6px; border-radius: 2px; margin-left: 6px; }
  .project-desc { font-size: 10.5px; color: #444; margin-top: 3px; line-height: 1.55; }
  .project-tech { font-size: 9.5px; color: #888; margin-top: 3px; }
  .edu-item { margin-bottom: 8px; }
  .edu-header { display: flex; justify-content: space-between; align-items: baseline; gap: 12px; }
  .edu-title { font-weight: 600; font-size: 11px; color: #333; }
  .edu-org { color: hsl(270,70%,45%); font-weight: 500; }
  .edu-year { font-size: 10px; color: #777; white-space: nowrap; }
  .edu-desc { font-size: 10px; color: #666; margin-top: 2px; line-height: 1.5; }
  .cert-item { display: flex; justify-content: space-between; align-items: baseline; gap: 12px; margin-bottom: 6px; }
  .cert-title { font-size: 10.5px; font-weight: 500; color: #333; }
  .cert-org { color: hsl(270,70%,45%); }
  .cert-year { font-size: 10px; color: #777; white-space: nowrap; }
  .skills-grid { display: flex; flex-wrap: wrap; gap: 6px 14px; }
  .skill-category { font-weight: 600; color: #333; font-size: 10.5px; }
  .skill-item { font-size: 10.5px; color: #444; }
  .avoid-break, .job, .project, .edu-item, .cert-item { break-inside: avoid; page-break-inside: avoid; }
</style>
<link rel="preconnect" href="https://fonts.googleapis.com"><link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@300..700&family=DM+Sans:ital,opsz,wght@0,9..40,100..1000;1,9..40,100..1000&display=swap" rel="stylesheet">
</head>
<body>
<div class="page">

  <div class="header avoid-break">
    <h1>Darshan Linge Gowda</h1>
    <div class="header-gradient"></div>
    <div class="contact-row">
      <span>darshanl1711@gmail.com</span>
      <span class="separator">|</span>
      <a href="https://www.linkedin.com/in/darshan-lingegowda-ai/">linkedin.com/in/darshan-lingegowda-ai</a>
      <span class="separator">|</span>
      <a href="https://github.com/DarshanLingegowda">github.com/DarshanLingegowda</a>
      <span class="separator">|</span>
      <span>Bangalore, India · Open to Remote (USA)</span>
    </div>
  </div>

  <div class="section avoid-break">
    <div class="section-title">Professional Summary</div>
    <div class="summary-text">AI/ML Engineer specialising in production LLM systems, RAG pipelines, and hybrid LLM + statistical ML approaches for high-stakes decision workflows. Proven impact in regulated environments: 80% accuracy improvement in ML-driven document processing for automotive-sector clients, 70% faster workflows via conversational multi-agent AI with human-in-the-loop design. Strong command of embeddings, vector search, and semantic retrieval — combined with 3+ years operating in regulated enterprise environments (healthcare, finance, industrial). Experienced with GCP/Vertex AI, LangChain, MLflow, Docker, and production observability tooling.</div>
  </div>

  <div class="section">
    <div class="section-title">Core Competencies</div>
    <div class="competencies-grid">
      <span class="competency-tag">Production LLM Systems</span>
      <span class="competency-tag">RAG &amp; Embeddings-Based Retrieval</span>
      <span class="competency-tag">Hybrid LLM + Statistical ML</span>
      <span class="competency-tag">Human-in-the-Loop Workflows</span>
      <span class="competency-tag">Vector Search &amp; Semantic Indexing</span>
      <span class="competency-tag">LLMOps &amp; Model Observability</span>
      <span class="competency-tag">Python &amp; FastAPI</span>
      <span class="competency-tag">Regulated Environment ML</span>
    </div>
  </div>

  <div class="section">
    <div class="section-title">Work Experience</div>

    <div class="job">
      <div class="job-header">
        <span class="job-company">PICA GmbH</span>
        <span class="job-period">2018 – 2021</span>
      </div>
      <div class="job-role">Software Developer · Munich, Germany</div>
      <ul>
        <li>Automated complex document processing (SAP invoice generation) for BMW service centres using ML-driven data validation — delivering <strong>80% accuracy improvement</strong> and 50% processing time reduction in a regulated, high-volume production environment.</li>
        <li>Built hybrid data validation pipelines combining rule-based logic with statistical ML models for structured and semi-structured document extraction — a direct analogue to insurance claims synthesis.</li>
        <li>Designed normalisation rules and automated quality checks for heterogeneous datasets across healthcare, mobility, and industrial domains.</li>
        <li>Strengthened system observability via structured error handling and logging, reducing recurring production failures — applies directly to LLM reliability engineering.</li>
      </ul>
    </div>

    <div class="job">
      <div class="job-header">
        <span class="job-company">AAM IT GmbH</span>
        <span class="job-period">2018</span>
      </div>
      <div class="job-role">Software Developer (Contract) · Munich, Germany</div>
      <ul>
        <li>Built Python scripts for real-time data mining, extraction, and correction against continuously-updated production datasets — experience transferable to LLM-driven information extraction.</li>
        <li>Integrated SOAP web services and automated ETL pipelines via Jenkins and Bash; developed C# APIs for complex multi-interface data flows.</li>
      </ul>
    </div>
  </div>

  <div class="section avoid-break">
    <div class="section-title">Selected AI/ML Projects</div>

    <div class="project">
      <span class="project-title">AI Invoice Agent — Human-in-the-Loop Multi-Agent System</span>
      <span class="project-badge">Gen AI Academy APAC · 2025</span>
      <div class="project-desc">Production conversational AI agent with human-in-the-loop design for invoice creation, payment tracking, and financial analytics. Multi-agent orchestration with Gemini 2.5 Flash; resolved critical LLM reliability issues in production. <strong>Impact: ~70% faster workflows.</strong></div>
      <div class="project-tech">Gemini 2.5 · LangChain · Multi-agent orchestration · Docker · Cloud Run · HITL design</div>
    </div>

    <div class="project">
      <span class="project-title">TravelAI — RAG Pipeline with Retrieval Evaluation</span>
      <span class="project-badge">Kaggle Capstone · 2025</span>
      <div class="project-desc">End-to-end RAG implementation: data ingestion, embedding-based retrieval indexing, hybrid semantic search, LLM grounding, and controlled experiments measuring retrieval relevance and consistency — hybrid LLM + retrieval approach directly applicable to medical document synthesis.</div>
      <div class="project-tech">RAG · Pinecone-style vector DB · Semantic search · Embeddings · LLM grounding · Python</div>
    </div>

    <div class="project">
      <span class="project-title">AI SRE Copilot — LLM Observability &amp; Reliability</span>
      <span class="project-badge">DEVPOST Hackathon · 2025</span>
      <div class="project-desc">Multimodal AI copilot for automated root-cause analysis using Vertex AI + Datadog. Demonstrates production LLM monitoring, structured output design, and reliability engineering.</div>
      <div class="project-tech">Vertex AI · Gemini · Datadog · FastAPI · Cloud Run · LLMOps</div>
    </div>

    <div class="project">
      <span class="project-title">Generative AI Search Agent</span>
      <span class="project-badge">Personal Project · Ongoing</span>
      <div class="project-desc">LLM-powered semantic retrieval agent: embeddings + vector DB, prompt chaining, contextual memory. <strong>~40% faster resolution vs. keyword search.</strong> Demonstrates hybrid retrieval + LLM approach.</div>
      <div class="project-tech">LangChain · Embeddings · Vector DB · OpenAI API · Docker · CI/CD</div>
    </div>
  </div>

  <div class="section avoid-break">
    <div class="section-title">Education</div>
    <div class="edu-item">
      <div class="edu-header">
        <span class="edu-title">M.Sc. Information &amp; Communication Systems · <span class="edu-org">TU Chemnitz, Germany</span></span>
        <span class="edu-year">2017 – 2021</span>
      </div>
      <div class="edu-desc">Thesis: IEEE-published — large-scale geospatial data processing with ML pipeline adopted by a Galileo Map Service Provider for production use.</div>
    </div>
    <div class="edu-item">
      <div class="edu-header">
        <span class="edu-title">B.E. Electronics &amp; Communication Engineering · <span class="edu-org">Visvesvaraya Technological University</span></span>
        <span class="edu-year">2013 – 2017</span>
      </div>
    </div>
  </div>

  <div class="section avoid-break">
    <div class="section-title">Certifications</div>
    <div class="cert-item">
      <span class="cert-title"><span class="cert-org">DataCamp</span> — Associate AI Engineer for Developers (LLMOps, OpenAI API, LangChain, Pinecone, Embeddings, Semantic Search)</span>
      <span class="cert-year">2025</span>
    </div>
    <div class="cert-item">
      <span class="cert-title"><span class="cert-org">MIT-IDSS</span> — Data Science &amp; Machine Learning · <span class="cert-org">HarvardX</span> CS50 Python · <span class="cert-org">Kaggle</span> GenAI &amp; AI Agents Intensive</span>
      <span class="cert-year">2024–2025</span>
    </div>
  </div>

  <div class="section avoid-break">
    <div class="section-title">Skills</div>
    <div class="skills-grid">
      <span class="skill-item"><span class="skill-category">LLMs &amp; GenAI:</span> OpenAI API, Gemini 2.5, LangChain, LangGraph, RAG, embeddings, vector DBs (Pinecone), prompt engineering, LLMOps</span>
      <span class="skill-item"><span class="skill-category">ML Engineering:</span> Supervised/unsupervised ML, statistical models, MLflow, evaluation design, feature engineering</span>
      <span class="skill-item"><span class="skill-category">Cloud / MLOps:</span> GCP, Vertex AI, Cloud Run, Cloud Build CI/CD, Docker, Datadog observability</span>
      <span class="skill-item"><span class="skill-category">Languages:</span> Python, FastAPI, Node.js, SQL, C#, Bash</span>
    </div>
  </div>

</div>
</body>
</html>
"@
 | Out-File -FilePath $htmlPath_evolutioniq -Encoding utf8NoBOM
$pdfPath_evolutioniq = Join-Path $outDir "cv-darshan-lingegowda-evolutioniq-2026-04-13.pdf"
Write-Host "Generating cv-darshan-lingegowda-evolutioniq-2026-04-13.pdf..."
& $chrome --headless=new --no-sandbox --disable-gpu --paper-width=8.5 --paper-height=11 `
    --print-to-pdf="$pdfPath_evolutioniq" --print-to-pdf-no-header `
    "$htmlPath_evolutioniq" 2>$null
if (Test-Path $pdfPath_evolutioniq) { Write-Host "  OK: cv-darshan-lingegowda-evolutioniq-2026-04-13.pdf" } else { Write-Warning "  FAILED: cv-darshan-lingegowda-evolutioniq-2026-04-13.pdf" }

# --- NATERA ---
$htmlPath_natera = Join-Path $htmlDir "cv-darshan-lingegowda-natera.html"
@"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Darshan Linge Gowda — CV</title>
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  html { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  body { font-family: 'DM Sans', sans-serif; font-size: 11px; line-height: 1.5; color: #1a1a2e; background: #ffffff; }
  .page { width: 100%; max-width: 8.5in; margin: 0 auto; padding: 2px 0; }
  .header { margin-bottom: 20px; }
  .header h1 { font-family: 'Space Grotesk', sans-serif; font-size: 28px; font-weight: 700; color: #1a1a2e; letter-spacing: -0.02em; margin-bottom: 6px; }
  .header-gradient { height: 2px; background: linear-gradient(to right, hsl(187,74%,32%), hsl(270,70%,45%)); border-radius: 1px; margin-bottom: 10px; }
  .contact-row { display: flex; flex-wrap: wrap; gap: 8px 14px; font-size: 10.5px; color: #555; }
  .contact-row a { color: #555; text-decoration: none; }
  .contact-row .separator { color: #ccc; }
  .section { margin-bottom: 18px; }
  .section-title { font-family: 'Space Grotesk', sans-serif; font-size: 12px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.06em; color: hsl(187,74%,32%); border-bottom: 1.5px solid #e2e2e2; padding-bottom: 4px; margin-bottom: 10px; }
  .summary-text { font-size: 11px; line-height: 1.7; color: #2f2f2f; }
  a { white-space: nowrap; }
  .competencies-grid { display: flex; flex-wrap: wrap; gap: 8px; }
  .competency-tag { font-size: 10px; font-weight: 500; color: hsl(187,74%,28%); background: hsl(187,40%,95%); padding: 4px 10px; border-radius: 3px; border: 1px solid hsl(187,40%,88%); }
  .job { margin-bottom: 14px; }
  .job-header { display: flex; justify-content: space-between; align-items: baseline; gap: 12px; margin-bottom: 4px; }
  .job-company { font-family: 'Space Grotesk', sans-serif; font-size: 12.5px; font-weight: 600; color: hsl(270,70%,45%); }
  .job-period { font-size: 10.5px; color: #777; white-space: nowrap; }
  .job-role { font-size: 11px; font-weight: 600; color: #333; margin-bottom: 6px; }
  .job ul { padding-left: 18px; margin-top: 6px; }
  .job li { font-size: 10.5px; line-height: 1.6; color: #333; margin-bottom: 4px; }
  .job li strong { font-weight: 600; }
  .project { margin-bottom: 12px; }
  .project-title { font-family: 'Space Grotesk', sans-serif; font-size: 11.5px; font-weight: 600; color: hsl(270,70%,45%); }
  .project-badge { font-size: 9px; font-weight: 500; color: hsl(187,74%,32%); background: hsl(187,40%,95%); padding: 1px 6px; border-radius: 2px; margin-left: 6px; }
  .project-desc { font-size: 10.5px; color: #444; margin-top: 3px; line-height: 1.55; }
  .project-tech { font-size: 9.5px; color: #888; margin-top: 3px; }
  .edu-item { margin-bottom: 8px; }
  .edu-header { display: flex; justify-content: space-between; align-items: baseline; gap: 12px; }
  .edu-title { font-weight: 600; font-size: 11px; color: #333; }
  .edu-org { color: hsl(270,70%,45%); font-weight: 500; }
  .edu-year { font-size: 10px; color: #777; white-space: nowrap; }
  .edu-desc { font-size: 10px; color: #666; margin-top: 2px; line-height: 1.5; }
  .cert-item { display: flex; justify-content: space-between; align-items: baseline; gap: 12px; margin-bottom: 6px; }
  .cert-title { font-size: 10.5px; font-weight: 500; color: #333; }
  .cert-org { color: hsl(270,70%,45%); }
  .cert-year { font-size: 10px; color: #777; white-space: nowrap; }
  .skills-grid { display: flex; flex-wrap: wrap; gap: 6px 14px; }
  .skill-category { font-weight: 600; color: #333; font-size: 10.5px; }
  .skill-item { font-size: 10.5px; color: #444; }
  .avoid-break, .job, .project, .edu-item, .cert-item { break-inside: avoid; page-break-inside: avoid; }
</style>
<link rel="preconnect" href="https://fonts.googleapis.com"><link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@300..700&family=DM+Sans:ital,opsz,wght@0,9..40,100..1000;1,9..40,100..1000&display=swap" rel="stylesheet">
</head>
<body>
<div class="page">

  <div class="header avoid-break">
    <h1>Darshan Linge Gowda</h1>
    <div class="header-gradient"></div>
    <div class="contact-row">
      <span>darshanl1711@gmail.com</span>
      <span class="separator">|</span>
      <a href="https://www.linkedin.com/in/darshan-lingegowda-ai/">linkedin.com/in/darshan-lingegowda-ai</a>
      <span class="separator">|</span>
      <a href="https://github.com/DarshanLingegowda">github.com/DarshanLingegowda</a>
      <span class="separator">|</span>
      <span>Bangalore, India · Open to Remote (USA)</span>
    </div>
  </div>

  <div class="section avoid-break">
    <div class="section-title">Professional Summary</div>
    <div class="summary-text">AI/ML Engineer with production experience building agentic AI platforms, RAG systems, and LLMOps pipelines for enterprise workloads. Strong command of LangChain, multi-agent orchestration, vector databases, and compliance-first AI design — applied in regulated environments including healthcare and financial services. Proven impact: 80% accuracy improvement in ML-driven data validation for automotive-sector clients; 70% faster workflows via production multi-agent AI. Experienced with GCP/Vertex AI, Datadog observability, MLflow experiment tracking, and containerised CI/CD deployments. M.Sc. in Information Systems (TU Chemnitz, Germany) with an IEEE-published ML thesis adopted by a production Galileo Map Service Provider.</div>
  </div>

  <div class="section">
    <div class="section-title">Core Competencies</div>
    <div class="competencies-grid">
      <span class="competency-tag">Agentic AI Architecture (LangChain, LangGraph)</span>
      <span class="competency-tag">RAG at Scale &amp; Vector DB Design</span>
      <span class="competency-tag">LLMOps &amp; Datadog Observability</span>
      <span class="competency-tag">Multi-Agent Orchestration (Google ADK)</span>
      <span class="competency-tag">MLflow &amp; Experiment Tracking</span>
      <span class="competency-tag">Compliance-First AI in Regulated Environments</span>
      <span class="competency-tag">GCP / Vertex AI / Cloud Run</span>
      <span class="competency-tag">Docker &amp; Cloud Build CI/CD</span>
    </div>
  </div>

  <div class="section">
    <div class="section-title">Work Experience</div>

    <div class="job">
      <div class="job-header">
        <span class="job-company">PICA GmbH</span>
        <span class="job-period">2018 – 2021</span>
      </div>
      <div class="job-role">Software Developer · Munich, Germany</div>
      <ul>
        <li>Built ML-driven validation and automation systems for BMW service centres — achieving <strong>80% accuracy improvement</strong> and 50% processing time reduction in a regulated, high-volume production environment (GDPR, data integrity compliance).</li>
        <li>Operated across healthcare, mobility, and industrial domains — directly relevant to HIPAA/CLIA-adjacent compliance requirements in clinical AI platforms.</li>
        <li>Designed normalisation rules, label quality checks, and automated anomaly detection pipelines for structured and semi-structured datasets; implemented observability via structured error handling and logging.</li>
      </ul>
    </div>

    <div class="job">
      <div class="job-header">
        <span class="job-company">AAM IT GmbH / Vorwerk Switzerland</span>
        <span class="job-period">2018</span>
      </div>
      <div class="job-role">Software Developer (Contract) · Munich, Germany</div>
      <ul>
        <li>Designed and executed large-scale database migration with evolving schema; built Python scripts for real-time data mining and ETL pipeline optimisation against continuously-updated datasets.</li>
        <li>Developed C# APIs and integrated SOAP web services via Jenkins/Bash; provided cross-timezone support ensuring pipeline stability.</li>
      </ul>
    </div>
  </div>

  <div class="section avoid-break">
    <div class="section-title">Selected AI/ML Projects</div>

    <div class="project">
      <span class="project-title">Multi-Agent Productivity Assistant — Agent Orchestration Platform</span>
      <span class="project-badge">Gen AI Academy APAC · 2025</span>
      <div class="project-desc">Architected multi-agent AI system (Gemini 2.0 Flash + Google ADK): Orchestrator Agent routes intent to four specialist sub-agents (Calendar, Task, Notes, Email) via real-time tool-use (Google Calendar MCP, Gmail MCP). Parallel agent execution handles multi-step workflows in a single turn. AlloyDB AI for persistent structured memory. Demonstrates the agentic AI platform architecture Natera requires.</div>
      <div class="project-tech">Gemini 2.0 Flash · Google ADK · LangGraph · Multi-agent orchestration · Cloud Run · AlloyDB AI</div>
    </div>

    <div class="project">
      <span class="project-title">AI Invoice Agent — Production Agentic System</span>
      <span class="project-badge">Gen AI Academy APAC · 2025</span>
      <div class="project-desc">Production multi-agent system automating invoice creation, payment tracking, and analytics using Gemini 2.5 Flash. Resolved critical LLM reliability issue in production. Deployed via Docker + Cloud Build CI/CD on Cloud Run. <strong>Impact: ~70% faster invoice workflows.</strong></div>
      <div class="project-tech">Gemini 2.5 · LangChain · Docker · Cloud Run · CI/CD · LLM reliability engineering</div>
    </div>

    <div class="project">
      <span class="project-title">AI SRE Copilot — LLMOps Observability</span>
      <span class="project-badge">DEVPOST Hackathon · 2025</span>
      <div class="project-desc">Multimodal AI copilot integrating Datadog metrics/logs with Vertex AI (Gemini) for automated root-cause analysis. FastAPI on Cloud Run. Demonstrates real-time LLMOps monitoring — cost tracking, drift detection, safety monitoring of AI outputs.</div>
      <div class="project-tech">Vertex AI · Gemini · Datadog · FastAPI · Cloud Run · LLMOps</div>
    </div>

    <div class="project">
      <span class="project-title">TravelAI — RAG Pipeline at Scale</span>
      <span class="project-badge">Kaggle Capstone · 2025</span>
      <div class="project-desc">End-to-end RAG implementation: ingestion pipelines, chunking/embedding policies, hybrid vector search, LLM grounding, retrieval relevance experiments. Demonstrates RAG-at-scale engineering across the full Natera LLMOps stack.</div>
      <div class="project-tech">RAG · Pinecone-style vector DB · Embeddings · Semantic search · LLM grounding</div>
    </div>
  </div>

  <div class="section avoid-break">
    <div class="section-title">Education</div>
    <div class="edu-item">
      <div class="edu-header">
        <span class="edu-title">M.Sc. Information &amp; Communication Systems · <span class="edu-org">TU Chemnitz, Germany</span></span>
        <span class="edu-year">2017 – 2021</span>
      </div>
      <div class="edu-desc">Thesis: IEEE-published — large-scale geospatial ML pipeline adopted by a Galileo Map Service Provider for production use.</div>
    </div>
    <div class="edu-item">
      <div class="edu-header">
        <span class="edu-title">B.E. Electronics &amp; Communication Engineering · <span class="edu-org">Visvesvaraya Technological University</span></span>
        <span class="edu-year">2013 – 2017</span>
      </div>
    </div>
  </div>

  <div class="section avoid-break">
    <div class="section-title">Certifications</div>
    <div class="cert-item">
      <span class="cert-title"><span class="cert-org">DataCamp</span> — Associate AI Engineer for Developers (LLMOps, OpenAI API, LangChain, Pinecone, Embeddings, Production AI)</span>
      <span class="cert-year">2025</span>
    </div>
    <div class="cert-item">
      <span class="cert-title"><span class="cert-org">MIT-IDSS</span> — Data Science &amp; Machine Learning · <span class="cert-org">Kaggle</span> Generative AI &amp; AI Agents Intensive</span>
      <span class="cert-year">2024–2025</span>
    </div>
  </div>

  <div class="section avoid-break">
    <div class="section-title">Skills</div>
    <div class="skills-grid">
      <span class="skill-item"><span class="skill-category">Agentic AI:</span> LangChain, LangGraph, Google ADK, multi-agent orchestration, tool-use, ReAct pattern, stateful agent design</span>
      <span class="skill-item"><span class="skill-category">LLMs &amp; RAG:</span> Gemini 2.5, OpenAI API, RAG pipelines, vector DBs (Pinecone), embeddings, prompt engineering, fine-tuning</span>
      <span class="skill-item"><span class="skill-category">MLOps:</span> MLflow, Datadog, GCP Vertex AI, Cloud Run, Cloud Build CI/CD, Docker, containerisation</span>
      <span class="skill-item"><span class="skill-category">Languages:</span> Python, FastAPI, Node.js, SQL, Bash</span>
    </div>
  </div>

</div>
</body>
</html>
"@
 | Out-File -FilePath $htmlPath_natera -Encoding utf8NoBOM
$pdfPath_natera = Join-Path $outDir "cv-darshan-lingegowda-natera-2026-04-13.pdf"
Write-Host "Generating cv-darshan-lingegowda-natera-2026-04-13.pdf..."
& $chrome --headless=new --no-sandbox --disable-gpu --paper-width=8.5 --paper-height=11 `
    --print-to-pdf="$pdfPath_natera" --print-to-pdf-no-header `
    "$htmlPath_natera" 2>$null
if (Test-Path $pdfPath_natera) { Write-Host "  OK: cv-darshan-lingegowda-natera-2026-04-13.pdf" } else { Write-Warning "  FAILED: cv-darshan-lingegowda-natera-2026-04-13.pdf" }

# --- TRMLABS ---
$htmlPath_trmlabs = Join-Path $htmlDir "cv-darshan-lingegowda-trmlabs.html"
@"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Darshan Linge Gowda — CV</title>
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  html { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  body { font-family: 'DM Sans', sans-serif; font-size: 11px; line-height: 1.5; color: #1a1a2e; background: #ffffff; }
  .page { width: 100%; max-width: 8.5in; margin: 0 auto; padding: 2px 0; }
  .header { margin-bottom: 20px; }
  .header h1 { font-family: 'Space Grotesk', sans-serif; font-size: 28px; font-weight: 700; color: #1a1a2e; letter-spacing: -0.02em; margin-bottom: 6px; }
  .header-gradient { height: 2px; background: linear-gradient(to right, hsl(187,74%,32%), hsl(270,70%,45%)); border-radius: 1px; margin-bottom: 10px; }
  .contact-row { display: flex; flex-wrap: wrap; gap: 8px 14px; font-size: 10.5px; color: #555; }
  .contact-row a { color: #555; text-decoration: none; }
  .contact-row .separator { color: #ccc; }
  .section { margin-bottom: 18px; }
  .section-title { font-family: 'Space Grotesk', sans-serif; font-size: 12px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.06em; color: hsl(187,74%,32%); border-bottom: 1.5px solid #e2e2e2; padding-bottom: 4px; margin-bottom: 10px; }
  .summary-text { font-size: 11px; line-height: 1.7; color: #2f2f2f; }
  a { white-space: nowrap; }
  .competencies-grid { display: flex; flex-wrap: wrap; gap: 8px; }
  .competency-tag { font-size: 10px; font-weight: 500; color: hsl(187,74%,28%); background: hsl(187,40%,95%); padding: 4px 10px; border-radius: 3px; border: 1px solid hsl(187,40%,88%); }
  .job { margin-bottom: 14px; }
  .job-header { display: flex; justify-content: space-between; align-items: baseline; gap: 12px; margin-bottom: 4px; }
  .job-company { font-family: 'Space Grotesk', sans-serif; font-size: 12.5px; font-weight: 600; color: hsl(270,70%,45%); }
  .job-period { font-size: 10.5px; color: #777; white-space: nowrap; }
  .job-role { font-size: 11px; font-weight: 600; color: #333; margin-bottom: 6px; }
  .job ul { padding-left: 18px; margin-top: 6px; }
  .job li { font-size: 10.5px; line-height: 1.6; color: #333; margin-bottom: 4px; }
  .job li strong { font-weight: 600; }
  .project { margin-bottom: 12px; }
  .project-title { font-family: 'Space Grotesk', sans-serif; font-size: 11.5px; font-weight: 600; color: hsl(270,70%,45%); }
  .project-badge { font-size: 9px; font-weight: 500; color: hsl(187,74%,32%); background: hsl(187,40%,95%); padding: 1px 6px; border-radius: 2px; margin-left: 6px; }
  .project-desc { font-size: 10.5px; color: #444; margin-top: 3px; line-height: 1.55; }
  .project-tech { font-size: 9.5px; color: #888; margin-top: 3px; }
  .edu-item { margin-bottom: 8px; }
  .edu-header { display: flex; justify-content: space-between; align-items: baseline; gap: 12px; }
  .edu-title { font-weight: 600; font-size: 11px; color: #333; }
  .edu-org { color: hsl(270,70%,45%); font-weight: 500; }
  .edu-year { font-size: 10px; color: #777; white-space: nowrap; }
  .edu-desc { font-size: 10px; color: #666; margin-top: 2px; line-height: 1.5; }
  .cert-item { display: flex; justify-content: space-between; align-items: baseline; gap: 12px; margin-bottom: 6px; }
  .cert-title { font-size: 10.5px; font-weight: 500; color: #333; }
  .cert-org { color: hsl(270,70%,45%); }
  .cert-year { font-size: 10px; color: #777; white-space: nowrap; }
  .skills-grid { display: flex; flex-wrap: wrap; gap: 6px 14px; }
  .skill-category { font-weight: 600; color: #333; font-size: 10.5px; }
  .skill-item { font-size: 10.5px; color: #444; }
  .avoid-break, .job, .project, .edu-item, .cert-item { break-inside: avoid; page-break-inside: avoid; }
</style>
<link rel="preconnect" href="https://fonts.googleapis.com"><link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@300..700&family=DM+Sans:ital,opsz,wght@0,9..40,100..1000;1,9..40,100..1000&display=swap" rel="stylesheet">
</head>
<body>
<div class="page">

  <div class="header avoid-break">
    <h1>Darshan Linge Gowda</h1>
    <div class="header-gradient"></div>
    <div class="contact-row">
      <span>darshanl1711@gmail.com</span>
      <span class="separator">|</span>
      <a href="https://www.linkedin.com/in/darshan-lingegowda-ai/">linkedin.com/in/darshan-lingegowda-ai</a>
      <span class="separator">|</span>
      <a href="https://github.com/DarshanLingegowda">github.com/DarshanLingegowda</a>
      <span class="separator">|</span>
      <span>Bangalore, India · Open to Remote (Worldwide)</span>
    </div>
  </div>

  <div class="section avoid-break">
    <div class="section-title">Professional Summary</div>
    <div class="summary-text">ML/LLMOps Engineer with hands-on experience building CI/CD pipelines for LLM deployment, containerised model infrastructure on GCP Cloud Run, and production observability systems using Datadog and Vertex AI. Delivered reusable CI/CD workflows for model training and deployment across multiple production systems — including a multi-agent orchestration platform and an LLM observability copilot. Proven impact: 80% accuracy improvement in production ML systems for automotive clients; infrastructure reliability-first approach from 3 years in regulated German enterprise environments. Experienced with MLflow for experiment tracking, vector databases for semantic retrieval, and Docker-based deployment at scale.</div>
  </div>

  <div class="section">
    <div class="section-title">Core Competencies</div>
    <div class="competencies-grid">
      <span class="competency-tag">CI/CD for LLM Training &amp; Deployment</span>
      <span class="competency-tag">Vector Database Management</span>
      <span class="competency-tag">Model Observability &amp; Monitoring</span>
      <span class="competency-tag">MLflow &amp; Model Registry</span>
      <span class="competency-tag">Docker &amp; Container Orchestration</span>
      <span class="competency-tag">GCP Cloud Run &amp; Cloud Build</span>
      <span class="competency-tag">LLMOps Infrastructure Design</span>
      <span class="competency-tag">Production LLM Reliability Engineering</span>
    </div>
  </div>

  <div class="section">
    <div class="section-title">Work Experience</div>

    <div class="job">
      <div class="job-header">
        <span class="job-company">PICA GmbH</span>
        <span class="job-period">2018 – 2021</span>
      </div>
      <div class="job-role">Software Developer · Munich, Germany</div>
      <ul>
        <li>Delivered <strong>80% accuracy improvement</strong> and 50% processing time reduction in ML-driven SAP invoice automation for BMW service centres — full pipeline ownership from data validation through production deployment and monitoring.</li>
        <li>Built automated CI/CD workflows and validation pipelines for ML models operating at high volume across regulated enterprise environments (healthcare, automotive, industrial).</li>
        <li>Implemented observability via structured error handling, logging, and automated failure detection — reducing recurring production failures; directly applicable to LLMOps monitoring.</li>
        <li>Operated across multi-tenant environments with strict data governance and compliance requirements.</li>
      </ul>
    </div>

    <div class="job">
      <div class="job-header">
        <span class="job-company">AAM IT GmbH</span>
        <span class="job-period">2018</span>
      </div>
      <div class="job-role">Software Developer (Contract) · Munich, Germany</div>
      <ul>
        <li>Designed and executed large-scale data migration with version-controlled schema evolution; built automated ETL pipelines via Jenkins and Bash — demonstrates model versioning and pipeline automation skills.</li>
        <li>Developed C# APIs for multi-interface integrations; provided cross-timezone production support ensuring pipeline stability.</li>
      </ul>
    </div>
  </div>

  <div class="section avoid-break">
    <div class="section-title">Selected AI/ML Projects</div>

    <div class="project">
      <span class="project-title">AI SRE Copilot — LLMOps Observability Infrastructure</span>
      <span class="project-badge">DEVPOST Hackathon · 2025</span>
      <div class="project-desc">Built LLMOps observability infrastructure integrating Datadog metrics/logs with Vertex AI (Gemini) for automated root-cause analysis. FastAPI service deployed serverlessly on Cloud Run with structured AI insights and real-time performance monitoring — directly models the observability tooling TRM Labs needs.</div>
      <div class="project-tech">Vertex AI · Gemini · Datadog · FastAPI · Cloud Run · LLMOps observability · Python</div>
    </div>

    <div class="project">
      <span class="project-title">Multi-Agent Productivity Assistant — CI/CD for Agentic Systems</span>
      <span class="project-badge">Gen AI Academy APAC · 2025</span>
      <div class="project-desc">Designed and deployed a multi-agent system with reusable CI/CD pipeline (Cloud Build + Docker) for containerised deployment on Cloud Run. AlloyDB AI for persistent memory. Demonstrates automated model deployment and versioning infrastructure at scale.</div>
      <div class="project-tech">Gemini 2.0 Flash · Google ADK · Cloud Build CI/CD · Docker · Cloud Run · AlloyDB AI</div>
    </div>

    <div class="project">
      <span class="project-title">AI Invoice Agent — Production LLM Reliability Engineering</span>
      <span class="project-badge">Gen AI Academy APAC · 2025</span>
      <div class="project-desc">Diagnosed and resolved critical production LLM reliability issue (model/API version mismatch) in a live multi-agent system — achieving stable low-latency responses at scale. <strong>Impact: ~70% faster workflows.</strong> Demonstrates the production reliability engineering mindset central to LLMOps.</div>
      <div class="project-tech">Gemini 2.5 Flash · LangChain · Docker · Cloud Run · CI/CD · LLM reliability</div>
    </div>

    <div class="project">
      <span class="project-title">Generative AI Search Agent — Vector DB &amp; Semantic Retrieval</span>
      <span class="project-badge">Personal Project · Ongoing</span>
      <div class="project-desc">LLMOps-focused semantic search agent: vector database management, embedding pipeline, prompt chaining, contextual memory. Containerised with Docker/CI-CD. <strong>~40% faster resolution vs. keyword search.</strong></div>
      <div class="project-tech">LangChain · Vector DB · Embeddings · OpenAI API · Docker · CI/CD · MLflow</div>
    </div>
  </div>

  <div class="section avoid-break">
    <div class="section-title">Education</div>
    <div class="edu-item">
      <div class="edu-header">
        <span class="edu-title">M.Sc. Information &amp; Communication Systems · <span class="edu-org">TU Chemnitz, Germany</span></span>
        <span class="edu-year">2017 – 2021</span>
      </div>
      <div class="edu-desc">Thesis: IEEE-published — ML pipeline for large-scale geospatial data processing adopted by a Galileo Map Service Provider for production use.</div>
    </div>
    <div class="edu-item">
      <div class="edu-header">
        <span class="edu-title">B.E. Electronics &amp; Communication Engineering · <span class="edu-org">Visvesvaraya Technological University</span></span>
        <span class="edu-year">2013 – 2017</span>
      </div>
    </div>
  </div>

  <div class="section avoid-break">
    <div class="section-title">Certifications</div>
    <div class="cert-item">
      <span class="cert-title"><span class="cert-org">DataCamp</span> — Associate AI Engineer for Developers (LLMOps, OpenAI API, LangChain, Pinecone, Embeddings, Production AI Systems)</span>
      <span class="cert-year">2025</span>
    </div>
    <div class="cert-item">
      <span class="cert-title"><span class="cert-org">MIT-IDSS</span> — Data Science &amp; Machine Learning · <span class="cert-org">Kaggle</span> Generative AI &amp; AI Agents Intensive</span>
      <span class="cert-year">2024–2025</span>
    </div>
  </div>

  <div class="section avoid-break">
    <div class="section-title">Skills</div>
    <div class="skills-grid">
      <span class="skill-item"><span class="skill-category">LLMOps / MLOps:</span> CI/CD for model deployment, MLflow, Langfuse-adjacent (Datadog), vector DBs, model registries, Docker, Kubernetes-ready containerisation</span>
      <span class="skill-item"><span class="skill-category">Cloud Infrastructure:</span> GCP, Vertex AI, Cloud Run, Cloud Build, serverless deployment, multi-region scaling</span>
      <span class="skill-item"><span class="skill-category">LLMs &amp; GenAI:</span> Gemini 2.5, OpenAI API, LangChain, LangGraph, RAG, prompt engineering, multi-agent orchestration</span>
      <span class="skill-item"><span class="skill-category">Languages:</span> Python, FastAPI, Node.js, SQL, Bash</span>
    </div>
  </div>

</div>
</body>
</html>
"@
 | Out-File -FilePath $htmlPath_trmlabs -Encoding utf8NoBOM
$pdfPath_trmlabs = Join-Path $outDir "cv-darshan-lingegowda-trmlabs-2026-04-13.pdf"
Write-Host "Generating cv-darshan-lingegowda-trmlabs-2026-04-13.pdf..."
& $chrome --headless=new --no-sandbox --disable-gpu --paper-width=8.5 --paper-height=11 `
    --print-to-pdf="$pdfPath_trmlabs" --print-to-pdf-no-header `
    "$htmlPath_trmlabs" 2>$null
if (Test-Path $pdfPath_trmlabs) { Write-Host "  OK: cv-darshan-lingegowda-trmlabs-2026-04-13.pdf" } else { Write-Warning "  FAILED: cv-darshan-lingegowda-trmlabs-2026-04-13.pdf" }

# --- QUORA ---
$htmlPath_quora = Join-Path $htmlDir "cv-darshan-lingegowda-quora.html"
@"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Darshan Linge Gowda — CV</title>
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  html { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  body { font-family: 'DM Sans', sans-serif; font-size: 11px; line-height: 1.5; color: #1a1a2e; background: #ffffff; }
  .page { width: 100%; max-width: 8.5in; margin: 0 auto; padding: 2px 0; }
  .header { margin-bottom: 20px; }
  .header h1 { font-family: 'Space Grotesk', sans-serif; font-size: 28px; font-weight: 700; color: #1a1a2e; letter-spacing: -0.02em; margin-bottom: 6px; }
  .header-gradient { height: 2px; background: linear-gradient(to right, hsl(187,74%,32%), hsl(270,70%,45%)); border-radius: 1px; margin-bottom: 10px; }
  .contact-row { display: flex; flex-wrap: wrap; gap: 8px 14px; font-size: 10.5px; color: #555; }
  .contact-row a { color: #555; text-decoration: none; }
  .contact-row .separator { color: #ccc; }
  .section { margin-bottom: 18px; }
  .section-title { font-family: 'Space Grotesk', sans-serif; font-size: 12px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.06em; color: hsl(187,74%,32%); border-bottom: 1.5px solid #e2e2e2; padding-bottom: 4px; margin-bottom: 10px; }
  .summary-text { font-size: 11px; line-height: 1.7; color: #2f2f2f; }
  a { white-space: nowrap; }
  .competencies-grid { display: flex; flex-wrap: wrap; gap: 8px; }
  .competency-tag { font-size: 10px; font-weight: 500; color: hsl(187,74%,28%); background: hsl(187,40%,95%); padding: 4px 10px; border-radius: 3px; border: 1px solid hsl(187,40%,88%); }
  .job { margin-bottom: 14px; }
  .job-header { display: flex; justify-content: space-between; align-items: baseline; gap: 12px; margin-bottom: 4px; }
  .job-company { font-family: 'Space Grotesk', sans-serif; font-size: 12.5px; font-weight: 600; color: hsl(270,70%,45%); }
  .job-period { font-size: 10.5px; color: #777; white-space: nowrap; }
  .job-role { font-size: 11px; font-weight: 600; color: #333; margin-bottom: 6px; }
  .job ul { padding-left: 18px; margin-top: 6px; }
  .job li { font-size: 10.5px; line-height: 1.6; color: #333; margin-bottom: 4px; }
  .job li strong { font-weight: 600; }
  .project { margin-bottom: 12px; }
  .project-title { font-family: 'Space Grotesk', sans-serif; font-size: 11.5px; font-weight: 600; color: hsl(270,70%,45%); }
  .project-badge { font-size: 9px; font-weight: 500; color: hsl(187,74%,32%); background: hsl(187,40%,95%); padding: 1px 6px; border-radius: 2px; margin-left: 6px; }
  .project-desc { font-size: 10.5px; color: #444; margin-top: 3px; line-height: 1.55; }
  .project-tech { font-size: 9.5px; color: #888; margin-top: 3px; }
  .edu-item { margin-bottom: 8px; }
  .edu-header { display: flex; justify-content: space-between; align-items: baseline; gap: 12px; }
  .edu-title { font-weight: 600; font-size: 11px; color: #333; }
  .edu-org { color: hsl(270,70%,45%); font-weight: 500; }
  .edu-year { font-size: 10px; color: #777; white-space: nowrap; }
  .edu-desc { font-size: 10px; color: #666; margin-top: 2px; line-height: 1.5; }
  .cert-item { display: flex; justify-content: space-between; align-items: baseline; gap: 12px; margin-bottom: 6px; }
  .cert-title { font-size: 10.5px; font-weight: 500; color: #333; }
  .cert-org { color: hsl(270,70%,45%); }
  .cert-year { font-size: 10px; color: #777; white-space: nowrap; }
  .skills-grid { display: flex; flex-wrap: wrap; gap: 6px 14px; }
  .skill-category { font-weight: 600; color: #333; font-size: 10.5px; }
  .skill-item { font-size: 10.5px; color: #444; }
  .avoid-break, .job, .project, .edu-item, .cert-item { break-inside: avoid; page-break-inside: avoid; }
</style>
<link rel="preconnect" href="https://fonts.googleapis.com"><link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@300..700&family=DM+Sans:ital,opsz,wght@0,9..40,100..1000;1,9..40,100..1000&display=swap" rel="stylesheet">
</head>
<body>
<div class="page">

  <div class="header avoid-break">
    <h1>Darshan Linge Gowda</h1>
    <div class="header-gradient"></div>
    <div class="contact-row">
      <span>darshanl1711@gmail.com</span>
      <span class="separator">|</span>
      <a href="https://www.linkedin.com/in/darshan-lingegowda-ai/">linkedin.com/in/darshan-lingegowda-ai</a>
      <span class="separator">|</span>
      <a href="https://github.com/DarshanLingegowda">github.com/DarshanLingegowda</a>
      <span class="separator">|</span>
      <span>Bangalore, India · Remote (Worldwide)</span>
    </div>
  </div>

  <div class="section avoid-break">
    <div class="section-title">Professional Summary</div>
    <div class="summary-text">AI Engineer with a track record of taking LLM-powered products from zero to production — building scalable multi-agent systems, LLM evaluation frameworks, and personalized AI-driven features on top of Gemini and OpenAI APIs. Deep understanding of LLM API integration, prompt engineering, and LLM workflow design, applied across conversational agents, RAG concierge systems, and multimodal observability tooling. Built two production agentic systems from scratch in 2025, both deployed on GCP Cloud Run with real users. Comfortable owning end-to-end ML systems across data pipelines, model training, and production integration. Available for async-first worldwide remote collaboration.</div>
  </div>

  <div class="section">
    <div class="section-title">Core Competencies</div>
    <div class="competencies-grid">
      <span class="competency-tag">LLM API Integration (Gemini, OpenAI)</span>
      <span class="competency-tag">Prompt Engineering &amp; LLM Evaluation</span>
      <span class="competency-tag">Zero-to-One Product Delivery</span>
      <span class="competency-tag">Scalable Multi-Agent LLM Workflows</span>
      <span class="competency-tag">RAG &amp; Semantic Search</span>
      <span class="competency-tag">End-to-End ML System Ownership</span>
      <span class="competency-tag">GCP / Vertex AI / Cloud Run</span>
      <span class="competency-tag">Production AI Deployment</span>
    </div>
  </div>

  <div class="section">
    <div class="section-title">Work Experience</div>

    <div class="job">
      <div class="job-header">
        <span class="job-company">PICA GmbH</span>
        <span class="job-period">2018 – 2021</span>
      </div>
      <div class="job-role">Software Developer · Munich, Germany</div>
      <ul>
        <li>Owned end-to-end ML systems for BMW service centres — from data pipelines and model design through production deployment and monitoring — delivering <strong>80% accuracy improvement</strong> in automated invoice generation.</li>
        <li>Built production backend systems across healthcare, mobility, and industrial domains; operated in regulated environments with strict data governance requirements.</li>
        <li>Strengthened production observability via structured error handling and logging; reduced recurring failures through systematic root-cause analysis.</li>
      </ul>
    </div>

    <div class="job">
      <div class="job-header">
        <span class="job-company">AAM IT GmbH / Vorwerk Switzerland</span>
        <span class="job-period">2018</span>
      </div>
      <div class="job-role">Software Developer (Contract) · Munich, Germany</div>
      <ul>
        <li>Built Python scripts for real-time data mining and ETL optimisation; developed C# APIs for multi-interface integrations. Cross-timezone support for production pipeline stability.</li>
      </ul>
    </div>
  </div>

  <div class="section avoid-break">
    <div class="section-title">Selected AI/ML Projects</div>

    <div class="project">
      <span class="project-title">Multi-Agent Productivity Assistant — LLM Workflow Design</span>
      <span class="project-badge">Gen AI Academy APAC · 2025</span>
      <div class="project-desc">Built from zero to production: Orchestrator Agent routes intent to four specialist sub-agents (Calendar, Task, Notes, Email) using Gemini 2.0 Flash + Google ADK. Real-time tool-use via Google Calendar MCP and Gmail MCP. Parallel agent execution resolves multi-step user requests in a single conversation turn. Demonstrates the AI-driven personalized feature development Poe requires.</div>
      <div class="project-tech">Gemini 2.0 Flash · Google ADK · LangGraph · Multi-agent LLM workflows · Cloud Run · AlloyDB AI</div>
    </div>

    <div class="project">
      <span class="project-title">AI Invoice Agent — Conversational AI from Zero to Production</span>
      <span class="project-badge">Gen AI Academy APAC · 2025</span>
      <div class="project-desc">Designed and shipped a conversational AI agent from scratch — Gemini 2.5 Flash reasoning layer powering invoice creation, payment tracking, and real-time analytics. Resolved LLM reliability issue in production (API version mismatch). <strong>Impact: ~70% faster workflows.</strong> Zero-to-one ownership across the full stack.</div>
      <div class="project-tech">Gemini 2.5 Flash · LangChain · React + Vite · Node.js · Docker · Cloud Run</div>
    </div>

    <div class="project">
      <span class="project-title">TravelAI — RAG Concierge with LLM Evaluation</span>
      <span class="project-badge">Kaggle Capstone · 2025</span>
      <div class="project-desc">End-to-end RAG system with semantic search, context tracking, LLM grounding, and controlled retrieval experiments — measuring relevance and consistency across retrieval configurations. Demonstrates LLM evaluation and scalable LLM workflow design.</div>
      <div class="project-tech">RAG · Vector DB · Embeddings · LLM grounding · Retrieval evaluation · Python</div>
    </div>

    <div class="project">
      <span class="project-title">AI SRE Copilot — Multimodal LLM Application</span>
      <span class="project-badge">DEVPOST Hackathon · 2025</span>
      <div class="project-desc">Multimodal AI application integrating Datadog with Vertex AI for root-cause analysis — demonstrates LLM workflow design for complex, data-rich inputs. FastAPI + Cloud Run deployment.</div>
      <div class="project-tech">Vertex AI · Gemini · Datadog · FastAPI · Cloud Run · Multimodal LLM</div>
    </div>
  </div>

  <div class="section avoid-break">
    <div class="section-title">Education</div>
    <div class="edu-item">
      <div class="edu-header">
        <span class="edu-title">M.Sc. Information &amp; Communication Systems · <span class="edu-org">TU Chemnitz, Germany</span></span>
        <span class="edu-year">2017 – 2021</span>
      </div>
      <div class="edu-desc">Thesis: IEEE-published — large-scale ML system for geospatial data processing adopted by a Galileo Map Service Provider for production use.</div>
    </div>
    <div class="edu-item">
      <div class="edu-header">
        <span class="edu-title">B.E. Electronics &amp; Communication Engineering · <span class="edu-org">Visvesvaraya Technological University</span></span>
        <span class="edu-year">2013 – 2017</span>
      </div>
    </div>
  </div>

  <div class="section avoid-break">
    <div class="section-title">Certifications</div>
    <div class="cert-item">
      <span class="cert-title"><span class="cert-org">DataCamp</span> — Associate AI Engineer for Developers (LLMOps, OpenAI API, LangChain, Pinecone, Embeddings, Production AI Systems)</span>
      <span class="cert-year">2025</span>
    </div>
    <div class="cert-item">
      <span class="cert-title"><span class="cert-org">MIT-IDSS</span> — Data Science &amp; Machine Learning · <span class="cert-org">Kaggle</span> Generative AI &amp; AI Agents Intensive</span>
      <span class="cert-year">2024–2025</span>
    </div>
  </div>

  <div class="section avoid-break">
    <div class="section-title">Skills</div>
    <div class="skills-grid">
      <span class="skill-item"><span class="skill-category">LLM APIs:</span> Gemini 2.5/2.0, OpenAI API, LangChain, LangGraph, Google ADK, prompt engineering, LLM evaluation, fine-tuning</span>
      <span class="skill-item"><span class="skill-category">AI Systems:</span> Multi-agent orchestration, RAG pipelines, vector DBs (Pinecone), embeddings, semantic search, ReAct pattern</span>
      <span class="skill-item"><span class="skill-category">Cloud / Infra:</span> GCP, Vertex AI, Cloud Run, Cloud Build CI/CD, Docker, MLflow, Datadog</span>
      <span class="skill-item"><span class="skill-category">Languages:</span> Python, FastAPI, Node.js, TypeScript, SQL, Bash</span>
    </div>
  </div>

</div>
</body>
</html>
"@
 | Out-File -FilePath $htmlPath_quora -Encoding utf8NoBOM
$pdfPath_quora = Join-Path $outDir "cv-darshan-lingegowda-quora-2026-04-13.pdf"
Write-Host "Generating cv-darshan-lingegowda-quora-2026-04-13.pdf..."
& $chrome --headless=new --no-sandbox --disable-gpu --paper-width=8.5 --paper-height=11 `
    --print-to-pdf="$pdfPath_quora" --print-to-pdf-no-header `
    "$htmlPath_quora" 2>$null
if (Test-Path $pdfPath_quora) { Write-Host "  OK: cv-darshan-lingegowda-quora-2026-04-13.pdf" } else { Write-Warning "  FAILED: cv-darshan-lingegowda-quora-2026-04-13.pdf" }

Write-Host ''
Write-Host 'Done. PDFs are in:' (Join-Path $repoRoot 'output')