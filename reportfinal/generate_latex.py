"""
LaTeX Report Generator – Nimma Yatri DevOps Project
Compiles with xelatex (MiKTeX) for real Times New Roman.
"""
import os, shutil, subprocess, sys

BASE = os.path.dirname(os.path.abspath(__file__))
PICS = os.path.join(BASE, "pics")
IMGS = os.path.join(BASE, "imgs")
TEX  = os.path.join(BASE, "report.tex")
PDF  = os.path.join(BASE, "DevOps_Project_Report_NimmaYatri.pdf")
XE   = r"C:\Users\Harsh\AppData\Local\Programs\MiKTeX\miktex\bin\x64\xelatex.exe"

os.makedirs(IMGS, exist_ok=True)

# ── copy & rename images ──────────────────────────────────────────────────────
IMGMAP = {
    "img01": "Screenshot 2026-05-18 183526.png",
    "img02": "Screenshot 2026-05-18 184426.png",
    "img03": "Screenshot 2026-05-18 185109.png",
    "img04": "Screenshot 2026-05-18 185812.png",
    "img05": "Screenshot 2026-05-18 185826.png",
    "img06": "Screenshot 2026-05-18 185844.png",
    "img07": "Screenshot 2026-05-18 190253.png",
    "img08": "Screenshot 2026-05-18 190306.png",
    "img09": "Screenshot 2026-05-18 190521.png",
    "img10": "Screenshot 2026-05-18 191813.png",
    "img11": "Screenshot 2026-05-18 191905.png",
    "img12": "Screenshot 2026-05-18 192316.png",
    "img13": "Screenshot 2026-05-18 202304.png",
    "img14": "Screenshot 2026-05-18 210619.png",
    "img15": "Screenshot 2026-05-18 212021.png",
    "img16": "Screenshot 2026-05-18 212040.png",
    "img17": "Screenshot 2026-05-18 212056.png",
    "img18": "Screenshot 2026-05-18 212111.png",
    "img19": "Screenshot 2026-05-18 212826.png",
    "img20": "Screenshot 2026-05-22 120051.png",
    "img21": "Screenshot 2026-05-22 120108.png",
    "img22": "Screenshot 2026-05-22 120330.png",
    "img23": "Screenshot 2026-05-22 120440.png",
    "img24": "Screenshot 2026-05-22 120841.png",
}
for k, v in IMGMAP.items():
    src = os.path.join(PICS, v)
    dst = os.path.join(IMGS, k + ".png")
    if os.path.exists(src):
        shutil.copy2(src, dst)

# ═════════════════════════════════════════════════════════════════════════════
# LaTeX content
# ═════════════════════════════════════════════════════════════════════════════

PREAMBLE = r"""
\documentclass[12pt,a4paper,oneside]{report}

%% ── Fonts ────────────────────────────────────────────────────────────────────
\usepackage{fontspec}
\setmainfont[
  BoldFont        = {Times New Roman Bold},
  ItalicFont      = {Times New Roman Italic},
  BoldItalicFont  = {Times New Roman Bold Italic}
]{Times New Roman}
\setmonofont[Scale=0.88]{Courier New}

%% ── Page layout ──────────────────────────────────────────────────────────────
\usepackage[
  a4paper,
  left=1.30in, right=1.05in, top=1.10in, bottom=1.05in,
  headheight=14pt, headsep=10pt
]{geometry}

%% ── Line spacing ─────────────────────────────────────────────────────────────
\usepackage{setspace}
\onehalfspacing

%% ── Chapter / Section headings ───────────────────────────────────────────────
\usepackage{titlesec}

\titleformat{\chapter}[display]
  {\centering\bfseries}
  {\fontsize{12}{14.4}\selectfont\MakeUppercase{Chapter \thechapter}}
  {4pt}
  {\centering\fontsize{16}{19.2}\selectfont\MakeUppercase}
  [\vspace{4pt}\rule{\textwidth}{0.6pt}]
\titlespacing*{\chapter}{0pt}{0pt}{18pt}

\titleformat{\section}
  {\bfseries\fontsize{14}{16.8}\selectfont}
  {\thesection}
  {0.55em}{}
\titlespacing*{\section}{0pt}{14pt}{5pt}

\titleformat{\subsection}
  {\bfseries\fontsize{12}{14.4}\selectfont}
  {\thesubsection}
  {0.5em}{}
\titlespacing*{\subsection}{0pt}{10pt}{3pt}

%% ── Table of Contents ────────────────────────────────────────────────────────
\usepackage{tocloft}
\renewcommand{\cfttoctitlefont}{\hfil\bfseries\fontsize{16}{19.2}\selectfont\MakeUppercase}
\renewcommand{\cftaftertoctitle}{\hfil\vspace{4pt}\\\rule{\textwidth}{0.6pt}}
\setlength{\cftbeforetoctitleskip}{0pt}
\setlength{\cftaftertoctitleskip}{14pt}
\renewcommand{\cftchapfont}{\bfseries}
\renewcommand{\cftchapdotsep}{\cftdotsep}
\renewcommand{\cftchappagefont}{\bfseries}
\setlength{\cftbeforechapskip}{7pt}
\setlength{\cftsecindent}{1.5em}
\renewcommand{\cftsecfont}{}
\renewcommand{\cftsecdotsep}{\cftdotsep}

%% ── Header / Footer ──────────────────────────────────────────────────────────
\usepackage{fancyhdr}
\pagestyle{fancy}
\fancyhf{}
\fancyhead[C]{\small\itshape
  Nimma Yatri\,---\,DevOps CI/CD Pipeline Report
  \;\textbar\;
  Ramaiah Institute of Technology}
\fancyfoot[C]{\small\thepage}
\renewcommand{\headrulewidth}{0.4pt}
\renewcommand{\footrulewidth}{0pt}
\fancypagestyle{plain}{
  \fancyhf{}
  \fancyfoot[C]{\small\thepage}
  \renewcommand{\headrulewidth}{0pt}
}

%% ── Colours ──────────────────────────────────────────────────────────────────
\usepackage[table,svgnames]{xcolor}
\definecolor{tblhdr}   {RGB}{31,  78, 121}
\definecolor{tblalt}   {RGB}{235,243,251}
\definecolor{codebg}   {RGB}{248,248,248}
\definecolor{codeborder}{RGB}{204,204,204}
\definecolor{imgbdr}   {RGB}{180,180,180}
\definecolor{navytext} {RGB}{31,  78, 121}

%% ── Tables ───────────────────────────────────────────────────────────────────
\usepackage{array, tabularx, booktabs, multirow, longtable}
\renewcommand{\arraystretch}{1.45}
\newcolumntype{L}[1]{>{\raggedright\arraybackslash}p{#1}}
\newcolumntype{C}[1]{>{\centering\arraybackslash}p{#1}}

%% ── Graphics ─────────────────────────────────────────────────────────────────
\usepackage{graphicx}
\usepackage{float}
\usepackage[labelfont=bf,
            font={small,it},
            justification=centering,
            skip=5pt]{caption}

%% ── Image-border frame ───────────────────────────────────────────────────────
\usepackage[framemethod=default]{mdframed}
\mdfdefinestyle{imgstyle}{
  linewidth=0.6pt,
  linecolor=imgbdr,
  backgroundcolor=gray!3,
  innertopmargin=5pt, innerbottommargin=5pt,
  innerleftmargin=5pt, innerrightmargin=5pt,
  skipabove=6pt, skipbelow=2pt
}
\newcommand{\reportfig}[2]{%
  \begin{figure}[H]
    \centering
    \begin{mdframed}[style=imgstyle]
      \centering
      \includegraphics[width=\linewidth,keepaspectratio]{imgs/#1}%
    \end{mdframed}
    \vspace{-2pt}
    \caption{#2}
  \end{figure}%
}

%% ── Code listings ────────────────────────────────────────────────────────────
\usepackage{listings}
\lstset{
  basicstyle        = \footnotesize\ttfamily,
  backgroundcolor   = \color{codebg},
  frame             = single,
  framesep          = 4pt,
  rulecolor         = \color{codeborder},
  breaklines        = true,
  keepspaces        = true,
  columns           = fullflexible,
  showstringspaces  = false,
  aboveskip         = 8pt,
  belowskip         = 6pt,
  xleftmargin       = 2pt,
  xrightmargin      = 2pt
}

%% ── Lists ────────────────────────────────────────────────────────────────────
\usepackage{enumitem}
\setlist[itemize]{
  leftmargin=1.6em, itemsep=2pt, parsep=0pt, topsep=3pt, partopsep=0pt,
  label=\textbullet
}
\setlist[enumerate]{
  leftmargin=1.8em, itemsep=2pt, parsep=0pt, topsep=3pt, partopsep=0pt
}

%% ── Para spacing ─────────────────────────────────────────────────────────────
\setlength{\parskip}{4pt}
\setlength{\parindent}{0pt}

%% ── Float spacing ────────────────────────────────────────────────────────────
\setlength{\textfloatsep}{10pt plus 2pt minus 2pt}
\setlength{\intextsep}{8pt plus 2pt minus 2pt}
\setlength{\floatsep}{8pt plus 2pt minus 2pt}

%% ── Page border (eso-pic + tikz) ─────────────────────────────────────────────
\usepackage{eso-pic}
\usepackage{tikz}
\AddToShipoutPictureBG{%
  \begin{tikzpicture}[remember picture,overlay]
    %% outer thick line
    \draw[line width=2.0pt, color={rgb,255:red,31;green,78;blue,121}]
      ([xshift=14pt,  yshift=-14pt] current page.north west)
      rectangle
      ([xshift=-14pt, yshift=14pt]  current page.south east);
    %% inner thin line
    \draw[line width=0.6pt, color={rgb,255:red,31;green,78;blue,121}]
      ([xshift=18pt,  yshift=-18pt] current page.north west)
      rectangle
      ([xshift=-18pt, yshift=18pt]  current page.south east);
  \end{tikzpicture}%
}

%% ── Overflow prevention ──────────────────────────────────────────────────────
\setlength{\emergencystretch}{4em}
\tolerance=2000
\hbadness=10000
\PassOptionsToPackage{hyphens}{url}

%% ── PDF metadata ─────────────────────────────────────────────────────────────
\usepackage[colorlinks=false,
            pdfborder={0 0 0},
            pdftitle={Nimma Yatri -- DevOps CI/CD Pipeline Report},
            pdfauthor={Harshendra M},
            bookmarks=true,
            bookmarksnumbered=true]{hyperref}
"""

# ─────────────────────────────────────────────────────────────────────────────
TITLE = r"""
\begin{document}

%% ── TITLE PAGE ───────────────────────────────────────────────────────────────
\begin{titlepage}
  \thispagestyle{empty}
  \centering\singlespacing

  \vspace*{10pt}
  {\bfseries\fontsize{18}{22}\selectfont RAMAIAH INSTITUTE OF TECHNOLOGY}\\[6pt]
  {\bfseries\fontsize{13}{16}\selectfont
    Department of Information Science and Engineering}\\[4pt]
  {\fontsize{12}{14}\selectfont Bengaluru\,--\,560\,054}\\[8pt]

  \rule{\textwidth}{1.8pt}\\[3pt]
  \rule{\textwidth}{0.5pt}\\[28pt]

  {\bfseries\fontsize{19}{23}\selectfont\color{navytext}
    DevOps Laboratory Project Report}\\[10pt]
  \rule{0.45\textwidth}{0.5pt}\\[10pt]

  {\bfseries\fontsize{15}{18}\selectfont\color{navytext}
    Nimma Yatri\,--\,Bengaluru Auto-Rickshaw Survival Tool}\\[8pt]

  {\fontsize{12}{15}\selectfont
    A Complete CI/CD Pipeline Implementation using\\[3pt]
    Jenkins\;\textbullet\;ESLint\;\textbullet\;Trivy\;\textbullet\;
    Docker\;\textbullet\;Vercel}\\[36pt]

  {\fontsize{12}{14}\selectfont Submitted by:}\\[14pt]

  \renewcommand{\arraystretch}{1.6}
  \begin{tabular}{r@{\hskip 6pt:\hskip 8pt}l}
    \textbf{Student Name} & Harshendra M \\
    \textbf{USN}          & 1MS22IS053 \\
    \textbf{Section}      & IS-A\enspace|\enspace 4th Semester \\
    \textbf{Subject}      & DevOps Laboratory (21ISL67) \\
  \end{tabular}
  \renewcommand{\arraystretch}{1.45}

  \vfill

  \rule{\textwidth}{0.5pt}\\[3pt]
  \rule{\textwidth}{1.8pt}\\[10pt]
  {\fontsize{12}{14}\selectfont Academic Year:\enspace 2025\,--\,2026}\\[4pt]
  {\fontsize{12}{14}\selectfont Date of Submission:\enspace May 2026}
\end{titlepage}

\clearpage
\setcounter{page}{1}

%% ── TABLE OF CONTENTS ────────────────────────────────────────────────────────
\pagestyle{plain}
\tableofcontents
\clearpage
\pagestyle{fancy}
"""

# ─────────────────────────────────────────────────────────────────────────────
CH1 = r"""
%% ══════════════════════════════════════════════════════════════════════════════
\chapter{Abstract}

This report presents the design, implementation, and evaluation of a complete
DevOps CI/CD pipeline built around \textbf{Nimma Yatri} --- a production-grade,
AI-powered Next.js web application that helps Bengaluru commuters navigate
auto-rickshaw fare disputes. The primary purpose of this project is to demonstrate
end-to-end automation of software delivery using industry-standard DevOps tools,
integrating source code management, automated testing, static code analysis,
vulnerability scanning, containerisation, and cloud deployment within a single,
cohesive pipeline.

The technology stack employed includes \textbf{GitHub} for source code management
and webhook-based change detection, \textbf{Jenkins 2.541.2} as the CI/CD automation
server, \textbf{ESLint} for static code quality analysis,
\textbf{Trivy 0.69.3} for both filesystem and Docker image vulnerability scanning,
\textbf{Docker Desktop 28.0.4} for containerising the Next.js application, and
\textbf{Docker Hub} as the container image registry. The deployment target is
\textbf{Vercel}, which provides seamless cloud hosting at the public URL
\texttt{https://nimmayatri.vercel.app}.

The CI/CD workflow is triggered automatically when a developer pushes code to the
\texttt{main} branch on GitHub. A GitHub webhook, tunnelled through \textbf{ngrok},
notifies Jenkins instantly, triggering the pipeline. Stages execute sequentially:
source checkout, dependency installation, ESLint code quality check, Next.js
production build, Trivy filesystem scan, Docker image build, Trivy image scan,
Docker Hub login, and Docker push. A secondary \emph{no-push} pipeline variant was
also configured for development testing.

The outcome confirmed successful automation: the Docker image
\mbox{\texttt{suicide768/nimmayatri-app:latest}} was published to Docker Hub,
both Trivy reports were archived as Jenkins build artefacts
(\texttt{trivy-fs-report.txt}: 31.66~KB; \texttt{trivy-image-report.txt}: 212.44~KB),
and the full pipeline completed in under 4 minutes from a GitHub push event ---
with \textbf{zero ESLint errors} reported.
The build was triggered automatically, confirmed by the Jenkins status
\emph{``Started by GitHub push by harshendram''}, demonstrating complete end-to-end
DevOps automation.
"""

# ─────────────────────────────────────────────────────────────────────────────
CH2 = r"""
\chapter{Objective}

\section{Why CI/CD Pipelines are Important}

Continuous Integration and Continuous Deployment (CI/CD) pipelines are the backbone
of modern software engineering. Traditional software delivery --- where integration,
testing, and deployment are manual, periodic activities --- leads to integration
failures, long release cycles, and unreliable deployments. CI/CD pipelines automate
the entire lifecycle from code commit to production, ensuring every change is built,
validated, and deployed consistently and quickly. This eliminates manual errors,
reduces risk, and provides immediate feedback to developers within minutes of a commit.

\section{Goals of Automation}

The specific automation goals defined for this project are:

\begin{itemize}
  \item Eliminate all manual build and deployment steps through a Jenkins pipeline
        triggered automatically by a GitHub push event.
  \item Enforce code quality standards at every commit using ESLint, blocking
        non-compliant code from advancing in the pipeline.
  \item Detect security vulnerabilities early by scanning both the npm dependency tree
        and the Docker container image with Trivy before code reaches production.
  \item Standardise deployment environments using Docker multi-stage builds,
        guaranteeing identical behaviour across development, CI, and production.
  \item Automate container image versioning (Jenkins build numbers) and publishing
        to Docker Hub for every successful build.
  \item Enable zero-touch pipeline execution through GitHub Webhooks, so a
        \texttt{git push} is the only developer action required.
\end{itemize}

\section{Benefits of DevOps Practices}

Adopting a DevOps culture and implementing CI/CD pipelines delivers the following
measurable benefits:

\begin{itemize}
  \item \textbf{Faster Time to Market:} Automated pipelines reduce deployment cycles
        from days to minutes, enabling frequent and reliable releases.
  \item \textbf{Improved Code Quality:} Automated linting and build checks catch
        errors before they reach production, reducing post-release defect rates.
  \item \textbf{Enhanced Security Posture:} Continuous vulnerability scanning with
        Trivy identifies known CVEs in dependencies and container images, enabling
        timely patching.
  \item \textbf{Environment Consistency:} Docker containerisation eliminates the
        ``works on my machine'' problem, ensuring the application runs identically
        everywhere.
  \item \textbf{Full Traceability:} Every pipeline execution is linked to a specific
        git commit, providing a complete audit trail from code change to deployed
        artefact.
  \item \textbf{Reduced Deployment Risk:} Consistent, automated build processes and
        Docker image versioning enable rapid rollback to any previous build.
\end{itemize}
"""

# ─────────────────────────────────────────────────────────────────────────────
CH3 = r"""
\chapter{Tools and Technologies Used}

The following table summarises all tools and technologies used in this project along
with their specific roles in the CI/CD pipeline.

\vspace{4pt}
\begin{longtable}{|L{0.32\textwidth}|L{0.63\textwidth}|}
\hline
\rowcolor{tblhdr}
\textcolor{white}{\textbf{Tool / Technology}} &
\textcolor{white}{\textbf{Purpose in the Pipeline}} \\
\hline
\endfirsthead
\hline
\rowcolor{tblhdr}
\textcolor{white}{\textbf{Tool / Technology}} &
\textcolor{white}{\textbf{Purpose in the Pipeline}} \\
\hline
\endhead
\hline
\endfoot
GitHub &
  Source code management, version control, and webhook-based change notification
  to Jenkins. \\
\rowcolor{tblalt}
Jenkins 2.541.2 &
  CI/CD automation server orchestrating all pipeline stages from checkout to
  Docker push. \\
ESLint (eslint-config-next) &
  Static code quality analysis enforcing \texttt{next/core-web-vitals} rules
  across the TypeScript codebase. \\
\rowcolor{tblalt}
Trivy 0.69.3 &
  Open-source vulnerability scanner for filesystem npm dependencies and Docker
  container images. \\
Docker Desktop 28.0.4 &
  Containerisation platform used to build a production-ready multi-stage Docker
  image. \\
\rowcolor{tblalt}
Docker Hub &
  Cloud container image registry storing versioned
  \texttt{suicide768/nimmayatri-app} images. \\
Node.js 18 (LTS) &
  JavaScript runtime for \texttt{npm install}, ESLint checks, and Next.js
  production builds. \\
\rowcolor{tblalt}
Next.js 14.2.21 &
  React-based full-stack web framework forming the application codebase. \\
ngrok &
  Secure tunnel exposing the local Jenkins server to GitHub for webhook
  delivery. \\
\rowcolor{tblalt}
Vercel &
  Cloud deployment platform hosting the live Next.js application at
  \texttt{nimmayatri.vercel.app}. \\
Git 2.44.0 &
  Distributed version control system for all source code tracking. \\
\rowcolor{tblalt}
npm &
  Node.js package manager for dependency installation and build script
  execution. \\
\hline
\end{longtable}
"""

# ─────────────────────────────────────────────────────────────────────────────
CH4 = r"""
\chapter{Project Tech Stack}

\section{Project Overview}

\textbf{Nimma Yatri} (meaning \emph{Your Journey} in Kannada) is an AI-powered web
application designed to help Bengaluru commuters navigate the notorious auto-rickshaw
fare dispute culture. The application provides:

\begin{itemize}
  \item AI-assisted fare calculation using Google Maps integration and a
        Scam-O-Meter\texttrademark\ gauge.
  \item Real-time voice and video negotiation support via the Gemini 2.5 Flash
        Multimodal Live API over WebSocket.
  \item Kannada phrase flashcards with Text-to-Speech audio playback for
        non-Kannada speakers.
  \item A Panic Button safety feature simulating a police call to deter scam drivers.
  \item An AI Chatbot (REST API) for text-based negotiation queries.
  \item Community intelligence from r/bangalore with location-tagged posts and
        curated tip cards.
  \item Multi-language support for 10 Indian languages via a React Context provider.
\end{itemize}

Built for the \textbf{Kiro Week 5 Challenge} (The Local Guide), the application is
deployed live at \texttt{https://nimmayatri.vercel.app} and serves as the subject
codebase for this DevOps pipeline implementation --- comprising 12 React components,
4 API routes, 4 custom hooks, and 6 utility libraries in over 8,500 lines of
TypeScript.

\section{Frontend Architecture}

The frontend is built on \textbf{Next.js 14 App Router} with TypeScript for type
safety. Key components include \texttt{FareCalculator} (Google Maps + Scam-O-Meter
animated gauge), \texttt{FloatingLiveAssistant} (WebSocket-based real-time AI),
\texttt{KannadaPhrases} (audio deck), \texttt{PanicButton}, \texttt{Chatbot}, and
\texttt{RedditPosts}. Styling uses Tailwind CSS 3.4 with Framer Motion animations.

\section{Backend and API Routes}

Next.js API routes serve as the backend layer. \texttt{/api/chatbot} integrates
Google Gemini 2.5 Flash with model fallback (Gemini 2.0). \texttt{/api/fare} computes
fares at Rs.\,30 base + Rs.\,15/km with a 1.5\texttimes\ night rate.
\texttt{/api/gemini-live} manages WebSocket configuration for multimodal AI sessions.
Google Maps Platform APIs (Places, Distance Matrix, Geocoding) handle all location
services.

\section{Technology Stack Summary}

\vspace{6pt}
\begin{center}
\begin{tabular}{|L{0.25\textwidth}|L{0.42\textwidth}|L{0.27\textwidth}|}
\hline
\rowcolor{tblhdr}
\textcolor{white}{\textbf{Layer}} &
\textcolor{white}{\textbf{Technology}} &
\textcolor{white}{\textbf{Version}} \\
\hline
Framework        & Next.js (App Router)                    & 14.2.21 \\
\rowcolor{tblalt}
Language         & TypeScript                              & 5.7.2 \\
Styling          & Tailwind CSS + Framer Motion            & 3.4.17 / 11.15.0 \\
\rowcolor{tblalt}
AI Engine        & Google Gemini Multimodal Live API       & 2.5 Flash \\
Maps             & Google Maps Platform                    & Places + Distance Matrix \\
\rowcolor{tblalt}
State            & React Context API                       & 18.3.1 \\
Icons            & Lucide React                            & 0.468.0 \\
\rowcolor{tblalt}
Runtime          & Node.js                                 & 18 LTS \\
Container        & Docker (multi-stage build)              & 28.0.4 \\
\rowcolor{tblalt}
Deployment       & Vercel -- Mumbai (bom1)                 & Latest \\
Performance      & Lighthouse Score                        & 94 / 100 \\
\hline
\end{tabular}
\end{center}

\section{Application Screenshots}

\reportfig{img02}{Nimma Yatri running on \texttt{localhost:3000} --- Fare Calculator
  (left panel) with Google Maps integration and Scam-O-Meter gauge, alongside the
  Kannada Phrases flashcard deck (right panel)}

\reportfig{img01}{VS Code terminal: \texttt{npm run lint} returns
  \checkmark\ \emph{No ESLint warnings or errors};
  \texttt{npm run build} completes the Next.js production build successfully;
  Docker version 28.0.4 confirmed}
"""

# ─────────────────────────────────────────────────────────────────────────────
CH5 = r"""
\chapter{GitHub Repository Details}

\section{Repository Information}

\vspace{6pt}
\begin{center}
\begin{tabular}{|L{0.32\textwidth}|L{0.63\textwidth}|}
\hline
\rowcolor{tblhdr}
\textcolor{white}{\textbf{Attribute}} &
\textcolor{white}{\textbf{Value}} \\
\hline
Repository Name      & \texttt{Devops\_Lab\_Assign} \\
\rowcolor{tblalt}
Owner                & harshendram (Harshendra M) \\
Clone URL            & {\small\texttt{https://github.com/harshendram/}\newline\texttt{Devops\_Lab\_Assign.git}} \\
\rowcolor{tblalt}
Primary Branch       & \texttt{main} \\
Visibility           & Public \\
\rowcolor{tblalt}
Language Composition & 97.1\% TypeScript, 2.5\% CSS \\
Total Deployments    & 6 (Vercel) \\
\rowcolor{tblalt}
Contributors         & 1 --- harshendram (Harshendra M) \\
\hline
\end{tabular}
\end{center}

\section{Repository Structure}

Key files and directories relevant to the DevOps pipeline:

\begin{lstlisting}
nimmayatri/
|-- app/                       # Next.js App Router pages + API routes
|   `-- api/                   # chatbot / fare / gemini / gemini-live
|-- components/                # 12 React UI components
|-- context/  hooks/  lib/     # State, custom hooks, utilities
|-- public/assets/             # Static assets
|-- Jenkinsfile                # Full CI/CD pipeline (with Docker push)
|-- Jenkinsfile-no-push        # Pipeline variant without Docker push
|-- Dockerfile                 # Multi-stage Docker build (3 stages)
|-- .eslintrc.json             # ESLint config (next/core-web-vitals)
|-- next.config.mjs            # Next.js config (standalone output)
|-- package.json               # Dependencies and build scripts
`-- README.md                  # Project documentation
\end{lstlisting}

\section{Branching Strategy}

The project follows a \textbf{single-branch strategy} with \texttt{main} as the
only active branch. Since this is a single-developer lab project, multi-branch GitFlow
is not required. All development commits and pipeline triggers operate on \texttt{main}.
Every push to \texttt{main} immediately activates the Jenkins pipeline via the
configured GitHub webhook.

\section{GitHub Repository Screenshot}

\reportfig{img24}{GitHub repository root view: commit
  \emph{``testing for vercel deployment''} by harshendram (cb350df), complete
  project file tree, 6 Vercel deployments, single contributor, 97.1\% TypeScript}

\section{GitHub Webhook Configuration}

A GitHub webhook was configured to notify Jenkins whenever code is pushed to the
repository. Because Jenkins runs on a local Windows machine without a public IP
address, \textbf{ngrok} was used to create a secure HTTPS tunnel, exposing the
Jenkins \texttt{/github-webhook/} endpoint to the internet.

\subsection{Setup Steps}

\begin{itemize}
  \item \textbf{Step 1:} Start ngrok: \texttt{ngrok http 8080} --- generates a public
        HTTPS URL (e.g., {\small\texttt{https://semiobliviously-unevaporative-neriah\allowbreak.ngrok-free.dev}}).
  \item \textbf{Step 2:} In GitHub: \textit{Settings} \textrightarrow\
        \textit{Webhooks} \textrightarrow\ \textit{Add webhook}.
  \item \textbf{Step 3:} Set Payload URL to
        \texttt{https://<ngrok-url>/github-webhook/}
  \item \textbf{Step 4:} Content type: \texttt{application/json};
        Trigger: \emph{Just the push event}.
  \item \textbf{Step 5:} Click \emph{Add webhook}. A green tick confirms successful
        delivery of the test ping.
  \item \textbf{Step 6:} In Jenkins, enable \emph{GitHub hook trigger for GITScm
        polling} in the job's Build Triggers section.
\end{itemize}

\reportfig{img14}{GitHub Settings \textrightarrow\ Webhooks \textrightarrow\ Manage
  Webhook: ngrok-generated Payload URL configured, Content type =
  \texttt{application/json}, SSL verification enabled, trigger = push event only}

\reportfig{img19}{ngrok session dashboard: Forwarding URL mapped to
  \texttt{http://127.0.0.1:8080}, Region: India, HTTP request log showing
  POST \texttt{/github-webhook/} \textrightarrow\ 200 OK --- webhook successfully
  delivered to Jenkins}
"""

# ─────────────────────────────────────────────────────────────────────────────
CH6 = r"""
\chapter{Jenkins Configuration}

\section{Jenkins Installation}

Jenkins 2.541.2 was installed on a Windows 11 machine using the official Jenkins LTS
Windows installer. The installer registers Jenkins as a Windows Service accessible at
\texttt{http://localhost:8080}. Initial setup involved unlocking Jenkins with the
auto-generated administrator password, installing suggested plugins, and creating an
admin user account.

\section{Required Plugins}

\vspace{6pt}
\begin{center}
\begin{tabular}{|L{0.36\textwidth}|L{0.59\textwidth}|}
\hline
\rowcolor{tblhdr}
\textcolor{white}{\textbf{Plugin}} &
\textcolor{white}{\textbf{Purpose}} \\
\hline
Pipeline &
  Enables declarative Jenkinsfile pipeline support. \\
\rowcolor{tblalt}
Git Plugin &
  Clones and polls GitHub repositories. \\
GitHub Integration Plugin &
  Enables webhook-triggered builds from GitHub push events. \\
\rowcolor{tblalt}
NodeJS Plugin &
  Provides Node.js 18 tool (configured as \texttt{Node18}) for npm commands. \\
Credentials Binding &
  Securely injects Docker Hub credentials into pipeline stages via
  \texttt{withCredentials}. \\
\rowcolor{tblalt}
Pipeline: Stage View &
  Visual stage-by-stage pipeline execution dashboard. \\
Workspace Cleanup &
  Cleans workspace between builds for consistent environments. \\
\hline
\end{tabular}
\end{center}

\section{Pipeline Jobs Created}

Two pipeline jobs were created in Jenkins:

\begin{itemize}
  \item \textbf{nimmayatri-pipeline:} Full CI/CD pipeline including Docker Hub login
        and image push. Uses \texttt{Jenkinsfile} from the repository root.
  \item \textbf{nimmayatri-pipeline-no-push:} Testing variant executing all stages
        except Docker Hub login and push. Uses \texttt{Jenkinsfile-no-push}.
\end{itemize}

Both jobs are configured as Pipeline type with: Definition = Pipeline script from SCM;
SCM = Git (\texttt{main} branch); GitHub hook trigger enabled.

\section{Jenkinsfile Explanation --- Full Pipeline}

The primary \texttt{Jenkinsfile} defines a declarative pipeline with 9 stages:

\begin{lstlisting}
pipeline {
    agent any
    tools { nodejs 'Node18' }
    environment {
        IMAGE_NAME = "suicide768/nimmayatri-app"
        IMAGE_TAG  = "${env.BUILD_NUMBER}"
    }
    stages {
        stage('Checkout')                { steps { git branch: 'main', url: '...' } }
        stage('Install Dependencies')    { steps { bat 'npm install' } }
        stage('ESLint Check')            { steps { bat 'npm run lint' } }
        stage('Build Next.js App')       { steps { bat 'npm run build' } }
        stage('Trivy Filesystem Scan')   { steps { bat 'trivy fs . > trivy-fs-report.txt' } }
        stage('Docker Build')            { steps { bat "docker build -t ..." } }
        stage('Trivy Docker Image Scan') { steps { bat "trivy image ... > trivy-image-report.txt" } }
        stage('Docker Hub Login')        { steps { withCredentials([...]) { ... } } }
        stage('Docker Push')             { steps { bat "docker push ..." } }
    }
    post {
        always  { archiveArtifacts artifacts: 'trivy-*.txt', allowEmptyArchive: true }
        success { echo 'Pipeline executed successfully!' }
        failure { echo 'Pipeline failed!' }
    }
}
\end{lstlisting}

\subsection{Stage-by-Stage Explanation}

\textbf{agent any:} Runs the pipeline on any available Jenkins build node --- in
this case, the local Windows controller node.

\textbf{tools \{ nodejs `Node18' \}:} Ensures Node.js 18 is on the PATH before
any stage executes, enabling all \texttt{npm} and \texttt{node} commands.

\textbf{environment block:} Defines \texttt{IMAGE\_NAME} =
\texttt{suicide768/nimmayatri-app} and \texttt{IMAGE\_TAG} = \texttt{BUILD\_NUMBER}.
Using the build number as a Docker tag provides automatic immutable versioning:
build~1 produces tag \texttt{:1}, build~2 produces tag \texttt{:2}, enabling
rollback to any previous build.

\textbf{Checkout:} Clones the \texttt{main} branch from GitHub. When triggered by
a webhook, Jenkins fetches the exact commit that triggered the build.

\textbf{Install Dependencies:} Runs \texttt{npm install}, installing all packages
from \texttt{package.json} including runtime and development dependencies.

\textbf{ESLint Check:} Executes \texttt{npm run lint} (next lint). If any ESLint
error is found, the stage fails immediately, halting the pipeline and preventing
non-compliant code from advancing.

\textbf{Build Next.js App:} Runs \texttt{npm run build} (next build), compiling
TypeScript, performing type checking, and generating the \texttt{.next/standalone}
output directory --- a self-contained Next.js server required for Docker deployment
(enabled by \texttt{output: 'standalone'} in \texttt{next.config.mjs}).

\textbf{Trivy Filesystem Scan:} Runs \texttt{trivy fs .} to scan
\texttt{package-lock.json} and \texttt{node\_modules} for CVEs using the NVD and
GitHub Advisory databases. Output saved to \texttt{trivy-fs-report.txt}.

\textbf{Docker Build:} Builds the multi-stage Docker image with two tags: the
build-number tag (e.g., \texttt{:2}) and \texttt{:latest}.

\textbf{Trivy Docker Image Scan:} Scans the built Docker image for OS-level (Alpine
Linux) and application-level CVEs. Output saved to \texttt{trivy-image-report.txt}.

\textbf{Docker Hub Login:} Uses \texttt{withCredentials} to inject Docker Hub
credentials (Jenkins Credentials ID: \texttt{dockerhub-creds}) as environment
variables, then performs a non-interactive login via \texttt{--password-stdin}
to avoid credentials appearing in logs.

\textbf{Docker Push:} Pushes both the build-number tag and \texttt{:latest} to
Docker Hub, making the containerised application publicly accessible.

\textbf{post/always block:} Archives all \texttt{trivy-*.txt} files as Jenkins build
artefacts regardless of pipeline outcome, providing permanent security scan records.

\section{Jenkinsfile-no-push Explanation}

The \texttt{Jenkinsfile-no-push} is identical to the full pipeline except the
\emph{Docker Hub Login} and \emph{Docker Push} stages are removed. This variant was
used during development to verify that build, lint, scan, and Docker build stages
work correctly without requiring Docker Hub credentials. It is suitable for pull
request validation or environments without registry access. Both variants archive
Trivy reports as build artefacts.

\section{Jenkins Pipeline Screenshots}

\reportfig{img18}{Jenkins Dashboard: \texttt{nimmayatri-pipeline} (Last Success:
  5\,min\,57\,sec, Build \#2) and \texttt{nimmayatri-pipeline-no-push} (Last Success:
  59\,min, Build \#2), both showing green success status}

\reportfig{img20}{Jenkins Builds history (user: Harshendra): all builds stable ---
  \texttt{nimmayatri-pipeline-no-push} \#1 and \#2, \texttt{nimmayatri-pipeline}
  \#1, \texttt{firstjob} \#1 --- all green checkmarks}

\reportfig{img21}{Jenkins project page for \texttt{nimmayatri-pipeline-no-push}:
  Last Successful Artefacts --- \texttt{trivy-fs-report.txt} (31.66\,KiB) and
  \texttt{trivy-image-report.txt} (212.44\,KiB); Build \#2 is Last stable, Last
  successful, and Last completed build (May 18, 2026 at 8:21\,PM)}

\reportfig{img17}{Jenkins Build \#2 status page: \emph{``Started by GitHub push by
  harshendram''} --- confirms automatic webhook-triggered execution; build took
  3\,min\,54\,sec}

\reportfig{img13}{Jenkins Build \#2 details: Build Artefacts section shows
  \texttt{trivy-fs-report.txt} (31.66\,KB) and \texttt{trivy-image-report.txt}
  (212.64\,KB), confirming both Trivy scans completed and archived}

\reportfig{img15}{Jenkins Console Output (top): \emph{``Started by GitHub push by
  harshendram''}, Jenkinsfile retrieved from git, pipeline stages beginning execution}

\reportfig{img16}{Jenkins Console Output (tail): Docker layers pushed, artefacts
  archived, \emph{``Pipeline executed successfully!''} --- \textbf{Finished: SUCCESS}}
"""

# ─────────────────────────────────────────────────────────────────────────────
CH7 = r"""
\chapter{Code Quality Analysis}

\section{Tool Used --- ESLint}

ESLint is the industry-standard JavaScript and TypeScript static analysis tool. For
this project, ESLint runs via Next.js's built-in lint integration (\texttt{next lint}),
which wraps ESLint with Next.js-specific rule sets. The configuration file
\texttt{.eslintrc.json} extends \textbf{next/core-web-vitals} --- the strictest
built-in Next.js preset, covering React best practices, Hooks rules, accessibility,
and Core Web Vitals performance patterns.

\section{Configuration}

\begin{lstlisting}
{
  "extends": "next/core-web-vitals",
  "rules": {
    "react/no-unescaped-entities": "off",
    "@next/next/no-img-element": "off"
  }
}
\end{lstlisting}

The \texttt{next/core-web-vitals} preset includes:

\begin{itemize}
  \item \textbf{eslint:recommended} --- Standard JS best practices
        (\texttt{no-unused-vars}, \texttt{no-undef}, etc.).
  \item \textbf{plugin:react/recommended} --- React-specific rules
        (prop-types, JSX correctness).
  \item \textbf{plugin:react-hooks/recommended} --- Rules of Hooks
        (\texttt{exhaustive-deps}, \texttt{rules-of-hooks}).
  \item \textbf{@next/next/core-web-vitals} --- Performance rules for
        Next.js-specific patterns.
  \item \textbf{plugin:jsx-a11y/recommended} --- Accessibility rules for JSX
        elements.
\end{itemize}

Two rules were explicitly disabled in this project:

\begin{itemize}
  \item \texttt{react/no-unescaped-entities} --- Disabled to accommodate Kannada
        text strings and apostrophes in JSX content that would otherwise generate
        false-positive errors.
  \item \texttt{@next/next/no-img-element} --- Disabled for specific components
        where the Next.js Image optimisation component is not applicable.
\end{itemize}

\section{ESLint in the Pipeline}

ESLint runs as \textbf{Stage 3} of the Jenkins pipeline (\texttt{npm run lint}) after
dependency installation but before the production build. If any ESLint error is
detected, the stage exits with a non-zero code, Jenkins marks the build as FAILED,
and all subsequent stages are skipped. This ensures no non-compliant code advances to
build, scan, or deployment stages.

\section{Analysis Results}

The ESLint analysis of the complete Nimma Yatri codebase returned \textbf{zero
warnings and zero errors} across all 12 React components, 4 API routes, 4 custom
hooks, and 6 utility libraries. Key validations confirmed by the clean lint result:

\begin{itemize}
  \item No unused variables or imports across the TypeScript codebase.
  \item Correct React Hooks usage --- no conditional hook calls, no missing
        dependency arrays.
  \item No deprecated Next.js patterns or performance anti-patterns.
  \item No JSX accessibility violations.
\end{itemize}

\reportfig{img01}{VS Code terminal: \texttt{npm run lint} returns
  \emph{\checkmark\ No ESLint warnings or errors}; \texttt{npm run build} confirms
  a clean Next.js production build; Docker version 28.0.4 verified}

\section{Issues Fixed During Development}

\begin{itemize}
  \item \textbf{Unescaped Kannada entities:} \texttt{react/no-unescaped-entities}
        errors caused by Kannada text in JSX. Resolution: rule disabled in
        \texttt{.eslintrc.json}.
  \item \textbf{Standard \texttt{<img>} elements:}
        \texttt{@next/next/no-img-element} warnings on specific components.
        Resolution: rule disabled where Next.js Image was not applicable.
  \item \textbf{Incomplete Hook dependency arrays:}
        \texttt{react-hooks/exhaustive-deps} warnings. Resolution: dependency
        arrays updated to include all referenced variables.
\end{itemize}
"""

# ─────────────────────────────────────────────────────────────────────────────
CH8 = r"""
\chapter{Dependency and Vulnerability Scanning}

\section{Tool Used --- Trivy}

Trivy (version 0.69.3) is an open-source comprehensive vulnerability scanner by
Aqua Security. It is the industry's most widely adopted container and filesystem
security scanner, detecting CVEs in OS packages, language-specific dependencies
(npm, pip, etc.), IaC misconfigurations, and secret leaks. Trivy was chosen for
this project due to its seamless CI pipeline integration, broad vulnerability
database coverage (NVD, GitHub Advisory Database, OS vendor advisories), and
ability to scan both project filesystems and Docker images.

\section{Scan Process}

Two Trivy scans are performed sequentially in the Jenkins pipeline:

\begin{itemize}
  \item \textbf{Stage 5 --- Filesystem Scan} (\texttt{trivy fs .}): Scans the
        project directory focusing on \texttt{package-lock.json} for npm dependency
        CVEs. Runs after the Next.js build while \texttt{node\_modules} are present.
        Output saved to \texttt{trivy-fs-report.txt}.
  \item \textbf{Stage 7 --- Docker Image Scan} (\texttt{trivy image <image:tag>}):
        Scans the built Docker image for vulnerabilities in Alpine Linux OS packages
        and Node.js runtime packages within the container. Output saved to
        \texttt{trivy-image-report.txt}.
\end{itemize}

\section{Filesystem Scan Results}

Trivy filesystem scan of \texttt{package-lock.json} identified \textbf{28
vulnerabilities}:

\vspace{6pt}
\begin{center}
\begin{tabular}{|C{0.3\textwidth}|C{0.25\textwidth}|}
\hline
\rowcolor{tblhdr}
\textcolor{white}{\textbf{Severity}} &
\textcolor{white}{\textbf{Count}} \\
\hline
CRITICAL & 1 \\
\rowcolor{tblalt}
HIGH     & 9 \\
MEDIUM   & 14 \\
\rowcolor{tblalt}
LOW      & 4 \\
UNKNOWN  & 0 \\
\rowcolor{tblalt}
\textbf{TOTAL} & \textbf{28} \\
\hline
\end{tabular}
\end{center}
\vspace{6pt}

Key findings (all in the \texttt{next} / Next.js 14.2.21 package):

\begin{itemize}
  \item \textbf{CVE-2025-29927 [CRITICAL]:} Authorization Bypass in Next.js
        Middleware. Fixed in v14.2.25. Highest priority remediation.
  \item \textbf{CVE-2026-44573 [HIGH]:} Middleware/Proxy bypass in Pages Router
        applications.
  \item \textbf{CVE-2026-44578 [HIGH]:} Server-side request forgery via WebSocket
        upgrades.
  \item \textbf{GHSA-5j59-xgg2-r9c4 [HIGH]:} Denial of Service via Server
        Components (incomplete fix).
  \item \textbf{GHSA-8h8q-6873-q5fj [HIGH]:} Denial of Service via Server
        Components.
  \item \textbf{CVE-2025-55173 [MEDIUM]:} Content Injection vulnerability in Image
        Optimisation.
\end{itemize}

\reportfig{img04}{Trivy filesystem scan summary: \texttt{package-lock.json} ---
  Total 28 vulnerabilities (0 UNKNOWN, 4 LOW, 14 MEDIUM, 9 HIGH, 1 CRITICAL)}

\reportfig{img05}{Trivy CVE detail table: CVE-2025-29927 (CRITICAL),
  CVE-2026-44573 (HIGH), CVE-2026-44578 (HIGH) with descriptions and affected
  Next.js versions}

\reportfig{img06}{Trivy scan continued: additional HIGH and MEDIUM severity CVEs
  in the Next.js dependency with full CVE identifiers and advisory links}

\section{Docker Image Scan Results}

The Trivy Docker image scan of \texttt{nimmayatri-app} (based on
\texttt{node:18-alpine} / Alpine 3.21.3) identified approximately \textbf{58
vulnerabilities}, including CVEs from Alpine Linux OS packages (BusyBox) in addition
to npm packages:

\vspace{6pt}
\begin{center}
\begin{tabular}{|C{0.3\textwidth}|C{0.25\textwidth}|}
\hline
\rowcolor{tblhdr}
\textcolor{white}{\textbf{Severity}} &
\textcolor{white}{\textbf{Count}} \\
\hline
CRITICAL & 4 \\
\rowcolor{tblalt}
HIGH     & 15 \\
MEDIUM   & 26 \\
\rowcolor{tblalt}
LOW      & 5 \\
UNKNOWN  & 0 \\
\rowcolor{tblalt}
\textbf{TOTAL} & \textbf{\textasciitilde 58} \\
\hline
\end{tabular}
\end{center}
\vspace{6pt}

Additional CVEs beyond the filesystem scan (from Alpine Linux packages):

\begin{itemize}
  \item \textbf{CVE-2024-58251 [MEDIUM]:} BusyBox --- local users can launch
        processes with altered network namespaces.
  \item \textbf{CVE-2025-46394 [LOW]:} BusyBox TAR archive filename handling
        vulnerability.
  \item Multiple CVEs in \texttt{node-pkg} packages inherited from the Alpine
        Node.js 18 runtime.
\end{itemize}

\reportfig{img07}{Trivy filesystem scan: detailed \texttt{node\_modules} package
  vulnerability list}

\reportfig{img08}{Trivy Docker image scan summary for \texttt{nimmayatri-app}
  (Alpine 3.21.3): Total \textasciitilde 58 vulnerabilities including 4 CRITICAL,
  BusyBox and npm package findings}

\reportfig{img09}{Trivy image scan output: OS family \texttt{alpine} detected,
  scanning process, CVE-2026-31082 and related entries with severity levels}

\section{Mitigation Steps}

\begin{itemize}
  \item \textbf{Upgrade Next.js} to \(\geq\)\,14.2.25 to resolve CVE-2025-29927
        (CRITICAL). This is the highest priority remediation action.
  \item \textbf{Update the Docker base image} from \texttt{node:18-alpine} to the
        latest patched version to resolve BusyBox CVEs.
  \item \textbf{Run \texttt{npm audit fix}} after the Next.js upgrade to resolve
        transitive dependency vulnerabilities automatically.
  \item \textbf{Trivy reports} are archived as Jenkins build artefacts for audit
        trail and periodic security review.
  \item \textbf{No secrets were detected} by Trivy's secret scanner, confirming
        proper use of \texttt{.env.local} (gitignored) and Vercel environment
        variable configuration.
\end{itemize}
"""

# ─────────────────────────────────────────────────────────────────────────────
CH9 = r"""
\chapter{Docker Hub Integration}

\section{Dockerfile Explanation}

The \texttt{Dockerfile} uses a \textbf{three-stage multi-stage build} to produce a
minimal, secure, production-ready image. Multi-stage builds keep build tools out of
the final image, significantly reducing image size and attack surface.

\begin{lstlisting}
# Stage 1: deps -- Install all npm dependencies
FROM node:18-alpine AS deps
WORKDIR /app
COPY package.json package-lock.json* ./
RUN npm install

# Stage 2: builder -- Build the Next.js application
FROM node:18-alpine AS builder
WORKDIR /app
COPY --from=deps /app/node_modules ./node_modules
COPY . .
RUN npm run build

# Stage 3: runner -- Minimal production image
FROM node:18-alpine AS runner
WORKDIR /app
ENV NODE_ENV=production

RUN addgroup -S nodejs && adduser -S nextjs -G nodejs

COPY --from=builder /app/next.config.mjs ./
COPY --from=builder /app/public ./public
COPY --from=builder --chown=nextjs:nodejs /app/.next/standalone ./
COPY --from=builder --chown=nextjs:nodejs /app/.next/static ./.next/static

USER nextjs
EXPOSE 3000
ENV PORT=3000
CMD ["node", "server.js"]
\end{lstlisting}

\subsection{Stage 1 --- deps}

Installs all npm dependencies. Using \texttt{node:18-alpine}
($\approx$\,5\,MB vs $\approx$\,900\,MB for the Debian image) minimises the base
size. Copying only \texttt{package.json} and \texttt{package-lock.json} first
leverages Docker layer caching --- if dependencies haven't changed, this expensive
layer is reused without reinstalling.

\subsection{Stage 2 --- builder}

Copies installed \texttt{node\_modules} and the full source code, then runs
\texttt{npm run build} (next build). The \texttt{next.config.mjs} specifies
\texttt{output: 'standalone'}, generating a \texttt{.next/standalone} directory --- a
self-contained Node.js server with only the files needed to run the app, excluding
the full \texttt{node\_modules} directory.

\subsection{Stage 3 --- runner}

The minimal production image. Security features: \textbf{(1)} A non-root user
\texttt{nextjs} is created and used to run the process, following the principle of
least privilege. \textbf{(2)} Only the standalone server, public assets, and static
files are copied --- full source code and \texttt{node\_modules} are
\emph{not} in the final image. \textbf{(3)} The container listens on port 3000 and
starts via \texttt{node server.js} (the standalone Next.js server).

\section{Docker Hub Repository Details}

\vspace{6pt}
\begin{center}
\begin{tabular}{|L{0.34\textwidth}|L{0.61\textwidth}|}
\hline
\rowcolor{tblhdr}
\textcolor{white}{\textbf{Attribute}} &
\textcolor{white}{\textbf{Details}} \\
\hline
Docker Hub Username  & \texttt{suicide768} \\
\rowcolor{tblalt}
Repository Name      & \texttt{suicide768/nimmayatri-app} \\
Visibility           & Public \\
\rowcolor{tblalt}
Image Tags           & Build number (\texttt{:1}, \texttt{:2}, \ldots) +
                       \texttt{:latest} always updated \\
Base Image           & \texttt{node:18-alpine} \\
\rowcolor{tblalt}
Authentication       & Jenkins Credentials (ID: \texttt{dockerhub-creds})
                       via \texttt{withCredentials} \\
Final Image Digest   & \texttt{sha256:ce8053b55a74bb3f85\ldots} \\
\hline
\end{tabular}
\end{center}

\section{Build and Push Commands}

\begin{lstlisting}
# Build with build-number tag AND latest tag
docker build -t suicide768/nimmayatri-app:2 -t suicide768/nimmayatri-app:latest .

# Non-interactive Docker Hub login (password via stdin -- never exposed in logs)
echo %DOCKER_PASS% | docker login -u %DOCKER_USER% --password-stdin

# Push both tags to Docker Hub
docker push suicide768/nimmayatri-app:2
docker push suicide768/nimmayatri-app:latest
\end{lstlisting}

\section{Image Tags and Versioning}

\begin{itemize}
  \item \textbf{Build number tag (\texttt{:1}, \texttt{:2}, \texttt{:3}\ldots):}
        Immutable, monotonically increasing. Enables rollback to any specific build
        (\texttt{docker pull suicide768/nimmayatri-app:1}).
  \item \textbf{\texttt{:latest} tag:} Always points to the most recently pushed
        image. Used as the default pull target when no tag is specified.
\end{itemize}

\section{Docker Hub Screenshots}

\reportfig{img03}{VS Code terminal: Docker multi-stage build (18 build steps
  completed); \texttt{docker run -p 3000:3000 nimmayatri-app} confirms the
  container starts successfully on port 3000}

\reportfig{img10}{Terminal: \texttt{docker push suicide768/nimmayatri-app} --- all
  image layers pushed; final digest \texttt{sha256:ce8053b55a74\ldots} confirms
  successful publication to Docker Hub}

\reportfig{img11}{Docker Hub web interface: \texttt{suicide768/nimmayatri-app}
  repository visible, last pushed 1 minute ago, Visibility: Public}

\reportfig{img12}{Terminal: \texttt{docker pull suicide768/nimmayatri-app} ---
  image pulled from Docker Hub; \emph{``Status: Image is up to date''} with digest
  verification confirms accessibility}
"""

# ─────────────────────────────────────────────────────────────────────────────
CH10 = r"""
\chapter{Deployment to Cloud Platform}

\section{Platform --- Vercel}

Vercel is the cloud deployment platform used for the Nimma Yatri application. As the
creators of Next.js, Vercel provides first-class Next.js support with zero-configuration
setup, automatic HTTPS, global CDN, and seamless GitHub integration. The application
is deployed in the \textbf{Mumbai (bom1)} region to minimise latency for Indian users.

\section{Deployment Steps}

\begin{itemize}
  \item \textbf{Step 1:} Connect GitHub repository to Vercel by authorising the
        Vercel GitHub app on the harshendram account.
  \item \textbf{Step 2:} Import the \texttt{Devops\_Lab\_Assign} repository into
        Vercel's dashboard.
  \item \textbf{Step 3:} Vercel auto-detects Next.js and configures build settings
        (\texttt{npm run build}, output: \texttt{.next}).
  \item \textbf{Step 4:} Add environment variables in Vercel project settings:
        \texttt{NEXT\_PUBLIC\_GOOGLE\_MAPS\_API\_KEY},
        \texttt{NEXT\_PUBLIC\_GEMINI\_API\_KEY}, \texttt{GEMINI\_API\_KEY}.
  \item \textbf{Step 5:} Click Deploy. Vercel builds and deploys to the global CDN
        with an automatic HTTPS certificate.
  \item \textbf{Step 6:} Every subsequent push to \texttt{main} triggers automatic
        Vercel re-deployment.
\end{itemize}

\section{Environment Configuration}

\vspace{6pt}
\begin{center}
\begin{tabular}{|L{0.46\textwidth}|L{0.49\textwidth}|}
\hline
\rowcolor{tblhdr}
\textcolor{white}{\textbf{Environment Variable}} &
\textcolor{white}{\textbf{Purpose}} \\
\hline
{\small\texttt{NEXT\_PUBLIC\_GOOGLE\_MAPS\_API\_KEY}} &
  Google Maps Platform --- Places, Distance Matrix, Geocoding APIs (client-side). \\
\rowcolor{tblalt}
{\small\texttt{NEXT\_PUBLIC\_GEMINI\_API\_KEY}} &
  Gemini API for client-side features (Gemini Live WebSocket). \\
{\small\texttt{GEMINI\_API\_KEY}} &
  Gemini API for server-side API routes (chatbot endpoint). \\
\hline
\end{tabular}
\end{center}

\section{Deployment Details}

\vspace{6pt}
\begin{center}
\begin{tabular}{|L{0.34\textwidth}|L{0.61\textwidth}|}
\hline
\rowcolor{tblhdr}
\textcolor{white}{\textbf{Attribute}} &
\textcolor{white}{\textbf{Value}} \\
\hline
Platform           & Vercel \\
\rowcolor{tblalt}
Production URL     & \texttt{https://nimmayatri.vercel.app} \\
Region             & Mumbai, India (bom1) \\
\rowcolor{tblalt}
Framework          & Next.js 14 (auto-detected) \\
Build Command      & \texttt{npm run build} \\
\rowcolor{tblalt}
HTTPS              & Automatic (Let's Encrypt) \\
CDN                & Vercel Edge Network (Global) \\
\rowcolor{tblalt}
Lighthouse Score   & 94 / 100 \\
First Load JS      & 299\,KB (vendor chunk: 270\,KB) \\
\rowcolor{tblalt}
Total Static Assets& 141.77\,MB (21 files) \\
Build Duration     & 52 seconds \\
\rowcolor{tblalt}
Total Deployments  & 6 \\
Trigger            & GitHub push to \texttt{main} branch \\
\hline
\end{tabular}
\end{center}

\section{Deployment Screenshots}

\reportfig{img22}{Vercel dashboard for the \texttt{nimmayatri} project: Deployment
  Status = \textbf{Ready} (green), Production Domain =
  \texttt{nimmayatri.vercel.app}, created May 19 2026 by harshendram from commit
  \texttt{cb350df}; Firewall: Active, 0 block events; Observability: 0\% Error Rate}

\reportfig{img23}{Vercel build logs: Next.js 14.2.21 production compilation output
  --- routing table (\texttt{/} = 21\,kB, \texttt{/\_not-found} = 184\,B, 4
  serverless API functions), First Load JS = 299\,kB, build time 52 seconds,
  141.77\,MB assets deployed successfully}
"""

# ─────────────────────────────────────────────────────────────────────────────
CH11 = r"""
\chapter{Automation using Webhooks}

\section{Overview}

Webhooks enable GitHub to actively push HTTP notifications to Jenkins the instant a
push event occurs, rather than Jenkins polling GitHub at intervals. This results in
near-instantaneous pipeline triggering --- typically within 2--5 seconds of a
\texttt{git push}. The ngrok tunnel bridges GitHub's internet-accessible webhook
endpoint to the local Jenkins server running on a Windows machine without a public
IP address.

\section{End-to-End Automation Flow}

\vspace{6pt}
\begin{center}
\renewcommand{\arraystretch}{1.5}
\begin{tabular}{|C{0.06\textwidth}|L{0.52\textwidth}|L{0.36\textwidth}|}
\hline
\rowcolor{tblhdr}
\textcolor{white}{\textbf{Step}} &
\textcolor{white}{\textbf{Action}} &
\textcolor{white}{\textbf{Component}} \\
\hline
1 & Developer runs: \texttt{git push origin main}
  & Git / Developer Machine \\
\rowcolor{tblalt}
2 & GitHub receives the push and identifies all webhooks configured for the repository
  & GitHub \\
3 & GitHub sends HTTP POST to the Payload URL with JSON body containing commit info,
    author, branch, and repository details
  & GitHub \textrightarrow\ ngrok \\
\rowcolor{tblalt}
4 & ngrok receives the HTTPS request and forwards it as HTTP to
    \texttt{localhost:8080/github-webhook/}
  & ngrok Tunnel \\
5 & Jenkins GitHub Integration Plugin validates the payload and identifies matching
    pipeline jobs
  & Jenkins \\
\rowcolor{tblalt}
6 & Jenkins triggers \texttt{nimmayatri-pipeline} (and \texttt{nimmayatri-pipeline-no-push}
    if configured)
  & Jenkins \\
7 & All 9 pipeline stages execute sequentially from Checkout to Docker Push
  & Jenkins Pipeline \\
\rowcolor{tblalt}
8 & Trivy reports archived; Docker image pushed to Docker Hub;
    build status set to SUCCESS
  & Jenkins / Docker Hub \\
\hline
\end{tabular}
\end{center}
\renewcommand{\arraystretch}{1.45}

\section{ngrok Configuration Details}

\vspace{6pt}
\begin{center}
\begin{tabular}{|L{0.34\textwidth}|L{0.61\textwidth}|}
\hline
\rowcolor{tblhdr}
\textcolor{white}{\textbf{Attribute}} &
\textcolor{white}{\textbf{Value}} \\
\hline
ngrok Account     & \texttt{lordhshiva@gmail.com} (Free Plan) \\
\rowcolor{tblalt}
Forwarding URL    & {\small\texttt{https://semiobliviously-unevaporative-}\newline\texttt{neriah.ngrok-free.dev}} \\
Local Target      & \texttt{http://127.0.0.1:8080} \\
\rowcolor{tblalt}
Region            & India (in) \\
Latency           & 16\,ms \\
\rowcolor{tblalt}
Webhook Endpoint  & \texttt{https://<ngrok-url>/github-webhook/} \\
HTTP Response     & 200 OK (Jenkins acknowledged delivery) \\
\hline
\end{tabular}
\end{center}

\section{Webhook Screenshots}

\reportfig{img14}{GitHub Settings \textrightarrow\ Webhooks \textrightarrow\ Manage
  Webhook: Payload URL = ngrok HTTPS URL, Content type =
  \texttt{application/json}, SSL verification enabled, trigger = Just the push event}

\reportfig{img19}{ngrok session: Forwarding URL mapped to
  \texttt{http://127.0.0.1:8080}, Region: India, HTTP request log showing
  POST \texttt{/github-webhook/} \textrightarrow\ 200 OK --- webhook successfully
  delivered to Jenkins}

\reportfig{img17}{Jenkins Build \#2 status: \emph{``Started by GitHub push by
  harshendram''} --- definitive proof of automatic webhook-triggered pipeline
  execution (manual triggers show \emph{``Started by user''} instead)}

\reportfig{img15}{Jenkins Console Output: \emph{``Started by GitHub push by
  harshendram''}; Jenkinsfile obtained from git at the triggering commit's HEAD;
  pipeline stages beginning execution}

\section{Verification of Automatic Execution}

Automatic webhook-based execution was verified through two independent sources:

\begin{itemize}
  \item \textbf{Jenkins build metadata:} The build status page for Build \#2 shows
        \emph{``Started by GitHub push by harshendram''}. This text is set
        exclusively by the Jenkins GitHub Integration Plugin upon receipt of a valid
        webhook payload --- it cannot appear for manual builds.
  \item \textbf{ngrok request log:} The ngrok dashboard at
        \texttt{http://127.0.0.1:4040} records
        \texttt{POST /github-webhook/} with HTTP 200 response, confirming GitHub
        delivered the payload and Jenkins acknowledged receipt.
\end{itemize}
"""

# ─────────────────────────────────────────────────────────────────────────────
CH12 = r"""
\chapter{Conclusion}

\section{Summary}

This project successfully designed and implemented a complete, production-quality
DevOps CI/CD pipeline for the Nimma Yatri Next.js application. The pipeline
integrates \textbf{GitHub} source control, \textbf{Jenkins} automation,
\textbf{ESLint} code quality enforcement, \textbf{Trivy} security scanning,
\textbf{Docker} multi-stage containerisation, \textbf{Docker Hub} image registry,
\textbf{Vercel} cloud deployment, and \textbf{GitHub webhook-based} automatic
triggering into a cohesive, fully automated software delivery system.

Both pipeline variants --- the full pipeline with Docker push (Build \#2: 3\,min
54\,sec) and the no-push testing variant (Build \#2: 1\,min 19\,sec) --- completed
successfully with green status. The Docker image \texttt{suicide768/nimmayatri-app:latest}
was published to Docker Hub, both Trivy reports were archived as Jenkins artefacts,
and webhook-based automation was confirmed by the \emph{``Started by GitHub push by
harshendram''} build trigger message.

\section{Lessons Learnt}

\begin{itemize}
  \item \textbf{Shift quality left:} Placing ESLint checks and Trivy scans early in
        the pipeline ensures issues are caught before they can reach production,
        minimising the cost of defect detection.

  \item \textbf{Security scanning reveals hidden vulnerabilities:} The Trivy scan
        uncovered 28 CVEs in \texttt{package-lock.json} and $\approx$\,58 in the
        Docker image, including a CRITICAL Next.js authorization bypass
        (CVE-2025-29927) that was not apparent from the application's functionality
        alone.

  \item \textbf{Multi-stage Dockerfiles are essential for production:} The
        three-stage build produced a significantly smaller and more secure image by
        excluding source code, build tools, and full \texttt{node\_modules} from
        the runner stage, running as a non-root user.

  \item \textbf{Local Jenkins requires bridging infrastructure:} Running Jenkins
        locally without a public IP required ngrok for webhook delivery. In a
        production DevOps setup, Jenkins would run on a cloud VM (AWS EC2, GCP,
        Azure) with a static IP, eliminating this dependency.

  \item \textbf{Jenkins Credentials store is non-negotiable:} Storing Docker Hub
        credentials in Jenkins and injecting them via \texttt{withCredentials}
        prevents credentials from appearing in console logs, build history, or
        source code.

  \item \textbf{Pipeline as Code (Jenkinsfile in SCM) is the correct model:}
        Version-controlling the Jenkinsfile alongside application code ensures
        pipeline changes are auditable, peer-reviewable, and always in sync with
        the code they deploy.

  \item \textbf{Two pipeline variants serve different purposes:} The no-push variant
        proved valuable for rapid development iteration, avoiding unnecessary Docker
        Hub image pollution during testing.

  \item \textbf{DevOps tools require investment but deliver compounding returns:}
        Initial configuration of Jenkins, Trivy, Docker multi-stage builds, and
        webhook integration required significant time. However, once configured, the
        system runs reliably and automatically for every subsequent commit with zero
        developer overhead.
\end{itemize}

\section{Future Enhancements}

\begin{itemize}
  \item Upgrade Next.js to $\geq$\,14.2.25 to resolve the CRITICAL CVE-2025-29927
        authorization bypass vulnerability.
  \item Add Jest unit and integration tests with a dedicated \emph{Test} stage in
        the pipeline.
  \item Integrate SonarQube Cloud for deeper code quality metrics (coverage,
        duplication, code smells).
  \item Migrate Jenkins to a cloud VM to eliminate the ngrok dependency and enable
        24/7 persistent pipeline availability.
  \item Add automated deployment from Docker Hub to Render or Railway as an
        additional cloud target.
  \item Implement Docker Compose for local multi-container development environment
        management.
\end{itemize}

\vspace{24pt}
\begin{center}
\rule{0.4\textwidth}{0.5pt}\\[6pt]
{\itshape End of Report}\\[4pt]
\rule{0.4\textwidth}{0.5pt}
\end{center}

\end{document}
"""

# ═════════════════════════════════════════════════════════════════════════════
# Assemble & write .tex
# ═════════════════════════════════════════════════════════════════════════════
full = (PREAMBLE + TITLE + CH1 + CH2 + CH3 + CH4 + CH5 +
        CH6 + CH7 + CH8 + CH9 + CH10 + CH11 + CH12)

with open(TEX, 'w', encoding='utf-8') as f:
    f.write(full)
print("Written: " + TEX)

# ═════════════════════════════════════════════════════════════════════════════
# Compile with xelatex (twice for TOC + refs)
# ═════════════════════════════════════════════════════════════════════════════
for run in range(1, 3):
    print(f"xelatex pass {run}/2 ...")
    res = subprocess.run(
        [XE, "-interaction=nonstopmode", "-output-directory", BASE, TEX],
        capture_output=True, text=True, cwd=BASE
    )
    raw_pdf_check = os.path.join(BASE, "report.pdf")
    if not os.path.exists(raw_pdf_check):
        log = res.stdout + res.stderr
        lines = log.splitlines()
        print("ERRORS (last 40 lines):")
        print("\n".join(lines[-40:]))
        sys.exit(1)
    print(f"  pass {run} OK")

# Rename output to final name
raw_pdf = os.path.join(BASE, "report.pdf")
if os.path.exists(raw_pdf):
    if os.path.exists(PDF):
        os.remove(PDF)
    os.rename(raw_pdf, PDF)
    size = os.path.getsize(PDF) / (1024*1024)
    print(f"PDF ready: {PDF}  ({size:.2f} MB)")
else:
    print("ERROR: report.pdf not found after compilation")
    sys.exit(1)
