# WaitSender

WaitSender is a Python desktop application designed to orchestrate, schedule, and automate outbound messaging workflows for B2B and B2C operational communication. The platform incorporates a hybrid dispatch engine (direct UI automation via WhatsApp Web and API integration via Green API), adaptive rate-limiting logic, and an LLM-assisted copy refinement module.

## Technical Highlights and Architecture

- Kinetic Cursor Interruption: Background thread that continuously samples mouse displacement (> 6 px). If manual input is detected while dispatch routines are active, the automation immediately halts to prevent focus loss, misplaced keystrokes, or accidental interaction.
- Probabilistic Batch Rate-Limiting: Batched message queue with pseudo-random inter-message and inter-batch delays to avoid deterministic execution signatures and minimize anti-abuse detection triggers.
- Zero-Loss Fault Handling: State preservation mechanism triggered on network disconnects or UI exceptions. The system generates an incident log (.txt) and exports an Excel workbook containing strictly the remaining unfulfilled records for immediate replay.
- Human-in-the-Loop LLM Integration: Contextual rewriting, copy variant generation, and grammar refinement powered by Groq, OpenAI, or Anthropic client libraries, ensuring end-user review before execution.
- Local Storage and Data Privacy: Contact records, logs, and generated files remain strictly on the host operating system, avoiding third-party server persistence.

## Tech Stack

- Language: Python 3.10+
- Graphical User Interface: Tkinter, TkinterDnD2
- Data Ingestion & Serialization: Pandas, OpenPyXL
- Automation & OS Interaction: PyAutoGUI, PowerShell Bridge
- Audio Input & Processing: SoundDevice, Scipy, SpeechRecognition
- AI & LLM Providers: Groq SDK, OpenAI SDK, Anthropic SDK
- Packaging: PyInstaller

## Repository Structure

WaitSender/
│
├── app.py                     # Main application logic and Tkinter UI
├── README.md                  # Project overview and setup documentation
├── requirements.txt           # Environment dependencies
├── .gitignore                 # Exclusion rules for local and transient files
└── docs/                      # Technical manuals and governance documentation
    ├── manual_operativo_SOP.md
    ├── guia_inicio_rapido.md
    └── eula_terminos_servicio.md

## Installation and Setup

1. Clone the repository:
```bash
git clone [https://github.com/wendyL2107/WaitSender.git](https://github.com/wendyL2107/WaitSender.git)
cd WaitSender
