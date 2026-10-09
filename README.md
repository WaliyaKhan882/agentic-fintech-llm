# Agentic FinTech LLM

A small domain-expert language model for security and reliability questions about reinforcement-learning-based financial systems.

I turned a 60-paper literature survey into a 3,280-example instruction dataset, then fine-tuned **Llama 3.2 1B Instruct** on it with LoRA so it can answer expert questions about RL trading agents, adversarial threats, reward hacking, and regulatory gaps in FinTech.

| | Link |
|---|---|
| Dataset | [WaliyaKhan882/agentic_fintech](https://huggingface.co/datasets/WaliyaKhan882/agentic_fintech) |
| LoRA adapter | [WaliyaKhan882/metallama-trl-ft-lora-adapter](https://huggingface.co/WaliyaKhan882/metallama-trl-ft-lora-adapter) |
| Merged model | [WaliyaKhan882/metallama-trl-ft-merged](https://huggingface.co/WaliyaKhan882/metallama-trl-ft-merged) |
| GGUF for Ollama | [WaliyaKhan882/metallama-merged-trl-ft-gguf](https://huggingface.co/WaliyaKhan882/metallama-merged-trl-ft-gguf) |

---

## Why build the dataset this way

A model trained on "summarise paper X" questions only helps people who already know the papers. So every question in this dataset is written at the **domain level** ("What adversarial risks arise when deploying multi-agent RL systems in live markets?") and never names a paper. Paper titles appear only in answers that are explicitly asked to list references.

Each core fact is asked in **five different phrasings** with the same answer, which teaches the model the fact rather than one wording of the question.

## What's in the dataset

| Section | Entries | What it covers |
|---|---|---|
| Per-paper domain Q&A | about 2,700 | 60 papers × 9 question types (concept, threat, solution, empirical evidence, dimension, future work, datasets, taxonomy, which papers) × 5 phrasings |
| Grouped paper lists | about 500 | "Which research addresses X?" grouped by pillar, dimension and mechanism |
| Synthesis questions | 80 | 16 cross-cutting topics such as goal misgeneralisation, systemic risk and evaluation metrics, × 5 phrasings |

Every record is in ChatML `messages` format (system, user, assistant), ready for supervised fine-tuning.

## Fine-tuning setup

| Setting | Value |
|---|---|
| Base model | `meta-llama/Llama-3.2-1B-Instruct` |
| Method | LoRA, supervised fine-tuning with Hugging Face TRL |
| LoRA rank / alpha / dropout | 16 / 32 / 0.05 |
| Target modules | `q_proj`, `k_proj`, `v_proj`, `o_proj` |
| Outputs | LoRA adapter, merged weights, and a GGUF build for running locally in Ollama |

The trained artifacts are published on Hugging Face (links above).

---

## Repository contents

```
agentic-fintech-llm/
├── build_dataset.py              # builds the instruction dataset from the survey files
└── data/
    └── agentic_fintech_sft.jsonl # the generated dataset (3,280 examples)
```

`build_dataset.py` reads two source files that belong to an unpublished literature survey and are not included here:

- `data/literature_matrix.xlsx`: one row per paper with 15 structured fields (venue, year, domain, pillar, dimension, mechanism, problem, solution, results, datasets, future work and more)
- `data/paper_summaries.docx`: a structured summary of each paper

With those files in place:

```bash
pip install openpyxl python-docx
python build_dataset.py
```

## Try the model with Ollama

Download `metallama-merged.gguf` from the GGUF repo above, then create a Modelfile that points to it:

```
FROM ./metallama-merged.gguf
```

```bash
ollama create agentic-fintech -f Modelfile
ollama run agentic-fintech "What is reward hacking in RL-based trading agents?"
```

## Limits

This is a 1B-parameter model trained on 60 papers. It is useful for explaining concepts from that literature, not for financial advice or trading decisions, and it can still state things confidently that are wrong. Check important answers against the original papers.

---

Built by [Waliya Khan](https://github.com/WaliyaKhan882) during my MS in Data Science at the University of the Punjab.
