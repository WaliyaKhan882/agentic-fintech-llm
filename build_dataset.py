"""
Dataset builder for fine-tuning meta-llama/Llama-3.2-1B-Instruct

DESIGN PHILOSOPHY:
  - Questions are DOMAIN-LEVEL — no paper titles in the user prompt.
  - A user who has never heard of the papers should get expert answers.
  - Paper titles appear ONLY in answers that explicitly list references.
  - Each core fact → 5 rephrasings of the question, same answer.
  - Includes: grouped "list papers on X" entries and general synthesis Q&A.

Structure:
  60 papers x 9 concept Q-types x 5 rephrasings  = ~2,700 per-paper entries
  + grouped "which papers studied X" list entries  =  ~300 entries
  + general domain synthesis questions             =   ~75 entries
  Total: ~3,000+

Output  : data/agentic_fintech_sft.jsonl
Format  : {"messages": [{"role":"system",...},{"role":"user",...},{"role":"assistant",...}]}
"""

import json, re, random
import openpyxl
from docx import Document
from collections import defaultdict

EXCEL_FILE = "data/literature_matrix.xlsx"   # survey matrix: one row per paper (not included)
DOCX_FILE  = "data/paper_summaries.docx"    # structured paper summaries (not included)

SYSTEM_PROMPT = (
    "You are a knowledgeable research assistant specializing in artificial "
    "intelligence, reinforcement learning, and financial technology (FinTech). "
    "Answer questions accurately and concisely based on published academic research "
    "on security and reliability gaps in RL-based agentic financial systems."
)

# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────
def clean(val):
    if val is None:
        return ""
    s = str(val).strip()
    return "" if s.lower() in ("none", "n/a", "-", "") else s

def strip_number(s):
    """Remove leading numbering like '6.2 ' or '6. ' from taxonomy names."""
    return re.sub(r"^\d+\.?\d*\s+", "", str(s)).strip()

def clean_domain(s):
    """'Fintech / Trading' -> 'FinTech and Trading'"""
    return re.sub(r"\s*/\s*", " and ", s).strip()

def norm(s):
    return re.sub(r"\s+", " ", s.lower().strip())

# ─────────────────────────────────────────────────────────────────────────────
# Load Excel
# ─────────────────────────────────────────────────────────────────────────────
wb = openpyxl.load_workbook(EXCEL_FILE)
ws = wb["Sheet1"]

papers = []
for i, row in enumerate(ws.iter_rows(values_only=True)):
    if i < 3:
        continue
    if not row[0]:
        continue
    papers.append({
        "title"        : clean(row[0]),
        "venue"        : clean(row[1]),
        "year"         : clean(row[2]),
        "domain"       : clean_domain(clean(row[6])),
        "sub_domain"   : clean_domain(clean(row[7])),
        "pillar"       : strip_number(clean(row[8])),
        "dimension"    : strip_number(clean(row[9])),
        "mechanism"    : strip_number(clean(row[10])),
        "reasoning"    : clean(row[11]),
        "problem"      : clean(row[12]),
        "solution"     : clean(row[13]),
        "results"      : clean(row[14]),
        "eval_measures": clean(row[16]),
        "dataset_field": clean(row[17]),
        "future"       : clean(row[19]),
        "notes"        : clean(row[20]),
        "summary"      : clean(row[21]),
    })

print(f"[Excel] {len(papers)} papers loaded")

# ─────────────────────────────────────────────────────────────────────────────
# Load DOCX
# ─────────────────────────────────────────────────────────────────────────────
doc = Document(DOCX_FILE)
docx_map = {}
cur_title, cur_steps, cur_key = None, {}, None

for para in doc.paragraphs:
    text  = para.text.strip()
    if not text:
        continue
    style = para.style.name.lower()
    is_title = (
        style == "heading 1"
        or (style == "normal"
            and para.runs
            and any(r.bold for r in para.runs if r.text.strip())
            and not text.lower().startswith("step"))
    )
    if is_title and not text.lower().startswith("step"):
        if cur_title and cur_steps:
            docx_map[cur_title] = dict(cur_steps)
        cur_title, cur_steps, cur_key = text, {}, None
        continue
    tl = text.lower()
    if   tl.startswith("step 1") or "abstract"     in tl[:30]: cur_key = "abstract"
    elif tl.startswith("step 2") or "introduction" in tl[:30]: cur_key = "intro"
    elif tl.startswith("step 3") or tl[:10].startswith("method"):  cur_key = "method"
    elif tl.startswith("step 4") or tl[:10].startswith("result"):  cur_key = "result"
    elif tl.startswith("step 5") or "discussion"   in tl[:15]:     cur_key = "discussion"
    elif cur_key:
        cur_steps[cur_key] = cur_steps.get(cur_key, "") + " " + text
        continue
    if cur_key:
        cur_steps[cur_key] = text

if cur_title and cur_steps:
    docx_map[cur_title] = cur_steps

print(f"[DOCX]  {len(docx_map)} papers loaded")

# ─────────────────────────────────────────────────────────────────────────────
# Merge Excel + DOCX
# ─────────────────────────────────────────────────────────────────────────────
for p in papers:
    en    = norm(p["title"])
    p["docx"] = {}
    for dt, dd in docx_map.items():
        dn = norm(dt)
        if en[:50] == dn[:50] or en in dn or dn in en:
            p["docx"] = dd
            break

print(f"[Merge] {len(papers)} merged records")

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 1 — Per-paper domain-knowledge Q&A
# Questions use mechanism / dimension / domain fields — NEVER the paper title.
# Paper titles appear only in "which_papers" answers.
# ─────────────────────────────────────────────────────────────────────────────

TEMPLATES = {

    # 1. What is this concept?
    "concept": [
        "What is {mechanism} and why does it matter in {domain}?",
        "Can you explain {mechanism} as it applies to AI-driven financial systems?",
        "What does {mechanism} mean in the context of reinforcement learning-based FinTech?",
        "Define {mechanism} and describe its relevance to autonomous financial agents.",
        "Describe {mechanism} and explain why it is studied in {sub_domain} research.",
    ],

    # 2. What threat / problem does it create?
    "threat": [
        "What security or reliability risk does {mechanism} pose to {domain} systems?",
        "How does {mechanism} threaten the dependability of AI-driven financial agents?",
        "What problems arise from {mechanism} in {sub_domain} applications?",
        "Why is {mechanism} a critical concern for autonomous FinTech agents?",
        "What are the consequences of {mechanism} for RL-based financial decision-making?",
    ],

    # 3. How is it addressed / defended?
    "solution": [
        "How can {mechanism} be mitigated in {domain} AI systems?",
        "What techniques have researchers proposed to defend against {mechanism} in financial AI?",
        "What approaches address {mechanism} in {sub_domain}?",
        "How do agentic FinTech systems handle the challenge of {mechanism}?",
        "What defensive strategies exist against {mechanism} in RL-based financial agents?",
    ],

    # 4. What do experiments show?
    "empirical": [
        "What experimental evidence exists about {mechanism} in {domain}?",
        "What do research results show about the impact of {mechanism} on financial AI performance?",
        "How effective are countermeasures against {mechanism} according to published studies?",
        "What quantitative findings have been reported for {mechanism} in {sub_domain}?",
        "What have studies demonstrated about the severity and handling of {mechanism} in {domain}?",
    ],

    # 5. What is this dimension / technical area?
    "dimension": [
        "What does {dimension} refer to as a technical challenge in agentic FinTech?",
        "Why is {dimension} an important concern for autonomous financial AI systems?",
        "Describe the {dimension} challenges that affect RL-based FinTech agents.",
        "How does {dimension} impact the design of reliable financial AI systems?",
        "What specific failure modes are grouped under {dimension} in FinTech AI research?",
    ],

    # 6. Future open problems
    "future": [
        "What open research problems remain around {mechanism} in {domain}?",
        "Where should future work focus to better handle {mechanism} in financial AI?",
        "What unsolved challenges in {dimension} are identified within {domain} research?",
        "How might future research improve resilience against {mechanism} in agentic FinTech?",
        "What directions do researchers propose for advancing work on {mechanism} in {domain}?",
    ],

    # 7. Datasets and evaluation
    "datasets": [
        "What data is typically used to study {mechanism} in {sub_domain}?",
        "Which datasets and benchmarks are relevant to evaluating {mechanism} in {domain}?",
        "On what kind of financial data has {mechanism} been empirically tested?",
        "What real-world datasets have researchers used to evaluate {mechanism}?",
        "What evaluation metrics and data sources are common in {mechanism} research?",
    ],

    # 8. Taxonomy / classification
    "taxonomy": [
        "How is {mechanism} classified in the security and reliability taxonomy of agentic FinTech?",
        "Under which research pillar and dimension does {mechanism} fall in RL-FinTech research?",
        "What high-level theme and technical area is {mechanism} associated with in financial AI?",
        "In a systematic survey of RL-based agentic FinTech gaps, how is {mechanism} categorized?",
        "What is the research classification for {mechanism} in autonomous FinTech systems?",
    ],

    # 9. Which papers studied this (paper title in ANSWER only, not in question)
    "which_papers": [
        "What research has specifically studied {mechanism} in {domain}?",
        "Are there published academic studies that investigate {mechanism} in {sub_domain}?",
        "What academic work addresses {mechanism} as a challenge in financial AI?",
        "Which studies focus on {mechanism} in the context of {sub_domain}?",
        "Has peer-reviewed research examined {mechanism} in {domain} applications?",
    ],
}


def compose_answer(fact_type, paper):
    prob = paper["problem"]
    sol  = paper["solution"]
    res  = paper["results"]
    fut  = paper["future"]
    ds   = paper["dataset_field"]
    ev   = paper["eval_measures"]
    dim  = paper["dimension"]
    pil  = paper["pillar"]
    mech = paper["mechanism"]
    smry = paper["summary"]
    rsn  = paper["reasoning"]
    docx = paper["docx"]

    if fact_type == "concept":
        parts = []
        if prob: parts.append(prob)
        if sol:  parts.append(f"To address this, researchers have proposed: {sol}")
        if res:  parts.append(f"Empirical evidence shows: {res}")
        return "\n\n".join(parts)

    if fact_type == "threat":
        return prob

    if fact_type == "solution":
        parts = []
        if sol: parts.append(sol)
        if res: parts.append(f"Results demonstrate: {res}")
        return "\n\n".join(parts)

    if fact_type == "empirical":
        parts = []
        if res: parts.append(res)
        if ev:  parts.append(f"Evaluation metrics used: {ev}")
        if ds:  parts.append(f"Dataset: {ds}")
        return "\n\n".join(parts)

    if fact_type == "dimension":
        base = docx.get("abstract", "") or docx.get("intro", "") or prob
        if dim and base:
            return f"{dim} in agentic FinTech concerns: {base[:600].strip()}"
        return base[:600].strip() if base else ""

    if fact_type == "future":
        return (fut or docx.get("discussion", "")).strip()[:600]

    if fact_type == "datasets":
        parts = []
        if ds: parts.append(f"Datasets: {ds}")
        if ev: parts.append(f"Evaluation metrics: {ev}")
        return "\n".join(parts)

    if fact_type == "taxonomy":
        parts = []
        if pil:  parts.append(f"Pillar (L1): {pil}")
        if dim:  parts.append(f"Dimension (L2): {dim}")
        if mech: parts.append(f"Mechanism (L3): {mech}")
        if rsn:  parts.append(f"Why this classification: {rsn}")
        return "\n".join(parts)

    if fact_type == "which_papers":
        desc  = smry or prob or docx.get("abstract", "")
        desc  = (desc[:280] + "...") if len(desc) > 280 else desc
        title = paper["title"]
        venue = paper["venue"]
        year  = paper["year"]
        ref   = f'"{title}"'
        if venue and year: ref += f" ({venue}, {year})"
        elif year:         ref += f" ({year})"
        return f"Key research in this area: {ref}.\n\n{desc}"

    return ""


# Build per-paper entries
dataset = []
skipped = 0

for paper in papers:
    mech = paper["mechanism"]
    dim  = paper["dimension"]
    dom  = paper["domain"]   or "FinTech"
    sub  = paper["sub_domain"] or "financial AI"

    if not mech and not dim:
        skipped += 1
        continue

    concept_token = mech or dim   # prefer specific mechanism; fall back to dimension

    fmt = {
        "mechanism" : concept_token,
        "dimension" : dim or concept_token,
        "domain"    : dom,
        "sub_domain": sub,
    }

    for fact_type, templates in TEMPLATES.items():
        answer = compose_answer(fact_type, paper).strip()
        if not answer:
            skipped += 1
            continue
        for q_tpl in templates:
            try:
                question = q_tpl.format(**fmt)
            except KeyError:
                continue
            dataset.append({
                "messages": [
                    {"role": "system",    "content": SYSTEM_PROMPT},
                    {"role": "user",      "content": question},
                    {"role": "assistant", "content": answer},
                ]
            })

print(f"[Per-paper]  {len(dataset)} entries  (skipped: {skipped} empty fields)")

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 2 — Grouped "list papers on topic X" questions
# ─────────────────────────────────────────────────────────────────────────────
def paper_entry(p):
    desc = p["summary"] or p["problem"] or p["docx"].get("abstract", "")
    desc = (desc[:160] + "...") if len(desc) > 160 else desc
    line = f'"{p["title"]}"'
    if p["venue"] and p["year"]: line += f' ({p["venue"]}, {p["year"]})'
    elif p["year"]:              line += f' ({p["year"]})'
    if desc: line += f"\n   → {desc}"
    return line

def add_list_entries(topic_label, matched_papers, q_templates):
    if not matched_papers:
        return
    answer = (
        f"Research papers on {topic_label}:\n\n"
        + "\n\n".join(f"{i+1}. {paper_entry(p)}" for i, p in enumerate(matched_papers))
    )
    for q in q_templates:
        dataset.append({
            "messages": [
                {"role": "system",    "content": SYSTEM_PROMPT},
                {"role": "user",      "content": q},
                {"role": "assistant", "content": answer},
            ]
        })

# Group by pillar
pillar_groups = defaultdict(list)
for p in papers:
    if p["pillar"]:
        pillar_groups[p["pillar"]].append(p)

for pillar, plist in pillar_groups.items():
    add_list_entries(pillar, plist, [
        f"List research papers that cover the '{pillar}' theme in agentic FinTech.",
        f"Which studies fall under the '{pillar}' pillar in RL-based financial AI research?",
        f"What papers have you studied that are classified under '{pillar}'?",
        f"Name the research works related to '{pillar}' in autonomous FinTech systems.",
        f"Can you list papers dealing with '{pillar}' in the context of financial AI?",
    ])

# Group by dimension
dim_groups = defaultdict(list)
for p in papers:
    if p["dimension"]:
        dim_groups[p["dimension"]].append(p)

for dim, plist in dim_groups.items():
    add_list_entries(dim, plist, [
        f"Which papers specifically study {dim} in financial AI systems?",
        f"List research works that address {dim} as a challenge in agentic FinTech.",
        f"What studies have investigated {dim} in the context of RL-based finance?",
        f"Name some papers covering {dim} in autonomous financial systems.",
        f"Can you point me to research on {dim} in FinTech AI?",
    ])

# Group by mechanism
mech_groups = defaultdict(list)
for p in papers:
    if p["mechanism"]:
        mech_groups[p["mechanism"]].append(p)

for mech, plist in mech_groups.items():
    dom_example = plist[0]["domain"] or "FinTech"
    add_list_entries(mech, plist, [
        f"What papers have studied {mech} in financial AI?",
        f"List research works that examine {mech} in {dom_example}.",
        f"Which studies specifically address {mech} in agentic FinTech?",
        f"Are there papers on {mech} as it relates to autonomous financial systems?",
        f"What research covers {mech} in RL-based FinTech?",
    ])

print(f"[Grouped]    {len(dataset)} entries so far")

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 3 — General synthesis / meta questions
# ─────────────────────────────────────────────────────────────────────────────
all_titles  = [p["title"] for p in papers]
all_pillars = sorted(set(p["pillar"]    for p in papers if p["pillar"]))
all_dims    = sorted(set(p["dimension"] for p in papers if p["dimension"]))
all_mechs   = sorted(set(p["mechanism"] for p in papers if p["mechanism"]))
all_domains = sorted(set(p["domain"]    for p in papers if p["domain"]))

sec_papers = [p for p in papers if
    "security" in (p["pillar"] + p["dimension"]).lower()
    or "adversar" in p["mechanism"].lower()
    or "attack"   in p["problem"].lower()
    or "poison"   in p["mechanism"].lower()]

rel_papers = [p for p in papers if
    "reliability" in (p["pillar"] + p["dimension"]).lower()
    or "hallucin"  in p["problem"].lower()
    or "goal mis"  in p["mechanism"].lower()]

rl_papers = [p for p in papers if
    "reinforcement" in (p["domain"] + p["sub_domain"]).lower()]

llm_papers = [p for p in papers if
    "llm" in (p["domain"] + p["sub_domain"] + p["solution"]).lower()
    or "language model" in (p["solution"] + p["problem"]).lower()]

def bullets(plist, field="problem", limit=12):
    out = []
    for p in plist[:limit]:
        val = p.get(field, "")
        val = (val[:120] + "...") if len(val) > 120 else val
        out.append(f"- {p['title']}" + (f": {val}" if val else ""))
    return "\n".join(out)


GENERAL_QA = [

    (   # Security gaps overview
        [
            "What are the main security gaps in RL-based agentic FinTech systems?",
            "List the key security vulnerabilities in autonomous financial AI.",
            "What security threats do reinforcement learning agents face in FinTech?",
            "Summarize the critical security risks in AI-driven financial systems.",
            "What are the most important security challenges for agentic FinTech?",
        ],
        "Key security gaps in RL-based agentic FinTech systems include:\n\n"
        + bullets(sec_papers, "problem")
        + "\n\nMajor threat categories: adversarial evasion attacks (FGSM, DeepFool), "
          "training-time data poisoning, oracle-based label manipulation, "
          "goal misgeneralization, and systemic risk from correlated multi-agent failures."
    ),

    (   # Reliability gaps overview
        [
            "What reliability gaps exist in agentic FinTech AI systems?",
            "What are the main reliability and robustness issues in RL-based financial agents?",
            "Summarize the reliability risks that autonomous financial AI faces.",
            "How does current research characterize reliability failures in RL-FinTech agents?",
            "What makes RL-based financial agents unreliable in real-world deployments?",
        ],
        "Reliability gaps in agentic FinTech include:\n\n"
        + bullets(rel_papers, "problem")
        + "\n\nCommon failure types: LLM hallucinations in financial queries, "
          "goal misgeneralization (agent pursues wrong proxy objective), "
          "reward misspecification, and brittleness under out-of-distribution market conditions."
    ),

    (   # Adversarial attacks deep dive
        [
            "What types of adversarial attacks threaten financial AI models?",
            "How are adversarial machine learning attacks used against FinTech systems?",
            "What is data poisoning in financial AI and how does it work?",
            "Explain adversarial attacks on RL-based trading and fraud detection systems.",
            "What adversarial exploits are documented in financial machine learning research?",
        ],
        "Adversarial attacks in financial AI research:\n\n"
        "1. Testing-time Evasion (FGSM, JSMA, DeepFool): small imperceptible perturbations "
        "are injected into input data (e.g., price deltas) to fool trained models. "
        "Studies show accuracy can drop from ~90% to ~40% across MLP, CNN, and LSTM architectures.\n\n"
        "2. Training-time Label Poisoning: attackers corrupt the training set by mislabeling "
        "fraudulent transactions as legitimate, embedding hidden backdoors in the model.\n\n"
        "3. Oracle-based Poisoning: the attacker replicates the defender's model to craft "
        "maximally effective poison samples. In simulated scenarios, this enabled theft of up "
        "to €1M; a proposed defense using synthetic fraud pre-training reduced losses by 31–100%.\n\n"
        "Key implication: financial AI must protect both its inference pipeline and its training data."
    ),

    (   # Goal misgeneralization
        [
            "What is goal misgeneralization in reinforcement learning?",
            "How does goal misgeneralization affect RL agents in financial systems?",
            "Why might an RL agent pursue the wrong goal even when trained correctly?",
            "Explain goal misgeneralization and its relevance to agentic FinTech.",
            "What happens when an RL trading agent generalizes to the wrong objective?",
        ],
        "Goal misgeneralization occurs when an RL agent correctly learns how to achieve a goal "
        "during training but internalizes a proxy goal rather than the intended one. "
        "Under distribution shift during deployment, the agent competently pursues this wrong objective.\n\n"
        "In FinTech: a trading agent trained to maximize returns might instead misgeneralize "
        "to 'mimic a specific bot's patterns' — a spurious correlation. In live markets, "
        "this produces stable-looking but incorrect behavior such as unnecessary trades or systemic loops.\n\n"
        "Key insight: an agent's competence is itself a risk factor — a highly capable agent "
        "will efficiently pursue the wrong goal. Correct reward specification during training "
        "is insufficient to guarantee correct goals during deployment."
    ),

    (   # Role of RL in FinTech
        [
            "How is reinforcement learning used in financial technology?",
            "What role does RL play in autonomous financial decision-making?",
            "Give an overview of reinforcement learning applications in FinTech.",
            "What RL-based systems are used in modern financial services?",
            "How are RL agents applied in trading, fraud detection, and risk management?",
        ],
        "Reinforcement learning in FinTech applications:\n\n"
        "- Algorithmic trading: RL agents learn buy/sell policies to maximize cumulative return "
        "while managing risk (Sharpe Ratio, Max Drawdown).\n"
        "- Fraud detection: adaptive RL models update detection policies in response to evolving fraud patterns.\n"
        "- Portfolio management: multi-agent RL frameworks coordinate positions across assets.\n"
        "- Anti-money laundering: RL optimizes transaction monitoring thresholds.\n"
        "- Risk assessment: agents process multimodal data (earnings calls, news, prices) "
        "to predict volatility.\n\n"
        + bullets(rl_papers, "solution", limit=8)
    ),

    (   # LLM agents in FinTech
        [
            "How are large language models used as financial AI agents?",
            "What LLM-based frameworks exist for autonomous financial decision-making?",
            "What is an LLM trading agent and how does it work?",
            "How do LLM agents differ from traditional RL-based financial systems?",
            "What advantages do LLM agents offer for FinTech applications?",
        ],
        "LLM-based financial agents use large language models as the reasoning core:\n\n"
        "- Memory: layered modules (shallow/deep) store recent news, historical patterns, "
        "and long-term fundamentals separately.\n"
        "- Multi-agent coordination: specialized analyst/researcher/trader agents debate "
        "and produce decisions, mimicking real trading firm structures.\n"
        "- RAG integration: Retrieval-Augmented Generation grounds responses in real documents, "
        "reducing hallucinations.\n"
        "- Explainability: unlike black-box DRL, LLM agents produce natural-language reasoning traces.\n\n"
        "Active research:\n"
        + bullets(llm_papers, "solution", limit=8)
    ),

    (   # Hallucinations in financial AI
        [
            "What is the hallucination problem in LLM-based financial AI?",
            "Why do LLM financial agents produce incorrect or fabricated answers?",
            "How is the hallucination risk addressed in FinTech AI systems?",
            "What causes hallucinations in LLM trading agents and how are they mitigated?",
            "How does RAG help reduce hallucinations in financial language models?",
        ],
        "Hallucination in LLM-based financial agents refers to generating plausible-sounding "
        "but factually wrong information — fabricated financial figures, non-existent regulations, "
        "or memory indices that do not exist.\n\n"
        "Causes: LLMs cannot reliably distinguish retrieved facts from statistical patterns; "
        "long context windows cause attention drift away from key facts.\n\n"
        "Mitigations:\n"
        "- GraphRAG: structures retrieved context as a knowledge graph, reducing noise. "
        "Demonstrated 6% hallucination reduction and 734× token reduction for compliance tasks.\n"
        "- HybridRAG: combines vector (semantic) and graph (symbolic) retrieval via KNN clustering.\n"
        "- SCMRAG (Self-Corrective Multihop RAG): multi-hop verification before finalizing answers.\n\n"
        "For regulated financial use cases, hallucination mitigation is critical to meet "
        "explainability requirements under frameworks like the EU AI Act."
    ),

    (   # Regulatory / compliance gaps
        [
            "What regulatory challenges do agentic FinTech AI systems face?",
            "How does AI regulation affect autonomous financial systems?",
            "What compliance gaps exist for RL-based agents in financial services?",
            "How does regulatory misalignment manifest in AI-driven FinTech?",
            "What does the EU AI Act mean for autonomous financial AI agents?",
        ],
        "Regulatory and compliance gaps in agentic FinTech:\n\n"
        "1. Black-box opacity: EU AI Act and GDPR require explainability for high-stakes decisions, "
        "but most RL and deep learning models lack interpretable decision rationales.\n"
        "2. Algorithmic discrimination: opaque credit-scoring algorithms can systematically "
        "exclude marginalized groups, violating fair lending laws.\n"
        "3. EU AI Act compliance: autonomous financial agents fall under 'high-risk AI,' requiring "
        "conformity assessments, audit trails, and human oversight — not natively built into RL.\n"
        "4. Regulatory arbitrage: rapid FinTech innovation outpaces specific regulatory standards.\n"
        "5. Multi-jurisdictional complexity: cross-border agentic payments face inconsistent rules."
    ),

    (   # Federated learning in FinTech
        [
            "What is federated learning and how is it applied in FinTech?",
            "How does federated learning help with privacy in financial AI?",
            "What role does federated learning play in financial risk management?",
            "Why is federated learning relevant to agentic FinTech systems?",
            "How is federated RLHF used to improve LLMs in financial applications?",
        ],
        "Federated learning (FL) allows financial institutions to collaboratively train AI models "
        "without sharing raw sensitive data — each party trains locally and shares only model updates.\n\n"
        "FinTech applications:\n"
        "- Fraud detection: banks train shared fraud models without exposing customer records.\n"
        "- AML: cross-institution pattern detection while preserving data privacy regulations.\n"
        "- Smart contract security: federated analysis of blockchain transaction patterns.\n"
        "- Federated RLHF: aggregating client preference signals to fine-tune LLMs without "
        "centralizing user conversations.\n\n"
        "Challenges: model poisoning by malicious participants, communication overhead, "
        "heterogeneous data distributions, and participant integrity verification."
    ),

    (   # Memory architectures in LLM agents
        [
            "What memory architectures are used in LLM-based trading agents?",
            "How do AI trading agents manage and retain financial knowledge over time?",
            "What is layered memory in the context of LLM financial agents?",
            "How does memory design affect the performance of autonomous financial agents?",
            "What role does the Ebbinghaus forgetting curve play in financial AI memory?",
        ],
        "Memory architectures in LLM-based financial agents mimic human cognitive structures:\n\n"
        "1. Layered memory (shallow → deep): short-term holds recent news/prices; "
        "intermediate holds company reports; deep memory holds long-term fundamentals. "
        "Inspired by the Ebbinghaus forgetting curve — less important memories decay over time.\n\n"
        "2. Vector/RAG retrieval: financial documents are embedded and retrieved semantically "
        "when relevant to a current decision.\n\n"
        "3. Reflection modules: agents periodically consolidate past experiences to update "
        "their beliefs about market conditions.\n\n"
        "4. Risk-adaptive character: memory systems encode risk profiles "
        "(risk-seeking, risk-averse, adaptive) that modulate decision behavior.\n\n"
        "Proper memory design is critical — agents with poor memory ignore important signals "
        "or get overwhelmed by noise."
    ),

    (   # Multi-agent RL in FinTech
        [
            "How is multi-agent reinforcement learning used in financial systems?",
            "What advantages does multi-agent RL offer for algorithmic trading?",
            "Why use multiple RL agents instead of a single agent in FinTech?",
            "What challenges does multi-agent RL face in financial applications?",
            "How do multi-agent systems improve financial decision-making?",
        ],
        "Multi-agent RL (MARL) distributes the financial decision task across specialized agents:\n\n"
        "Advantages:\n"
        "- Specialization: separate agents for technical, sentiment, fundamental analysis, "
        "and risk management; a manager agent makes final decisions.\n"
        "- Debate mechanisms: bullish and bearish agents present opposing views, reducing bias.\n"
        "- Portfolio coordination: agents balance diversification and correlation risk.\n\n"
        "Challenges:\n"
        "- Coordination cost: inter-agent communication overhead scales with agent count.\n"
        "- Value estimation variance in multi-agent settings (addressed by estimator variance reduction).\n"
        "- Hallucination risk increases in multi-asset MARL due to more complex context.\n"
        "- Systemic risk: correlated agent behavior can amplify market instability."
    ),

    (   # Research taxonomy
        [
            "What is the research taxonomy used to classify security and reliability gaps in agentic FinTech?",
            "How are research topics organized in a systematic review of RL-based FinTech AI?",
            "What are the main pillars, dimensions, and mechanisms in agentic FinTech gap analysis?",
            "Describe the L1-L2-L3 classification framework for RL-FinTech security research.",
            "What categories are used to organize the literature on agentic FinTech AI gaps?",
        ],
        "The taxonomy uses a three-level hierarchy:\n\n"
        "L1 Pillars (high-level themes):\n"
        + "\n".join(f"  - {p}" for p in all_pillars)
        + "\n\nL2 Dimensions (technical areas within each pillar):\n"
        + "\n".join(f"  - {d}" for d in all_dims[:20])
        + "\n\nL3 Mechanisms (specific algorithms, exploits, or techniques):\n"
        + "\n".join(f"  - {m}" for m in all_mechs[:20])
        + "\n\nEach paper is assigned Pillar → Dimension → Mechanism to precisely locate "
          "its contribution within the security and reliability gap landscape."
    ),

    (   # Meta: list papers you know
        [
            "What research papers on agentic FinTech security have you been trained on?",
            "List the papers in your knowledge base about RL-based financial AI.",
            "Name 5 research papers you have learned about on autonomous financial systems.",
            "Which research papers on agentic FinTech are you familiar with?",
            "What academic works covering FinTech AI security are part of your knowledge?",
        ],
        f"I have knowledge of {len(all_titles)} research papers covering security and "
        f"reliability gaps in RL-based agentic FinTech. A selection:\n\n"
        + "\n".join(
            f"{i+1}. \"{p['title']}\""
            + (f" ({p['venue']}, {p['year']})" if p["venue"] and p["year"]
               else f" ({p['year']})" if p["year"] else "")
            for i, p in enumerate(papers[:20])
        )
        + f"\n\n...and {len(all_titles) - 20} more. Topics span adversarial attacks, "
          "goal misgeneralization, LLM agent architectures, regulatory compliance, "
          "federated learning, RAG systems, MARL, and agentic FinTech benchmarks."
    ),

    (   # Evaluation metrics
        [
            "What performance metrics are commonly used to evaluate financial AI agents?",
            "How is an RL trading agent's performance measured in research?",
            "What evaluation metrics matter most in agentic FinTech research?",
            "How do researchers quantify the success of autonomous financial systems?",
            "What are the standard KPIs for RL-based trading and risk agents?",
        ],
        "Common evaluation metrics in financial AI research:\n\n"
        "Trading performance:\n"
        "- Cumulative Return: total portfolio growth over the test period.\n"
        "- Sharpe Ratio: risk-adjusted return (return / volatility); higher is better.\n"
        "- Maximum Drawdown: worst peak-to-trough loss; lower is better.\n"
        "- Volatility: standard deviation of periodic returns.\n\n"
        "Security metrics:\n"
        "- Attack Success Rate / Evasion Rate: fraction of adversarial samples that fool the model.\n"
        "- Accuracy and F1-Score under attack: performance degradation under adversarial input.\n"
        "- Financial loss prevented: monetary impact for fraud and AML systems.\n\n"
        "LLM and agent quality:\n"
        "- Faithfulness Score: accuracy of generated output relative to retrieved facts.\n"
        "- Hallucination Rate: fraction of generated claims that are factually wrong.\n"
        "- IFEval: instruction following accuracy.\n"
        "- Token efficiency: computational cost of generating a response."
    ),

    (   # Agentic AI definition
        [
            "What is agentic AI and how does it differ from traditional AI systems?",
            "Define agentic FinTech and explain its key characteristics.",
            "What makes an AI system 'agentic' in the context of financial services?",
            "What capabilities define an agentic AI system in FinTech?",
            "How does agentic AI change how financial services are delivered?",
        ],
        "Agentic AI systems are autonomous agents that perceive their environment, plan "
        "multi-step actions, use tools, and pursue goals with minimal human intervention.\n\n"
        "Key characteristics:\n"
        "- Autonomy: acts independently to achieve assigned financial objectives.\n"
        "- Tool use: calls APIs, executes trades, queries databases, runs code.\n"
        "- Memory: maintains state across interactions.\n"
        "- Reasoning: uses LLMs or RL policies to plan action sequences.\n"
        "- Self-healing: some systems detect and recover from failures automatically.\n\n"
        "Difference from traditional AI: traditional ML models passively predict given input. "
        "Agentic AI actively takes actions, interacts with external systems, and adapts over time.\n\n"
        "FinTech applications: autonomous trading, real-time fraud monitoring, personalized "
        "banking assistants, compliance checking, and cross-border payment orchestration."
    ),

    (   # Systemic risk
        [
            "What systemic risks do agentic AI systems create in financial markets?",
            "How can AI agents acting autonomously destabilize financial systems?",
            "What is systemic risk in agentic FinTech and why does it matter?",
            "How do correlated AI agent behaviors lead to market instability?",
            "What oversight mechanisms are proposed to manage systemic risk from financial AI?",
        ],
        "Systemic risk in agentic FinTech arises when many autonomous AI agents collectively "
        "destabilize markets or institutions:\n\n"
        "Mechanisms:\n"
        "- Herding: agents trained on similar data execute the same trades simultaneously, "
        "amplifying price movements.\n"
        "- Feedback loops: agents react to each other's actions, creating self-reinforcing "
        "cycles such as flash crashes or liquidity crises.\n"
        "- Correlated failures: shared architectures or training data cause many agents "
        "to fail under the same market conditions.\n"
        "- Speed asymmetry: AI agents operate at microsecond timescales, outpacing human oversight.\n\n"
        "Proposed mitigations:\n"
        "- Circuit breakers and position limits enforced at the infrastructure level.\n"
        "- Agent-based regulatory frameworks that simulate collective behavior.\n"
        "- Diversity requirements (different model architectures) to reduce correlation.\n"
        "- Real-time monitoring of agent decision patterns for anomaly detection."
    ),
]

before_general = len(dataset)
for q_list, answer in GENERAL_QA:
    for q in q_list:
        dataset.append({
            "messages": [
                {"role": "system",    "content": SYSTEM_PROMPT},
                {"role": "user",      "content": q},
                {"role": "assistant", "content": answer.strip()},
            ]
        })

print(f"[General]    {len(dataset) - before_general} entries added")
print(f"[Dataset]    Total: {len(dataset)} entries")

# ─────────────────────────────────────────────────────────────────────────────
# Save
# ─────────────────────────────────────────────────────────────────────────────
OUTPUT = "data/agentic_fintech_sft.jsonl"
with open(OUTPUT, "w", encoding="utf-8") as f:
    for entry in dataset:
        f.write(json.dumps(entry, ensure_ascii=False) + "\n")

print(f"[Saved]      {OUTPUT}  ({len(dataset)} lines)")

# ─────────────────────────────────────────────────────────────────────────────
# Sanity check
# ─────────────────────────────────────────────────────────────────────────────
all_questions = [e["messages"][1]["content"] for e in dataset]

# Check no paper title leaked into a user question
title_set   = {norm(p["title"]) for p in papers}
title_leaks = [q for q in all_questions if norm(q) in title_set]

no_quotes = sum(1 for q in all_questions if "'" not in q and '"' not in q)

print(f"\n── Sanity check ────────────────────────────────────────────────────────")
print(f"  Paper title leaks in questions : {len(title_leaks)}  (target: 0)")
print(f"  Questions with no quotes       : {no_quotes}  (pure domain questions)")
print(f"  Questions with quotes          : {len(all_questions) - no_quotes}  (taxonomy names etc.)")

# Preview 5 random domain-level entries
random.seed(42)
print("\n── 5 random sample Q&A ─────────────────────────────────────────────────")
for entry in random.sample(dataset, 5):
    q = entry["messages"][1]["content"]
    a = entry["messages"][2]["content"]
    print(f"\nQ: {q}")
    print(f"A: {a[:200]}{'...' if len(a) > 200 else ''}")
