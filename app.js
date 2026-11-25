if (typeof pdfjsLib !== "undefined" && pdfjsLib.GlobalWorkerOptions) {
    pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
}

function esc(str) {
    return String(str)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
}

// Technology-Domain Knowledge Sample (Cloud Infrastructure & AI Stack)
const TECH_SAMPLE_TEXT = `# Kubernetes Cluster
Category: Container Orchestration Platform
Runtime Engine: containerd
Kubernetes Cluster is deployed in AWS Mumbai Region.
Kubernetes Cluster uses Automated Horizontal Scaler.

# API Gateway
API Gateway is part of Ingress Mesh.
Ingress Mesh is part of Kubernetes Cluster.
API Gateway requires TLS 1.3 Encryption.
Protocol: REST and gRPC

# Vector Database
Vector Database requires NVMe SSD Storage.
Indexing Algorithm: HNSW Graph
Vector Database is part of AI Inference Stack.

# GPU Worker Node
GPU Worker Node uses TensorRT Runtime.
GPU Worker Node is part of Kubernetes Cluster.

# High Traffic Spike
High Traffic Spike causes CPU Saturation.
CPU Saturation requires Automated Horizontal Scaler.

# Domain Production Rules
IF ?X causes CPU Saturation THEN ?X is Autoscaling Trigger
IF ?X uses Automated Horizontal Scaler THEN ?X is High Availability Ready
IF ?X uses TensorRT Runtime THEN ?X is Hardware Accelerated`;

const TECH_SAMPLE_JSON = JSON.stringify({
    "Kubernetes Cluster": {
        "Category": "Container Orchestration Platform",
        "Runtime Engine": "containerd",
        "deployed_in": "AWS Mumbai Region",
        "uses": "Automated Horizontal Scaler"
    },
    "API Gateway": {
        "part_of": "Ingress Mesh",
        "requires": "TLS 1.3 Encryption",
        "Protocol": "REST and gRPC"
    },
    "Ingress Mesh": {
        "part_of": "Kubernetes Cluster"
    },
    "Vector Database": {
        "requires": "NVMe SSD Storage",
        "Indexing Algorithm": "HNSW Graph",
        "part_of": "AI Inference Stack"
    },
    "GPU Worker Node": {
        "uses": "TensorRT Runtime",
        "part_of": "Kubernetes Cluster"
    },
    "High Traffic Spike": {
        "causes": "CPU Saturation"
    },
    "Rules": [
        "IF ?X causes CPU Saturation THEN ?X is Autoscaling Trigger",
        "IF ?X uses Automated Horizontal Scaler THEN ?X is High Availability Ready",
        "IF ?X uses TensorRT Runtime THEN ?X is Hardware Accelerated"
    ]
}, null, 2);

class MultiFormatAdapter {
    static async decode(file) {
        if (!file) throw new Error("No file selected.");
        const ext = file.name.toLowerCase().split(".").pop();
        if (ext === "pdf") return await this.fromPDF(file);
        if (ext === "docx") return await this.fromDOCX(file);
        if (ext === "doc") throw new Error("Old .doc files are unsupported. Save as .docx or .pdf.");
        if (ext === "xlsx" || ext === "xls") return await this.fromExcel(file);
        if (ext === "csv" || ext === "tsv") return this.fromCSV(await file.text(), ext === "tsv" ? "\t" : ",");
        if (ext === "json") return this.fromJSON(await file.text());
        if (ext === "xml" || ext === "html" || ext === "htm") return this.fromMarkup(await file.text(), ext === "xml");
        return await file.text();
    }

    static async fromPDF(file) {
        if (typeof pdfjsLib === "undefined") throw new Error("PDF reader unavailable offline.");
        const pdf = await pdfjsLib.getDocument({ data: await file.arrayBuffer() }).promise;
        let pages = [];
        for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);
            const content = await page.getTextContent();
            let lastY = null, line = "", lines = [];
            for (const item of content.items) {
                const y = item.transform ? Math.round(item.transform[5]) : null;
                if (lastY !== null && y !== null && Math.abs(y - lastY) > 4) {
                    if (line.trim()) lines.push(line.trim());
                    line = item.str;
                } else {
                    line += (line && !line.endsWith(" ") ? " " : "") + item.str;
                }
                lastY = y;
            }
            if (line.trim()) lines.push(line.trim());
            pages.push(lines.join("\n"));
        }
        const out = pages.join("\n").trim();
        if (!out) throw new Error("Scanned image PDF has no selectable text.");
        return out;
    }

    static async fromDOCX(file) {
        if (typeof mammoth === "undefined") throw new Error("Word reader unavailable offline.");
        const res = await mammoth.extractRawText({ arrayBuffer: await file.arrayBuffer() });
        const out = (res.value || "").trim();
        if (!out) throw new Error("Word document is empty.");
        return out;
    }

    static async fromExcel(file) {
        if (typeof XLSX === "undefined") throw new Error("Spreadsheet reader unavailable offline.");
        const wb = XLSX.read(await file.arrayBuffer(), { type: "array" });
        let lines = [];
        wb.SheetNames.forEach(s => {
            const rows = XLSX.utils.sheet_to_json(wb.Sheets[s], { defval: "" });
            rows.forEach((r, idx) => {
                const keys = Object.keys(r);
                if (keys.length === 3 && keys.map(k => k.toLowerCase()).includes("subject")) {
                    lines.push(`${r[keys[0]]} | ${r[keys[1]]} | ${r[keys[2]]}`);
                } else if (keys.length >= 2) {
                    lines.push(`# ${String(r[keys[0]] || `Node_${idx + 1}`).trim()}`);
                    keys.slice(1).forEach(k => { if (String(r[k]).trim()) lines.push(`${k}: ${r[k]}`); });
                }
            });
        });
        return lines.join("\n");
    }

    static splitRow(line, sep) {
        const res = []; let cur = "", q = false;
        for (const c of line) {
            if (c === '"') q = !q;
            else if (c === sep && !q) { res.push(cur.trim()); cur = ""; }
            else cur += c;
        }
        res.push(cur.trim());
        return res;
    }

    static fromCSV(raw, sep) {
        const rows = raw.split(/\r?\n/).map(r => r.trim()).filter(Boolean);
        if (rows.length < 2) return raw;
        const headers = this.splitRow(rows[0], sep);
        let lines = [];
        for (let i = 1; i < rows.length; i++) {
            const cols = this.splitRow(rows[i], sep);
            if (cols.length >= 2) {
                lines.push(`# ${cols[0] || `Row_${i}`}`);
                for (let j = 1; j < headers.length; j++) if (cols[j]) lines.push(`${headers[j]}: ${cols[j]}`);
            }
        }
        return lines.join("\n");
    }

    static fromJSON(raw) {
        const obj = JSON.parse(raw);
        let lines = [];
        const walk = (node, topic) => {
            if (Array.isArray(node)) {
                node.forEach(item => {
                    if (typeof item === "string") lines.push(/^IF\s+/i.test(item) ? item : `${topic}: ${item}`);
                    else if (item && typeof item === "object") walk(item, topic);
                });
            } else if (node && typeof node === "object") {
                for (const [k, v] of Object.entries(node)) {
                    if (v && typeof v === "object" && !Array.isArray(v)) { lines.push(`# ${k}`); walk(v, k); }
                    else if (Array.isArray(v)) {
                        if (k.toLowerCase() === "rules") v.forEach(r => lines.push(String(r)));
                        else v.forEach(val => lines.push(`${k}: ${val}`));
                    } else lines.push(`${k}: ${v}`);
                }
            }
        };
        walk(obj, "Document");
        return lines.join("\n");
    }

    static fromMarkup(raw, isXML) {
        const doc = new DOMParser().parseFromString(raw, isXML ? "text/xml" : "text/html");
        if (doc.querySelector("parsererror")) throw new Error("Invalid XML/HTML tags.");
        let lines = [];
        if (!isXML) {
            doc.querySelectorAll("h1, h2, h3, h4, p, li, tr").forEach(el => {
                const tag = el.tagName.toLowerCase();
                const txt = el.textContent.replace(/\s+/g, " ").trim();
                if (!txt) return;
                if (tag.startsWith("h")) lines.push(`# ${txt}`);
                else if (tag === "tr") {
                    const cells = Array.from(el.querySelectorAll("td, th")).map(c => c.textContent.trim());
                    if (cells.length >= 2) lines.push(`${cells[0]}: ${cells.slice(1).join(" - ")}`);
                } else lines.push(txt);
            });
        } else {
            const walk = el => {
                const kids = Array.from(el.children);
                if (!kids.length) { const v = el.textContent.trim(); if (v) lines.push(`${el.tagName}: ${v}`); }
                else { lines.push(`# ${el.getAttribute("name") || el.getAttribute("id") || el.tagName}`); kids.forEach(walk); }
            };
            if (doc.documentElement) walk(doc.documentElement);
        }
        return lines.join("\n");
    }
}

class DeterministicTechEngine {
    constructor() {
        this.topics = {};
        this.facts = [];
        this.rules = [];
        this.deduced = [];
    }

    clear() {
        this.topics = {};
        this.facts = [];
        this.rules = [];
        this.deduced = [];
    }

    addFact(subject, relation, detail, deduced = false) {
        const s = String(subject || "").trim();
        const r = String(relation || "").trim().toLowerCase().replace(/\s+/g, "_");
        const d = String(detail || "").trim().replace(/\.$/, "");
        if (!s || !r || !d) return;

        const exists = this.facts.some(
            f => f.s.toLowerCase() === s.toLowerCase() && f.r === r && f.d.toLowerCase() === d.toLowerCase()
        );
        if (!exists) {
            this.facts.push({ s, r, d, deduced });
            if (deduced) this.deduced.push({ s, r, d });
        }

        const key = s.toLowerCase();
        if (!this.topics[key]) this.topics[key] = { name: s, details: {} };
        if (!this.topics[key].details[r]) this.topics[key].details[r] = [];
        if (!this.topics[key].details[r].includes(d)) this.topics[key].details[r].push(d);
    }

    parseRuleClause(clause) {
        const clean = clause.trim().replace(/\.$/, "");
        const verbs = [
            " is deployed in ", " is part of ", " is a subset of ", " is defined as ",
            " requires ", " causes ", " uses ", " produces ", " enables ", " supports ",
            " part_of ", " subset_of ", " deployed_in ", " has ", " is "
        ];
        for (const v of verbs) {
            const idx = clean.toLowerCase().indexOf(v);
            if (idx !== -1) {
                return {
                    s: clean.slice(0, idx).trim(),
                    r: v.trim().toLowerCase().replace(/\s+/g, "_"),
                    d: clean.slice(idx + v.length).trim()
                };
            }
        }
        return null;
    }

    build(rawText) {
        this.clear();
        let text = (rawText || "").trim();
        if (!text) return;

        if ((text.startsWith("{") && text.endsWith("}")) || (text.startsWith("[") && text.endsWith("]"))) {
            try { text = MultiFormatAdapter.fromJSON(text); } catch (e) {}
        }

        const lines = text.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
        let curTopic = "System Architecture";

        const sentencePatterns = [
            { rx: /^(.+?)\s+is\s+defined\s+as\s+(.+)$/i, rel: "definition" },
            { rx: /^(.+?)\s+is\s+(?:deployed|located|hosted)\s+in\s+(.+)$/i, rel: "deployed_in" },
            { rx: /^(.+?)\s+is\s+part\s+of\s+(.+)$/i, rel: "part_of" },
            { rx: /^(.+?)\s+is\s+a\s+subset\s+of\s+(.+)$/i, rel: "subset_of" },
            { rx: /^(.+?)\s+requires\s+(.+)$/i, rel: "requires" },
            { rx: /^(.+?)\s+uses\s+(.+)$/i, rel: "uses" },
            { rx: /^(.+?)\s+produces\s+(.+)$/i, rel: "produces" },
            { rx: /^(.+?)\s+causes\s+(.+)$/i, rel: "causes" },
            { rx: /^(.+?)\s+enables\s+(.+)$/i, rel: "enables" },
            { rx: /^(.+?)\s+supports\s+(.+)$/i, rel: "supports" },
            { rx: /^(.+?)\s+has\s+(.+)$/i, rel: "has" },
            { rx: /^(.+?)\s+is\s+a\s+(.+)$/i, rel: "category" },
            { rx: /^(.+?)\s+is\s+(.+)$/i, rel: "is" }
        ];

        for (const line of lines) {
            const rm = line.match(/^IF\s+(.+?)\s+THEN\s+(.+)$/i);
            if (rm) {
                const a = this.parseRuleClause(rm[1]);
                const c = this.parseRuleClause(rm[2]);
                if (a && c) this.rules.push({ if: a, then: c });
                continue;
            }

            const pipe = line.split("|").map(p => p.trim());
            if (pipe.length === 3 && pipe[0] && pipe[1] && pipe[2]) {
                this.addFact(pipe[0], pipe[1], pipe[2]);
                continue;
            }

            if (line.startsWith("#") || (line.endsWith(":") && line.split(/\s+/).length <= 6)) {
                curTopic = line.replace(/^#+/, "").replace(/:$/, "").trim();
                continue;
            }

            const kv = line.match(/^[-*]?\s*([A-Za-z0-9 _/()-]{2,40}):\s*(.+)$/);
            if (kv) {
                this.addFact(curTopic, kv[1], kv[2]);
                continue;
            }

            const clean = line.replace(/^[-*\s]+/, "").replace(/\.$/, "");
            let ok = false;
            for (const p of sentencePatterns) {
                const m = clean.match(p.rx);
                if (m) {
                    this.addFact(m[1], p.rel, m[2]);
                    ok = true;
                    break;
                }
            }
            if (!ok) this.addFact(curTopic, "note", clean);
        }

        this.forwardChain();
    }

    forwardChain() {
        let added = true, guard = 0;
        while (added && guard < 60) {
            added = false;
            guard++;
            for (const rule of this.rules) {
                for (const f of [...this.facts]) {
                    if (f.r === rule.if.r && f.d.toLowerCase() === rule.if.d.toLowerCase()) {
                        const sub = rule.if.s === "?X" ? f.s : rule.if.s;
                        const target = rule.then.s === "?X" ? sub : rule.then.s;
                        const exists = this.facts.some(
                            x => x.s.toLowerCase() === target.toLowerCase() &&
                                 x.r === rule.then.r &&
                                 x.d.toLowerCase() === rule.then.d.toLowerCase()
                        );
                        if (!exists) {
                            this.addFact(target, rule.then.r, rule.then.d, true);
                            added = true;
                        }
                    }
                }
            }
        }
    }

    bfsShortestPath(startTopic, endTopic) {
        const start = startTopic.toLowerCase();
        const end = endTopic.toLowerCase();
        const queue = [[start, []]];
        const seen = new Set([start]);

        while (queue.length) {
            const [cur, path] = queue.shift();
            if (cur === end && path.length) return path;
            for (const f of this.facts) {
                if (f.s.toLowerCase() === cur) {
                    const nxt = f.d.toLowerCase();
                    if (!seen.has(nxt)) {
                        seen.add(nxt);
                        queue.push([nxt, [...path, `${f.s} (${f.r.replace(/_/g, " ")}) -> ${f.d}`]]);
                    }
                }
            }
        }
        return null;
    }

    ask(question) {
        const clean = (question || "").trim().replace(/\?+$/, "").trim();
        const low = clean.toLowerCase();
        const steps = [];

        if (!clean) {
            return { found: false, answer: "Please enter a question first.", method: "Empty Input", steps: ["No question text entered."], score: 0 };
        }
        if (!this.facts.length) {
            return { found: false, answer: "Knowledge Base is empty. Load a document first.", method: "Empty KB", steps: ["0 facts available."], score: 0 };
        }

        const escRx = s => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
        const matchedTopics = Object.entries(this.topics)
            .filter(([k]) => k !== "system architecture" && new RegExp(`(^|[^a-z0-9])${escRx(k)}([^a-z0-9]|$)`, "i").test(low))
            .map(([, v]) => v.name)
            .sort((a, b) => b.length - a.length);

        if (matchedTopics.length >= 2) {
            const [t1, t2] = matchedTopics;
            steps.push(`Identified two entities: "${t1}" and "${t2}".`);
            const chain = this.bfsShortestPath(t1, t2) || this.bfsShortestPath(t2, t1);
            if (chain) {
                steps.push("Connected both entities via shortest-path Breadth-First Search (BFS).");
                return {
                    found: true,
                    answer: `Yes, they are connected: ${chain.join(" -> ")}.`,
                    method: "Multi-Hop BFS Graph Search",
                    steps: [...steps, ...chain],
                    score: 100
                };
            }
        }

        let targetRels = [];
        if (/^(what is|define|explain)/i.test(low)) {
            targetRels = ["definition", "category", "is", "note"];
            steps.push("Intent classified: Definition / State lookup.");
        } else if (/^(where|location|deployed|hosted)/i.test(low)) {
            targetRels = ["deployed_in", "location", "part_of"];
            steps.push("Intent classified: Deployment / Location lookup.");
        } else if (/^(why|cause)/i.test(low)) {
            targetRels = ["causes", "requires"];
            steps.push("Intent classified: Causal / Dependency lookup.");
        } else if (low.includes("require") || low.includes("need")) {
            targetRels = ["requires", "uses"];
            steps.push("Intent classified: Requirement lookup.");
        }

        if (matchedTopics.length) {
            const topic = matchedTopics[0];
            const obj = this.topics[topic.toLowerCase()];
            steps.push(`Matched entity frame: "${topic}".`);

            const sortedRels = Object.keys(obj.details).sort((a, b) => b.length - a.length);
            for (const rel of sortedRels) {
                const readable = rel.replace(/_/g, " ");
                const relRx = new RegExp(`(^|[^a-z0-9])(${escRx(rel)}|${escRx(readable)})([^a-z0-9]|$)`, "i");
                if (relRx.test(low) && rel !== "is" && rel !== "has") {
                    const vals = obj.details[rel];
                    const isRule = vals.some(v => this.deduced.some(df => df.s === topic && df.r === rel && df.d === v));
                    steps.push(`Matched slot "${readable}" -> ${vals.join(", ")}.`);
                    return {
                        found: true,
                        answer: `The ${readable} of ${topic} is: ${vals.join(", ")}.`,
                        method: isRule ? "Forward-Chaining Rule Deduction" : "Direct Frame Slot Lookup",
                        steps,
                        score: 100
                    };
                }
            }

            for (const rel of targetRels) {
                if (obj.details[rel]) {
                    const vals = obj.details[rel];
                    const isRule = vals.some(v => this.deduced.some(df => df.s === topic && df.r === rel && df.d === v));
                    steps.push(`Matched intent slot "${rel.replace(/_/g, " ")}" -> ${vals.join(", ")}.`);
                    return {
                        found: true,
                        answer: `${topic}: ${vals.join(", ")}.`,
                        method: isRule ? "Forward-Chaining Rule Deduction" : "Intent Slot Lookup",
                        steps,
                        score: 100
                    };
                }
            }

            const summary = Object.entries(obj.details).map(([k, v]) => `${k.replace(/_/g, " ")}: ${v.join(", ")}`);
            steps.push("Returned complete entity profile.");
            return {
                found: true,
                answer: `${topic} — ${summary.join("; ")}.`,
                method: "Complete Entity Frame",
                steps,
                score: 95
            };
        }

        const stop = new Set(["what", "where", "when", "which", "does", "have", "about", "tell", "give", "the", "and", "for", "with", "from"]);
        const tokens = (low.match(/[a-z0-9]+/g) || []).filter(w => w.length > 3 && !stop.has(w));
        let best = null, hits = 0;

        for (const f of this.facts) {
            const str = `${f.s} ${f.r.replace(/_/g, " ")} ${f.d}`.toLowerCase();
            const h = tokens.reduce((acc, t) => acc + (new RegExp(`(^|[^a-z0-9])${escRx(t)}([^a-z0-9]|$)`, "i").test(str) ? 1 : 0), 0);
            if (h > hits) { hits = h; best = f; }
        }

        if (best && hits >= 2) {
            steps.push(`Matched ${hits} exact keywords in fact: "${best.s} ${best.r.replace(/_/g, " ")} ${best.d}".`);
            return {
                found: true,
                answer: `${best.s} ${best.r.replace(/_/g, " ")} ${best.d}.`,
                method: "Exact Keyword Fact Match",
                steps,
                score: 90
            };
        }

        return {
            found: false,
            answer: "No matching knowledge found in the document for this query.",
            method: "Safe Abstention (Zero Hallucination)",
            steps: ["Verified against all entity frames, semantic triples, and production rules; abstained to avoid guessing."],
            score: 0
        };
    }

    evaluate() {
        const tests = [
            { q: "What is the Runtime Engine of Kubernetes Cluster?", exp: "containerd" },
            { q: "Where is Kubernetes Cluster deployed?", exp: "AWS Mumbai Region" },
            { q: "What does Vector Database require?", exp: "NVMe SSD Storage" },
            { q: "How is API Gateway related to Kubernetes Cluster?", exp: "Ingress Mesh" },
            { q: "What is High Traffic Spike?", exp: "Autoscaling Trigger" },
            { q: "What is Kubernetes Cluster?", exp: "High Availability Ready" },
            { q: "What is GPU Worker Node?", exp: "Hardware Accelerated" },
            { q: "What is the clock speed of Quantum Processor?", exp: "No matching knowledge found" }
        ];
        let passed = 0;
        const rows = tests.map(t => {
            const r = this.ask(t.q);
            const ok = r.answer.toLowerCase().includes(t.exp.toLowerCase());
            if (ok) passed++;
            return { ...t, actual: r.answer, method: r.method, ok };
        });
        return { total: tests.length, passed, pct: Math.round((passed / tests.length) * 100), rows };
    }
}

const engine = new DeterministicTechEngine();

function switchSub(id, btn) {
    document.querySelectorAll(".sub-view").forEach(v => v.classList.remove("active"));
    document.querySelectorAll(".t-btn").forEach(b => b.classList.remove("active"));
    document.getElementById(id).classList.add("active");
    btn.classList.add("active");
}

function setStatus(msg, isErr = false) {
    const el = document.getElementById("status-bar");
    el.textContent = msg;
    el.className = isErr ? "status-pill err" : "status-pill";
}

async function handleFile(ev) {
    const file = ev.target.files[0];
    if (!file) return;
    setStatus(`Parsing "${file.name}"...`);
    try {
        const text = await MultiFormatAdapter.decode(file);
        document.getElementById("editor").value = text;
        compileKB();
        setStatus(`Loaded "${file.name}" — ${engine.facts.length} facts extracted.`);
    } catch (e) {
        setStatus(`Error: ${e.message}`, true);
    }
    ev.target.value = "";
}

function loadTechSample() {
    document.getElementById("editor").value = TECH_SAMPLE_TEXT;
    compileKB();
    setStatus("Ready — Cloud & AI Architecture sample loaded.");
}

function loadTechJSON() {
    document.getElementById("editor").value = TECH_SAMPLE_JSON;
    compileKB();
    setStatus("Ready — JSON Infrastructure Spec parsed into Knowledge Base.");
}

function compileKB() {
    engine.build(document.getElementById("editor").value);
    document.getElementById("k-topics").textContent = Object.keys(engine.topics).length;
    document.getElementById("k-facts").textContent = engine.facts.length;
    document.getElementById("k-rules").textContent = engine.rules.length;
    document.getElementById("k-deduced").textContent = engine.deduced.length;
    renderTable();
    renderCards();
}

function renderTable() {
    const q = document.getElementById("search-facts").value.trim().toLowerCase();
    const tbody = document.getElementById("facts-tbody");
    tbody.innerHTML = "";
    const list = q ? engine.facts.filter(f => `${f.s} ${f.r.replace(/_/g, " ")} ${f.d}`.toLowerCase().includes(q)) : engine.facts;
    if (!list.length) {
        tbody.innerHTML = `<tr><td colspan="4" style="text-align:center;padding:1rem;">No matching facts found.</td></tr>`;
        return;
    }
    list.forEach(f => {
        const badge = f.deduced
            ? '<span class="tag tag-amber">Rule Deduced</span>'
            : '<span class="tag tag-emerald">Extracted</span>';
        tbody.innerHTML += `<tr><td><strong>${esc(f.s)}</strong></td><td><span class="tag">${esc(f.r.replace(/_/g, " "))}</span></td><td>${esc(f.d)}</td><td>${badge}</td></tr>`;
    });
}

function renderCards() {
    const grid = document.getElementById("entity-grid");
    grid.innerHTML = "";
    Object.values(engine.topics).forEach(t => {
        let rows = "";
        for (const [k, vals] of Object.entries(t.details)) {
            rows += `<div class="kv-row"><span style="color:var(--text-muted)">${esc(k.replace(/_/g, " "))}</span><strong style="color:#fff">${esc(vals.join(", "))}</strong></div>`;
        }
        grid.innerHTML += `<div class="entity-box"><div class="entity-title"><span>${esc(t.name)}</span><span class="tag">${Object.keys(t.details).length}</span></div>${rows}</div>`;
    });
}

function presetAsk(q) {
    document.getElementById("q-input").value = q;
    executeAsk();
}

function executeAsk() {
    const res = engine.ask(document.getElementById("q-input").value);
    const card = document.getElementById("ans-card");
    card.style.display = "block";
    card.className = res.found ? "answer-card" : "answer-card not-found";

    const m = document.getElementById("ans-method");
    m.textContent = res.method;
    m.className = res.found ? "tag" : "tag tag-amber";

    document.getElementById("ans-conf").textContent = res.found ? `${res.score}% Accuracy` : "Abstained";
    document.getElementById("ans-text").textContent = res.answer;
    document.getElementById("ans-trace").innerHTML =
        "<strong style='color:var(--cyan)'>Symbolic Inference Trace:</strong><br>" +
        res.steps.map(s => `• ${esc(s)}`).join("<br>");
}

function runBenchmark() {
    const rep = engine.evaluate();
    const banner = document.getElementById("bench-summary");
    banner.textContent = `Benchmark Result: ${rep.passed} / ${rep.total} Passed (${rep.pct}% Accuracy)`;
    const tbody = document.getElementById("bench-tbody");
    tbody.innerHTML = "";
    rep.rows.forEach(r => {
        const st = r.ok ? '<span class="tag tag-emerald">PASS</span>' : '<span class="tag tag-rose">FAIL</span>';
        tbody.innerHTML += `<tr><td><strong>${esc(r.q)}</strong></td><td>${esc(r.exp)}</td><td>${esc(r.actual)}</td><td><span class="tag">${esc(r.method)}</span></td><td>${st}</td></tr>`;
    });
}

window.addEventListener("DOMContentLoaded", () => {
    loadTechSample();
    runBenchmark();
    const dz = document.getElementById("drop-zone");
    ["dragenter", "dragover"].forEach(e => dz.addEventListener(e, ev => { ev.preventDefault(); dz.classList.add("over"); }));
    ["dragleave", "drop"].forEach(e => dz.addEventListener(e, ev => { ev.preventDefault(); dz.classList.remove("over"); }));
    dz.addEventListener("drop", ev => {
        const f = ev.dataTransfer && ev.dataTransfer.files[0];
        if (f) {
            const dt = new DataTransfer();
            dt.items.add(f);
            const inp = document.getElementById("file-input");
            inp.files = dt.files;
            handleFile({ target: inp });
        }
    });
});
