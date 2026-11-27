// Live findings report for one session, as plain HTML (or Markdown with &format=md).
// Built for people and for AI assistants: everything is ready-made text, no app code needs to run.
//
//   /api/report?session=<session id>&key=<report key>[&only=issues][&format=md]
//
// The report key is set by an admin in Edit session → Report link. No key (or the wrong one) = 404.
const SUPABASE_URL = process.env.SUPABASE_URL || "https://badylykbhmujmeaysixg.supabase.co";
const SUPABASE_KEY = process.env.SUPABASE_KEY || "sb_publishable_lPrrkR3Nodq6LMOrgRnLAQ_6VgKXFkY";

const STATUS_LABEL = { Pass: "Pass", Fail: "Fail", Blocked: "Blocked", "Needs Review": "Needs review", "Not Tested": "Not run" };
const ISSUE = new Set(["Fail", "Blocked", "Needs Review"]);
const STEP_STATE = { done: "Done", review: "Works but needs review", blocked: "Unable to proceed" };
const CHECK_LINE = /^\s*(?:✓|✔|☑|☐|\[\s?[xX✓]?\s?\])\s*/;

async function rest(path) {
  const r = await fetch(`${SUPABASE_URL}/rest/v1/${path}`, {
    headers: { apikey: SUPABASE_KEY, Authorization: `Bearer ${SUPABASE_KEY}` },
  });
  if (!r.ok) {
    const t = await r.text();
    const e = new Error(t || r.statusText);
    e.status = r.status;
    throw e;
  }
  return r.json();
}
const optional = (p) => rest(p).catch(() => []);

const esc = (s) => String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const has = (v) => v != null && String(v).trim() !== "" && String(v).trim() !== "—";
const expectedItems = (txt) => {
  if (!txt || !String(txt).split(/\n/).some((l) => CHECK_LINE.test(l))) return null;
  const items = [];
  for (const l of String(txt).split(/\n/)) {
    if (CHECK_LINE.test(l)) items.push(l.replace(CHECK_LINE, "").trim());
    else if (l.trim() && items.length) items[items.length - 1] += ` ${l.trim()}`;
  }
  return items.filter(Boolean);
};
const fixLabel = (m) =>
  m.status === "in_progress" ? "In progress" : m.status === "fixed" ? (m.fix_stage === "not_merged" ? "Fixed (not on the test site yet)" : "Ready to retest") : m.status || "Note";
const day = (t) => (t ? new Date(t).toISOString().replace("T", " ").slice(0, 16) + " UTC" : "");

export default async function handler(req, res) {
  const q = req.query || {};
  const sessionId = String(q.session || "");
  const key = String(q.key || "");
  const md = String(q.format || "").toLowerCase() === "md";
  const onlyIssues = String(q.only || "").toLowerCase() === "issues";
  res.setHeader("Cache-Control", "no-store");
  res.setHeader("X-Robots-Tag", "noindex, nofollow");
  const notFound = () => {
    res.status(404).setHeader("Content-Type", "text/plain; charset=utf-8");
    res.send("Report not found. The link may have been reset or turned off.");
  };
  if (!/^[0-9a-f-]{36}$/i.test(sessionId) || key.length < 20) return notFound();

  let session;
  try {
    [session] = await rest(`sessions?id=eq.${sessionId}&select=*`);
  } catch {
    return notFound();
  }
  if (!session || !session.report_key || session.report_key !== key) return notFound();

  const [tests, findings, marks, stepReplies] = await Promise.all([
    rest(`test_cases?session_id=eq.${sessionId}&select=*&order=position.asc`),
    rest(`findings?session_id=eq.${sessionId}&select=*`),
    rest(`done_marks?session_id=eq.${sessionId}&select=*`),
    optional(`dev_step_notes?session_id=eq.${sessionId}&select=*`),
  ]);

  const testers = Array.isArray(session.testers) ? session.testers : [];
  const byTest = (rows) => rows.reduce((a, r) => ((a[r.test_case_id] ||= []).push(r), a), {});
  const F = byTest(findings), M = byTest(marks), D = byTest(stepReplies);

  // A test counts as an issue if any tester marked it Fail, Blocked or Needs review.
  const isIssue = (t) => (F[t.id] || []).some((f) => ISSUE.has(f.status));
  const shown = onlyIssues ? tests.filter(isIssue) : tests;
  const counts = {};
  findings.forEach((f) => (counts[f.status || "Not Tested"] = (counts[f.status || "Not Tested"] || 0) + 1));

  // Build a simple document model, then print it as HTML or Markdown.
  const blocks = [];
  const H = (lvl, text) => blocks.push({ t: "h", lvl, text });
  const P = (text) => blocks.push({ t: "p", text });
  const KV = (k, v) => blocks.push({ t: "kv", k, v });
  const UL = (items) => items.length && blocks.push({ t: "ul", items });
  const LINK = (label, url) => ({ label, url });

  H(1, `${session.name || "UAT session"}: findings report`);
  P(`Generated ${day(new Date())}. This page is live: it always shows the latest results.`);
  KV("Tests", `${tests.length}${onlyIssues ? ` (showing ${shown.length} with issues)` : ""}`);
  KV("Testers", testers.join(", ") || "—");
  KV("Results logged", Object.entries(counts).map(([s, n]) => `${STATUS_LABEL[s] || s} ${n}`).join(" · ") || "None yet");
  const issueList = tests.filter(isIssue);
  if (issueList.length) {
    H(2, "Needs attention");
    UL(
      issueList.map((t) => {
        const st = (F[t.id] || []).filter((f) => ISSUE.has(f.status)).map((f) => `${f.tester_name}: ${STATUS_LABEL[f.status]}${f.bug_severity ? ` (${f.bug_severity})` : ""}`);
        return `${t.test_id || "—"} ${t.title || ""} · ${st.join("; ")}`;
      }),
    );
  }
  H(2, onlyIssues ? "Tests with issues" : "All tests");

  for (const t of shown) {
    H(3, `${t.test_id || "—"}: ${t.title || "Untitled"}`);
    const meta = [t.module && `Module: ${t.module}`, has(t.feature) && `Feature: ${t.feature}`, t.priority && `Priority: ${t.priority}`, t.focus && `Focus area: ${t.focus}`, t.dev_branch_name && `Dev branch: ${t.dev_branch_name}`, t.initial_tester && `Assigned to: ${t.initial_tester}`].filter(Boolean);
    if (meta.length) P(meta.join(" · "));
    if (has(t.pre_conditions)) KV("Pre-conditions", t.pre_conditions);
    if (has(t.test_data)) KV("Test data", t.test_data);
    const steps = Array.isArray(t.steps) ? t.steps.filter((s) => String(s || "").trim()) : [];
    if (steps.length) {
      blocks.push({ t: "label", text: "Steps" });
      blocks.push({ t: "ol", items: steps });
    }
    const exp = expectedItems(t.expected_result);
    if (exp) {
      blocks.push({ t: "label", text: "Expected result" });
      UL(exp);
    } else if (has(t.expected_result)) KV("Expected result", t.expected_result);
    if (has(t.downstream_check)) KV("Downstream check", t.downstream_check);
    if (has(t.edge_case)) KV("Edge case", t.edge_case);

    const results = (F[t.id] || []).filter((f) => (f.status && f.status !== "Not Tested") || has(f.notes) || Object.keys(f.step_notes || {}).length);
    if (!results.length) P("Tester results: not run yet.");
    for (const f of results) {
      H(4, `Result from ${f.tester_name}: ${STATUS_LABEL[f.status] || f.status || "Not run"}${f.bug_severity ? ` · Severity: ${f.bug_severity}` : ""}`);
      if (f.updated_at) P(`Last edited ${day(f.updated_at)}${f.updated_by ? ` by ${f.updated_by}` : ""}`);
      if (has(f.notes)) KV("Tester notes", f.notes);
      if (has(f.other_observations)) KV("Other observations", f.other_observations);
      const checks = f.step_checks || {};
      const sn = f.step_notes || {};
      const stepLines = steps
        .map((s, i) => {
          const parts = [];
          if (STEP_STATE[checks[i]]) parts.push(STEP_STATE[checks[i]]);
          if (has(sn[i])) parts.push(`note: ${sn[i]}`);
          const replies = (D[t.id] || []).filter((r) => r.step_key === String(i) && has(r.note));
          replies.forEach((r) => parts.push(`reply from ${r.developer_name} (developer): ${r.note}`));
          return parts.length ? `Step ${i + 1}: ${parts.join(" · ")}` : null;
        })
        .filter(Boolean);
      (exp || []).forEach((item, i) => {
        const k = `exp-${i}`;
        const parts = [];
        if (checks[k] === "done") parts.push("As expected");
        else if (checks[k] === "blocked") parts.push("Not as expected");
        else if (checks[k] === "review") parts.push("Unsure");
        if (has(sn[k])) parts.push(`note: ${sn[k]}`);
        if (parts.length) stepLines.push(`Expected "${item}": ${parts.join(" · ")}`);
      });
      if (stepLines.length) {
        blocks.push({ t: "label", text: "Step by step" });
        UL(stepLines);
      }
      const shots = (f.screenshots || []).filter((s) => s && s.url);
      const recs = (f.recordings || []).filter((r) => r && r.url);
      if (shots.length || recs.length) {
        blocks.push({ t: "label", text: "Evidence" });
        UL([
          ...shots.map((s, i) => LINK(`Fig. ${i + 1} (screenshot)`, s.url)),
          ...recs.map((r, i) => LINK(`Recording ${i + 1}`, r.url)),
        ]);
      }
    }
    const dm = (M[t.id] || []).filter((m) => m.status || has(m.note));
    if (dm.length) {
      blocks.push({ t: "label", text: "Developers" });
      UL(dm.map((m) => `${m.developer_name}: ${fixLabel(m)}${m.done_at ? ` (${day(m.done_at)})` : ""}${has(m.note) ? `. Note: ${m.note}` : ""}`));
    }
  }

  const inline = (x) => (typeof x === "object" && x ? `[${x.label}](${x.url})` : String(x));
  if (md) {
    const out = blocks
      .map((b) => {
        if (b.t === "h") return `${"#".repeat(b.lvl)} ${b.text}\n`;
        if (b.t === "p") return `${b.text}\n`;
        if (b.t === "label") return `**${b.text}**`;
        if (b.t === "kv") return `**${b.k}:** ${String(b.v).replace(/\n/g, "  \n")}\n`;
        if (b.t === "ul") return b.items.map((i) => `- ${inline(i).replace(/\n/g, " ")}`).join("\n") + "\n";
        if (b.t === "ol") return b.items.map((i, n) => `${n + 1}. ${String(i).replace(/\n/g, " ")}`).join("\n") + "\n";
        return "";
      })
      .join("\n");
    res.status(200).setHeader("Content-Type", "text/markdown; charset=utf-8");
    return res.send(out);
  }
  const li = (x) => (typeof x === "object" && x ? `<a href="${esc(x.url)}">${esc(x.label)}</a>` : esc(x).replace(/\n/g, "<br>"));
  const body = blocks
    .map((b) => {
      if (b.t === "h") return `<h${b.lvl}>${esc(b.text)}</h${b.lvl}>`;
      if (b.t === "p") return `<p class="meta">${esc(b.text)}</p>`;
      if (b.t === "label") return `<p class="label">${esc(b.text)}</p>`;
      if (b.t === "kv") return `<p><strong>${esc(b.k)}:</strong> ${esc(b.v).replace(/\n/g, "<br>")}</p>`;
      if (b.t === "ul") return `<ul>${b.items.map((i) => `<li>${li(i)}</li>`).join("")}</ul>`;
      if (b.t === "ol") return `<ol>${b.items.map((i) => `<li>${li(i)}</li>`).join("")}</ol>`;
      return "";
    })
    .join("\n");
  const mdLink = `?${new URLSearchParams({ ...q, format: "md" }).toString()}`;
  const html = `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><meta name="robots" content="noindex, nofollow">
<title>${esc(session.name || "UAT")}: findings report</title>
<style>body{font:15px/1.6 -apple-system,"Segoe UI",system-ui,sans-serif;max-width:880px;margin:0 auto;padding:32px 20px;color:#15181D;background:#fff}h1{font-size:24px;margin:0 0 4px}h2{font-size:19px;margin:36px 0 8px;padding-top:16px;border-top:2px solid #15181D}h3{font-size:16px;margin:28px 0 4px;padding-top:12px;border-top:1px solid #E2E5EA}h4{font-size:14px;margin:16px 0 4px}p{margin:4px 0}.meta{color:#59616D;font-size:13px}.label{font-weight:600;margin-top:10px}ul,ol{margin:4px 0 8px;padding-left:22px}a{color:#2F7FE0}</style></head>
<body><p class="meta">Also available as <a href="${esc(mdLink)}">plain text (Markdown)</a>.</p>
${body}
</body></html>`;
  res.status(200).setHeader("Content-Type", "text/html; charset=utf-8");
  return res.send(html);
}
