// Test Case Collaborator — editable source (Pro + Light editions).
// Rebuilt on 2026-09-25 from the live build (GitHub commit db05d95, 2 Jun 2026).
// Edit this file, then run `npm install` once and `npm run build` to produce index.html.
// Note: local variable names inside functions are short (e, t, a...) because they
// were recovered from the compiled app; component and helper names are descriptive.

import React, {
  Fragment,
  createContext,
  useContext,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
} from "react";
import * as ReactDOM from "react-dom/client";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import JSZip from "jszip";
import { createClient } from "@supabase/supabase-js";
import {
  Clock,
  CircleCheck,
  CircleX,
  CircleAlert,
  AlertTriangle,
  Paperclip,
  Image as ImageIcon,
  Circle as CircleIcon,
  Type,
  Lock,
  BookOpen,
  RefreshCw,
  CircleHelp,
  ExternalLink,
  Check,
  Copy,
  PanelLeftClose,
  Search,
  X,
  Filter,
  Users,
  PanelLeftOpen,
  ChevronLeft,
  ChevronRight,
  Home,
  RotateCcw,
  Download,
  FileText,
  FileDown,
  Plus,
  Settings2,
  GripVertical,
  Pencil,
  Trash2,
  Upload,
  ArrowLeft,
  Monitor,
  GitBranch,
} from "lucide-react";
var EDITION = "Pro";
var SUPABASE_URL = "https://badylykbhmujmeaysixg.supabase.co";
var SUPABASE_KEY = "sb_publishable_lPrrkR3Nodq6LMOrgRnLAQ_6VgKXFkY";
var supabase = EDITION === "Pro" ? createClient(SUPABASE_URL, SUPABASE_KEY) : null;
var STATUS_OPTIONS = [
  {
    value: "Not Tested",
    color: "bg-gray-700 text-gray-300 border-gray-600",
    dot: "bg-gray-500",
    icon: Clock,
  },
  {
    value: "Pass",
    color: "bg-green-900/40 text-green-300 border-green-700",
    dot: "bg-green-500",
    icon: CircleCheck,
  },
  {
    value: "Fail",
    color: "bg-red-900/40 text-red-300 border-red-700",
    dot: "bg-red-500",
    icon: CircleX,
  },
  {
    value: "Blocked",
    color: "bg-orange-900/40 text-orange-300 border-orange-700",
    dot: "bg-orange-500",
    icon: CircleAlert,
  },
  {
    value: "Needs Review",
    color: "bg-yellow-900/40 text-yellow-300 border-yellow-700",
    dot: "bg-yellow-500",
    icon: CircleHelp,
  },
];
var getStatusConfig = (e) => STATUS_OPTIONS.find((t) => t.value === e) || STATUS_OPTIONS[0];
var SURFACE_DARK = "rgb(20 24 35)";
var STORAGE_LIMIT_BYTES = 5242880;
var STORAGE_WARN_AT = 0.8;
var makeFinding = (e = {}) => ({
  status: "Not Tested",
  testedBy: "",
  notes: "",
  bugSeverity: "",
  otherObservations: "",
  screenshots: [],
  recordings: [],
  stepNotes: {},
  stepChecks: {},
  stepNoteOpen: {},
  ...e,
});
var isEmptyValue = (e) => {
  if (e == null) {
    return true;
  }
  let t = String(e).trim();
  return t === "" || t === "—" || t === "nan" || t === "NaN";
};
var parseTesterMap = (e) => {
  if (!e) {
    return {};
  }
  let t = e.match(/\[([^\]]+)\]/);
  if (!t) {
    return {};
  }
  let a = {};
  t[1].split(",").forEach((r) => {
    let o = r.trim().match(/^(\S+)\s*-\s*(.+)$/);
    if (o) {
      a[o[1].trim()] = o[2].trim();
    }
  });
  return a;
};
var isFailureStatus = (e) => {
  if (!e) {
    return false;
  }
  let t = String(e).trim().toLowerCase();
  return t === "fail" || t === "blocked" || t === "fail/blocked" || t === "fail/ blocked";
};
// Anything a developer may need to look at: failed, blocked, or "works but needs review".
var isIssueStatus = (e) => isFailureStatus(e) || String(e || "").trim().toLowerCase() === "needs review";
var stepsToText = (e) =>
  (e || []).map((t, a) => `${a + 1}. ${t}`).join(`
`);
var prettyArrows = (e) => (typeof e == "string" && e.indexOf("->") !== -1 ? e.replace(/->/g, "→") : e);
var timeAgo = (e) => {
  if (!e) {
    return "";
  }
  let t = Date.now() - new Date(e).getTime();
  if (t < 0) {
    return "just now";
  }
  let a = Math.floor(t / 1000);
  if (a < 45) {
    return "just now";
  }
  let r = Math.floor(a / 60);
  if (r < 60) {
    return `${r}m ago`;
  }
  let o = Math.floor(r / 60);
  if (o < 24) {
    return `${o}h ago`;
  }
  let s = Math.floor(o / 24);
  if (s < 30) {
    return `${s}d ago`;
  } else {
    return new Date(e).toLocaleDateString();
  }
};
var parseRecordingUrl = (e) => {
  let t = String(e || "").trim();
  if (!t) {
    return null;
  }
  let a = t.match(/loom\.com\/(?:share|embed)\/([a-zA-Z0-9]+)/);
  if (a) {
    return {
      host: "loom",
      label: "Loom",
      embed: `https://www.loom.com/embed/${a[1]}`,
    };
  }
  a = t.match(/drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (a) {
    return {
      host: "drive",
      label: "Google Drive",
      embed: `https://drive.google.com/file/d/${a[1]}/preview`,
    };
  }
  a = t.match(/drive\.google\.com\/open\?id=([a-zA-Z0-9_-]+)/);
  if (a) {
    return {
      host: "drive",
      label: "Google Drive",
      embed: `https://drive.google.com/file/d/${a[1]}/preview`,
    };
  }
  a = t.match(/(?:youtube\.com\/watch\?(?:[^&]*&)*v=|youtube\.com\/embed\/)([a-zA-Z0-9_-]{6,})/);
  if (a) {
    return {
      host: "youtube",
      label: "YouTube",
      embed: `https://www.youtube.com/embed/${a[1]}`,
    };
  }
  a = t.match(/youtu\.be\/([a-zA-Z0-9_-]{6,})/);
  if (a) {
    return {
      host: "youtube",
      label: "YouTube",
      embed: `https://www.youtube.com/embed/${a[1]}`,
    };
  }
  a = t.match(/vimeo\.com\/(?:video\/)?(\d+)/);
  if (a) {
    return {
      host: "vimeo",
      label: "Vimeo",
      embed: `https://player.vimeo.com/video/${a[1]}`,
    };
  }
  if (/\.(mp4|webm|mov|m4v|ogg)(\?|#|$)/i.test(t)) {
    return {
      host: "file",
      label: "Video file",
      embed: null,
      isVideoFile: true,
    };
  }
  try {
    return {
      host: "link",
      label: new URL(t).hostname.replace(/^www\./, ""),
      embed: null,
    };
  } catch {
    return {
      host: "link",
      label: "Link",
      embed: null,
    };
  }
};
var CSV_GENERATION_PROMPT =
  'You are generating a User Acceptance Testing (UAT) plan as a CSV file. The CSV will be reviewed by non-technical stakeholders (product owners, QA, business users) and executed manually by testers. Follow the rules below exactly. Consistency across UAT sheets matters more than cleverness — these conventions exist so any tester on any team can pick up the sheet and execute it without translation.\n\n## OUTPUT FORMAT\n\nOutput ONE CSV file and nothing else — no preamble, no explanation, no markdown fences. The first row is the header, exactly:\n\nID,Module,Feature,Title,Tester [1- <Name1>, 2- <Name2>, 3- <Name3>],Pre-conditions,Test Data,Steps,Expected Result,Downstream Check,Edge Case,Pass/Fail,Tester Notes,Bug Severity (P0/ P1 / P2),Other Observations\n\nSubstitute the real tester names supplied in APP CONTEXT into the Tester column header. If no testers are supplied, leave the placeholders.\n\nEvery subsequent row is one test case. CSV rules:\n- Comma-separated. Quote any field that contains a comma, newline, or double quote.\n- Inside quoted fields, escape double quotes by doubling them (`""`).\n- Newlines inside Steps / Test Data / Expected Result / Pre-conditions cells are real newline characters inside the quoted field. Multi-line cells are normal — Steps almost always spans multiple lines.\n- Leave Pass/Fail, Tester Notes, Bug Severity, and Other Observations EMPTY. Tester gets the assigned tester\'s NUMBER (1, 2, 3...) — author-assigned, not filled in by the tester. These are reviewed and updated by testers during execution.\n- For optional columns where no value applies, write `—` (em-dash) rather than leaving completely empty — keeps cell counts consistent for spreadsheet review.\n\n## COLUMN DEFINITIONS\n\n**ID** — `<phase-code>-<area-code><nn>` format. Phase code identifies the testing phase or release (`T1`, `T2`, `T3` for sequenced UAT cycles, or `R4` for a release). Area code is a **1–3 letter** shortcode for the feature area — examples seen in real sheets: `L` (Leads), `Q` (Quotes), `P` (Projects), `PO` (Purchase Orders), `ME` (Multi-entity), `AB` (mAIson Bot), `PCH` (Portal Config Homeowner). `nn` is a two-digit sequence within that area, starting at `01`. Examples: `T1-L03`, `T2-PO07`, `T3-PCH02`. Sequence within an area; do not reset across modules. Use the same phase code for an entire test cycle. Same letter may mean different things in different phases — that\'s fine, area codes are scoped to phase.\n\n**Module** — A high-level feature/epic name, usually a single word in Title Case (`Finance`, `Projects`, `Customization`, `Leads`, `Quotes`, `Homeowner`, `Notifications`, `Integration`, `Tools`). Hyphenated compounds are fine (`Multi-entity`), as is product-native capitalization (`mAIson Bot`, `Portal Config`). Group related tests under the same Module value across consecutive rows; do not skip back and forth. A typical UAT sheet has 10–20 modules.\n\n**Feature** — A sub-area inside the Module, sentence case. Mix of styles is fine and reflects what testers will see:\n- Imperative verb phrase: `Create ticket`, `Edit line item`, `Approve invoice`, `Reject with comment`\n- Noun phrase naming a UI surface or workflow: `Status lifecycle`, `Bell drawer`, `Settings page`, `Activity tab`, `AR Aging`\n- UI variant in parens to disambiguate two paths to the same thing: `Create lead (drawer)` vs `Create lead (page)`\n- Colon for sub-action or function call: `Tool call: overdue tasks`, `Write tool: create_lead`\n- Named product component: `EntityPicker reports`, `Composer milestone`, `Gantt drag-resize`\nUse `—` when the test spans the whole module or isn\'t scoped to a single sub-feature (e.g. an end-to-end integration test).\n\n**Title** — A one-line description (no period at the end) that names WHAT is being verified, not how. Be specific and dense:\n- Good: `Create escalation via /escalations/new`\n- Good: `Move escalation through statuses`\n- Good: `Record commission payout`\n- Bad: `Test that creating an escalation works` (verbose, says "test that")\n- Use `→` for sequence, `+` to join multiple verified things in one test, `—` (em-dash) to attach a qualifier or scope.\n\n**Tester [1- <Name1>, 2- <Name2>, ...]** — The header lists each tester\'s number→name mapping. Each row\'s cell contains just the NUMBER (`1`, `2`, `3`) of the assigned tester. Distribute load roughly evenly across testers, but cluster related cases (same Module/Feature) under the same tester so they keep context.\n\n**Pre-conditions** — Compact. Name accounts/data/feature flags needed. Examples:\n- `On /escalations page`\n- `Ticket from E01 exists`\n- `Admin account. Commission entry accrued.`\n- `Feature flag "new_checkout" enabled.`\nUse `—` if no setup is required beyond a logged-in user.\n\n**Test Data** — Concrete sample inputs the tester should use. Multi-line if multiple fields. Examples:\n- `Subject: T1 - Test escalation\\nPriority: Medium\\nCategory: Timeline`\n- `Amount: $500\\nDate: today\\nReference: TEST-001`\nUse `—` if the test doesn\'t require seeded data.\n\n**Steps** — A numbered list, one step per line, beginning each line with `<n>. ` (digit, period, space). Style rules:\n- Imperative voice: `Click`, `Open`, `Navigate to`, `Confirm`. Never `You should click`.\n- Concrete UI paths: `/escalations → \'New Ticket\'`, `Settings → Team → "Invite"`.\n- Quote literal UI text in single quotes inside steps (which avoid the need to double-escape inside the CSV).\n- Inline edge-case probes count as steps: `Try to save without required fields. Confirm BLOCKED with field-level errors.`\n\n**Expected Result** — Prose, not a list. Describe the observable outcome in 1–3 sentences. Style rules:\n- Lead with the success state, then qualifying detail.\n- Mention IDs/formats that should appear: `Ticket created. Auto-generated ticket_id (T-YYYY-NNNNNN). Toast confirms.`\n- Quote literal UI strings the tester should see (banner copy, button text, toast text, error messages).\n- Use SHOUTY caps sparingly for invariants: `BLOCKED`, `ALWAYS`, `regardless of toggle`.\n\n**Downstream Check** — What to verify outside the immediate flow once the action succeeds. Examples: a DB column, an audit log entry, an email sent, a downstream system updated, an ID format conformance. Use `—` if there\'s nothing to verify downstream.\n\n**Edge Case** — Specific edge cases or boundary conditions worth probing in this test. Short note, not a step. Examples: `empty cart`, `max length input`, `cleared on reopen`. Use `—` if no edge case applies.\n\n**Pass/Fail, Tester Notes, Bug Severity (P0/ P1 / P2), Other Observations** — Always blank in the authored CSV. Testers fill these in during execution.\n\nReference (for the AUTHOR\'s awareness only — do not pre-fill these):\n- **Pass/Fail** values testers use in practice: `Pass`, `Fail`, `Blocked`, `Fail/ Blocked` (when a failure also blocks downstream tests), `Partial Pass` (some sub-checks pass), or empty (untested). Authors leave it empty.\n- **Bug Severity** values: `P0` (critical / blocker / data corruption), `P1` (major functional break), `P2` (minor / cosmetic). Only filled when status is `Fail` or `Blocked`, and even then frequently left empty — only ~1% of rows tend to carry a severity in observed sheets.\n- **Tester Notes** captures the main observation when something fails or behaves unexpectedly.\n- **Other Observations** is for side-channel notes the tester wants to flag — UX issues, performance hiccups, unrelated bugs spotted during testing.\n\n## CONTENT RULES\n\n1. **One test = one verifiable behavior.** If you\'re tempted to write `and also verify X` for an unrelated check, that\'s a second test case. Combining is allowed when checks share setup and flow.\n\n2. **Cover the negative path.** When a feature has a guard (refusal, disabled button, hidden control, validation error), add a step that probes it and an Expected Result that names the guard.\n\n3. **Cover the toggle/branch matrix.** When a feature has variants (ON/OFF, role-gated, default/non-default), write steps for each within the same test if they share setup, otherwise as adjacent tests.\n\n4. **Cross-surface verification when relevant.** For features touching multiple surfaces (in-app + email + public page + API), list each surface as a step and require Expected Result to confirm consistency.\n\n5. **Reference ticket/spec codes inline** when the test maps to a known spec item — in the Title or as a Steps prefix.\n\n6. **Sequence tests logically.** Setup tests come before tests that depend on them. Use Pre-conditions to make dependencies explicit (`Ticket from E01 exists`).\n\n## TONE\n\nTerse, technical, no fluff. Read like an engineer wrote it for another engineer to execute. No "please", no "make sure to", no marketing adjectives. UI strings get quotes; everything else is bare.\n\n## WHAT YOU WILL RECEIVE\n\nBelow this prompt the requester will paste two sections:\n\n**APP CONTEXT** — info about the product so you can ground the tests:\n- What the app does (one paragraph)\n- Tester names and their number assignment for the Tester column header (e.g. "Testers: 1-Ruchit, 2-Oliva, 3-Dane")\n- Any standing concepts (multi-tenant? feature flags? environments?) that recur across tests\n- Module/Feature label conventions if continuing an existing sheet (e.g. "use Module=Finance, starting T2-C04")\n\n**FEATURE SCOPE** — what to generate tests for. Any of these work:\n- A feature spec / PRD section\n- A list of tickets or PRs with brief descriptions\n- A bullet list of changes shipped\n- Release notes / changelog entries\n- A raw `git diff`, PR diff, or commit-range summary\n- Source code of changed files\n- A mix of the above\n\nWhen given code or a diff, treat it as the source of truth for behavior. Read what changed, infer the user-visible outcome, and write tests for the outcome — not for the code structure.\n\nYour job: derive Module + Feature labels, enumerate test cases that fully cover the behavior (happy path + negative paths + toggle branches + cross-surface consistency), assign each to a tester by number, and output the CSV.\n\nIf APP CONTEXT is missing details:\n- Default Pre-conditions to the minimum realistic setup — never write `N/A` or `None`. Use `—` only when truly nothing is required.\n- Default Test Data to `—` when the test action doesn\'t need seeded inputs.\n- Pick sensible Module/Feature labels from the feature scope wording.\n- Distribute testers round-robin if no assignment guidance is given.\n\nDo not invent features the input doesn\'t describe — but DO add the standard guard/negative-path checks any responsible UAT plan would include for the described feature.\n\nAfter the CSV, output nothing.\n\n---\n\n# APP CONTEXT (paste below)\n\n<one-paragraph app description>\n<list of personas / test accounts>\n<recurring concepts: multi-tenant, feature flags, environments>\n<starting TC number and module label conventions, if continuing an existing sheet>\n\n# FEATURE SCOPE (paste below)\n\n<paste ANY of: spec / tickets / PR notes / release notes / git diff / PR diff / source code of changed files>';
// Matches app paths like /quotes/<id>/preview or /tools/x?leadId=<lead id>.
// Placeholders in angle brackets may contain spaces (e.g. <lead id>) and stay inside the chip.
var PATH_TOKEN_RE =
  /(?<![a-zA-Z0-9_\/:.])\/[a-zA-Z][a-zA-Z0-9_-]*(?:\/(?:[a-zA-Z0-9_.-]|<[a-zA-Z0-9_][a-zA-Z0-9_ .-]{0,38}>)+)*\/?(?:\?(?:[a-zA-Z0-9_=&.%\-]|<[a-zA-Z0-9_][a-zA-Z0-9_ .-]{0,38}>)+)?/g;
var BaseUrlContext = createContext("");
function PathChipText({ text: e }) {
  let t = useContext(BaseUrlContext);
  let a = t ? t.replace(/\/+$/, "") : "";
  let r = a ? `${a}${e}` : e;
  let o = !!a;
  let [s, l] = useState(false);
  let i = async (u) => {
    u.stopPropagation();
    u.preventDefault();
    try {
      await navigator.clipboard.writeText(e);
      l(true);
      setTimeout(() => l(false), 1200);
    } catch (d) {
      console.warn("Copy to clipboard failed:", d);
    }
  };
  if (o) {
    return (
      <a
        href={r}
        target="_blank"
        rel="noopener noreferrer"
        onClick={(u) => u.stopPropagation()}
        title={`Open in new tab: ${r}`}
        className="inline-flex items-center gap-1 font-mono text-[12px] border rounded-sm px-1 py-[1px] cursor-pointer transition-colors align-baseline break-all no-underline bg-gray-800 text-indigo-300 border-gray-700/70 hover:bg-gray-700 hover:border-indigo-600 hover:text-indigo-200"
      >
        <span className="hover:underline underline-offset-2 decoration-indigo-400/50">{e}</span>
        <ExternalLink className="w-2.5 h-2.5 flex-shrink-0 opacity-60" />
      </a>
    );
  } else {
    return (
      <code
        role="button"
        tabIndex={0}
        onClick={i}
        onKeyDown={(u) => {
          if (u.key === "Enter" || u.key === " ") {
            u.preventDefault();
            i(u);
          }
        }}
        title={s ? "Copied" : `Click to copy: ${e}`}
        className={`inline-flex items-center gap-1 font-mono text-[12px] border rounded-sm px-1 py-[1px] cursor-pointer transition-colors align-baseline break-all select-none ${s ? "bg-emerald-900/50 text-emerald-200 border-emerald-700" : "bg-gray-800 text-indigo-300 border-gray-700/70 hover:bg-gray-700 hover:border-indigo-600 hover:text-indigo-200"}`}
      >
        <span>{e}</span>
        {s ? (
          <Check className="w-3 h-3 flex-shrink-0" />
        ) : (
          <Copy className="w-2.5 h-2.5 flex-shrink-0 opacity-50" />
        )}
      </code>
    );
  }
}
// "[Fig. N]" in notes links to the Nth screenshot of the test (see FigureContext / MentionTextarea).
var FIG_TOKEN_RE = /\[Fig\. (\d+)\]/g;
var FigureContext = createContext(null);
function FigChip({ n: e }) {
  let t = useContext(FigureContext);
  let a = t?.shots?.[e - 1];
  if (!a || !t.open) {
    return `[Fig. ${e}]`;
  }
  return (
    <button
      type="button"
      onClick={(r) => {
        r.preventDefault();
        r.stopPropagation();
        t.open(e - 1);
      }}
      className="inline-flex items-center gap-1 align-baseline mx-0.5 px-1.5 py-px text-[11px] font-mono font-semibold bg-sky-950/60 text-sky-300 border border-sky-800 rounded-sm transition-colors hover:bg-sky-900/60 hover:text-sky-100"
      title={`Open ${a.name || `Fig. ${e}`}`}
    >
      <ImageIcon className="w-3 h-3" />
      Fig. {e}
    </button>
  );
}
// Renumbers "[Fig. N]" references after screenshot number `e` (1-based) is removed.
var renumberFigRefs = (e, t) =>
  typeof e != "string"
    ? e
    : e.replace(FIG_TOKEN_RE, (a, r) => {
        let o = parseInt(r, 10);
        if (o === t) return "[image removed]";
        return o > t ? `[Fig. ${o - 1}]` : a;
      });
var normalizePath = (e) => {
  if (e == null) {
    return e;
  }
  let t = String(e);
  if (!t.includes("[Fig. ")) {
    return normalizePathOnly(t);
  }
  let a = [];
  let r = 0;
  let o;
  for (FIG_TOKEN_RE.lastIndex = 0; (o = FIG_TOKEN_RE.exec(t)) !== null;) {
    if (o.index > r) {
      a.push(<Fragment key={`t${r}`}>{normalizePathOnly(t.substring(r, o.index))}</Fragment>);
    }
    a.push(<FigChip n={parseInt(o[1], 10)} key={`f${o.index}`} />);
    r = o.index + o[0].length;
  }
  if (r < t.length) {
    a.push(<Fragment key={`t${r}`}>{normalizePathOnly(t.substring(r))}</Fragment>);
  }
  return a;
};
// Textarea that offers the test's screenshots when you type "@" (e.g. "@F" → Fig. 1), inserting "[Fig. N]".
function MentionTextarea({ value: e, onChange: t, onKeyDown: a, onBlur: r, showRefs: o = true, ...s }) {
  let l = useContext(FigureContext);
  let i = useRef(null);
  let [u, d] = useState(null);
  let [c, p] = useState(0);
  let h = l?.shots || [];
  let b = (x) => {
    let g = x.selectionStart ?? x.value.length;
    let v = x.value.slice(0, g);
    let k = /(^|\s)@(fig\.?\s?\d{0,3}|[^\s@]{0,20})$/i.exec(v);
    if (k) {
      d({ start: g - k[2].length - 1, end: g, query: k[2] });
    } else {
      d(null);
    }
    p(0);
  };
  let _ = u
    ? h
        .map((x, g) => ({ idx: g, name: x.name || "" }))
        .filter((x) => {
          let g = u.query.toLowerCase().replace(/[^a-z0-9]/g, "");
          if (!g) return true;
          let v = `fig${x.idx + 1}`;
          return v.startsWith(g) || String(x.idx + 1).startsWith(g) || x.name.toLowerCase().replace(/[^a-z0-9]/g, "").includes(g);
        })
    : [];
  let S = (x) => {
    if (!u) return;
    let g = `[Fig. ${x + 1}] `;
    let v = (e || "").slice(0, u.start) + g + (e || "").slice(u.end);
    let k = u.start + g.length;
    d(null);
    t({ target: { value: v } });
    requestAnimationFrame(() => {
      let C = i.current;
      if (C) {
        C.focus();
        C.setSelectionRange(k, k);
      }
    });
  };
  let I = [];
  if (o && e && e.includes("[Fig. ")) {
    let x;
    for (FIG_TOKEN_RE.lastIndex = 0; (x = FIG_TOKEN_RE.exec(e)) !== null;) {
      let g = parseInt(x[1], 10);
      if (!I.includes(g)) I.push(g);
    }
  }
  return (
    <div className="relative">
      <textarea
        ref={i}
        value={e}
        onChange={(x) => {
          t(x);
          b(x.target);
        }}
        onKeyDown={(x) => {
          if (u) {
            if (x.key === "Escape") {
              x.preventDefault();
              x.stopPropagation();
              d(null);
              return;
            }
            if (_.length > 0) {
              if (x.key === "ArrowDown") {
                x.preventDefault();
                p((g) => (g + 1) % _.length);
                return;
              }
              if (x.key === "ArrowUp") {
                x.preventDefault();
                p((g) => (g - 1 + _.length) % _.length);
                return;
              }
              if (x.key === "Enter" || x.key === "Tab") {
                x.preventDefault();
                S(_[Math.min(c, _.length - 1)].idx);
                return;
              }
            }
          }
          a?.(x);
        }}
        onClick={(x) => b(x.target)}
        onBlur={(x) => {
          setTimeout(() => d(null), 120);
          r?.(x);
        }}
        {...s}
      />
      {u && (
        <div className="absolute left-0 right-0 top-full mt-1 z-30 bg-gray-900 border border-gray-700 rounded-sm shadow-2xl overflow-hidden">
          {h.length === 0 ? (
            <div className="px-3 py-2 text-xs text-gray-400">
              No screenshots on this test yet. Upload one under Screenshots, then type @ to link it.
            </div>
          ) : _.length === 0 ? (
            <div className="px-3 py-2 text-xs text-gray-400">No screenshot matches "{u.query}".</div>
          ) : (
            <div className="max-h-56 overflow-y-auto py-1">
              {_.map((x, g) => (
                <button
                  type="button"
                  key={x.idx}
                  onMouseDown={(v) => {
                    v.preventDefault();
                    S(x.idx);
                  }}
                  onMouseEnter={() => p(g)}
                  className={`w-full flex items-center gap-2.5 px-2.5 py-1.5 text-left text-xs ${g === c ? "bg-gray-800 text-gray-100" : "text-gray-300"}`}
                >
                  <img
                    src={h[x.idx].url}
                    alt=""
                    className="w-10 h-7 object-cover rounded-sm border border-gray-700 flex-shrink-0 bg-gray-950"
                  />
                  <span className="font-mono font-semibold text-sky-300 flex-shrink-0">Fig. {x.idx + 1}</span>
                  <span className="truncate text-gray-500">{x.name}</span>
                </button>
              ))}
            </div>
          )}
          <div className="px-2.5 py-1 border-t border-gray-800 text-[10px] text-gray-500">
            ↑↓ to choose · Enter to insert · Esc to close
          </div>
        </div>
      )}
      {I.length > 0 && (
        <div className="flex flex-wrap items-center gap-1 mt-1">
          <span className="text-[10px] text-gray-500">Linked:</span>
          {I.map((x) => (
            <FigChip n={x} key={x} />
          ))}
        </div>
      )}
    </div>
  );
}
// Full web addresses (https://…) are shown as one plain link, never split into path chips.
var FULL_URL_RE = /\bhttps?:\/\/[^\s<>"'`]+/g;
var normalizePathOnly = (e) => {
  if (e == null) {
    return e;
  }
  let t = String(e);
  if (!t.includes("/")) {
    return t;
  }
  if (/https?:\/\//i.test(t)) {
    let a = [];
    let r = 0;
    let o;
    for (FULL_URL_RE.lastIndex = 0; (o = FULL_URL_RE.exec(t)) !== null;) {
      let l = o[0].replace(/[.,;:!?)\]]+$/, "");
      if (o.index > r) {
        a.push(<Fragment key={`s${r}`}>{normalizePathOnlyPlain(t.substring(r, o.index))}</Fragment>);
      }
      a.push(
        <a
          key={`u${o.index}`}
          href={l}
          target="_blank"
          rel="noopener noreferrer"
          onClick={(i) => i.stopPropagation()}
          className="text-sky-300 underline decoration-sky-700 underline-offset-2 break-all hover:text-sky-200"
        >
          {l}
        </a>,
      );
      r = o.index + l.length;
      FULL_URL_RE.lastIndex = r;
    }
    if (r < t.length) {
      a.push(<Fragment key={`s${r}`}>{normalizePathOnlyPlain(t.substring(r))}</Fragment>);
    }
    return a;
  }
  return normalizePathOnlyPlain(t);
};
var normalizePathOnlyPlain = (e) => {
  let t = String(e);
  if (!t.includes("/")) {
    return t;
  }
  let a = [];
  let r = 0;
  let o;
  for (PATH_TOKEN_RE.lastIndex = 0; (o = PATH_TOKEN_RE.exec(t)) !== null;) {
    if (o.index > r) {
      a.push(t.substring(r, o.index));
    }
    a.push(<PathChipText text={o[0]} key={`p${o.index}`} />);
    r = o.index + o[0].length;
  }
  if (r === 0) {
    return t;
  } else {
    if (r < t.length) {
      a.push(t.substring(r));
    }
    return a;
  }
};
var parseSteps = (e) => {
  if (!e || typeof e != "string") {
    return [];
  }
  let t = e.trim();
  if (!t) {
    return [];
  }
  let a = /^\s*(?:\(?\d+[\.\)\:]|step\s+\d+[\.\)\:]?|[-*•])\s+/i;
  let r = (i) => i.replace(a, "").trim();
  let o = t
    .split(/\n+/)
    .map((i) => i.trim())
    .filter((i) => i.length > 0);
  if (o.length === 0) {
    return [];
  }
  if (!o.some((i) => a.test(i))) {
    return o;
  }
  let l = [];
  for (let i of o) {
    if (a.test(i)) {
      l.push(r(i));
    } else if (l.length > 0) {
      l[l.length - 1] += `
${i}`;
    } else {
      l.push(i);
    }
  }
  return l.filter((i) => i.length > 0);
};
// ---------------------------------------------------------------------------
// Test case import: CSV files and multi-tab spreadsheets (.xlsx).
// Column headers are matched loosely, so "Test" works as "Title",
// "Preconditions" as "Pre-conditions", "Expected result" as "Expected Result", etc.
// ---------------------------------------------------------------------------
var normalizeHeader = (e) => String(e ?? "").toLowerCase().replace(/[^a-z0-9]/g, "");
var HEADER_ALIASES = {
  testId: ["id", "testid", "testcaseid", "tcid", "caseid", "testno", "no", "step", "stepno", "stepid"],
  module: ["module", "epic"],
  feature: ["feature", "featurearea", "subfeature", "area", "scene"],
  title: ["title", "test", "testcase", "testtitle", "testcasetitle", "testname", "scenario", "testscenario", "trythis"],
  preConditions: ["preconditions", "precondition", "prerequisites", "prerequisite", "setup"],
  testData: ["testdata", "data"],
  steps: ["steps", "teststeps", "stepstoreproduce", "procedure", "how", "howto", "instructions"],
  expectedResult: ["expectedresult", "expectedresults", "expected", "expectedoutcome", "youshouldsee", "shouldsee"],
  downstreamCheck: ["downstreamcheck", "downstreamchecks", "downstream"],
  edgeCase: ["edgecase", "edgecases"],
  devBranchName: ["devbranchname", "devbranch", "branch", "branchname"],
  focus: ["focus", "focusarea", "testfocus", "category", "testtype"],
  priority: ["priority", "prio", "testpriority"],
  // Folded into other fields: "Log in as" → Pre-conditions, "Files to use" → Test Data.
  loginAs: ["loginas", "youare", "persona", "loginuser"],
  filesToUse: ["filestouse", "files", "samplefiles", "filesneeded"],
};
var resolveColumns = (e) => {
  let t = e.map((r) => String(r ?? "").trim());
  let a = t.map(normalizeHeader);
  let r = {};
  for (let [o, l] of Object.entries(HEADER_ALIASES)) {
    let i = a.findIndex((u) => u && l.includes(u));
    if (i >= 0) {
      r[o] = t[i];
    }
  }
  let o =
    t.find((l) => /^tester\s*\[/i.test(l)) ||
    t.find((l, i) => ["tester", "testers", "assignee", "assignedto", "assignedtester"].includes(a[i]));
  if (o) {
    r.tester = o;
  }
  return r;
};
var fileBaseName = (e) =>
  String(e || "")
    .replace(/\.[a-z0-9]+$/i, "")
    .replace(/[_]/g, " ")
    .trim();
// rows: array of objects keyed by header text, each with __rowNum (sheet/CSV row number)
var rowsToTestCases = (e, t, a) => {
  let r = resolveColumns(t);
  let o = [];
  if (!r.testId) o.push("ID");
  if (!r.title && !r.steps) o.push("Title (or Test) / Steps");
  if (o.length > 0) {
    return {
      error: `Missing required column(s): ${o.join(", ")}.`,
    };
  }
  let l = parseTesterMap(r.tester);
  let i = (d, c) => (r[c] ? String(d[r[c]] ?? "").trim() : "");
  let u = e
    .filter((d) => i(d, "title") || i(d, "steps"))
    .map((d, c) => ({
      position: c,
      rowNum: d.__rowNum,
      testId: i(d, "testId"),
      module: i(d, "module") || a.fallbackModule || "",
      feature: i(d, "feature"),
      title: i(d, "title"),
      preConditions: [i(d, "loginAs") && `Log in as: ${i(d, "loginAs")}`, i(d, "preConditions")]
        .filter(Boolean)
        .join("\n"),
      testData: [i(d, "testData"), i(d, "filesToUse") && `Files: ${i(d, "filesToUse").replace(/\s*\n\s*/g, ", ")}`]
        .filter(Boolean)
        .join("\n"),
      steps: parseSteps(i(d, "steps")),
      expectedResult: i(d, "expectedResult"),
      downstreamCheck: i(d, "downstreamCheck"),
      edgeCase: i(d, "edgeCase"),
      devBranchName: i(d, "devBranchName"),
      focus: i(d, "focus"),
      priority: i(d, "priority"),
      initialTester: (() => {
        let p = i(d, "tester");
        if (!p || p === "—") {
          return "";
        }
        return l[p] || p;
      })(),
    }));
  if (u.length === 0) {
    return {
      error: "No test case rows found (rows need a Title/Test or Steps value).",
    };
  }
  let d = Object.values(l);
  if (d.length === 0) {
    // No "Tester [1 - Name, ...]" mapping: use names typed directly in the tester column.
    d = Array.from(new Set(u.map((c) => c.initialTester).filter((c) => c && !/^\d+$/.test(c))));
  }
  return {
    testCases: u,
    testers: d,
    testerColHeader: r.tester || "Tester",
  };
};
var parseCsvFile = (e) =>
  new Promise((t) => {
    if (!e) {
      t({
        error: "No file provided.",
      });
      return;
    }
    Papa.parse(e, {
      header: true,
      dynamicTyping: false,
      skipEmptyLines: true,
      transformHeader: (a) => a.trim(),
      complete: (a) => {
        let r = (a.errors || []).filter(
          (p) => p.type === "Quotes" || p.type === "FieldMismatch" || p.code === "UndetectableDelimiter",
        );
        if (r.length > 0) {
          t({
            error: `CSV could not be parsed: ${r[0].message || r[0].code}. Make sure your file is a valid CSV (no markdown fences, no preamble text).`,
          });
          return;
        }
        let o = (a.meta?.fields || []).map((p) => p.trim());
        let l = a.data.map((p, h) => ({
          ...p,
          __rowNum: h + 2,
        }));
        t(
          rowsToTestCases(l, o, {
            fallbackModule: fileBaseName(e.name),
          }),
        );
      },
      error: (a) =>
        t({
          error: `Parse failed: ${a.message || "unknown error"}`,
        }),
    });
  });
// "Read me": instructions shown to testers in a side panel. Stored as light markup in sessions.readme:
//   "# Heading" lines, "Term :: description" lines, plain paragraphs, blank lines between groups.
var README_TAB_NAMES = /(start\s*here|read\s*me|instruction|guide|how\s*to)/i;
var sheetRowsToReadme = (e) => {
  let t = [];
  let a = true;
  e.forEach((r) => {
    let o = (r || []).map((l) => String(l ?? "").replace(/\s+$/g, "").trim()).filter(Boolean);
    if (o.length === 0) {
      if (t.length && t[t.length - 1] !== "") t.push("");
      return;
    }
    if (o.length === 1) {
      let l = o[0].replace(/\n+/g, " ");
      if (a) {
        t.push(`# ${l}`);
      } else if (l.length <= 40 && !/[.!?:,…]$/.test(l)) {
        if (t.length && t[t.length - 1] !== "") t.push("");
        t.push(`## ${l}`);
      } else {
        t.push(l);
      }
    } else {
      t.push(`${o[0].replace(/\n+/g, " ")} :: ${o.slice(1).join(" · ").replace(/\n+/g, " ")}`);
    }
    a = false;
  });
  while (t.length && t[t.length - 1] === "") t.pop();
  return t.join("\n");
};
var README_SETUP_SQL = "alter table sessions add column if not exists readme text;";
var isReadmeSetupError = (e) => /readme/i.test(String(e?.message || e || ""));
// Tabs whose names look like instructions, overviews, issue logs or question lists start unticked
// (the admin can still tick them).
var NON_TEST_TAB_NAMES =
  /(start\s*here|summary|overview|read\s*me|instruction|cover|guide|index|dashboard|issue|bug|question|not\s*built|\blog\b|changelog)/i;
// Returns one entry per tab: { name, testCases, testers, testerColHeader } or { name, skipped }
var parseWorkbookFile = async (e) => {
  let t;
  try {
    let a = await e.arrayBuffer();
    t = XLSX.read(a, {
      type: "array",
    });
  } catch (a) {
    return [
      {
        name: e.name,
        skipped: `Couldn't open this file: ${a.message || "unknown error"}`,
      },
    ];
  }
  let a = t.Workbook?.Sheets || [];
  return t.SheetNames.map((r) => {
    let o = t.Sheets[r];
    let l = a.find((h) => h.name === r)?.Hidden;
    if (!o || !o["!ref"]) {
      return {
        name: r,
        skipped: "Empty tab",
      };
    }
    let i = XLSX.utils.decode_range(o["!ref"]).s.r;
    let u = XLSX.utils.sheet_to_json(o, {
      header: 1,
      defval: "",
      raw: false,
      blankrows: true,
    });
    let d = u.slice(0, 25).findIndex((h) => {
      let b = resolveColumns(h);
      return b.testId && (b.title || b.steps);
    });
    if (d < 0) {
      return {
        name: r,
        skipped: "No test table found",
        readmeText: README_TAB_NAMES.test(r.trim()) ? sheetRowsToReadme(u) : undefined,
      };
    }
    let c = u[d].map((h) => String(h ?? "").trim());
    let p = u.slice(d + 1).map((h, b) => {
      let _ = {
        __rowNum: i + d + 2 + b,
      };
      c.forEach((S, I) => {
        if (S && !(S in _)) {
          _[S] = h[I] ?? "";
        }
      });
      return _;
    });
    let h = rowsToTestCases(p, c, {
      fallbackModule: r.trim(),
    });
    if (h.error) {
      return {
        name: r,
        skipped: h.error,
      };
    }
    return {
      name: r,
      ...h,
      hidden: !!l,
      defaultOff: !!l || NON_TEST_TAB_NAMES.test(r.trim()),
    };
  });
};
// Turns a list of dropped files into import "sources" (one per CSV file or spreadsheet tab).
var parseImportFiles = async (e) => {
  let t = [];
  for (let a of e) {
    if (/\.(xlsx|xlsm|xls|ods)$/i.test(a.name)) {
      t.push(...(await parseWorkbookFile(a)));
    } else {
      let r = await parseCsvFile(a);
      t.push(
        r.error
          ? {
              name: a.name,
              skipped: r.error,
            }
          : {
              name: a.name,
              ...r,
            },
      );
    }
  }
  return t.map((a, r) => ({
    ...a,
    key: `${r}:${a.name}`,
    include: !a.skipped && !a.defaultOff,
  }));
};
// Combines the ticked sources into one list of test cases for the session.
var combineImportSources = (e) => {
  let t = e.filter((l) => l.include && l.testCases);
  if (t.length === 0) {
    return null;
  }
  let a = t.flatMap((l) => l.testCases).map((l, i) => ({
    ...l,
    position: i,
  }));
  let r = Array.from(new Set(t.flatMap((l) => l.testers || [])));
  let o = t.find((l) => l.testerColHeader && l.testerColHeader !== "Tester")?.testerColHeader || "Tester";
  let l = {};
  a.forEach((i) => {
    if (i.testId) l[i.testId] = (l[i.testId] || 0) + 1;
  });
  return {
    testCases: a,
    testers: r,
    testerColHeader: o,
    duplicateIds: Object.keys(l).filter((i) => l[i] > 1),
  };
};
// ---------------------------------------------------------------------------
// Re-upload an updated sheet into an existing session ("resync").
// Tests are matched by module (tab) + test ID; results always stay attached.
// ---------------------------------------------------------------------------
var RESYNC_FIELDS = [
  "module",
  "feature",
  "title",
  "pre_conditions",
  "test_data",
  "steps",
  "expected_result",
  "downstream_check",
  "edge_case",
  "dev_branch_name",
  "focus",
  "priority",
];
var RESYNC_FIELD_LABELS = {
  module: "tab",
  feature: "feature",
  title: "title",
  pre_conditions: "preconditions",
  test_data: "test data",
  steps: "steps",
  expected_result: "expected result",
  downstream_check: "downstream check",
  edge_case: "edge case",
  dev_branch_name: "branch",
  focus: "focus area",
  priority: "priority",
};
var importedTestToRow = (e) => ({
  test_id: e.testId,
  module: e.module,
  feature: e.feature,
  title: e.title,
  pre_conditions: e.preConditions,
  test_data: e.testData,
  steps: e.steps,
  expected_result: e.expectedResult,
  downstream_check: e.downstreamCheck,
  edge_case: e.edgeCase,
  dev_branch_name: e.devBranchName,
  focus: e.focus || null,
  priority: e.priority || null,
  row_num: e.rowNum,
  initial_tester: e.initialTester,
});
var resyncNorm = (e) =>
  String(e ?? "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
var resyncFieldValue = (e) => {
  if (e == null) return "";
  if (Array.isArray(e)) return JSON.stringify(e.map((t) => String(t ?? "").replace(/\s+/g, " ").trim()));
  if (typeof e === "object") return JSON.stringify(e);
  return String(e).replace(/\r\n/g, "\n").trim();
};
// Moves step ticks/notes to wherever the same wording now sits; drops them where the wording changed.
var remapChecksByText = (e, t, a, r, o) => {
  let l = (u) => resyncNorm(u);
  let i = {};
  let d = new Set();
  e.forEach((u, c) => {
    let p = -1;
    if (c < t.length && l(t[c]) === l(u) && !d.has(c)) {
      p = c;
    } else {
      p = t.findIndex((h, b) => !d.has(b) && l(h) === l(u));
    }
    if (p >= 0) {
      d.add(p);
      i[c] = p;
    }
  });
  let s = (u) => {
    let c = {};
    let p = 0;
    for (let [h, b] of Object.entries(u || {})) {
      let _ = o(h);
      if (_ === null) {
        c[h] = b;
        continue;
      }
      if (i[_] !== undefined) {
        c[r(i[_])] = b;
      } else {
        p++;
      }
    }
    return [c, p];
  };
  let [n, m] = s(a.checks);
  let [f, g] = s(a.notes);
  return { checks: n, notes: f, dropped: m + g };
};
var stepKeyIndex = (e) => (/^\d+$/.test(e) ? parseInt(e, 10) : null);
var expKeyIndex = (e) => {
  let t = /^exp-(\d+)$/.exec(e);
  return t ? parseInt(t[1], 10) : null;
};
var expectedItemsOf = (e) => parseExpectedChecklist(e)?.items || [];
// existing: all test_cases rows in the session (including removed). sources: from parseImportFiles.
var computeResyncPlan = (e, t) => {
  let a = t.filter((v) => v.include && v.testCases);
  let r = a.flatMap((v) => v.testCases.map((k) => ({ ...k, source: v.name })));
  let o = new Set(r.map((v) => resyncNorm(v.module)));
  let s = (v, k) => {
    let C = {};
    v.forEach((m) => {
      let f = resyncNorm(k(m));
      if (f) C[f] = (C[f] || 0) + 1;
    });
    return C;
  };
  let l = s(e, (v) => v.test_id);
  let i = s(r, (v) => v.testId);
  let u = new Set();
  let d = r.map((v) => {
    let k = resyncNorm(v.testId);
    let C = resyncNorm(v.module);
    let m = e.find(
      (w) =>
        !u.has(w.id) &&
        resyncNorm(w.module) === C &&
        (k ? resyncNorm(w.test_id) === k : !resyncNorm(w.test_id) && resyncNorm(w.title) === resyncNorm(v.title)),
    );
    let f = false;
    if (!m && k && i[k] === 1 && l[k] === 1) {
      m = e.find((w) => !u.has(w.id) && resyncNorm(w.test_id) === k);
      f = !!m;
    }
    if (m) u.add(m.id);
    let g = importedTestToRow(v);
    if (!m) return { incoming: v, row: g, kind: "new" };
    let w = RESYNC_FIELDS.filter((T) => resyncFieldValue(m[T]) !== resyncFieldValue(g[T]));
    return {
      incoming: v,
      row: g,
      existing: m,
      kind: "match",
      moved: f || resyncNorm(m.module) !== C,
      restored: !!m.removed_at,
      changedFields: w,
      contentChanged: w.includes("steps") || w.includes("expected_result"),
    };
  });
  let c = e.filter((v) => !u.has(v.id) && !v.removed_at && o.has(resyncNorm(v.module)));
  let p = e.filter((v) => !u.has(v.id) && !v.removed_at && !o.has(resyncNorm(v.module)));
  // Final order: keep the session's tab order; ticked tabs follow the sheet's order inside them.
  let h = [...e].filter((v) => !v.removed_at).sort((v, k) => (v.position ?? 0) - (k.position ?? 0));
  let b = [];
  let _ = new Set();
  h.forEach((v) => {
    let k = resyncNorm(v.module);
    if (!_.has(k)) {
      _.add(k);
      b.push({ key: k, label: v.module });
    }
  });
  let S = [];
  r.forEach((v) => {
    let k = resyncNorm(v.module);
    if (!S.some((C) => C.key === k)) S.push({ key: k, label: v.module });
  });
  S.forEach((v) => {
    if (_.has(v.key)) return;
    let k = d.find((m) => resyncNorm(m.incoming.module) === v.key && m.existing && !m.existing.removed_at);
    let C = k ? b.findIndex((m) => m.key === resyncNorm(k.existing.module)) : -1;
    if (C >= 0) b.splice(C + 1, 0, v);
    else b.push(v);
    _.add(v.key);
  });
  let I = [];
  b.forEach((v) => {
    if (o.has(v.key)) {
      d.filter((k) => resyncNorm(k.incoming.module) === v.key).forEach((k) => I.push(k));
    } else {
      p.filter((k) => resyncNorm(k.module) === v.key).forEach((k) => I.push({ kind: "keep", existing: k }));
    }
  });
  let x = new Map();
  I.forEach((v, k) => {
    if (v.existing) x.set(v.existing.id, k);
    v.position = k;
  });
  let perSource = a.map((v) => {
    let k = d.filter((m) => m.incoming.source === v.name);
    let C = (m) => k.filter(m).length;
    let m = e.some((f) => !f.removed_at && k.some((g) => resyncNorm(g.incoming.module) === resyncNorm(f.module)));
    return {
      name: v.name,
      isNewTab: !m && !k.some((f) => f.kind === "match"),
      total: k.length,
      added: C((f) => f.kind === "new"),
      updated: C((f) => f.kind === "match" && f.changedFields.length > 0 && !f.moved && !f.restored),
      moved: C((f) => f.kind === "match" && f.moved),
      restored: C((f) => f.kind === "match" && f.restored),
      removed: c.filter((f) => k.some((g) => resyncNorm(g.incoming.module) === resyncNorm(f.module))).length,
    };
  });
  let leftAlone = [];
  p.forEach((v) => {
    if (!leftAlone.includes(v.module)) leftAlone.push(v.module);
  });
  return {
    entries: d,
    ordered: I,
    added: d.filter((v) => v.kind === "new"),
    updated: d.filter((v) => v.kind === "match" && v.changedFields.length > 0 && !v.moved && !v.restored),
    moved: d.filter((v) => v.kind === "match" && v.moved),
    restored: d.filter((v) => v.kind === "match" && v.restored),
    unchanged: d.filter((v) => v.kind === "match" && v.changedFields.length === 0 && !v.restored),
    removed: c,
    leftAlone,
    leftAloneCount: p.length,
    perSource,
    hasChanges:
      d.some((v) => v.kind === "new" || v.restored || v.changedFields.length > 0) ||
      c.length > 0 ||
      I.some((v) => v.existing && v.existing.position !== v.position),
  };
};
// Writes the plan: inserts new tests, updates changed ones, hides removed ones, and keeps
// every finding attached (step/expected ticks follow their wording).
var applyResyncPlan = async (e, t, a) => {
  let r = new Date().toISOString();
  let o = {};
  (a || []).forEach((m) => {
    o[m.test_case_id] = m;
  });
  let l = [];
  let i = [];
  let u = [];
  let d = 0;
  let c = 0;
  t.ordered.forEach((m) => {
    if (m.kind === "new") {
      l.push({ session_id: e.id, position: m.position, ...m.row });
      return;
    }
    let f = m.existing;
    let g = { ...f, position: m.position };
    if (m.kind === "match") {
      let { initial_tester: w, ...T } = m.row;
      Object.assign(g, T, { removed_at: null });
      let A = o[f.id];
      if (m.contentChanged && A) {
        let j = { checks: A.step_checks || {}, notes: A.step_notes || {} };
        let M = remapChecksByText(f.steps || [], m.row.steps || [], j, (V) => String(V), stepKeyIndex);
        let W = remapChecksByText(
          expectedItemsOf(f.expected_result),
          expectedItemsOf(m.row.expected_result),
          M,
          (V) => `exp-${V}`,
          expKeyIndex,
        );
        let F = M.dropped + W.dropped;
        if (JSON.stringify(W.checks) !== JSON.stringify(A.step_checks || {}) || JSON.stringify(W.notes) !== JSON.stringify(A.step_notes || {})) {
          u.push({ ...A, step_checks: W.checks, step_notes: W.notes });
        }
        d += F;
        let se = (A.status && A.status !== "Not Tested") || Object.keys(A.step_checks || {}).length > 0 || Object.keys(A.step_notes || {}).length > 0;
        if (se) {
          g.content_updated_at = r;
          c++;
        }
      }
    }
    if (JSON.stringify(g) !== JSON.stringify(f)) i.push(g);
  });
  t.removed.forEach((m) => {
    i.push({ ...m, removed_at: r, position: t.ordered.length + (m.position ?? 0) });
  });
  let p = async (m, f, g) => {
    for (let w = 0; w < f.length; w += 200) {
      let { error: T } = await g(supabase.from(m), f.slice(w, w + 200));
      if (T) throw T;
    }
  };
  await p("test_cases", i, (m, f) => m.upsert(f));
  await p("test_cases", l, (m, f) => m.insert(f));
  await p("findings", u, (m, f) => m.upsert(f));
  return { inserted: l.length, updated: i.length, droppedTicks: d, flagged: c };
};
// True when the sheet changed this test's steps/expected result after the tester last touched it.
var sheetUpdateUnseen = (e, t) =>
  !!e?.content_updated_at && (!t?.updated_at || new Date(t.updated_at) < new Date(e.content_updated_at));
function LightApp() {
  let [e, t] = useState([]);
  let [a, r] = useState({});
  let [o, s] = useState(0);
  let [l, i] = useState(true);
  let [u, d] = useState("");
  let [c, p] = useState("All");
  let [h, b] = useState(null);
  let [_, S] = useState(false);
  let [I, x] = useState("all");
  let [g, v] = useState(false);
  let [k, C] = useState(false);
  let [m, f] = useState([]);
  let [w, T] = useState("");
  let [A, j] = useState(false);
  let [M, W] = useState(false);
  let [V, F] = useState(false);
  let [se, E] = useState(null);
  let [D, O] = useState(null);
  let [B, ne] = useState(false);
  let [he, ie] = useState(0);
  let [ue, pe] = useState("");
  let [me, Y] = useState(null);
  let fe = useRef(null);
  let Ce = useRef(null);
  let be = useRef(null);
  let xe = useRef(null);
  useEffect(() => {
    (async () => {
      try {
        let L = await window.storage.get("test_data");
        if (L?.value) {
          let R = JSON.parse(L.value);
          let J = (R.testCases || []).map((Z, le) => ({
            ...Z,
            __rowNum: Z.__rowNum || le + 2,
            __steps:
              Z.__steps && Array.isArray(Z.__steps) && Z.__steps.length > 0 ? Z.__steps : parseSteps(Z.Steps),
          }));
          t(J);
          r(R.findings || {});
          s(R.currentIdx || 0);
          f(R.testers || []);
          T(R.lastUsedTester || "");
          if (Array.isArray(R.focusedIds) && R.focusedIds.length > 0) {
            b(new Set(R.focusedIds));
          }
        }
      } catch {
      } finally {
        v(true);
      }
    })();
  }, []);
  useEffect(() => {
    x(h ? "focused" : "all");
  }, [h]);
  useEffect(() => {
    if (!g || e.length === 0) {
      return;
    }
    let L = setTimeout(async () => {
      try {
        await window.storage.set(
          "test_data",
          JSON.stringify({
            testCases: e,
            findings: a,
            currentIdx: o,
            testers: m,
            lastUsedTester: w,
            focusedIds: h ? [...h] : null,
          }),
        );
      } catch (R) {
        console.error("Save failed", R);
      }
    }, 500);
    return () => clearTimeout(L);
  }, [e, a, o, m, w, h, g]);
  let $ = (L, R) =>
    L
      ? new Promise((J) => {
          Papa.parse(L, {
            header: true,
            dynamicTyping: false,
            skipEmptyLines: true,
            transformHeader: (Z) => Z.trim(),
            complete: (Z) => {
              let le = (Z.errors || []).filter(
                (Te) =>
                  Te.type === "Quotes" || Te.type === "FieldMismatch" || Te.code === "UndetectableDelimiter",
              );
              if (le.length > 0) {
                J({
                  error: `CSV could not be parsed: ${le[0].message || le[0].code}. Make sure your file is a valid CSV (no markdown fences, no preamble text).`,
                });
                return;
              }
              let ke = (Z.meta?.fields || []).map((Te) => Te.trim());
              let Ee = ["Module", "ID", "Title", "Steps"].filter((Te) => !ke.includes(Te));
              if (Ee.length > 0) {
                J({
                  error: `Missing required column(s): ${Ee.join(", ")}. The first row of your CSV must contain the headers: ID, Module, Feature, Title, Tester [...], Pre-conditions, Test Data, Steps, Expected Result, Downstream Check, Edge Case, Pass/Fail, Tester Notes, Bug Severity (P0/ P1 / P2), Other Observations.`,
                });
                return;
              }
              let at = ke.find((Te) => /^Tester\s*\[/i.test(Te)) || "Tester";
              let He = parseTesterMap(at);
              let Ne = Object.values(He);
              let _e = Z.data
                .map((Te, wa) => ({
                  ...Te,
                  __rowNum: wa + 2,
                }))
                .filter((Te) =>
                  Object.entries(Te).some(([wa, gr]) => !wa.startsWith("__") && gr && String(gr).trim()),
                )
                .map((Te, wa) => ({
                  ...Te,
                  __id: Te.ID || `row-${wa}`,
                  __steps: parseSteps(Te.Steps),
                  __testerCol: at,
                }));
              if (_e.length === 0) {
                J({
                  error:
                    "The CSV has the correct headers but no test case rows were found. Make sure each row has data in at least the Module, ID, Title, or Steps columns.",
                });
                return;
              }
              t(_e);
              let Ae = {};
              _e.forEach((Te) => {
                let wa = Te["Pass/Fail"];
                let gr = isEmptyValue(wa) ? "Not Tested" : String(wa).trim();
                let Ba = Te[at];
                let xr = isEmptyValue(Ba) ? "" : He[String(Ba).trim()] || String(Ba).trim();
                let vi = isEmptyValue(Te["Tester Notes"]) ? "" : Te["Tester Notes"];
                let mn = isEmptyValue(Te["Bug Severity (P0/ P1 / P2)"])
                  ? ""
                  : String(Te["Bug Severity (P0/ P1 / P2)"]).trim();
                let Xg = isEmptyValue(Te["Other Observations"]) ? "" : Te["Other Observations"];
                let Nc = Te.Recordings;
                let Qg = isEmptyValue(Nc)
                  ? []
                  : String(Nc)
                      .split(/[;\n]/)
                      .map((yi) => yi.trim())
                      .filter(Boolean)
                      .map((yi) => ({
                        url: yi,
                      }));
                Ae[Te.__id] = makeFinding({
                  status: gr,
                  testedBy: xr,
                  notes: vi,
                  bugSeverity: mn,
                  otherObservations: Xg,
                  recordings: Qg,
                });
              });
              r(Ae);
              let hr = [...new Set([...(R || []), ...Ne])];
              f(hr);
              s(0);
              b(null);
              J({
                success: true,
                count: _e.length,
              });
            },
            error: (Z) => {
              J({
                error: "Error reading CSV file: " + Z.message,
              });
            },
          });
        })
      : {
          error: "No file provided.",
        };
  let H = (L, R, J) => {
    r((Z) => {
      let ke = {
        ...(Z[L] || makeFinding()),
        [R]: J,
      };
      if (R === "status" && !isFailureStatus(J)) {
        ke.bugSeverity = "";
      }
      return {
        ...Z,
        [L]: ke,
      };
    });
  };
  let K = (L, R, J) => {
    r((Z) => {
      let le = Z[L] || makeFinding();
      return {
        ...Z,
        [L]: {
          ...le,
          stepNotes: {
            ...(le.stepNotes || {}),
            [R]: J,
          },
        },
      };
    });
  };
  let te = (L, R) => {
    r((J) => {
      let Z = J[L] || makeFinding();
      let le = (Z.stepChecks || {})[R];
      let ke = {
        ...(Z.stepChecks || {}),
      };
      if (le === "done" || le === true) {
        delete ke[R];
      } else {
        ke[R] = "done";
      }
      return {
        ...J,
        [L]: {
          ...Z,
          stepChecks: ke,
        },
      };
    });
  };
  let P = (L, R) => {
    H(L, "testedBy", R);
    if (R) {
      T(R);
    }
  };
  let X = (L) => {
    t((R) =>
      R.map((J, Z) => {
        if (Z !== o) {
          return J;
        }
        let le = L([...(J.__steps || [])]);
        return {
          ...J,
          __steps: le,
          Steps: stepsToText(le),
        };
      }),
    );
  };
  let ve = (L) => {
    let R = e[o]?.__id;
    if (!R) {
      return;
    }
    let J = (Z) => {
      if (!Z) {
        return {};
      }
      let le = [];
      Object.keys(Z).forEach((Ee) => {
        let at = parseInt(Ee, 10);
        if (!isNaN(at)) {
          le[at] = Z[Ee];
        }
      });
      let ke = L(le);
      let Ge = {};
      ke.forEach((Ee, at) => {
        if (Ee !== undefined) {
          Ge[at] = Ee;
        }
      });
      return Ge;
    };
    r((Z) => {
      let le = Z[R];
      if (le) {
        return {
          ...Z,
          [R]: {
            ...le,
            stepChecks: J(le.stepChecks),
            stepNotes: J(le.stepNotes),
            stepNoteOpen: J(le.stepNoteOpen),
          },
        };
      } else {
        return Z;
      }
    });
  };
  let de = (L, R) =>
    X((J) => {
      J[L] = R;
      return J;
    });
  let N = () => X((L) => [...L, ""]);
  let G = (L) => {
    X((R) => R.filter((J, Z) => Z !== L));
    ve((R) => {
      R.splice(L, 1);
      return R;
    });
  };
  let Q = (L, R) => {
    if (L !== R && L != null && R != null) {
      X((J) => {
        let [Z] = J.splice(L, 1);
        J.splice(R, 0, Z);
        return J;
      });
      ve((J) => {
        let [Z] = J.splice(L, 1);
        J.splice(R, 0, Z);
        return J;
      });
    }
  };
  useEffect(() => {
    if (e.length === 0) {
      return;
    }
    let L = e[o];
    if (!L) {
      return;
    }
    let R = a[L.__id];
    if (R && !R.testedBy && w) {
      H(L.__id, "testedBy", w);
    }
  }, [o, e.length]);
  let z = async (L, R) => {
    let Z = Array.from(R)
      .filter((ke) => ke.type.startsWith("image/"))
      .map(
        (ke) =>
          new Promise((Ge) => {
            let Ee = new FileReader();
            Ee.onload = (at) =>
              Ge({
                name: ke.name,
                data: at.target.result,
                size: ke.size,
                addedAt: Date.now() + Math.random(),
              });
            Ee.readAsDataURL(ke);
          }),
      );
    let le = await Promise.all(Z);
    r((ke) => {
      let Ge = ke[L] || makeFinding();
      return {
        ...ke,
        [L]: {
          ...Ge,
          screenshots: [...(Ge.screenshots || []), ...le],
        },
      };
    });
  };
  let re = (L, R) => {
    r((J) => {
      let Z = J[L];
      if (Z) {
        return {
          ...J,
          [L]: {
            ...Z,
            screenshots: Z.screenshots.filter((le, ke) => ke !== R),
          },
        };
      } else {
        return J;
      }
    });
  };
  useEffect(() => {
    let L = (R) => {
      if (e.length === 0) {
        return;
      }
      let J = R.clipboardData?.items;
      if (!J) {
        return;
      }
      let Z = [];
      for (let le of J) {
        if (le.type.startsWith("image/")) {
          let ke = le.getAsFile();
          if (ke) {
            Z.push(ke);
          }
        }
      }
      if (Z.length > 0) {
        R.preventDefault();
        let le = e[o]?.__id;
        if (le) {
          z(le, Z);
        }
      }
    };
    window.addEventListener("paste", L);
    return () => window.removeEventListener("paste", L);
  }, [e, o]);
  let ce = async () => {
    W(false);
    t([]);
    r({});
    s(0);
    f([]);
    T("");
    d("");
    p("All");
    b(null);
    j(false);
    ie(0);
    Y(null);
    try {
      await window.storage.delete("test_data");
    } catch {}
    try {
      let L = await window.storage.list("");
      if (L?.keys) {
        for (let R of L.keys) {
          try {
            await window.storage.delete(R);
          } catch {}
        }
      }
    } catch {}
  };
  let Le = () => {
    try {
      console.log("Starting HTML export (interactive)...");
      let L = (_e) =>
        String(_e ?? "").replace(
          /[&<>"']/g,
          (Ae) =>
            ({
              "&": "&amp;",
              "<": "&lt;",
              ">": "&gt;",
              '"': "&quot;",
              "'": "&#39;",
            })[Ae],
        );
      let R = I === "focused" && h ? e.filter((_e) => h.has(_e.__id)) : e;
      let J = {
        generatedAt: new Date().toISOString(),
        totalAllTests: e.length,
        isFocused: I === "focused" && !!h,
        tests: R.map((_e) => {
          let Ae = a[_e.__id] || makeFinding();
          return {
            id: _e.ID || `TC-${_e.__rowNum || ""}`,
            rowNum: _e.__rowNum,
            title: _e.Title || "Untitled",
            module: _e.Module || "",
            feature: _e.Feature || "",
            preConditions: _e["Pre-conditions"] || "",
            testData: _e["Test Data"] || "",
            steps: _e.__steps || [],
            expectedResult: _e["Expected Result"] || "",
            downstreamCheck: _e["Downstream Check"] || "",
            edgeCase: _e["Edge Case"] || "",
            status: Ae.status || "Not Tested",
            bugSeverity: (isFailureStatus(Ae.status) && Ae.bugSeverity) || "",
            testedBy: Ae.testedBy || "",
            notes: Ae.notes || "",
            otherObservations: Ae.otherObservations || "",
            stepChecks: Ae.stepChecks || {},
            stepNotes: Ae.stepNotes || {},
            screenshots: (Ae.screenshots || []).map((hr) => ({
              data: hr.data,
              name: hr.name,
            })),
            recordings: (Ae.recordings || []).map((hr) => ({
              url: hr.url,
            })),
          };
        }),
      };
      let Z = JSON.stringify(J).replace(/<\/script/gi, "<\\/script");
      let le = `Test Case Report \u2014 ${new Date().toLocaleDateString()}`;
      let ke = `Generated ${new Date().toLocaleString()} \xB7 ${R.length} test${R.length === 1 ? "" : "s"}${I === "focused" && h ? ` (focused subset of ${e.length})` : ""} \xB7 ${EDITION} edition`;
      let Ee = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${L(le)}</title>
<style>
  *, *::before, *::after { box-sizing: border-box; }
  html, body { margin: 0; padding: 0; height: 100%; background: #030712; color: #f3f4f6; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, system-ui, sans-serif; font-size: 14px; line-height: 1.5; }
  /* Dark-mode scrollbars to match the app theme */
  * { scrollbar-width: thin; scrollbar-color: #374151 transparent; }
  *::-webkit-scrollbar { width: 10px; height: 10px; }
  *::-webkit-scrollbar-track { background: transparent; }
  *::-webkit-scrollbar-thumb { background: #374151; border-radius: 2px; border: 2px solid transparent; background-clip: padding-box; }
  *::-webkit-scrollbar-thumb:hover { background: #4b5563; background-clip: padding-box; border: 2px solid transparent; }
  *::-webkit-scrollbar-corner { background: transparent; }
  .app { display: flex; flex-direction: column; height: 100vh; overflow: hidden; }
  .app-header { display: flex; align-items: center; justify-content: space-between; gap: 16px; padding: 14px 24px; background: #111827; border-bottom: 1px solid #1f2937; flex-wrap: wrap; flex-shrink: 0; }
  .app-header h1 { margin: 0; font-size: 16px; font-weight: 700; color: #f3f4f6; }
  .app-header .subtitle { font-size: 12px; color: #9ca3af; margin-top: 2px; }
  .progress-block { display: flex; flex-direction: column; align-items: flex-end; gap: 4px; min-width: 200px; }
  #progress-text { font-size: 12px; color: #9ca3af; font-variant-numeric: tabular-nums; }
  .progress-bar { width: 200px; height: 4px; background: #1f2937; border-radius: 2px; overflow: hidden; }
  #progress-fill { height: 100%; background: #818cf8; width: 0%; transition: width 0.2s; }

  .layout { flex: 1; display: grid; grid-template-columns: 320px 1fr; min-height: 0; }
  .sidebar { background: #111827; border-right: 1px solid #1f2937; display: flex; flex-direction: column; min-height: 0; overflow: hidden; }
  .sidebar-controls { padding: 12px; border-bottom: 1px solid #1f2937; display: flex; flex-direction: column; gap: 8px; flex-shrink: 0; }
  #search-input, #filter-select { width: 100%; background: #1f2937; border: 1px solid #374151; color: #f3f4f6; padding: 8px 10px; font-size: 13px; border-radius: 2px; font-family: inherit; }
  #filter-select { appearance: none; -webkit-appearance: none; -moz-appearance: none; padding-right: 28px; background-image: url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 12 12'><path d='M3 4.5l3 3 3-3' stroke='%239ca3af' stroke-width='1.5' fill='none' stroke-linecap='round' stroke-linejoin='round'/></svg>"); background-repeat: no-repeat; background-position: right 10px center; background-size: 10px; cursor: pointer; }
  #filter-select::-ms-expand { display: none; }
  #filter-select option { background: #1f2937; color: #f3f4f6; }
  #search-input:focus, #filter-select:focus { outline: none; border-color: #818cf8; }
  #search-input::placeholder { color: #6b7280; }
  .test-list { flex: 1; overflow-y: auto; min-height: 0; }
  .empty { padding: 24px; text-align: center; color: #6b7280; font-size: 13px; }

  .sb-row { width: 100%; display: flex; align-items: flex-start; gap: 10px; padding: 10px 12px; background: transparent; border: none; border-bottom: 1px solid #1f2937; border-left: 3px solid transparent; color: inherit; cursor: pointer; text-align: left; font-family: inherit; transition: background 0.1s; }
  .sb-row:hover { background: #1f2937; }
  .sb-row.active { background: #1e1b4b; border-left-color: #818cf8; }
  .sb-row.done { opacity: 0.55; }
  .sb-row.done.active { opacity: 1; }
  .sb-dot { width: 8px; height: 8px; border-radius: 50%; margin-top: 6px; flex-shrink: 0; }
  .dot-Pass { background: #22c55e; }
  .dot-Fail, .dot-Fail-Blocked { background: #ef4444; }
  .dot-Blocked { background: #f97316; }
  .dot-Partial-Pass { background: #eab308; }
  .dot-In-Progress { background: #3b82f6; }
  .dot-Not-Tested { background: #4b5563; }
  .sb-body { flex: 1; min-width: 0; }
  .sb-meta { display: flex; align-items: center; gap: 8px; margin-bottom: 2px; font-size: 11px; color: #9ca3af; }
  .sb-id { font-family: ui-monospace, SFMono-Regular, Menlo, monospace; color: #818cf8; }
  .sb-mod { color: #6b7280; }
  .sb-title { font-size: 13px; color: #e5e7eb; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .sb-row.active .sb-title { color: #f3f4f6; font-weight: 600; }
  .sb-check { width: 16px; height: 16px; border: 1.5px solid #374151; border-radius: 2px; display: flex; align-items: center; justify-content: center; font-size: 10px; color: transparent; flex-shrink: 0; margin-top: 2px; }
  .sb-check.on { background: #4f46e5; border-color: #4f46e5; color: white; }

  .detail-panel { padding: 32px 36px; overflow-y: auto; max-width: 100%; min-height: 0; }
  .placeholder { color: #6b7280; padding: 60px 24px; text-align: center; font-size: 14px; }
  .detail-inner { max-width: 1400px; }
  .detail-cols { display: grid; grid-template-columns: minmax(0, 1fr) minmax(0, 1fr); gap: 36px; align-items: start; }
  .detail-col { min-width: 0; }
  @media (min-width: 1201px) {
    .detail-col > .field:first-child, .detail-col > .findings:first-child { margin-top: 0; }
  }
  .detail-header { display: flex; justify-content: space-between; align-items: flex-start; gap: 16px; padding-bottom: 16px; border-bottom: 1px solid #1f2937; margin-bottom: 20px; flex-wrap: wrap; }
  .breadcrumb { display: flex; align-items: center; gap: 6px; font-size: 12px; color: #9ca3af; font-family: ui-monospace, SFMono-Regular, Menlo, monospace; margin-bottom: 6px; flex-wrap: wrap; }
  .row-num { background: #1f2937; color: #9ca3af; padding: 1px 6px; border-radius: 2px; font-size: 10px; margin-right: 6px; }
  .bc-id { color: #818cf8; }
  .bc-sep { color: #4b5563; }
  .detail-title { margin: 0; font-size: 22px; font-weight: 700; color: #f3f4f6; line-height: 1.3; }
  .detail-head-right { display: flex; align-items: center; gap: 12px; flex-wrap: wrap; }
  .badge-stack { display: inline-flex; gap: 6px; align-items: center; }

  .badge { display: inline-block; padding: 3px 10px; border-radius: 2px; font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.4px; }
  .badge-Pass { background: rgba(34, 197, 94, 0.15); color: #4ade80; border: 1px solid rgba(34, 197, 94, 0.3); }
  .badge-Fail, .badge-Fail-Blocked { background: rgba(239, 68, 68, 0.15); color: #f87171; border: 1px solid rgba(239, 68, 68, 0.3); }
  .badge-Blocked { background: rgba(249, 115, 22, 0.15); color: #fb923c; border: 1px solid rgba(249, 115, 22, 0.3); }
  .badge-Partial-Pass { background: rgba(234, 179, 8, 0.15); color: #facc15; border: 1px solid rgba(234, 179, 8, 0.3); }
  .badge-In-Progress { background: rgba(59, 130, 246, 0.15); color: #60a5fa; border: 1px solid rgba(59, 130, 246, 0.3); }
  .badge-Not-Tested { background: rgba(75, 85, 99, 0.2); color: #9ca3af; border: 1px solid rgba(75, 85, 99, 0.4); }
  .badge-P0 { background: rgba(239, 68, 68, 0.2); color: #fca5a5; border: 1px solid rgba(239, 68, 68, 0.4); }
  .badge-P1 { background: rgba(249, 115, 22, 0.2); color: #fdba74; border: 1px solid rgba(249, 115, 22, 0.4); }
  .badge-P2 { background: rgba(234, 179, 8, 0.2); color: #fde047; border: 1px solid rgba(234, 179, 8, 0.4); }

  .done-toggle { display: inline-flex; align-items: center; gap: 8px; padding: 6px 12px; background: #1f2937; border: 1px solid #374151; border-radius: 2px; cursor: pointer; user-select: none; font-size: 12px; color: #d1d5db; transition: all 0.1s; }
  .done-toggle:hover { background: #374151; }
  .done-toggle.on { background: #4f46e5; border-color: #4f46e5; color: white; }
  .done-toggle input { display: none; }
  .done-box { width: 14px; height: 14px; border: 1.5px solid currentColor; border-radius: 2px; display: inline-flex; align-items: center; justify-content: center; font-size: 9px; line-height: 1; }
  .done-toggle.on .done-box { background: white; color: #4f46e5; }
  .done-label { font-weight: 600; }

  .field { margin-bottom: 18px; }
  .field-label { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.7px; color: #9ca3af; margin-bottom: 6px; }
  .field-value { color: #e5e7eb; white-space: pre-wrap; line-height: 1.55; word-wrap: break-word; }
  .field-highlight .field-label { color: #fbbf24; }
  .field-highlight .field-value { background: rgba(251, 191, 36, 0.08); border-left: 3px solid #f59e0b; padding: 10px 14px; border-radius: 2px; color: #fde68a; }

  code.path { font-family: ui-monospace, SFMono-Regular, Menlo, monospace; font-size: 12px; background: #1f2937; color: #a5b4fc; border: 1px solid #374151; border-radius: 2px; padding: 1px 5px; word-break: break-all; }

  .steps { display: flex; flex-direction: column; gap: 6px; }
  .step { display: flex; gap: 10px; align-items: flex-start; padding: 10px 12px; background: #111827; border: 1px solid #1f2937; border-radius: 2px; }
  .step-done { background: rgba(34, 197, 94, 0.06); border-color: rgba(34, 197, 94, 0.25); }
  .step-blocked { background: rgba(239, 68, 68, 0.06); border-color: rgba(239, 68, 68, 0.25); }
  .step-mark { width: 18px; height: 18px; border-radius: 2px; border: 1.5px solid #374151; display: inline-flex; align-items: center; justify-content: center; flex-shrink: 0; font-size: 11px; font-weight: bold; color: transparent; }
  .step-done .step-mark { background: #16a34a; border-color: #16a34a; color: white; }
  .step-blocked .step-mark { background: #dc2626; border-color: #dc2626; color: white; }
  .step-num { color: #818cf8; font-weight: 700; font-family: ui-monospace, SFMono-Regular, Menlo, monospace; min-width: 22px; font-size: 13px; flex-shrink: 0; padding-top: 1px; }
  .step-body { flex: 1; min-width: 0; }
  .step-text { color: #e5e7eb; line-height: 1.55; }
  .step-blocked-tag { color: #fca5a5; font-size: 11px; font-weight: 600; margin-left: 4px; }
  .step-note { margin-top: 6px; background: rgba(251, 191, 36, 0.08); border-left: 3px solid #f59e0b; padding: 6px 10px; font-size: 12px; color: #fde68a; border-radius: 2px; white-space: pre-wrap; }

  .findings { margin-top: 28px; padding: 18px; background: #111827; border: 1px solid #1f2937; border-radius: 2px; }
  .findings-title { margin: 0 0 16px 0; font-size: 12px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.7px; color: #a5b4fc; }
  .finding-row { margin-bottom: 14px; }
  .finding-row:last-child { margin-bottom: 0; }
  .finding-value { background: #0a0f1d; border: 1px solid #1f2937; padding: 8px 12px; border-radius: 2px; color: #e5e7eb; white-space: pre-wrap; line-height: 1.55; word-wrap: break-word; }
  .finding-value.empty em { color: #6b7280; font-style: italic; }
  .finding-value-inline { display: flex; gap: 6px; align-items: center; }

  .rec-list { display: flex; flex-direction: column; gap: 14px; margin-top: 8px; }
  .rec { background: #0a0f1d; border: 1px solid #1f2937; border-radius: 2px; overflow: hidden; }
  .rec-embed { position: relative; width: 100%; padding-bottom: 56.25%; height: 0; background: #000; }
  .rec-embed iframe, .rec-embed video { position: absolute; top: 0; left: 0; width: 100%; height: 100%; border: 0; }
  .rec-meta { display: flex; align-items: center; gap: 10px; padding: 8px 12px; background: #1f2937; border-top: 1px solid #1f2937; font-size: 12px; }
  .rec-host-link .rec-meta, .rec:not(:has(.rec-embed)) .rec-meta { border-top: none; }
  .rec-badge { font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.4px; padding: 2px 8px; background: #111827; color: #a5b4fc; border: 1px solid #374151; border-radius: 2px; flex-shrink: 0; }
  .rec-host-loom .rec-badge { color: #c4b5fd; border-color: rgba(196,181,253,0.3); }
  .rec-host-youtube .rec-badge { color: #fca5a5; border-color: rgba(252,165,165,0.3); }
  .rec-host-vimeo .rec-badge { color: #67e8f9; border-color: rgba(103,232,249,0.3); }
  .rec-host-drive .rec-badge { color: #fde047; border-color: rgba(253,224,71,0.3); }
  .rec-link { color: #9ca3af; text-decoration: none; word-break: break-all; flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .rec-link:hover { color: #a5b4fc; text-decoration: underline; }
  .ss-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(220px, 1fr)); gap: 12px; margin-top: 8px; }
  .ss { position: relative; background: #0a0f1d; border: 1px solid #1f2937; border-radius: 2px; overflow: hidden; cursor: zoom-in; margin: 0; transition: transform 0.1s; }
  .ss:hover { transform: translateY(-1px); border-color: #374151; }
  .ss-fig { position: absolute; top: 6px; left: 6px; background: #4f46e5; color: white; font-size: 10px; font-weight: 700; padding: 2px 6px; border-radius: 2px; z-index: 1; }
  .ss img { width: 100%; height: auto; display: block; }
  .ss figcaption { padding: 6px 8px; font-size: 11px; color: #9ca3af; word-break: break-all; background: #1f2937; }

  .lightbox { position: fixed; inset: 0; background: rgba(0, 0, 0, 0.92); display: none; align-items: center; justify-content: center; z-index: 9999; padding: 24px; cursor: zoom-out; }
  .lightbox.open { display: flex; }
  .lightbox img { max-width: 100%; max-height: 90vh; border-radius: 2px; box-shadow: 0 20px 50px rgba(0,0,0,0.6); }
  .lightbox-close { position: absolute; top: 16px; right: 16px; width: 40px; height: 40px; background: rgba(255,255,255,0.1); color: white; border: none; border-radius: 2px; cursor: pointer; font-size: 22px; }
  .lightbox-close:hover { background: rgba(255,255,255,0.2); }
  .lightbox-caption { position: absolute; bottom: 24px; left: 50%; transform: translateX(-50%); color: #d1d5db; background: rgba(0,0,0,0.7); padding: 8px 16px; border-radius: 2px; font-size: 13px; max-width: 80%; text-align: center; word-break: break-all; }

  @media (max-width: 1200px) {
    .detail-cols { grid-template-columns: minmax(0, 1fr); gap: 0; }
  }

  @media (max-width: 768px) {
    .app { height: auto; min-height: 100vh; overflow: visible; }
    .layout { grid-template-columns: 1fr; }
    .sidebar { max-height: 50vh; border-right: none; border-bottom: 1px solid #1f2937; }
    .detail-panel { padding: 20px 16px; overflow-y: visible; }
    .detail-title { font-size: 18px; }
    .progress-block { min-width: auto; align-items: flex-start; }
    .progress-bar { width: 100%; }
  }
</style>
</head>
<body>
  <div class="app">
    <header class="app-header">
      <div>
        <h1>${L(le)}</h1>
        <div class="subtitle">${L(ke)}</div>
      </div>
      <div class="progress-block">
        <span id="progress-text">0 / 0 done</span>
        <div class="progress-bar"><div id="progress-fill"></div></div>
      </div>
    </header>
    <div class="layout">
      <aside class="sidebar">
        <div class="sidebar-controls">
          <input type="search" id="search-input" placeholder="Search ID, title, module\u2026">
          <select id="filter-select">
            <option value="all">All tests</option>
            <option value="failed">Failed only</option>
            <option value="notdone">Not yet done</option>
            <option value="done">Done</option>
          </select>
        </div>
        <div id="test-list" class="test-list"></div>
      </aside>
      <main id="detail-panel" class="detail-panel"></main>
    </div>
  </div>

  <div class="lightbox" id="lightbox">
    <button class="lightbox-close" id="lightbox-close">\xD7</button>
    <img id="lightbox-img" src="" alt="">
    <div class="lightbox-caption" id="lightbox-caption"></div>
  </div>

  <script type="application/json" id="report-data">${Z}<\/script>
  <script>
(function() {
  var data = JSON.parse(document.getElementById('report-data').textContent);
  var tests = data.tests;
  var STORAGE_KEY = 'tcr-report-done-v1';

  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function(c) {
      return ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'})[c];
    });
  }
  function isEmpty(v) {
    if (v == null) return true;
    var t = String(v).trim();
    return t === '' || t === '\u2014';
  }
  function isFailure(s) {
    if (!s) return false;
    var t = String(s).trim().toLowerCase();
    return t === 'fail' || t === 'blocked' || t === 'fail/blocked' || t === 'fail/ blocked';
  }
  function statusClass(s) {
    return String(s || 'Not Tested').replace(/[\\s/]+/g, '-').replace(/[^a-zA-Z0-9-]/g, '');
  }
  // Inline-code styling for URL-like paths (matches live app)
  var PATH_RE = /(?<![a-zA-Z0-9_])\\/[a-zA-Z][a-zA-Z0-9_-]*(?:\\/(?:[a-zA-Z0-9_-]+|<[a-zA-Z0-9_-]+>))*(?:\\?[a-zA-Z0-9_=&<>\\-]+)?/g;
  function fmt(text) {
    if (text == null) return '';
    return esc(text).replace(PATH_RE, function(m) { return '<code class="path">' + m + '</code>'; });
  }

  // Detect popular video hosts and build an embed URL. Mirrors the helper in
  // the live app so behaviour stays consistent.
  function detectVideoHost(rawUrl) {
    var url = String(rawUrl || '').trim();
    if (!url) return null;
    var m;
    m = url.match(/loom\\.com\\/(?:share|embed)\\/([a-zA-Z0-9]+)/);
    if (m) return { host: 'loom', label: 'Loom', embed: 'https://www.loom.com/embed/' + m[1] };
    m = url.match(/drive\\.google\\.com\\/file\\/d\\/([a-zA-Z0-9_-]+)/);
    if (m) return { host: 'drive', label: 'Google Drive', embed: 'https://drive.google.com/file/d/' + m[1] + '/preview' };
    m = url.match(/drive\\.google\\.com\\/open\\?id=([a-zA-Z0-9_-]+)/);
    if (m) return { host: 'drive', label: 'Google Drive', embed: 'https://drive.google.com/file/d/' + m[1] + '/preview' };
    m = url.match(/(?:youtube\\.com\\/watch\\?(?:[^&]*&)*v=|youtube\\.com\\/embed\\/)([a-zA-Z0-9_-]{6,})/);
    if (m) return { host: 'youtube', label: 'YouTube', embed: 'https://www.youtube.com/embed/' + m[1] };
    m = url.match(/youtu\\.be\\/([a-zA-Z0-9_-]{6,})/);
    if (m) return { host: 'youtube', label: 'YouTube', embed: 'https://www.youtube.com/embed/' + m[1] };
    m = url.match(/vimeo\\.com\\/(?:video\\/)?(\\d+)/);
    if (m) return { host: 'vimeo', label: 'Vimeo', embed: 'https://player.vimeo.com/video/' + m[1] };
    if (/\\.(mp4|webm|mov|m4v|ogg)(\\?|#|$)/i.test(url)) {
      return { host: 'file', label: 'Video file', embed: null, isVideoFile: true };
    }
    try {
      var hostName = new URL(url).hostname.replace(/^www\\./, '');
      return { host: 'link', label: hostName, embed: null };
    } catch (e) {
      return { host: 'link', label: 'Link', embed: null };
    }
  }

  function renderRecordings(list) {
    if (!list || list.length === 0) return '';
    var items = list.map(function(rec) {
      var info = detectVideoHost(rec.url);
      var label = (info && info.label) || 'Link';
      var host = (info && info.host) || 'link';
      var embedPart = '';
      if (info && info.embed) {
        embedPart = '<div class="rec-embed"><iframe src="' + esc(info.embed) + '" frameborder="0" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture; fullscreen" allowfullscreen></iframe></div>';
      } else if (info && info.isVideoFile) {
        embedPart = '<div class="rec-embed"><video controls preload="metadata" src="' + esc(rec.url) + '"></video></div>';
      }
      return '<div class="rec rec-host-' + esc(host) + '">'
        + embedPart
        + '<div class="rec-meta">'
        +   '<span class="rec-badge">' + esc(label) + '</span>'
        +   '<a class="rec-link" href="' + esc(rec.url) + '" target="_blank" rel="noopener noreferrer" title="' + esc(rec.url) + '">' + esc(rec.url) + '</a>'
        + '</div>'
        + '</div>';
    }).join('');
    return '<div class="rec-list">' + items + '</div>';
  }

  // State
  var doneMap = {};
  try {
    var raw = localStorage.getItem(STORAGE_KEY);
    if (raw) doneMap = JSON.parse(raw) || {};
  } catch (e) {}
  var selectedId = null;
  var search = '';
  var filter = 'all';

  // DOM refs
  var listEl = document.getElementById('test-list');
  var detailEl = document.getElementById('detail-panel');
  var progressTextEl = document.getElementById('progress-text');
  var progressFillEl = document.getElementById('progress-fill');
  var searchInput = document.getElementById('search-input');
  var filterSelect = document.getElementById('filter-select');

  function saveDone() {
    try { localStorage.setItem(STORAGE_KEY, JSON.stringify(doneMap)); } catch (e) {}
  }

  function getVisibleTests() {
    return tests.filter(function(t) {
      if (search) {
        var q = search.toLowerCase();
        var hay = (t.id + ' ' + t.title + ' ' + t.module + ' ' + t.feature).toLowerCase();
        if (hay.indexOf(q) === -1) return false;
      }
      if (filter === 'failed' && !isFailure(t.status)) return false;
      if (filter === 'notdone' && doneMap[t.id]) return false;
      if (filter === 'done' && !doneMap[t.id]) return false;
      return true;
    });
  }

  function renderSidebar() {
    var visible = getVisibleTests();
    if (visible.length === 0) {
      listEl.innerHTML = '<div class="empty">No tests match the current filter.</div>';
      return;
    }
    listEl.innerHTML = visible.map(function(t) {
      var done = !!doneMap[t.id];
      var active = t.id === selectedId;
      var sc = statusClass(t.status);
      return '<button type="button" class="sb-row' + (active ? ' active' : '') + (done ? ' done' : '') + '" data-id="' + esc(t.id) + '">'
        + '<span class="sb-dot dot-' + sc + '" title="' + esc(t.status) + '"></span>'
        + '<div class="sb-body">'
        +   '<div class="sb-meta"><span class="sb-id">' + esc(t.id) + '</span>'
        +   (t.module ? '<span class="sb-mod">' + esc(t.module) + '</span>' : '')
        +   '</div>'
        +   '<div class="sb-title">' + esc(t.title) + '</div>'
        + '</div>'
        + '<span class="sb-check' + (done ? ' on' : '') + '" aria-label="' + (done ? 'Done' : 'Not done') + '">' + (done ? '\u2713' : '') + '</span>'
        + '</button>';
    }).join('');
    Array.prototype.forEach.call(listEl.querySelectorAll('.sb-row'), function(btn) {
      btn.addEventListener('click', function() { selectTest(btn.getAttribute('data-id')); });
    });
  }

  function renderField(label, value, opts) {
    if (isEmpty(value)) return '';
    var highlight = opts && opts.highlight;
    return '<section class="field' + (highlight ? ' field-highlight' : '') + '">'
      + '<div class="field-label">' + esc(label) + '</div>'
      + '<div class="field-value">' + fmt(value) + '</div>'
      + '</section>';
  }

  function renderSteps(t) {
    if (!t.steps || t.steps.length === 0) return '';
    var rows = t.steps.map(function(step, idx) {
      var raw = (t.stepChecks || {})[idx];
      var state = (raw === 'done' || raw === true) ? 'done' : (raw === 'blocked' ? 'blocked' : 'none');
      var note = (t.stepNotes || {})[idx];
      var mark = state === 'done' ? '\u2713' : (state === 'blocked' ? '\u2715' : '');
      var blockedSuffix = state === 'blocked' ? ' <span class="step-blocked-tag">(unable to proceed)</span>' : '';
      return '<div class="step step-' + state + '">'
        + '<div class="step-mark">' + mark + '</div>'
        + '<div class="step-num">' + (idx + 1) + '.</div>'
        + '<div class="step-body"><div class="step-text">' + fmt(step) + blockedSuffix + '</div>'
        + (note ? '<div class="step-note"><strong>Note:</strong> ' + fmt(note) + '</div>' : '')
        + '</div></div>';
    }).join('');
    return '<section class="field"><div class="field-label">Steps</div><div class="steps">' + rows + '</div></section>';
  }

  function renderFindings(t) {
    var hasSeverity = isFailure(t.status) && t.bugSeverity;
    var screenshots = t.screenshots && t.screenshots.length > 0
      ? '<div class="ss-grid">'
        + t.screenshots.map(function(s, i) {
            return '<figure class="ss" data-src="' + esc(s.data) + '" data-name="' + esc(s.name) + '" data-fig="Fig. ' + (i + 1) + '">'
              + '<span class="ss-fig">Fig. ' + (i + 1) + '</span>'
              + '<img src="' + esc(s.data) + '" alt="' + esc(s.name) + '">'
              + '<figcaption>' + esc(s.name) + '</figcaption>'
              + '</figure>';
          }).join('')
        + '</div>'
      : '';
    var recordings = renderRecordings(t.recordings);
    return '<section class="findings">'
      + '<h3 class="findings-title">Tester Findings</h3>'
      + (t.testedBy ? '<div class="finding-row"><div class="field-label">Tester</div><div class="field-value finding-value">' + esc(t.testedBy) + '</div></div>' : '')
      + '<div class="finding-row"><div class="field-label">Tester Notes</div><div class="field-value finding-value' + (t.notes ? '' : ' empty') + '">' + (t.notes ? fmt(t.notes) : '<em>No notes</em>') + '</div></div>'
      + (t.otherObservations ? '<div class="finding-row"><div class="field-label">Other Observations</div><div class="field-value finding-value">' + fmt(t.otherObservations) + '</div></div>' : '')
      + (recordings ? '<div class="finding-row"><div class="field-label">Recordings (' + t.recordings.length + ')</div>' + recordings + '</div>' : '')
      + (screenshots ? '<div class="finding-row"><div class="field-label">Screenshots (' + t.screenshots.length + ')</div>' + screenshots + '</div>' : '')
      + '<div class="finding-row"><div class="field-label">Final Status</div><div class="finding-value-inline"><span class="badge badge-' + statusClass(t.status) + '">' + esc(t.status) + '</span>'
      + (hasSeverity ? ' <span class="badge badge-' + esc(t.bugSeverity) + '">' + esc(t.bugSeverity) + '</span>' : '')
      + '</div></div>'
      + '</section>';
  }

  function renderDetail() {
    if (!selectedId) {
      detailEl.innerHTML = '<div class="placeholder">Select a test case from the sidebar to view its details.</div>';
      return;
    }
    var t = tests.find(function(x) { return x.id === selectedId; });
    if (!t) {
      detailEl.innerHTML = '<div class="placeholder">Test not found.</div>';
      return;
    }
    var done = !!doneMap[t.id];
    var hasSeverity = isFailure(t.status) && t.bugSeverity;
    var rowNum = t.rowNum ? '<span class="row-num" title="CSV row ' + t.rowNum + '">Row ' + t.rowNum + '</span>' : '';
    detailEl.innerHTML = '<div class="detail-inner">'
      + '<div class="detail-header">'
      +   '<div class="detail-head-left">'
      +     '<div class="breadcrumb">' + rowNum + '<span class="bc-id">' + esc(t.id) + '</span>'
      +       (t.module ? '<span class="bc-sep">\xB7</span><span>' + esc(t.module) + '</span>' : '')
      +       (t.feature && !isEmpty(t.feature) ? '<span class="bc-sep">\xB7</span><span>' + esc(t.feature) + '</span>' : '')
      +     '</div>'
      +     '<h2 class="detail-title">' + esc(t.title) + '</h2>'
      +   '</div>'
      +   '<div class="detail-head-right">'
      +     '<span class="badge-stack">'
      +       '<span class="badge badge-' + statusClass(t.status) + '">' + esc(t.status) + '</span>'
      +       (hasSeverity ? '<span class="badge badge-' + esc(t.bugSeverity) + '">' + esc(t.bugSeverity) + '</span>' : '')
      +     '</span>'
      +     '<label class="done-toggle' + (done ? ' on' : '') + '">'
      +       '<input type="checkbox" id="done-cb"' + (done ? ' checked' : '') + '>'
      +       '<span class="done-box">' + (done ? '\u2713' : '') + '</span>'
      +       '<span class="done-label">' + (done ? 'Done' : 'Mark as done') + '</span>'
      +     '</label>'
      +   '</div>'
      + '</div>'
      + '<div class="detail-cols">'
      +   '<div class="detail-col detail-col-left">'
      +     renderField('Pre-conditions', t.preConditions)
      +     renderField('Test Data', t.testData)
      +     renderSteps(t)
      +     renderField('Expected Result', t.expectedResult, { highlight: true })
      +     renderField('Downstream Check', t.downstreamCheck)
      +     renderField('Edge Case', t.edgeCase)
      +   '</div>'
      +   '<div class="detail-col detail-col-right">'
      +     renderFindings(t)
      +   '</div>'
      + '</div>'
      + '</div>';

    var cb = document.getElementById('done-cb');
    if (cb) {
      cb.addEventListener('change', function() {
        if (cb.checked) doneMap[t.id] = Date.now();
        else delete doneMap[t.id];
        saveDone();
        renderSidebar();
        renderDetail();
        renderProgress();
      });
    }
  }

  function renderProgress() {
    var total = tests.length;
    var done = tests.filter(function(t) { return doneMap[t.id]; }).length;
    progressTextEl.textContent = done + ' / ' + total + ' done';
    progressFillEl.style.width = total > 0 ? ((done / total) * 100) + '%' : '0%';
  }

  function selectTest(id) {
    selectedId = id;
    history.replaceState(null, '', '#' + encodeURIComponent(id));
    renderSidebar();
    renderDetail();
    if (detailEl) detailEl.scrollTop = 0;
    if (window.matchMedia('(max-width: 768px)').matches) {
      // Scroll the page up so the detail is visible on mobile
      window.scrollTo({ top: detailEl.offsetTop - 8, behavior: 'smooth' });
    }
  }

  // Event wiring
  searchInput.addEventListener('input', function(e) {
    search = e.target.value;
    renderSidebar();
  });
  filterSelect.addEventListener('change', function(e) {
    filter = e.target.value;
    renderSidebar();
  });

  // Lightbox
  var lightbox = document.getElementById('lightbox');
  var lightboxImg = document.getElementById('lightbox-img');
  var lightboxCaption = document.getElementById('lightbox-caption');
  document.addEventListener('click', function(e) {
    var ss = e.target.closest && e.target.closest('.ss');
    if (ss) {
      lightboxImg.src = ss.getAttribute('data-src');
      var fig = ss.getAttribute('data-fig');
      var name = ss.getAttribute('data-name');
      lightboxCaption.textContent = fig + ' \u2014 ' + name;
      lightbox.classList.add('open');
      return;
    }
    if (e.target === lightbox || (e.target.closest && e.target.closest('#lightbox-close'))) {
      lightbox.classList.remove('open');
      lightboxImg.src = '';
    }
  });
  document.addEventListener('keydown', function(e) {
    if (e.key === 'Escape' && lightbox.classList.contains('open')) {
      lightbox.classList.remove('open');
      lightboxImg.src = '';
    }
    // Quick nav with j/k or arrow keys (when not typing)
    var typing = document.activeElement && (document.activeElement.tagName === 'INPUT' || document.activeElement.tagName === 'TEXTAREA');
    if (typing) return;
    var visible = getVisibleTests();
    if (visible.length === 0) return;
    var idx = visible.findIndex(function(t) { return t.id === selectedId; });
    if (e.key === 'j' || e.key === 'ArrowDown') {
      e.preventDefault();
      selectTest(visible[Math.min(idx + 1, visible.length - 1)].id);
    } else if (e.key === 'k' || e.key === 'ArrowUp') {
      e.preventDefault();
      selectTest(visible[Math.max(idx - 1, 0)].id);
    } else if (e.key === 'd') {
      // Toggle done with 'd' key
      var t = tests.find(function(x) { return x.id === selectedId; });
      if (!t) return;
      if (doneMap[t.id]) delete doneMap[t.id];
      else doneMap[t.id] = Date.now();
      saveDone();
      renderSidebar();
      renderDetail();
      renderProgress();
    }
  });

  // Init from URL hash
  var hash = location.hash.replace(/^#/, '');
  if (hash) {
    try {
      var decoded = decodeURIComponent(hash);
      if (tests.find(function(t) { return t.id === decoded; })) selectedId = decoded;
    } catch (e) {}
  }
  if (!selectedId && tests.length > 0) selectedId = tests[0].id;

  renderSidebar();
  renderDetail();
  renderProgress();
})();
<\/script>
</body>
</html>`;
      let at = new Blob([Ee], {
        type: "text/html;charset=utf-8",
      });
      let He = URL.createObjectURL(at);
      let Ne = document.createElement("a");
      Ne.href = He;
      Ne.download = `openmaison-uat-report-${I === "focused" && h ? "focused-" : ""}${new Date().toISOString().split("T")[0]}.html`;
      document.body.appendChild(Ne);
      Ne.click();
      document.body.removeChild(Ne);
      setTimeout(() => URL.revokeObjectURL(He), 1000);
      C(false);
      console.log("HTML export complete");
    } catch (L) {
      console.error("Export error:", L);
      alert("Export failed: " + L.message);
    }
  };
  let We = () => {
    try {
      console.log("Starting CSV export...");
      let L = I === "focused" && h ? e.filter((Ee) => h.has(Ee.__id)) : e;
      let R = (e[0] && e[0].__testerCol) || "Tester";
      let J = L.map((Ee, at) => {
        let He = a[Ee.__id] || {};
        let Ne = (He.screenshots || [])
          .map((Ba, xr) => `${Ee.ID || `TC-${at + 1}`}_${xr + 1}_${Ba.name}`)
          .join("; ");
        let _e = (Ee.__steps || []).map((Ba, xr) => {
          let vi = (He.stepChecks || {})[xr] ? "[✓]" : "[ ]";
          let mn = (He.stepNotes || {})[xr] || "";
          return `${vi} Step ${xr + 1}${mn ? ": " + mn : ""}`;
        }).join(`
`);
        let { __id: Ae, __steps: hr, __testerCol: Te, __rowNum: wa, ...gr } = Ee;
        return {
          ...gr,
          [R]: He.testedBy || gr[R] || "",
          "Pass/Fail": He.status && He.status !== "Not Tested" ? He.status : "",
          "Tester Notes": He.notes || "",
          "Bug Severity (P0/ P1 / P2)": He.bugSeverity || "",
          "Other Observations": He.otherObservations || "",
          Recordings: (He.recordings || []).map((Ba) => Ba.url).join("; "),
          "Step Notes": _e,
          Screenshots: Ne,
        };
      });
      let Z = Papa.unparse(J);
      let le = new Blob([Z], {
        type: "text/csv;charset=utf-8",
      });
      let ke = URL.createObjectURL(le);
      let Ge = document.createElement("a");
      Ge.href = ke;
      Ge.download = `openmaison-uat-findings-${I === "focused" && h ? "focused-" : ""}${new Date().toISOString().split("T")[0]}.csv`;
      document.body.appendChild(Ge);
      Ge.click();
      document.body.removeChild(Ge);
      setTimeout(() => URL.revokeObjectURL(ke), 1000);
      L.forEach((Ee, at) => {
        let He = a[Ee.__id];
        if (He?.screenshots) {
          He.screenshots.forEach((Ne, _e) => {
            let Ae = document.createElement("a");
            Ae.href = Ne.data;
            Ae.download = `${Ee.ID || `TC-${at + 1}`}_${_e + 1}_${Ne.name}`;
            document.body.appendChild(Ae);
            Ae.click();
            document.body.removeChild(Ae);
          });
        }
      });
      C(false);
      console.log("CSV export complete");
    } catch (L) {
      console.error("CSV export error:", L);
      alert("CSV export failed: " + L.message);
    }
  };
  let St = e.filter((L) => {
    if (h && !h.has(L.__id)) {
      return false;
    }
    let R = a[L.__id]?.status || "Not Tested";
    if (c !== "All" && R !== c) {
      return false;
    }
    if (!u) {
      return true;
    }
    let J = u.toLowerCase();
    return (
      (L.ID || "").toLowerCase().includes(J) ||
      (L.Title || "").toLowerCase().includes(J) ||
      (L.Module || "").toLowerCase().includes(J) ||
      (L.Feature || "").toLowerCase().includes(J)
    );
  });
  let q = e[o];
  let oe = q ? a[q.__id] || makeFinding() : null;
  let $t = STATUS_OPTIONS.reduce((L, R) => {
    L[R.value] = e.filter((J) => (a[J.__id]?.status || "Not Tested") === R.value).length;
    return L;
  }, {});
  let Qe = e.length - $t["Not Tested"];
  useEffect(() => {
    if (xe.current) {
      xe.current.scrollTop = 0;
    }
    F(false);
    E(null);
    O(null);
    pe("");
  }, [o]);
  if (!g) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-gray-950">
        <div className="text-gray-400">Loading...</div>
      </div>
    );
  }
  if (e.length === 0 || A) {
    return (
      <CsvUploadScreen
        onLoad={async (L, R) => {
          let J = await $(L, R);
          if (J?.success) {
            j(false);
          }
          return J;
        }}
        fileInputRef={Ce}
        hasExistingData={e.length > 0}
        onBack={() => j(false)}
        existingTesters={m}
      />
    );
  }
  let ya = getStatusConfig(oe?.status);
  let Ht = q.__steps || [];
  let xi = (L) => {
    let R = (oe.stepChecks || {})[L];
    if (R === "done" || R === true) {
      return "done";
    } else if (R === "blocked") {
      return "blocked";
    } else {
      return "none";
    }
  };
  let Kg = Ht.filter((L, R) => xi(R) === "done").length;
  let _c = Ht.filter((L, R) => xi(R) === "blocked").length;
  return (
    <div className="h-screen flex bg-gray-950 overflow-hidden text-gray-100 relative">
      <style>{`
        ::-webkit-scrollbar { width: 10px; height: 10px; }
        ::-webkit-scrollbar-track { background: #0a0a0f; }
        ::-webkit-scrollbar-thumb { background: #2a2a35; border-radius: 5px; border: 2px solid #0a0a0f; }
        ::-webkit-scrollbar-thumb:hover { background: #3a3a48; }
        ::-webkit-scrollbar-corner { background: #0a0a0f; }
        * { scrollbar-color: #2a2a35 #0a0a0f; scrollbar-width: thin; }
        .step-btn-preview { opacity: 0.6; transition: opacity 0.15s; }
        .step-btn:hover .step-btn-preview { opacity: 1; }
        .step-row-pencil { opacity: 0; transition: opacity 0.15s; }
        .step-row:hover .step-row-pencil { opacity: 1; }
        select.dark-select {
          appearance: none;
          -webkit-appearance: none;
          -moz-appearance: none;
          background-image: url("data:image/svg+xml;charset=utf-8,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='%239ca3af' stroke-width='2.5' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpolyline points='6 9 12 15 18 9'%3E%3C/polyline%3E%3C/svg%3E");
          background-repeat: no-repeat;
          background-position: right 0.75rem center;
          background-size: 12px;
          padding-right: 2.25rem !important;
        }
        select.dark-select::-ms-expand { display: none; }
        /* Sidebar \u2014 mobile: overlay from top, desktop: side panel */
        .sidebar-panel { width: 0; height: 100%; }
        .sidebar-panel.sidebar-open { width: 100%; position: absolute; top: 0; left: 0; z-index: 40; height: 100%; border-bottom: 1px solid rgb(31 41 55); }
        @media (min-width: 768px) {
          .sidebar-panel { position: relative !important; height: 100% !important; }
          .sidebar-panel.sidebar-open { width: 20rem; }
        }
      `}</style>
      <div
        className={`${l ? "sidebar-open" : ""} sidebar-panel transition-all duration-200 bg-gray-900 border-r border-gray-800 flex flex-col overflow-hidden flex-shrink-0`}
      >
        <div className="p-4 border-b border-gray-800">
          <div className="flex items-center justify-between mb-3">
            <div className="flex items-center gap-2">
              <button
                onClick={() => i(false)}
                className="p-1 hover:bg-gray-800 rounded-md"
                title="Collapse sidebar"
              >
                <PanelLeftClose className="w-4 h-4 text-gray-400" />
              </button>
              <h2 className="font-bold text-gray-100">Test Cases</h2>
              <span
                className={`text-[9px] font-bold uppercase tracking-widest px-1.5 py-0.5 rounded-sm border ${EDITION === "Pro" ? "bg-emerald-900/40 text-emerald-300 border-emerald-700" : "bg-gray-800 text-gray-500 border-gray-700"}`}
              >
                {EDITION}
              </span>
            </div>
            <span className="text-xs text-gray-400">
              {Qe} / {e.length} done
            </span>
          </div>
          <div className="w-full bg-gray-800 rounded-full h-1.5 mb-3">
            <div
              className="bg-indigo-500 h-1.5 rounded-full transition-all"
              style={{
                width: `${(Qe / e.length) * 100}%`,
              }}
            />
          </div>
          {B ? (
            <div className="relative">
              <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-gray-500" />
              <input
                ref={fe}
                type="text"
                placeholder="Search..."
                value={u}
                onChange={(L) => d(L.target.value)}
                onKeyDown={(L) => {
                  if (L.key === "Escape") {
                    d("");
                    ne(false);
                  }
                }}
                className="w-full pl-9 pr-8 py-2 text-sm bg-gray-800 border border-gray-700 text-gray-100 placeholder-gray-500 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-transparent"
              />
              <button
                onClick={() => {
                  d("");
                  ne(false);
                }}
                className="absolute right-2 top-1/2 -translate-y-1/2 p-1 text-gray-500 hover:text-gray-300 rounded"
                title="Close search"
              >
                <X className="w-3.5 h-3.5" />
              </button>
            </div>
          ) : (
            <div className="flex items-center gap-2">
              <button
                onClick={() => {
                  ne(true);
                  setTimeout(() => fe.current?.focus(), 0);
                }}
                className="flex-shrink-0 flex items-center justify-center bg-gray-800 border border-gray-700 text-gray-400 hover:text-indigo-300 hover:border-indigo-600 hover:bg-gray-700 rounded-md transition-colors"
                style={{
                  height: "38px",
                  width: "38px",
                }}
                title="Search test cases"
              >
                <Search className="w-4 h-4" />
              </button>
              <button
                onClick={() => S(true)}
                className={`flex-shrink-0 flex items-center justify-center border rounded-md transition-colors ${h ? "bg-indigo-600 border-indigo-600 text-white hover:bg-indigo-500" : "bg-gray-800 border-gray-700 text-gray-400 hover:text-indigo-300 hover:border-indigo-600 hover:bg-gray-700"}`}
                style={{
                  height: "38px",
                  width: "38px",
                }}
                title={h ? `Focus filter active (${h.size} tests)` : "Focus on a subset of tests"}
              >
                <Filter className="w-4 h-4" />
              </button>
              <select
                value={c}
                onChange={(L) => p(L.target.value)}
                className="dark-select flex-1 min-w-0 px-3 text-sm bg-gray-800 border border-gray-700 text-gray-100 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500"
                style={{
                  height: "38px",
                }}
              >
                <option value="All">All Statuses ({e.length})</option>
                {STATUS_OPTIONS.map((L) => (
                  <option value={L.value} key={L.value}>
                    {L.value} ({$t[L.value]})
                  </option>
                ))}
              </select>
            </div>
          )}
          {h && (
            <div className="mt-2 flex items-center gap-2 px-3 py-1.5 bg-indigo-950/40 border border-indigo-800/50 rounded-sm text-xs">
              <Filter className="w-3 h-3 text-indigo-400 flex-shrink-0" />
              <span className="text-indigo-200">
                Focused: {h.size} of {e.length}
              </span>
              <button
                onClick={() => b(null)}
                className="ml-auto text-indigo-400 hover:text-indigo-200 transition-colors"
                title="Clear focus filter"
              >
                <X className="w-3 h-3" />
              </button>
            </div>
          )}
        </div>
        <div className="flex-1 overflow-y-auto">
          {St.map((L) => {
            let R = e.indexOf(L);
            let J = a[L.__id] || {};
            let Z = getStatusConfig(J.status);
            let le = R === o;
            return (
              <button
                onClick={() => {
                  s(R);
                  if (window.matchMedia("(max-width: 767px)").matches) {
                    i(false);
                  }
                }}
                className={`w-full text-left px-4 py-3 border-b border-gray-800 hover:bg-gray-800 transition-colors ${le ? "bg-indigo-950/50 border-l-4 border-l-indigo-500" : "border-l-4 border-l-transparent"}`}
                key={L.__id}
              >
                <div className="flex items-start gap-2">
                  <div className={`w-2 h-2 rounded-full mt-1.5 flex-shrink-0 ${Z.dot}`} />
                  <div className="flex-1 min-w-0">
                    <div className="flex items-center gap-2 mb-0.5">
                      <span className="text-xs font-mono text-gray-400 whitespace-nowrap flex-shrink-0">
                        {L.ID}
                      </span>
                      {L.Module && <span className="text-xs text-gray-500 truncate">{L.Module}</span>}
                    </div>
                    <div
                      className={`text-sm truncate ${le ? "font-semibold text-indigo-300" : "text-gray-200"}`}
                    >
                      {L.Title || "Untitled"}
                    </div>
                    {!isEmptyValue(L.Feature) && (
                      <div className="mt-1">
                        <span className="inline-block text-[10px] font-medium uppercase tracking-wide px-1.5 py-0.5 bg-gray-800 text-gray-400 border border-gray-700 rounded-sm">
                          {L.Feature}
                        </span>
                      </div>
                    )}
                  </div>
                </div>
              </button>
            );
          })}
          {St.length === 0 && (
            <div className="p-6 text-center text-sm text-gray-500">No matching test cases</div>
          )}
        </div>
        {(me || he > STORAGE_LIMIT_BYTES * STORAGE_WARN_AT) && (
          <div
            className={`p-3 border-t text-xs ${me ? "bg-red-950/40 border-red-800 text-red-300" : "bg-amber-950/40 border-amber-800 text-amber-300"}`}
          >
            <div className="flex items-start gap-2">
              <CircleAlert className="w-3.5 h-3.5 flex-shrink-0 mt-0.5" />
              <div className="min-w-0">
                {me ? (
                  <Fragment>
                    <strong>Save failed:</strong> {me}
                  </Fragment>
                ) : (
                  <Fragment>
                    <strong>Storage {Math.round((he / STORAGE_LIMIT_BYTES) * 100)}% full</strong>
                    <div className="mt-0.5 opacity-90">
                      {(he / 1024 / 1024).toFixed(2)} MB used. Export & remove screenshots if you can.
                    </div>
                  </Fragment>
                )}
              </div>
            </div>
          </div>
        )}
        {m.length > 0 && (
          <div className="p-3 border-t border-gray-800 text-xs text-gray-400">
            <Users className="w-3 h-3 inline mr-1" />
            {m.length} tester{m.length !== 1 ? "s" : ""}: {m.join(", ")}
          </div>
        )}
      </div>
      <div className="flex-1 flex flex-col min-w-0">
        <div className="bg-gray-900 border-b border-gray-800 flex-shrink-0">
          <div className="px-3 sm:px-4 py-2 sm:py-3 flex items-center gap-2 sm:gap-3">
            {!l && (
              <button
                onClick={() => i(true)}
                className="p-2 hover:bg-gray-800 rounded-md flex-shrink-0"
                title="Show test cases"
              >
                <PanelLeftOpen className="w-5 h-5 text-gray-300" />
              </button>
            )}
            <div className="flex-1 min-w-0 hidden md:flex items-center gap-3">
              <span className="text-xs font-mono bg-indigo-950/60 text-indigo-300 px-2 py-1 rounded border border-indigo-800 flex-shrink-0">
                {q.ID}
              </span>
              <div className="min-w-0 flex-1">
                <div className="text-base font-semibold text-gray-100 truncate">{q.Title || "Untitled"}</div>
                {q.Module && <div className="text-xs text-gray-400 truncate">{q.Module}</div>}
              </div>
              <span
                className={`hidden lg:inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-semibold border flex-shrink-0 ${ya.color}`}
              >
                <ya.icon className="w-3.5 h-3.5" />
                {oe.status}
              </span>
            </div>
            <div className="flex-1 text-center md:hidden">
              <span className="text-xs text-gray-400">
                {o + 1} / {e.length}
              </span>
            </div>
            <div className="flex items-center gap-1 sm:gap-2 flex-shrink-0">
              <span className="text-xs text-gray-500 hidden lg:inline">
                {o + 1} / {e.length}
              </span>
              <button
                onClick={() => s(Math.max(0, o - 1))}
                disabled={o === 0}
                className="p-2 text-gray-300 hover:bg-gray-800 rounded-md disabled:opacity-30 disabled:cursor-not-allowed"
              >
                <ChevronLeft className="w-5 h-5" />
              </button>
              <button
                onClick={() => s(Math.min(e.length - 1, o + 1))}
                disabled={o === e.length - 1}
                className="p-2 text-gray-300 hover:bg-gray-800 rounded-md disabled:opacity-30 disabled:cursor-not-allowed"
              >
                <ChevronRight className="w-5 h-5" />
              </button>
              <div className="w-px h-6 bg-gray-700 mx-0.5 sm:mx-1" />
              <HoverHint label="Back to upload screen">
                <button
                  onClick={() => j(true)}
                  className="p-2 text-gray-300 hover:bg-gray-800 hover:text-indigo-400 rounded-md transition-colors"
                >
                  <Home className="w-5 h-5" />
                </button>
              </HoverHint>
              <HoverHint label="Reset all data">
                <button
                  onClick={() => W(true)}
                  className="p-2 text-gray-300 hover:bg-red-950 hover:text-red-400 rounded-md transition-colors"
                >
                  <RotateCcw className="w-5 h-5" />
                </button>
              </HoverHint>
              <div className="w-px h-6 bg-gray-700 mx-0.5 sm:mx-1" />
              <div className="relative">
                <button
                  onClick={() => C(!k)}
                  className="flex items-center gap-1.5 sm:gap-2 px-2.5 sm:px-3 py-2 bg-indigo-600 text-white text-sm font-medium rounded-md hover:bg-indigo-500"
                >
                  <Download className="w-4 h-4" />
                  <span className="hidden sm:inline">Export</span>
                </button>
                {k && (
                  <Fragment>
                    <div className="fixed inset-0 z-10" onClick={() => C(false)} />
                    <div className="absolute right-0 top-full mt-1 bg-gray-800 border border-gray-700 rounded-md shadow-xl z-20 w-64">
                      {h && (
                        <div className="px-3 py-2 border-b border-gray-700">
                          <div className="text-[10px] font-semibold uppercase tracking-wider text-gray-400 mb-1.5">
                            Scope
                          </div>
                          <div className="inline-flex w-full border border-gray-700 rounded-sm overflow-hidden">
                            <button
                              onClick={() => x("focused")}
                              className={`flex-1 px-2 py-1 text-xs transition-colors ${I === "focused" ? "bg-indigo-600 text-white" : "bg-gray-900 text-gray-400 hover:text-gray-200 hover:bg-gray-700"}`}
                            >
                              Focused ({h.size})
                            </button>
                            <button
                              onClick={() => x("all")}
                              className={`flex-1 px-2 py-1 text-xs transition-colors border-l border-gray-700 ${I === "all" ? "bg-indigo-600 text-white" : "bg-gray-900 text-gray-400 hover:text-gray-200 hover:bg-gray-700"}`}
                            >
                              All ({e.length})
                            </button>
                          </div>
                        </div>
                      )}
                      <button
                        onClick={Le}
                        className="w-full text-left px-4 py-2.5 text-sm text-gray-100 hover:bg-gray-700 flex items-center gap-2"
                      >
                        <FileText className="w-4 h-4 text-indigo-400" />
                        <div>
                          <div className="font-medium">HTML Report</div>
                          <div className="text-xs text-gray-400">With embedded screenshots</div>
                        </div>
                      </button>
                      <button
                        onClick={We}
                        className="w-full text-left px-4 py-2.5 text-sm text-gray-100 hover:bg-gray-700 flex items-center gap-2 border-t border-gray-700"
                      >
                        <FileDown className="w-4 h-4 text-indigo-400" />
                        <div>
                          <div className="font-medium">CSV + Screenshots</div>
                          <div className="text-xs text-gray-400">Findings CSV + image files</div>
                        </div>
                      </button>
                    </div>
                  </Fragment>
                )}
              </div>
            </div>
          </div>
          <div className="md:hidden px-3 pb-2.5 pt-0 flex items-center gap-2">
            <span className="text-xs font-mono bg-indigo-950/60 text-indigo-300 px-2 py-0.5 rounded border border-indigo-800 flex-shrink-0">
              {q.ID}
            </span>
            <div className="min-w-0 flex-1">
              <div className="text-sm font-semibold text-gray-100 truncate">{q.Title || "Untitled"}</div>
              {q.Module && <div className="text-[11px] text-gray-400 truncate">{q.Module}</div>}
            </div>
            <span
              className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-semibold border flex-shrink-0 ${ya.color}`}
            >
              <ya.icon className="w-3 h-3" />
              <span className="hidden xs:inline">{oe.status}</span>
            </span>
          </div>
        </div>
        <div ref={xe} className="flex-1 overflow-hidden">
          <style>{`
            @media (min-width: 1200px) {
              .split-container { flex-direction: row !important; overflow: hidden !important; }
              .split-column { overflow-y: auto !important; }
            }
          `}</style>
          <div className="split-container h-full flex flex-col gap-3 p-3 overflow-y-auto">
            <div className="split-column flex-1 min-w-0">
              <div className="bg-gray-900 rounded-lg border border-gray-800 shadow-sm overflow-hidden">
                <div
                  className="px-5 border-b border-gray-800 sticky top-0 z-10 flex items-center"
                  style={{
                    backgroundColor: SURFACE_DARK,
                    height: "52px",
                  }}
                >
                  <h2 className="font-bold text-gray-100">Details & Steps</h2>
                </div>
                <div className="p-5 space-y-4">
                  {(!isEmptyValue(q.Feature) || !isEmptyValue(q["Pre-conditions"])) && (
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                      {!isEmptyValue(q.Feature) && <InfoField label="Feature" value={q.Feature} />}
                      {!isEmptyValue(q["Pre-conditions"]) && (
                        <InfoField label="Pre-conditions" value={q["Pre-conditions"]} />
                      )}
                    </div>
                  )}
                  {!isEmptyValue(q["Test Data"]) && <InfoField label="Test Data" value={q["Test Data"]} />}
                  {Ht.length === 0 && q.Steps && (
                    <InfoField label="Steps (raw — parser found no steps)" value={q.Steps} />
                  )}
                  {Ht.length === 0 && !q.Steps && (
                    <div className="text-center py-3">
                      <button
                        onClick={() => {
                          N();
                          F(true);
                        }}
                        className="inline-flex items-center gap-2 px-3 py-1.5 border border-dashed border-gray-700 text-gray-400 hover:border-indigo-500 hover:text-indigo-300 hover:bg-indigo-950/20 rounded-md text-sm transition-colors"
                      >
                        <Plus className="w-4 h-4" /> Start a checklist
                      </button>
                    </div>
                  )}
                  {Ht.length > 0 && (
                    <div>
                      <div className="flex items-center justify-between mb-2">
                        <div className="text-xs font-semibold uppercase tracking-wider text-gray-400">
                          Steps Checklist
                        </div>
                        <div className="flex items-center gap-2">
                          <div className="text-xs text-gray-400">
                            {Kg} done{_c > 0 ? <span className="text-red-400"> · {_c} blocked</span> : null} ·{" "}
                            {Ht.length} total
                          </div>
                          <button
                            onClick={() => F(!V)}
                            className={`flex items-center gap-1 px-2 py-0.5 rounded text-xs font-medium border transition-colors ${V ? "bg-indigo-600 border-indigo-600 text-white" : "border-gray-700 text-gray-400 hover:text-indigo-300 hover:border-indigo-700"}`}
                            title={V ? "Done editing" : "Edit steps"}
                          >
                            {V ? (
                              <Fragment>
                                <Check className="w-3 h-3" />
                                Done
                              </Fragment>
                            ) : (
                              <Fragment>
                                <Settings2 className="w-3 h-3" />
                                Edit
                              </Fragment>
                            )}
                          </button>
                        </div>
                      </div>
                      <div className="space-y-1.5">
                        {Ht.map((L, R) => {
                          let J = xi(R);
                          let Z = J === "done";
                          let le = J === "blocked";
                          let ke = (oe.stepNotes || {})[R] || "";
                          let Ge = (oe.stepNoteOpen || {})[R];
                          let Ee = Ge || ke.length > 0;
                          return (
                            <div
                              draggable={V}
                              onDragStart={
                                V
                                  ? (Ne) => {
                                      E(R);
                                      Ne.dataTransfer.effectAllowed = "move";
                                      try {
                                        Ne.dataTransfer.setData("text/plain", String(R));
                                      } catch {}
                                    }
                                  : undefined
                              }
                              onDragOver={
                                V
                                  ? (Ne) => {
                                      Ne.preventDefault();
                                      Ne.dataTransfer.dropEffect = "move";
                                      if (D !== R) {
                                        O(R);
                                      }
                                    }
                                  : undefined
                              }
                              onDragLeave={
                                V
                                  ? () => {
                                      if (D === R) {
                                        O(null);
                                      }
                                    }
                                  : undefined
                              }
                              onDrop={
                                V
                                  ? (Ne) => {
                                      Ne.preventDefault();
                                      if (se != null && se !== R) {
                                        Q(se, R);
                                      }
                                      E(null);
                                      O(null);
                                    }
                                  : undefined
                              }
                              onDragEnd={
                                V
                                  ? () => {
                                      E(null);
                                      O(null);
                                    }
                                  : undefined
                              }
                              className={`step-row border rounded-md px-3 py-2 transition-colors ${Z ? "bg-green-950/40 border-green-800" : le ? "bg-red-950/40 border-red-800" : "border-gray-700"} ${V ? "cursor-move" : ""} ${se === R ? "opacity-40" : ""} ${D === R && se !== R ? "ring-2 ring-indigo-500" : ""}`}
                              style={
                                !Z && !le
                                  ? {
                                      backgroundColor: SURFACE_DARK,
                                    }
                                  : {}
                              }
                              key={R}
                            >
                              <div className="flex items-start gap-2">
                                {V && <GripVertical className="w-4 h-4 text-gray-500 flex-shrink-0 mt-1" />}
                                <span className="font-bold text-indigo-400 flex-shrink-0 text-sm leading-relaxed">
                                  {R + 1}.
                                </span>
                                <div className="flex-1 min-w-0">
                                  {V ? (
                                    <textarea
                                      value={L}
                                      onChange={(Ne) => de(R, Ne.target.value)}
                                      placeholder="Step description..."
                                      rows={1}
                                      className="w-full px-2 py-1 text-sm bg-gray-900 border border-gray-700 text-gray-100 placeholder-gray-500 rounded focus:outline-none focus:ring-1 focus:ring-indigo-400 focus:border-indigo-400 resize-y"
                                      onMouseDown={(Ne) => Ne.stopPropagation()}
                                    />
                                  ) : (
                                    <Fragment>
                                      <div className="flex items-start gap-2">
                                        <span
                                          className={`text-sm whitespace-pre-wrap flex-1 leading-relaxed ${Z ? "text-gray-400" : le ? "text-red-200" : "text-gray-100"}`}
                                        >
                                          {normalizePath(L)}
                                        </span>
                                        {!Ee && (
                                          <button
                                            onClick={() =>
                                              H(q.__id, "stepNoteOpen", {
                                                ...(oe.stepNoteOpen || {}),
                                                [R]: true,
                                              })
                                            }
                                            className="step-row-pencil flex-shrink-0 p-1 text-gray-500 hover:text-indigo-400 hover:bg-gray-800 rounded transition-colors"
                                            title="Add note"
                                          >
                                            <Pencil className="w-3.5 h-3.5" />
                                          </button>
                                        )}
                                      </div>
                                      {Ee && (
                                        <div className="relative mt-1.5">
                                          <textarea
                                            value={ke}
                                            onChange={(Ne) => K(q.__id, R, Ne.target.value)}
                                            placeholder="Notes for this step..."
                                            rows={1}
                                            autoFocus={Ge && !ke}
                                            className="w-full px-2 py-1 pr-7 text-sm bg-gray-900 border border-gray-700 text-gray-100 placeholder-gray-500 rounded focus:outline-none focus:ring-1 focus:ring-indigo-400 focus:border-indigo-400 resize-y"
                                          />
                                          {!ke && (
                                            <button
                                              onClick={() =>
                                                H(q.__id, "stepNoteOpen", {
                                                  ...(oe.stepNoteOpen || {}),
                                                  [R]: false,
                                                })
                                              }
                                              className="absolute top-1 right-1 p-0.5 text-gray-500 hover:text-gray-300 rounded"
                                              title="Hide note field"
                                            >
                                              <X className="w-3 h-3" />
                                            </button>
                                          )}
                                        </div>
                                      )}
                                    </Fragment>
                                  )}
                                </div>
                                {V ? (
                                  <button
                                    onClick={() => G(R)}
                                    className="flex-shrink-0 p-1.5 text-gray-500 hover:text-red-400 hover:bg-red-950/40 rounded transition-colors"
                                    title="Delete step"
                                  >
                                    <Trash2 className="w-4 h-4" />
                                  </button>
                                ) : (
                                  <div className="flex flex-col gap-1 flex-shrink-0 ml-1">
                                    <button
                                      onClick={() => te(q.__id, R)}
                                      className={`step-btn rounded-md border-2 flex items-center justify-center transition-all ${Z ? "bg-green-600 border-green-600 text-white hover:bg-green-500" : "border-gray-600 hover:border-green-500 hover:bg-green-950/40"}`}
                                      style={{
                                        width: "26px",
                                        height: "26px",
                                      }}
                                      title={Z ? "Click to uncheck" : "Mark as done"}
                                    >
                                      <CircleCheck
                                        className={`w-4 h-4 ${Z ? "" : "step-btn-preview text-green-500"}`}
                                      />
                                    </button>
                                    <button
                                      onClick={() => {
                                        r((Ne) => {
                                          let _e = Ne[q.__id] || makeFinding();
                                          let Ae = {
                                            ...(_e.stepChecks || {}),
                                          };
                                          if (le) {
                                            delete Ae[R];
                                          } else {
                                            Ae[R] = "blocked";
                                          }
                                          return {
                                            ...Ne,
                                            [q.__id]: {
                                              ..._e,
                                              stepChecks: Ae,
                                            },
                                          };
                                        });
                                      }}
                                      className={`step-btn rounded-md border-2 flex items-center justify-center transition-all ${le ? "bg-red-600 border-red-600 text-white hover:bg-red-500" : "border-gray-600 hover:border-red-500 hover:bg-red-950/40"}`}
                                      style={{
                                        width: "26px",
                                        height: "26px",
                                      }}
                                      title={le ? "Click to clear" : "Mark as unable to proceed"}
                                    >
                                      <X className={`w-4 h-4 ${le ? "" : "step-btn-preview text-red-500"}`} />
                                    </button>
                                  </div>
                                )}
                              </div>
                            </div>
                          );
                        })}
                        {V && (
                          <button
                            onClick={N}
                            className="w-full flex items-center justify-center gap-2 px-3 py-2 border-2 border-dashed border-gray-700 text-gray-400 hover:border-indigo-500 hover:text-indigo-300 hover:bg-indigo-950/20 rounded-md text-sm transition-colors"
                          >
                            <Plus className="w-4 h-4" /> Add step
                          </button>
                        )}
                      </div>
                    </div>
                  )}
                  {!isEmptyValue(q["Expected Result"]) && (
                    <div>
                      <div className="text-xs font-bold uppercase tracking-wider text-amber-400 mb-1.5 flex items-center gap-1.5">
                        <CircleCheck className="w-3.5 h-3.5" />
                        Expected Result
                      </div>
                      <div className="text-sm whitespace-pre-wrap bg-amber-950/30 border-l-4 border-amber-500 px-3 py-2 rounded text-amber-100">
                        {normalizePath(q["Expected Result"])}
                      </div>
                    </div>
                  )}
                  {!isEmptyValue(q["Downstream Check"]) && (
                    <InfoField label="Downstream Check" value={q["Downstream Check"]} />
                  )}
                  {!isEmptyValue(q["Edge Case"]) && <InfoField label="Edge Case" value={q["Edge Case"]} />}
                </div>
              </div>
            </div>
            <div className="split-column flex-1 min-w-0">
              <div className="bg-gray-900 rounded-lg border border-gray-800 shadow-sm overflow-hidden">
                <div
                  className="px-6 border-b border-gray-800 flex items-center"
                  style={{
                    backgroundColor: "rgb(20 24 35)",
                    height: "52px",
                  }}
                >
                  <h2 className="font-bold text-gray-100">Findings</h2>
                </div>
                <div className="p-5 space-y-4 divide-y divide-gray-800/60">
                  <div>
                    <label className="block text-xs font-semibold uppercase tracking-wider text-gray-400 mb-1.5">
                      Tester
                    </label>
                    {m.length > 0 ? (
                      <select
                        value={oe.testedBy}
                        onChange={(L) => P(q.__id, L.target.value)}
                        className="dark-select w-full px-3 py-2 bg-gray-800 border border-gray-700 text-gray-100 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500"
                      >
                        <option value="">— Select tester —</option>
                        {m.map((L) => (
                          <option value={L} key={L}>
                            {L}
                          </option>
                        ))}
                      </select>
                    ) : (
                      <input
                        type="text"
                        value={oe.testedBy}
                        onChange={(L) => P(q.__id, L.target.value)}
                        placeholder="Your name"
                        className="w-full px-3 py-2 bg-gray-800 border border-gray-700 text-gray-100 placeholder-gray-500 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-transparent"
                      />
                    )}
                  </div>
                  <div className="pt-4">
                    <label className="block text-xs font-semibold uppercase tracking-wider text-gray-400 mb-1.5">
                      Tester Notes
                    </label>
                    <textarea
                      value={oe.notes}
                      onChange={(L) => H(q.__id, "notes", L.target.value)}
                      rows={4}
                      placeholder="Observations, summary of issues found, deviations from expected result..."
                      className="w-full px-3 py-2 bg-gray-800 border border-gray-700 text-gray-100 placeholder-gray-500 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-transparent resize-y"
                    />
                  </div>
                  <div className="pt-4">
                    <label className="block text-xs font-semibold uppercase tracking-wider text-gray-400 mb-1.5">
                      Other Observations
                    </label>
                    <textarea
                      value={oe.otherObservations}
                      onChange={(L) => H(q.__id, "otherObservations", L.target.value)}
                      rows={3}
                      placeholder="Anything else worth flagging — UX, performance, side effects..."
                      className="w-full px-3 py-2 bg-gray-800 border border-gray-700 text-gray-100 placeholder-gray-500 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-transparent resize-y"
                    />
                  </div>
                  <div className="pt-4">
                    <label className="block text-xs font-semibold uppercase tracking-wider text-gray-400 mb-1.5">
                      Recordings{" "}
                      {oe.recordings?.length > 0 && (
                        <span className="normal-case font-normal text-gray-500">
                          ({oe.recordings.length})
                        </span>
                      )}
                    </label>
                    <div className="flex gap-2">
                      <input
                        type="url"
                        value={ue}
                        onChange={(L) => pe(L.target.value)}
                        onKeyDown={(L) => {
                          if (L.key === "Enter") {
                            L.preventDefault();
                            let R = ue.trim();
                            if (!R) {
                              return;
                            }
                            let J = [
                              ...(oe.recordings || []),
                              {
                                url: R,
                              },
                            ];
                            H(q.__id, "recordings", J);
                            pe("");
                          }
                        }}
                        placeholder="Paste Loom / Drive / YouTube / Vimeo / video URL"
                        className="flex-1 px-3 py-2 bg-gray-800 border border-gray-700 text-gray-100 placeholder-gray-500 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-transparent text-sm"
                      />
                      <button
                        type="button"
                        onClick={() => {
                          let L = ue.trim();
                          if (!L) {
                            return;
                          }
                          let R = [
                            ...(oe.recordings || []),
                            {
                              url: L,
                            },
                          ];
                          H(q.__id, "recordings", R);
                          pe("");
                        }}
                        disabled={!ue.trim()}
                        className="px-3 py-2 bg-indigo-600 text-white text-sm font-medium rounded-md hover:bg-indigo-500 disabled:bg-gray-700 disabled:text-gray-500 disabled:cursor-not-allowed"
                      >
                        Add
                      </button>
                    </div>
                    {oe.recordings?.length > 0 && (
                      <div className="mt-2 space-y-1.5">
                        {oe.recordings.map((L, R) => {
                          let Z = parseRecordingUrl(L.url)?.label || "Link";
                          return (
                            <div
                              className="flex items-center gap-2 px-2.5 py-1.5 bg-gray-800/60 border border-gray-700 rounded-sm group"
                              key={R}
                            >
                              <span className="text-[10px] font-semibold uppercase tracking-wider px-1.5 py-0.5 bg-gray-900 text-indigo-300 border border-gray-700 rounded-sm flex-shrink-0">
                                {Z}
                              </span>
                              <a
                                href={L.url}
                                target="_blank"
                                rel="noopener noreferrer"
                                className="text-xs text-gray-300 hover:text-indigo-300 truncate flex-1 min-w-0"
                                title={L.url}
                              >
                                {L.url}
                              </a>
                              <button
                                type="button"
                                onClick={() => {
                                  let le = (oe.recordings || []).filter((ke, Ge) => Ge !== R);
                                  H(q.__id, "recordings", le);
                                }}
                                className="opacity-40 group-hover:opacity-100 hover:text-red-400 text-gray-400 text-sm leading-none flex-shrink-0 px-1"
                                aria-label="Remove recording"
                              >
                                ×
                              </button>
                            </div>
                          );
                        })}
                      </div>
                    )}
                  </div>
                  <div className="pt-4">
                    <label className="block text-xs font-semibold uppercase tracking-wider text-gray-400 mb-1.5">
                      Screenshots{" "}
                      {oe.screenshots.length > 0 && (
                        <span className="normal-case font-normal text-gray-500">
                          ({oe.screenshots.length})
                        </span>
                      )}
                    </label>
                    <div
                      onClick={() => be.current?.click()}
                      onDragOver={(L) => {
                        L.preventDefault();
                        L.currentTarget.classList.add("border-indigo-500", "bg-indigo-950/30");
                      }}
                      onDragLeave={(L) => {
                        L.currentTarget.classList.remove("border-indigo-500", "bg-indigo-950/30");
                      }}
                      onDrop={(L) => {
                        L.preventDefault();
                        L.currentTarget.classList.remove("border-indigo-500", "bg-indigo-950/30");
                        if (L.dataTransfer.files.length > 0) {
                          z(q.__id, L.dataTransfer.files);
                        }
                      }}
                      className="border-2 border-dashed border-gray-700 rounded-md p-5 text-center cursor-pointer hover:border-indigo-500 hover:bg-indigo-950/20 transition-all"
                    >
                      <ImageIcon className="w-7 h-7 text-gray-500 mx-auto mb-1.5" />
                      <p className="text-sm text-gray-300">Click, drag, or paste screenshots</p>
                      <p className="text-xs text-gray-500 mt-0.5">💡 Tip: paste with Ctrl/Cmd+V</p>
                      <input
                        ref={be}
                        type="file"
                        accept="image/*"
                        multiple={true}
                        className="hidden"
                        onChange={(L) => {
                          if (L.target.files.length > 0) {
                            z(q.__id, L.target.files);
                          }
                          L.target.value = "";
                        }}
                      />
                    </div>
                    {oe.screenshots.length > 0 && (
                      <div className="grid grid-cols-2 gap-3 mt-3">
                        {oe.screenshots.map((L, R) => (
                          <div
                            className="relative group border border-gray-700 rounded-md overflow-hidden bg-gray-800"
                            key={R}
                          >
                            <div className="absolute top-2 left-2 z-10 bg-indigo-600 text-white text-xs font-bold px-2 py-0.5 rounded shadow-md">
                              Fig. {R + 1}
                            </div>
                            <img src={L.data} alt={L.name} className="w-full h-32 object-cover" />
                            <div className="absolute inset-0 bg-black/0 group-hover:bg-black/60 transition-colors flex items-center justify-center">
                              <button
                                onClick={() => re(q.__id, R)}
                                className="opacity-0 group-hover:opacity-100 bg-red-600 text-white p-2 rounded-full hover:bg-red-500 transition-all"
                              >
                                <Trash2 className="w-4 h-4" />
                              </button>
                            </div>
                            <div className="px-2 py-1 text-xs text-gray-400 truncate bg-gray-900 border-t border-gray-700">
                              {L.name}
                            </div>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                  <div className="pt-4">
                    <div className="flex items-center justify-between mb-1.5">
                      <label className="block text-xs font-semibold uppercase tracking-wider text-gray-400">
                        Final Status
                      </label>
                      {oe.status && oe.status !== "Not Tested" && (
                        <button
                          onClick={() => H(q.__id, "status", "Not Tested")}
                          className="text-xs text-gray-500 hover:text-red-400 transition-colors flex items-center gap-1"
                          title="Clear status"
                        >
                          <X className="w-3 h-3" />
                          Clear
                        </button>
                      )}
                    </div>
                    <div className="grid grid-cols-2 gap-1.5">
                      {STATUS_OPTIONS.filter((L) => L.value !== "Not Tested").map((L) => {
                        let R = L.icon;
                        let J = oe.status === L.value;
                        return (
                          <button
                            onClick={() => H(q.__id, "status", L.value)}
                            className={`flex items-center justify-center gap-1 px-2 py-2 rounded-md text-xs font-medium border transition-all ${J ? L.color + " ring-2 ring-offset-1 ring-offset-gray-900 ring-indigo-400" : "bg-gray-800 text-gray-300 border-gray-700 hover:bg-gray-700"}`}
                            key={L.value}
                          >
                            <R className="w-3.5 h-3.5 flex-shrink-0" />
                            <span className="truncate">{L.value}</span>
                          </button>
                        );
                      })}
                    </div>
                  </div>
                  {isFailureStatus(oe.status) && (
                    <div className="pt-4">
                      <div className="flex items-center justify-between mb-1.5">
                        <label className="block text-xs font-semibold uppercase tracking-wider text-gray-400">
                          Bug Severity
                        </label>
                        {oe.bugSeverity && (
                          <button
                            onClick={() => H(q.__id, "bugSeverity", "")}
                            className="text-xs text-gray-500 hover:text-red-400 transition-colors flex items-center gap-1"
                            title="Clear severity"
                          >
                            <X className="w-3 h-3" />
                            Clear
                          </button>
                        )}
                      </div>
                      <div className="grid grid-cols-3 gap-1.5">
                        {[
                          {
                            value: "P0",
                            label: "P0 · Critical",
                            color: "bg-red-700 text-white border-red-600",
                          },
                          {
                            value: "P1",
                            label: "P1 · Major",
                            color: "bg-orange-700 text-white border-orange-600",
                          },
                          {
                            value: "P2",
                            label: "P2 · Minor",
                            color: "bg-yellow-700 text-white border-yellow-600",
                          },
                        ].map((L) => {
                          let R = oe.bugSeverity === L.value;
                          return (
                            <button
                              onClick={() => H(q.__id, "bugSeverity", L.value)}
                              className={`flex items-center justify-center px-2 py-2 rounded-md text-xs font-medium border transition-all ${R ? L.color + " ring-2 ring-offset-1 ring-offset-gray-900 ring-indigo-400" : "bg-gray-800 text-gray-300 border-gray-700 hover:bg-gray-700"}`}
                              key={L.value}
                            >
                              <span className="truncate">{L.label}</span>
                            </button>
                          );
                        })}
                      </div>
                    </div>
                  )}
                  <div className="flex justify-between pt-4">
                    <button
                      onClick={() => s(Math.max(0, o - 1))}
                      disabled={o === 0}
                      className="flex items-center gap-2 px-4 py-2 bg-gray-800 border border-gray-700 text-gray-200 rounded-md hover:bg-gray-700 disabled:opacity-30 disabled:cursor-not-allowed font-medium text-sm"
                    >
                      <ChevronLeft className="w-4 h-4" /> Previous
                    </button>
                    <button
                      onClick={() => s(Math.min(e.length - 1, o + 1))}
                      disabled={o === e.length - 1}
                      className="flex items-center gap-2 px-4 py-2 bg-indigo-600 text-white rounded-md hover:bg-indigo-500 disabled:opacity-30 disabled:cursor-not-allowed font-medium text-sm"
                    >
                      Next <ChevronRight className="w-4 h-4" />
                    </button>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
      <FocusModeModal
        open={_}
        onClose={() => S(false)}
        testCases={e}
        currentFocusedIds={h}
        onApply={(L) => b(L)}
      />
      {M && (
        <div
          className="fixed top-0 left-0 right-0 bottom-0 flex items-center justify-center p-4"
          style={{
            zIndex: 99999,
            backgroundColor: "rgba(0,0,0,0.75)",
          }}
          onClick={() => W(false)}
        >
          <div
            className="bg-gray-900 rounded-lg shadow-2xl max-w-md w-full overflow-hidden border-2 border-red-700"
            onClick={(L) => L.stopPropagation()}
          >
            <div className="p-6">
              <div className="flex items-start gap-3 mb-4">
                <div className="flex-shrink-0 w-10 h-10 bg-red-950 rounded-full flex items-center justify-center border border-red-800">
                  <CircleAlert className="w-5 h-5 text-red-400" />
                </div>
                <div className="flex-1">
                  <h3 className="text-lg font-bold text-gray-100 mb-1">Reset all data?</h3>
                  <p className="text-sm text-gray-400">
                    This will permanently delete everything from your current session.
                  </p>
                </div>
              </div>
              <ul
                className="text-sm text-gray-300 space-y-1 mb-5"
                style={{
                  paddingLeft: "52px",
                }}
              >
                <li>• All {e.length} test cases</li>
                <li>• All findings, notes, and step checkboxes</li>
                <li>• All uploaded screenshots</li>
                <li>• Tester names ({m.length})</li>
              </ul>
              <div className="bg-amber-950/40 border border-amber-800 rounded-md p-3 text-sm text-amber-300 mb-5">
                <strong>This cannot be undone.</strong> Consider exporting your findings first if you want to
                keep them.
              </div>
              <div className="flex justify-between gap-2">
                <button
                  onClick={ce}
                  className="px-4 py-2 bg-red-600 text-white rounded-md hover:bg-red-500 font-medium text-sm"
                >
                  Yes, reset everything
                </button>
                <button
                  onClick={() => W(false)}
                  className="px-4 py-2 bg-gray-800 border border-gray-700 text-gray-200 rounded-md hover:bg-gray-700 font-medium text-sm"
                >
                  Cancel
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
function HoverHint({ label: e, children: t }) {
  let [a, r] = useState(false);
  return (
    <div className="relative inline-flex" onMouseEnter={() => r(true)} onMouseLeave={() => r(false)}>
      {t}
      {a && (
        <div className="absolute top-full left-1/2 -translate-x-1/2 mt-1.5 px-2 py-1 bg-black text-white text-xs rounded whitespace-nowrap pointer-events-none z-50 shadow-lg border border-gray-700">
          {e}
          <div className="absolute bottom-full left-1/2 -translate-x-1/2 w-0 h-0 border-l-4 border-r-4 border-b-4 border-transparent border-b-black" />
        </div>
      )}
    </div>
  );
}
function CsvUploadScreen({ onLoad: e, fileInputRef: t, hasExistingData: a, onBack: r, existingTesters: o }) {
  let [s, l] = useState("");
  let [i, u] = useState(o || []);
  let [d, c] = useState(null);
  let [p, h] = useState(false);
  let [b, _] = useState(false);
  let [S, I] = useState(null);
  let x = async () => {
    let m = false;
    try {
      if (navigator.clipboard?.writeText) {
        await navigator.clipboard.writeText(CSV_GENERATION_PROMPT);
        m = true;
      }
    } catch (f) {
      console.warn("Clipboard API failed, falling back:", f);
    }
    if (!m) {
      try {
        let f = document.createElement("textarea");
        f.value = CSV_GENERATION_PROMPT;
        f.style.position = "fixed";
        f.style.opacity = "0";
        document.body.appendChild(f);
        f.focus();
        f.select();
        m = document.execCommand("copy");
        document.body.removeChild(f);
      } catch (f) {
        console.error("Fallback copy failed:", f);
      }
    }
    if (m) {
      h(true);
      setTimeout(() => h(false), 2500);
    } else {
      alert(
        "Could not copy automatically. The prompt has been logged to the browser console — open DevTools to grab it.",
      );
      console.log(`CSV_GENERATION_PROMPT:

${CSV_GENERATION_PROMPT}`);
    }
  };
  let g = () => {
    let m = s.trim();
    if (!m) {
      return;
    }
    let w = m
      .split(",")
      .map((T) => T.trim())
      .filter((T) => T.length > 0)
      .filter((T) => !i.includes(T));
    if (w.length > 0) {
      u([...i, ...w]);
    }
    l("");
  };
  let v = (m) => {
    u(i.filter((f) => f !== m));
  };
  let k = (m) => {
    c(m);
    I(null);
  };
  let C = async () => {
    if (!d || b) {
      return;
    }
    _(true);
    I(null);
    let m = await e(d, i);
    _(false);
    if (m?.error) {
      I(m.error);
    }
  };
  return (
    <div className="min-h-screen bg-gradient-to-br from-gray-950 via-gray-900 to-indigo-950 flex items-center justify-center p-6">
      <style>{`
        ::-webkit-scrollbar { width: 10px; height: 10px; }
        ::-webkit-scrollbar-track { background: #0a0a0f; }
        ::-webkit-scrollbar-thumb { background: #2a2a35; border-radius: 5px; border: 2px solid #0a0a0f; }
        ::-webkit-scrollbar-thumb:hover { background: #3a3a48; }
        ::-webkit-scrollbar-corner { background: #0a0a0f; }
        * { scrollbar-color: #2a2a35 #0a0a0f; scrollbar-width: thin; }
      `}</style>
      <div className="max-w-2xl w-full">
        <div className="text-center mb-8">
          <div className="inline-flex items-center justify-center w-16 h-16 bg-indigo-600 rounded-2xl mb-4 shadow-lg shadow-indigo-900/50">
            <FileText className="w-8 h-8 text-white" />
          </div>
          <h1 className="text-3xl font-bold text-gray-100 mb-2">
            Test Case Report Compiler
            <span
              className={`ml-2 align-middle inline-block text-[10px] font-bold uppercase tracking-widest px-2 py-0.5 rounded-sm border ${EDITION === "Pro" ? "bg-emerald-900/40 text-emerald-300 border-emerald-700" : "bg-gray-800 text-gray-400 border-gray-700"}`}
            >
              {EDITION}
            </span>
          </h1>
        </div>
        {a && (
          <div className="bg-amber-950/40 border border-amber-800 rounded-lg p-3 mb-4 text-sm text-amber-300">
            <strong>Heads up:</strong> Uploading a new CSV will replace your current test cases and findings.
            Use the button below to return to your session.
          </div>
        )}
        <div
          onClick={() => !d && t.current?.click()}
          onDragOver={(m) => {
            m.preventDefault();
            if (!d) {
              m.currentTarget.classList.add("border-indigo-500", "bg-indigo-950/30");
            }
          }}
          onDragLeave={(m) => {
            m.currentTarget.classList.remove("border-indigo-500", "bg-indigo-950/30");
          }}
          onDrop={(m) => {
            m.preventDefault();
            m.currentTarget.classList.remove("border-indigo-500", "bg-indigo-950/30");
            if (m.dataTransfer.files[0]) {
              k(m.dataTransfer.files[0]);
            }
          }}
          className={`bg-gray-900 border-2 border-dashed rounded-2xl p-10 text-center transition-all shadow-sm mb-4 ${d ? (S ? "border-red-600 bg-red-950/20" : "border-green-600 bg-green-950/20") : "border-gray-700 cursor-pointer hover:border-indigo-500 hover:bg-indigo-950/20"}`}
        >
          {d ? (
            <Fragment>
              {S ? (
                <CircleAlert className="w-12 h-12 text-red-500 mx-auto mb-3" />
              ) : (
                <CircleCheck className="w-12 h-12 text-green-500 mx-auto mb-3" />
              )}
              <p className="text-lg font-medium text-gray-100 mb-1">{d.name}</p>
              <p className="text-sm text-gray-400 mb-4">{(d.size / 1024).toFixed(1)} KB</p>
              {S && (
                <div className="bg-red-950/60 border border-red-700 rounded-md p-3 mb-4 text-sm text-red-200 text-left">
                  <div className="flex items-start gap-2">
                    <CircleAlert className="w-4 h-4 flex-shrink-0 mt-0.5" />
                    <div>
                      <strong className="font-semibold">Couldn't load this CSV</strong>
                      <div className="mt-1 opacity-90">{S}</div>
                    </div>
                  </div>
                </div>
              )}
              <div className="flex gap-2 justify-center">
                <button
                  onClick={(m) => {
                    m.stopPropagation();
                    k(null);
                  }}
                  className="px-4 py-2 bg-gray-800 border border-gray-700 text-gray-200 rounded-md hover:bg-gray-700 text-sm font-medium"
                  disabled={b}
                >
                  Choose different file
                </button>
                {!S && (
                  <button
                    onClick={(m) => {
                      m.stopPropagation();
                      C();
                    }}
                    className="px-6 py-2 bg-indigo-600 text-white rounded-md hover:bg-indigo-500 text-sm font-medium disabled:opacity-60 disabled:cursor-not-allowed"
                    disabled={b}
                  >
                    {b ? "Loading…" : "Start Testing →"}
                  </button>
                )}
              </div>
            </Fragment>
          ) : (
            <Fragment>
              <Upload className="w-12 h-12 text-gray-500 mx-auto mb-3" />
              <p className="text-lg font-medium text-gray-200 mb-1">Drop your CSV file here</p>
              <p className="text-sm text-gray-400">or click to browse</p>
              <input
                ref={t}
                type="file"
                accept=".csv"
                className="hidden"
                onChange={(m) => k(m.target.files[0])}
              />
            </Fragment>
          )}
        </div>
        <div className="text-xs text-gray-500 text-center mb-4">
          Expected columns: ID, Module, Feature, Title, Tester [...], Pre-conditions, Test Data, Steps,
          Expected Result, Downstream Check, Edge Case, Pass/Fail, Tester Notes, Bug Severity, Other
          Observations
        </div>
        <div className="text-center mb-4">
          <button
            onClick={x}
            className={`inline-flex items-center gap-2 px-3 py-1.5 rounded-md text-xs font-medium border transition-all ${p ? "bg-green-950/40 border-green-700 text-green-300" : "bg-gray-900 border-gray-700 text-gray-400 hover:text-indigo-300 hover:border-indigo-700 hover:bg-indigo-950/30"}`}
            title="Copy a prompt you can paste into Claude or ChatGPT to generate a CSV that fits this app"
          >
            {p ? (
              <Fragment>
                <CircleCheck className="w-3.5 h-3.5" />
                Prompt copied — paste into Claude/ChatGPT
              </Fragment>
            ) : (
              <Fragment>
                <FileText className="w-3.5 h-3.5" />
                Don't have a CSV? Copy AI prompt to generate one
              </Fragment>
            )}
          </button>
        </div>
        {a && (
          <button
            onClick={r}
            className="w-full mb-4 flex items-center justify-center gap-2 px-4 py-3 bg-gray-900 border-2 border-indigo-700 text-indigo-300 rounded-lg hover:bg-indigo-950/40 hover:border-indigo-500 transition-all font-semibold shadow-sm"
          >
            <ChevronLeft className="w-5 h-5" />
            Back to current session
          </button>
        )}
        <div className="bg-gray-900 rounded-xl p-6 border border-gray-800 shadow-sm mb-4">
          <div className="flex items-center gap-2 mb-3">
            <Users className="w-5 h-5 text-indigo-400" />
            <h2 className="font-semibold text-gray-100">
              Tester Names <span className="text-sm font-normal text-gray-400">(optional)</span>
            </h2>
          </div>
          <p className="text-sm text-gray-400 mb-3">
            Add one or more tester names. You'll be able to pick from this list on each test case page.
          </p>
          <div className="flex gap-2 mb-3">
            <input
              type="text"
              value={s}
              onChange={(m) => l(m.target.value)}
              onKeyDown={(m) => {
                if (m.key === "Enter") {
                  m.preventDefault();
                  g();
                }
              }}
              placeholder="Enter name(s), comma-separated (e.g. Alice, Bob, Carol)"
              className="flex-1 px-3 py-2 bg-gray-800 border border-gray-700 text-gray-100 placeholder-gray-500 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-transparent"
            />
            <button
              onClick={g}
              disabled={!s.trim()}
              className="flex items-center gap-1 px-4 py-2 bg-indigo-600 text-white rounded-md hover:bg-indigo-500 disabled:opacity-40 disabled:cursor-not-allowed font-medium text-sm"
            >
              <Plus className="w-4 h-4" /> Add
            </button>
          </div>
          {i.length > 0 && (
            <div className="flex flex-wrap gap-2">
              {i.map((m) => (
                <span
                  className="inline-flex items-center gap-1.5 bg-indigo-950/60 text-indigo-300 px-3 py-1 rounded-full text-sm font-medium border border-indigo-800"
                  key={m}
                >
                  {m}
                  <button onClick={() => v(m)} className="hover:bg-indigo-800 rounded-full p-0.5">
                    <X className="w-3 h-3" />
                  </button>
                </span>
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
function FocusModeModal({ open: e, onClose: t, testCases: a, currentFocusedIds: r, onApply: o }) {
  let [s, l] = useState(() => new Set());
  let [i, u] = useState("");
  let [d, c] = useState("flat");
  let [p, h] = useState("");
  let [b, _] = useState("");
  let S =
    EDITION === "Pro"
      ? {
          solid: "bg-emerald-600 text-white hover:bg-emerald-500",
          tab: "bg-emerald-600 text-white",
          check: "bg-emerald-600 border-emerald-600",
          checkPartial: "bg-emerald-600/40 border-emerald-500",
          link: "text-emerald-300 hover:text-emerald-200",
          inputFocus: "focus:border-emerald-500",
        }
      : {
          solid: "bg-indigo-600 text-white hover:bg-indigo-500",
          tab: "bg-indigo-600 text-white",
          check: "bg-indigo-600 border-indigo-600",
          checkPartial: "bg-indigo-600/40 border-indigo-500",
          link: "text-indigo-300 hover:text-indigo-200",
          inputFocus: "focus:border-indigo-500",
        };
  useEffect(() => {
    if (e) {
      l(r ? new Set(r) : new Set(a.map((D) => D.__id)));
      u("");
      h("");
      _("");
    }
  }, [e, r, a]);
  if (!e) {
    return null;
  }
  let I = i.trim().toLowerCase();
  let x = I
    ? a.filter(
        (D) =>
          (D.ID || "").toLowerCase().includes(I) ||
          (D.Title || "").toLowerCase().includes(I) ||
          (D.Module || "").toLowerCase().includes(I) ||
          (D.Feature || "").toLowerCase().includes(I),
      )
    : a;
  let g = {};
  let v = [];
  x.forEach((D) => {
    let O = (D.Module || "").trim() || "Other";
    if (!g[O]) {
      g[O] = [];
      v.push(O);
    }
    g[O].push(D);
  });
  let k = v.filter((D) => D !== "Other").concat(v.includes("Other") ? ["Other"] : []);
  let C = (D) => {
    l((O) => {
      let B = new Set(O);
      if (B.has(D)) {
        B.delete(D);
      } else {
        B.add(D);
      }
      return B;
    });
  };
  let m = (D) => {
    let O = D.every((B) => s.has(B.__id));
    l((B) => {
      let ne = new Set(B);
      D.forEach((he) => {
        if (O) {
          ne.delete(he.__id);
        } else {
          ne.add(he.__id);
        }
      });
      return ne;
    });
  };
  let f = () => l(new Set(a.map((D) => D.__id)));
  let w = () => l(new Set());
  let T = () => {
    o(null);
    t();
  };
  let A = () => {
    if (s.size !== 0) {
      if (s.size === a.length) {
        o(null);
      } else {
        o(s);
      }
      t();
    }
  };
  let j = a.length ? Math.min(...a.map((D) => D.__rowNum || 0).filter((D) => D > 0)) : 0;
  let M = a.length ? Math.max(...a.map((D) => D.__rowNum || 0)) : 0;
  let W = parseInt(p, 10);
  let V = parseInt(b, 10);
  let F = !isNaN(W) && !isNaN(V);
  let se = () => {
    if (!F) {
      return;
    }
    let D = Math.min(W, V);
    let O = Math.max(W, V);
    let B = a.filter((ne) => ne.__rowNum >= D && ne.__rowNum <= O).map((ne) => ne.__id);
    l(new Set(B));
  };
  let E = (D) => {
    if (D.key === "Enter") {
      D.preventDefault();
      se();
    }
  };
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/70 p-4" onClick={t}>
      <div
        className="bg-gray-900 border border-gray-700 rounded-sm w-full max-w-2xl max-h-[90vh] flex flex-col shadow-2xl"
        onClick={(D) => D.stopPropagation()}
      >
        <div className="flex items-center justify-between px-5 py-3 border-b border-gray-800 flex-shrink-0">
          <div>
            <h2 className="text-base font-bold text-gray-100">Focus tests</h2>
            <div className="text-xs text-gray-400 mt-0.5">
              {s.size} of {a.length} selected
            </div>
          </div>
          <button onClick={t} className="text-gray-400 hover:text-gray-100 p-1" title="Close">
            <X className="w-5 h-5" />
          </button>
        </div>
        <div className="px-5 py-3 border-b border-gray-800 flex items-center gap-2 flex-wrap flex-shrink-0">
          <div className="relative flex-1 min-w-[200px]">
            <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-gray-500" />
            <input
              type="text"
              placeholder="Search ID, title, module, feature..."
              value={i}
              onChange={(D) => u(D.target.value)}
              className={`w-full pl-9 pr-3 py-1.5 bg-gray-800 border border-gray-700 text-gray-100 placeholder-gray-500 rounded-sm text-sm focus:outline-none ${S.inputFocus}`}
            />
          </div>
          <div className="inline-flex border border-gray-700 rounded-sm overflow-hidden">
            <button
              onClick={() => c("flat")}
              className={`px-2.5 py-1.5 text-xs transition-colors ${d === "flat" ? S.tab : "bg-gray-800 text-gray-400 hover:text-gray-200 hover:bg-gray-700"}`}
              title="Flat list in CSV order"
            >
              Flat
            </button>
            <button
              onClick={() => c("grouped")}
              className={`px-2.5 py-1.5 text-xs transition-colors border-l border-gray-700 ${d === "grouped" ? S.tab : "bg-gray-800 text-gray-400 hover:text-gray-200 hover:bg-gray-700"}`}
              title="Group tests by Module"
            >
              Grouped
            </button>
          </div>
          <button onClick={f} className={`text-xs px-2 py-1.5 ${S.link}`}>
            Select all
          </button>
          <button onClick={w} className="text-xs text-gray-400 hover:text-gray-200 px-2 py-1.5">
            Select none
          </button>
        </div>
        <div className="px-5 py-2.5 border-b border-gray-800 flex items-center gap-2 flex-wrap flex-shrink-0 bg-gray-950/40">
          <span className="text-xs text-gray-400 font-medium">Row</span>
          <input
            type="text"
            inputMode="numeric"
            value={p}
            onChange={(D) => h(D.target.value.replace(/\D/g, ""))}
            onKeyDown={E}
            placeholder={String(j)}
            className={`w-14 px-2 py-1 bg-gray-800 border border-gray-700 text-gray-100 placeholder-gray-600 rounded-sm text-xs text-center focus:outline-none ${S.inputFocus}`}
          />
          <span className="text-xs text-gray-500">to</span>
          <input
            type="text"
            inputMode="numeric"
            value={b}
            onChange={(D) => _(D.target.value.replace(/\D/g, ""))}
            onKeyDown={E}
            placeholder={String(M)}
            className={`w-14 px-2 py-1 bg-gray-800 border border-gray-700 text-gray-100 placeholder-gray-600 rounded-sm text-xs text-center focus:outline-none ${S.inputFocus}`}
          />
          <button
            onClick={se}
            disabled={!F}
            className={`text-xs px-2.5 py-1 rounded-sm ${S.solid} disabled:bg-gray-800 disabled:text-gray-500 disabled:cursor-not-allowed transition-colors font-medium`}
          >
            Select range
          </button>
          <span className="text-[10px] text-gray-500 ml-auto">
            Rows {j}–{M} available
          </span>
        </div>
        {d === "flat" && (
          <div className="px-3 py-1.5 border-b border-gray-800 flex items-stretch gap-2 text-[10px] uppercase tracking-wider text-gray-500 font-semibold flex-shrink-0 bg-gray-900">
            <div className="w-3.5 flex-shrink-0" aria-hidden="true" />
            <div className="w-14 flex-shrink-0 text-right whitespace-nowrap self-center">Cell Row</div>
            <div className="w-px bg-gray-700" aria-hidden="true" />
            <div className="flex-1 self-center">Test Cases</div>
          </div>
        )}
        <div className="flex-1 overflow-y-auto px-3 py-3">
          {k.length === 0 && (
            <div className="text-center text-gray-500 py-8 text-sm">No tests match your search.</div>
          )}
          {d === "flat" &&
            x.map((D) => {
              let O = s.has(D.__id);
              return (
                <button
                  onClick={() => C(D.__id)}
                  className="w-full flex items-start gap-2 px-3 py-1.5 hover:bg-gray-800/60 transition-colors text-left rounded-sm"
                  key={D.__id}
                >
                  <div
                    className={`w-3.5 h-3.5 rounded-sm border-2 flex items-center justify-center flex-shrink-0 mt-0.5 ${O ? S.check : "border-gray-600"}`}
                  >
                    {O && <Check className="w-2.5 h-2.5 text-white" />}
                  </div>
                  <div
                    className="w-14 flex-shrink-0 text-right text-[11px] font-mono text-gray-500 pt-px tabular-nums select-none"
                    title={`CSV row ${D.__rowNum}`}
                  >
                    Row {D.__rowNum}
                  </div>
                  <div className="flex-1 min-w-0">
                    <div className="text-[11px] text-gray-500">
                      <span className="font-mono text-gray-400">{D.ID}</span>
                      <span> · {D.Module || "—"}</span>
                      {D.Feature && String(D.Feature).trim() !== "—" && <span> · {D.Feature}</span>}
                    </div>
                    <div className="text-sm text-gray-200 truncate">{D.Title || "Untitled"}</div>
                  </div>
                </button>
              );
            })}
          {d === "grouped" &&
            k.map((D) => {
              let O = g[D];
              let B = O.filter((ie) => s.has(ie.__id)).length;
              let ne = B === O.length;
              let he = B > 0 && !ne;
              return (
                <div className="mb-4 last:mb-0" key={D}>
                  <button
                    onClick={() => m(O)}
                    className="w-full flex items-center gap-2 px-3 py-2 bg-gray-800 hover:bg-gray-750 border border-gray-700 rounded-sm transition-colors"
                  >
                    <div
                      className={`w-4 h-4 rounded-sm border-2 flex items-center justify-center flex-shrink-0 ${ne ? S.check : he ? S.checkPartial : "border-gray-500"}`}
                    >
                      {ne && <Check className="w-3 h-3 text-white" />}
                      {he && <div className="w-2 h-0.5 bg-white" />}
                    </div>
                    <span className="font-semibold text-sm text-gray-100">{D}</span>
                    <span className="text-xs text-gray-400 ml-auto">
                      {B} / {O.length}
                    </span>
                  </button>
                  <div className="ml-2 mt-1">
                    {O.map((ie) => {
                      let ue = s.has(ie.__id);
                      return (
                        <button
                          onClick={() => C(ie.__id)}
                          className="w-full flex items-start gap-2 px-3 py-1.5 hover:bg-gray-800/60 transition-colors text-left rounded-sm"
                          key={ie.__id}
                        >
                          <div
                            className={`w-3.5 h-3.5 rounded-sm border-2 flex items-center justify-center flex-shrink-0 mt-0.5 ${ue ? S.check : "border-gray-600"}`}
                          >
                            {ue && <Check className="w-2.5 h-2.5 text-white" />}
                          </div>
                          <div
                            className="w-14 flex-shrink-0 text-right text-[11px] font-mono text-gray-500 pt-px tabular-nums select-none"
                            title={`CSV row ${ie.__rowNum}`}
                          >
                            Row {ie.__rowNum}
                          </div>
                          <div className="flex-1 min-w-0">
                            <div className="text-[11px] text-gray-500">
                              <span className="font-mono text-gray-400">{ie.ID}</span>
                              <span> · {ie.Module || "—"}</span>
                              {ie.Feature && String(ie.Feature).trim() !== "—" && (
                                <span> · {ie.Feature}</span>
                              )}
                            </div>
                            <div className="text-sm text-gray-200 truncate">{ie.Title || "Untitled"}</div>
                          </div>
                        </button>
                      );
                    })}
                  </div>
                </div>
              );
            })}
        </div>
        <div className="px-5 py-3 border-t border-gray-800 flex items-center justify-between gap-2 flex-shrink-0">
          {r ? (
            <button onClick={T} className="text-xs text-gray-400 hover:text-gray-200">
              Clear focus filter
            </button>
          ) : (
            <span />
          )}
          <div className="flex gap-2">
            <button
              onClick={t}
              className="px-3 py-1.5 text-sm text-gray-300 hover:text-gray-100 hover:bg-gray-800 rounded-sm transition-colors"
            >
              Cancel
            </button>
            <button
              onClick={A}
              disabled={s.size === 0}
              className={`px-4 py-1.5 text-sm rounded-sm font-medium transition-colors ${S.solid} disabled:bg-gray-700 disabled:text-gray-500 disabled:cursor-not-allowed`}
            >
              {s.size === a.length ? "Show all" : `Apply (${s.size})`}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
function InfoField({ label: e, value: t, highlight: a }) {
  return (
    <div>
      <div className="text-xs font-semibold uppercase tracking-wider text-gray-400 mb-1">{e}</div>
      <div
        className={`text-sm whitespace-pre-wrap ${a ? "bg-amber-950/30 border-l-4 border-amber-500 px-3 py-2 rounded text-amber-100" : "text-gray-200"}`}
      >
        {normalizePath(t)}
      </div>
    </div>
  );
}
function App() {
  if (EDITION !== "Pro") {
    return <LightApp />;
  }
  let [e, t] = useState(() => readRoute());
  useEffect(() => {
    let r = () => t(readRoute());
    window.addEventListener("popstate", r);
    return () => window.removeEventListener("popstate", r);
  }, []);
  let a;
  if (e.role === "tester") {
    a = <TesterEntry sessionId={e.sessionId} />;
  } else if (e.role === "dev") {
    a = <DevEntry sessionId={e.sessionId} />;
  } else if (e.role === "picker") {
    a = <PortalPicker sessionId={e.sessionId} />;
  } else {
    a = (
      <AdminGate>
        <AdminPortal />
      </AdminGate>
    );
  }
  return (
    <Fragment>
      {a}
      <GlobalTooltip />
    </Fragment>
  );
}
function GlobalTooltip() {
  let [e, t] = useState(null);
  let a = useRef(null);
  let r = useRef({
    el: null,
    saved: null,
  });
  useEffect(() => {
    if (typeof window === "undefined" || !window.matchMedia("(hover: hover)").matches) {
      return;
    }
    let o = () => {
      let { el: u, saved: d } = r.current;
      if (u && d !== null) {
        try {
          u.setAttribute("title", d);
        } catch {}
        try {
          u.removeAttribute("data-tt-saved");
        } catch {}
      }
      r.current = {
        el: null,
        saved: null,
      };
    };
    let s = (u) => {
      let d = u.target;
      while (
        d &&
        d !== document.body &&
        (!d.getAttribute || d.getAttribute("title") == null || d.getAttribute("title") === "")
      ) {
        d = d.parentElement;
      }
      if (!d || d === document.body || d === r.current.el) {
        return;
      }
      let c = d.tagName;
      if (c === "INPUT" || c === "TEXTAREA" || c === "SELECT") {
        return;
      }
      o();
      let p = d.getAttribute("title");
      r.current = {
        el: d,
        saved: p,
      };
      d.setAttribute("data-tt-saved", p);
      d.removeAttribute("title");
      let h = d.getBoundingClientRect();
      t({
        text: p,
        anchorTop: h.top,
        anchorBottom: h.bottom,
        anchorCx: h.left + h.width / 2,
      });
    };
    let l = (u) => {
      let { el: d } = r.current;
      if (!d) {
        return;
      }
      let c = u.relatedTarget;
      if (!c || !d.contains || !d.contains(c)) {
        o();
        t(null);
      }
    };
    let i = () => {
      o();
      t(null);
    };
    document.addEventListener("mouseover", s);
    document.addEventListener("mouseout", l);
    window.addEventListener("scroll", i, true);
    window.addEventListener("blur", i);
    window.addEventListener("resize", i);
    document.addEventListener("keydown", (u) => {
      if (u.key === "Escape") {
        i();
      }
    });
    return () => {
      document.removeEventListener("mouseover", s);
      document.removeEventListener("mouseout", l);
      window.removeEventListener("scroll", i, true);
      window.removeEventListener("blur", i);
      window.removeEventListener("resize", i);
      o();
    };
  }, []);
  useLayoutEffect(() => {
    if (!e || !a.current) {
      return;
    }
    let o = a.current;
    let s = o.offsetWidth;
    let l = o.offsetHeight;
    let i = window.innerWidth;
    let u = window.innerHeight;
    let d = 6;
    let c = e.anchorBottom + d;
    if (c + l > u - 4) {
      c = Math.max(4, e.anchorTop - l - d);
    }
    let p = e.anchorCx - s / 2;
    p = Math.max(4, Math.min(p, i - s - 4));
    o.style.top = `${c}px`;
    o.style.left = `${p}px`;
    o.style.opacity = "1";
  }, [e]);
  if (e) {
    return (
      <div
        ref={a}
        role="tooltip"
        style={{
          position: "fixed",
          top: 0,
          left: 0,
          opacity: 0,
          zIndex: 9999,
          pointerEvents: "none",
        }}
        className="bg-gray-950 border border-gray-700 text-gray-100 text-xs px-2 py-1 rounded-sm shadow-lg max-w-xs whitespace-pre-line"
      >
        {e.text}
      </div>
    );
  } else {
    return null;
  }
}
function readRoute() {
  let e = new URLSearchParams(window.location.search);
  let t = e.get("role");
  let a = e.get("session") || null;
  if (a && (t === "tester" || t === "dev")) {
    return {
      role: t,
      sessionId: a,
    };
  } else if (a) {
    return {
      role: "picker",
      sessionId: a,
    };
  } else {
    return {
      role: "admin",
      sessionId: null,
    };
  }
}
var randomCode = () => String(Math.floor(1000 + Math.random() * 9000));
var portalUrl = (e, t) => `${window.location.origin + window.location.pathname}?role=${e}&session=${t}`;
var sessionUrl = (e) => `${window.location.origin + window.location.pathname}?session=${e}`;
var navigateTo = (e) => {
  window.location.assign(e);
};
// ---------------------------------------------------------------------------
// Admin lock: a 4-digit code stored in Supabase (table app_settings, key "admin_code").
// The code can't be read through the app; it's only checked by the check_admin_code()
// database function. Change it in Supabase > Table Editor > app_settings.
// ---------------------------------------------------------------------------
var ADMIN_CODE_KEY = "uat-pro-admin-code";
var AdminLockContext = createContext(null);
var checkAdminCode = async (e) => {
  let { data: t, error: a } = await supabase.rpc("check_admin_code", { code: e });
  if (a) {
    let r = new Error(a.message || "Check failed");
    r.setup = /check_admin_code|PGRST202|does not exist|schema cache/i.test(`${a.code || ""} ${a.message || ""}`);
    throw r;
  }
  return t === true;
};
function AdminGate({ children: e }) {
  let [t, a] = useState("checking");
  let [r, o] = useState("");
  let [s, l] = useState("");
  let [i, u] = useState(false);
  let d = useRef(null);
  let c = () => {
    try {
      localStorage.removeItem(ADMIN_CODE_KEY);
    } catch {}
    o("");
    l("");
    a("locked");
  };
  useEffect(() => {
    let h = false;
    (async () => {
      let b = null;
      try {
        b = localStorage.getItem(ADMIN_CODE_KEY);
      } catch {}
      if (!b) {
        a("locked");
        return;
      }
      try {
        let _ = await checkAdminCode(b);
        if (!h) a(_ ? "open" : "locked");
        if (!_) c();
      } catch (_) {
        if (!h) a(_.setup ? "setup" : "locked");
      }
    })();
    return () => {
      h = true;
    };
  }, []);
  useEffect(() => {
    if (t === "locked") setTimeout(() => d.current?.focus(), 50);
  }, [t]);
  let p = async (h) => {
    if (i || h.length !== 4) return;
    u(true);
    l("");
    try {
      if (await checkAdminCode(h)) {
        try {
          localStorage.setItem(ADMIN_CODE_KEY, h);
        } catch {}
        a("open");
      } else {
        l("That code isn't right.");
        o("");
        setTimeout(() => d.current?.focus(), 30);
      }
    } catch (b) {
      if (b.setup) a("setup");
      else l(b.message || "Couldn't check the code. Try again.");
    }
    u(false);
  };
  if (t === "open") {
    return <AdminLockContext.Provider value={{ lock: c }}>{e}</AdminLockContext.Provider>;
  }
  return (
    <div className="min-h-screen bg-gradient-to-br from-gray-950 via-gray-900 to-indigo-950 flex items-center justify-center p-6">
      <div className="w-full max-w-xs text-center">
        <div className="w-11 h-11 bg-emerald-600 rounded-sm flex items-center justify-center mx-auto mb-4">
          <FileText className="w-6 h-6 text-white" />
        </div>
        <div className="font-bold text-gray-100 text-lg">Test Case Collaborator</div>
        <div className="text-xs text-gray-500 mb-6">Admin</div>
        {t === "checking" ? (
          <div className="text-sm text-gray-400">Checking…</div>
        ) : t === "setup" ? (
          <div className="text-left p-3 bg-amber-950/40 border border-amber-700 rounded-sm text-xs text-amber-100 leading-relaxed">
            The admin code isn't set up in Supabase yet. Run the admin-code SQL from supabase-schema.sql in the
            Supabase SQL Editor, then reload this page.
          </div>
        ) : (
          <form
            onSubmit={(h) => {
              h.preventDefault();
              p(r);
            }}
          >
            <label htmlFor="admin-code" className="block text-sm text-gray-300 mb-3">
              Enter the 4-digit admin code
            </label>
            <input
              id="admin-code"
              ref={d}
              value={r}
              onChange={(h) => {
                let b = h.target.value.replace(/\D/g, "").slice(0, 4);
                o(b);
                l("");
                if (b.length === 4) p(b);
              }}
              inputMode="numeric"
              autoComplete="off"
              pattern="[0-9]*"
              maxLength={4}
              disabled={i}
              aria-invalid={!!s}
              className={`w-40 mx-auto block text-center font-mono text-2xl tracking-[0.6em] pl-[0.6em] py-2.5 bg-gray-900 border text-gray-100 rounded-sm focus:outline-none focus:ring-2 focus:ring-emerald-500 ${s ? "border-red-600" : "border-gray-700"}`}
              placeholder="••••"
            />
            <div className="h-5 mt-2 text-xs">
              {i ? (
                <span className="text-gray-400">Checking…</span>
              ) : s ? (
                <span className="text-red-300">{s}</span>
              ) : null}
            </div>
          </form>
        )}
      </div>
    </div>
  );
}
function AdminLockButton() {
  let e = useContext(AdminLockContext);
  if (!e) return null;
  return (
    <button
      onClick={e.lock}
      className="flex items-center gap-1.5 px-3 py-2 text-sm text-gray-400 border border-gray-700 rounded-sm hover:text-gray-100 hover:bg-gray-800"
      title="Lock the admin page on this device"
    >
      <Lock className="w-4 h-4" /> Lock
    </button>
  );
}
function AdminPortal() {
  let [e, t] = useState("list");
  let [a, r] = useState([]);
  let [o, s] = useState({});
  let [l, i] = useState(true);
  let [u, d] = useState("");
  let [c, p] = useState(null);
  let [h, b] = useState(null);
  let [_, S] = useState(0);
  useEffect(() => {
    if (e !== "list") {
      return;
    }
    let I = false;
    i(true);
    d("");
    (async () => {
      let { data: x, error: g } = await supabase
        .from("sessions")
        .select("id, name, created_at, testers, developers, tester_code, developer_code")
        .order("created_at", {
          ascending: false,
        });
      if (!I) {
        if (g) {
          d(g.message);
        } else {
          r(x || []);
        }
        i(false);
      }
    })();
    return () => {
      I = true;
    };
  }, [e, _]);
  useEffect(() => {
    if (e !== "list" || a.length === 0) {
      return;
    }
    let I = false;
    (async () => {
      let x = a.map((w) => w.id);
      let [g, v] = await Promise.all([
        supabase.from("test_cases").select("id, session_id, removed_at").in("session_id", x),
        supabase.from("findings").select("session_id, status, test_case_id").in("session_id", x),
      ]);
      if (g.error) {
        // Database without the removed_at column yet.
        g = await supabase.from("test_cases").select("id, session_id").in("session_id", x);
      }
      if (I) {
        return;
      }
      let k = {};
      let rm = new Set();
      (g.data || []).forEach((w) => {
        if (w.removed_at) {
          rm.add(w.id);
          return;
        }
        k[w.session_id] = (k[w.session_id] || 0) + 1;
      });
      let C = {};
      let m = {};
      (v.data || []).forEach((w) => {
        if (rm.has(w.test_case_id)) {
          return;
        }
        let T = w.status || "Not Tested";
        if (T !== "Not Tested") {
          C[w.session_id] = (C[w.session_id] || 0) + 1;
        }
        if (isFailureStatus(T)) {
          m[w.session_id] = (m[w.session_id] || 0) + 1;
        }
      });
      let f = {};
      x.forEach((w) => {
        f[w] = {
          total: k[w] || 0,
          tested: C[w] || 0,
          failed: m[w] || 0,
          loaded: true,
        };
      });
      s(f);
    })();
    return () => {
      I = true;
    };
  }, [e, a]);
  if (e === "create") {
    return (
      <CreateSessionWizard
        onCancel={() => t("list")}
        onCreated={(I) => {
          p(I);
          t("created");
        }}
      />
    );
  } else if (e === "created" && c) {
    return <SessionCreatedScreen session={c} onDone={() => t("list")} />;
  } else {
    return (
      <div className="min-h-screen bg-gradient-to-br from-gray-950 via-gray-900 to-indigo-950 flex flex-col">
        <header className="px-8 py-5 border-b border-gray-800 bg-gray-950/60 backdrop-blur">
          <div className="max-w-5xl mx-auto flex items-center justify-between">
            <div className="flex items-center gap-3">
              <div className="w-9 h-9 bg-emerald-600 rounded-sm flex items-center justify-center">
                <FileText className="w-5 h-5 text-white" />
              </div>
              <div>
                <div className="font-bold text-gray-100 leading-tight">Test Case Collaborator</div>
                <div className="text-[11px] text-gray-500 leading-tight">Admin</div>
              </div>
            </div>
            <div className="flex items-center gap-2">
              <AdminLockButton />
              <button
                onClick={() => t("create")}
                className="flex items-center gap-2 px-3 py-2 bg-emerald-600 text-white text-sm font-medium rounded-sm hover:bg-emerald-500"
              >
                <Plus className="w-4 h-4" /> New session
              </button>
            </div>
          </div>
        </header>
        <main className="flex-1 max-w-5xl w-full mx-auto px-8 py-10">
          <h2 className="text-xl font-bold text-gray-100 mb-1">Sessions</h2>
          <p className="text-sm text-gray-400 mb-6">
            Each session is one UAT campaign with its own test cases, testers, and developers.
          </p>
          {l ? (
            <div className="py-12 text-center text-gray-500 text-sm">Loading…</div>
          ) : u ? (
            <div className="p-4 bg-red-950/40 border border-red-800 rounded-sm text-sm text-red-300">
              <strong>Could not load sessions:</strong> {u}
              <div className="mt-2 text-xs text-red-300/80">
                If this is the first time running the app, you may need to run the schema SQL in Supabase
                first.
              </div>
            </div>
          ) : a.length === 0 ? (
            <div className="py-16 text-center border border-dashed border-gray-800 rounded-sm">
              <div className="text-gray-400 text-sm mb-4">No sessions yet.</div>
              <button
                onClick={() => t("create")}
                className="inline-flex items-center gap-2 px-4 py-2 bg-emerald-600 text-white text-sm font-medium rounded-sm hover:bg-emerald-500"
              >
                <Plus className="w-4 h-4" /> Create your first session
              </button>
            </div>
          ) : (
            <div className="space-y-2">
              {a.map((I) => (
                <SessionCard
                  session={I}
                  progress={o[I.id]}
                  onShowUrls={() => {
                    p(I);
                    t("created");
                  }}
                  onEdit={() => b(I)}
                  key={I.id}
                />
              ))}
            </div>
          )}
        </main>
        {h && (
          <EditSessionModal
            session={h}
            onClose={() => b(null)}
            onSaved={() => {
              b(null);
              S((I) => I + 1);
            }}
          />
        )}
      </div>
    );
  }
}
function SessionCard({ session: e, progress: t, onShowUrls: a, onEdit: r }) {
  let o = e.created_at ? new Date(e.created_at).toLocaleString() : "";
  let s = e.created_at ? timeAgo(e.created_at) : "";
  let l = (e.testers || []).length;
  let i = (e.developers || []).length;
  let u = t || {
    total: 0,
    tested: 0,
    failed: 0,
    loaded: false,
  };
  return (
    <div className="flex items-center gap-4 p-4 bg-gray-900/60 border border-gray-800 rounded-sm hover:border-gray-700">
      <div className="flex-1 min-w-0">
        <div className="font-medium text-gray-100 truncate">{e.name}</div>
        <div className="text-xs text-gray-500 mt-0.5 flex flex-wrap items-center gap-x-2 gap-y-0.5">
          <span title={o}>{s}</span>
          <span className="text-gray-700">·</span>
          <span>
            {l} tester{l === 1 ? "" : "s"}
          </span>
          <span className="text-gray-700">·</span>
          <span>
            {i} dev{i === 1 ? "" : "s"}
          </span>
          {u.loaded && u.total > 0 && (
            <Fragment>
              <span className="text-gray-700">·</span>
              <span>
                <span className="text-gray-300">{u.tested}</span>/{u.total} tested
              </span>
              {u.failed > 0 && (
                <Fragment>
                  <span className="text-gray-700">·</span>
                  <span className="text-red-400">{u.failed} failed</span>
                </Fragment>
              )}
            </Fragment>
          )}
        </div>
      </div>
      <button
        onClick={r}
        className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-gray-300 hover:text-gray-100 bg-gray-800/60 hover:bg-gray-800 border border-gray-700 rounded-sm"
        title="Edit session name and add testers/developers"
      >
        <Pencil className="w-3.5 h-3.5" /> Edit
      </button>
      <button
        onClick={a}
        className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-emerald-300 hover:text-emerald-200 bg-emerald-950/40 hover:bg-emerald-900/40 border border-emerald-900 rounded-sm"
      >
        <ExternalLink className="w-3.5 h-3.5" /> Share link
      </button>
    </div>
  );
}
// Import from a Google Sheets link. Only works for sheets shared as "Anyone with the link can view".
// The download goes through the "fetch-sheet" Supabase Edge Function (Google blocks direct browser downloads).
var looksLikeSheetLink = (e) => /docs\.google\.com\/spreadsheets\/(?:u\/\d+\/)?d\/[a-zA-Z0-9_-]{20,}/.test(String(e || ""));
var fetchSheetFromLink = async (e) => {
  let t;
  try {
    t = await fetch(`${SUPABASE_URL}/functions/v1/fetch-sheet`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        apikey: SUPABASE_KEY,
      },
      body: JSON.stringify({
        url: e,
      }),
    });
  } catch (r) {
    throw new Error("Couldn't reach the sheet importer. Check your connection, or that the fetch-sheet function is set up in Supabase.");
  }
  if (!t.ok) {
    let r = await t.json().catch(() => ({}));
    let o = r.error || r.message || r.msg || "";
    if (/function was not found|not_found/i.test(`${r.code || ""} ${o}`) && !r.error) {
      throw new Error("Link import isn't set up yet: the fetch-sheet function hasn't been added in Supabase.");
    }
    if (t.status === 401 && !r.error) {
      throw new Error("Link import is blocked by a setting: turn off \"Verify JWT\" on the fetch-sheet function in Supabase.");
    }
    throw new Error(o || `Import failed (${t.status}).`);
  }
  let a = await t.arrayBuffer();
  let r = "Google Sheet.xlsx";
  try {
    r = decodeURIComponent(t.headers.get("X-File-Name") || r);
  } catch (o) {}
  if (!/\.xlsx$/i.test(r)) r += ".xlsx";
  return {
    name: r,
    arrayBuffer: async () => a,
  };
};
function SheetLinkImport({ onFile: e, disabled: t }) {
  let [a, r] = useState("");
  let [o, s] = useState(false);
  let [l, i] = useState("");
  let u = async () => {
    let d = a.trim();
    if (!d || o) return;
    if (!looksLikeSheetLink(d)) {
      i("Paste a link that starts with https://docs.google.com/spreadsheets/…");
      return;
    }
    s(true);
    i("");
    try {
      let c = await fetchSheetFromLink(d);
      await e(c);
      r("");
    } catch (c) {
      i(c.message || "Import failed");
    }
    s(false);
  };
  return (
    <div className="mt-2">
      <div className="flex items-center gap-2 text-[11px] text-gray-500 mb-1.5">
        <span className="h-px flex-1 bg-gray-800" />
        or import from a Google Sheets link
        <span className="h-px flex-1 bg-gray-800" />
      </div>
      <div className="flex gap-2">
        <input
          type="url"
          value={a}
          onChange={(d) => {
            r(d.target.value);
            i("");
          }}
          onKeyDown={(d) => {
            if (d.key === "Enter") {
              d.preventDefault();
              u();
            }
          }}
          disabled={t || o}
          placeholder="https://docs.google.com/spreadsheets/d/…"
          autoCapitalize="off"
          autoComplete="off"
          spellCheck="false"
          className="flex-1 min-w-0 px-3 py-1.5 bg-gray-800 border border-gray-700 text-gray-100 placeholder-gray-500 rounded-sm text-base sm:text-sm focus:outline-none focus:border-emerald-500"
        />
        <button
          onClick={u}
          disabled={t || o || !a.trim()}
          className="flex-shrink-0 px-3 py-1.5 text-xs font-medium bg-emerald-600 text-white rounded-sm hover:bg-emerald-500 disabled:bg-gray-800 disabled:text-gray-500"
        >
          {o ? "Fetching…" : "Import"}
        </button>
      </div>
      {l ? (
        <div className="text-[11px] text-red-300 mt-1.5">{l}</div>
      ) : (
        <div className="text-[11px] text-gray-500 mt-1.5">
          Only for sheets shared as "Anyone with the link can view". Private sheets: download as .xlsx and drop it
          above.
        </div>
      )}
    </div>
  );
}
// ---------------------------------------------------------------------------
// Sample files: admins attach files (e.g. a zip) to a session; testers download them.
// Files live in the public "session-files" storage bucket under <session id>/,
// and the list is kept in sessions.sample_files.
// ---------------------------------------------------------------------------
var SAMPLE_BUCKET = "session-files";
var SAMPLE_MAX_BYTES = 50 * 1024 * 1024;
var SAMPLE_SETUP_SQL = `alter table sessions add column if not exists sample_files jsonb default '[]'::jsonb;

insert into storage.buckets (id, name, public, file_size_limit)
values ('session-files', 'session-files', true, 52428800)
on conflict (id) do nothing;

do $$
begin
  if not exists (
    select 1 from pg_policies
    where schemaname = 'storage' and tablename = 'objects'
      and policyname = 'session-files anon insert'
  ) then
    create policy "session-files anon insert"
    on storage.objects for insert
    to anon
    with check (bucket_id = 'session-files');
  end if;
  if not exists (
    select 1 from pg_policies
    where schemaname = 'storage' and tablename = 'objects'
      and policyname = 'session-files anon delete'
  ) then
    create policy "session-files anon delete"
    on storage.objects for delete
    to anon
    using (bucket_id = 'session-files');
  end if;
end $$;`;
var formatBytes = (e) => {
  if (!e && e !== 0) return "";
  if (e < 1024) return `${e} B`;
  if (e < 1048576) return `${Math.round(e / 1024)} KB`;
  return `${(e / 1048576).toFixed(e < 10485760 ? 1 : 0)} MB`;
};
var isSampleSetupError = (e) =>
  /bucket not found|sample_files|row-level security|violates row-level/i.test(String(e?.message || e || ""));
var SAMPLE_SETUP_MESSAGE =
  "Sample files need a one-time Supabase setup. Run the SQL from the setup note (Edit session → Sample files), then try again.";
var uploadSampleFiles = async (e, t) => {
  let a = [];
  for (let r of t) {
    if (r.size > SAMPLE_MAX_BYTES) {
      throw new Error(`${r.name} is ${formatBytes(r.size)}. Files must be under 50 MB. Split the zip or leave out the largest files.`);
    }
    let o = r.name.replace(/[^a-zA-Z0-9._-]+/g, "-").replace(/-+/g, "-");
    let s = `${e}/${Date.now()}-${o}`;
    let { error: l } = await supabase.storage.from(SAMPLE_BUCKET).upload(s, r, {
      contentType: r.type || "application/octet-stream",
      upsert: false,
    });
    if (l) {
      throw isSampleSetupError(l) ? new Error(SAMPLE_SETUP_MESSAGE) : l;
    }
    let {
      data: { publicUrl: i },
    } = supabase.storage.from(SAMPLE_BUCKET).getPublicUrl(s, {
      download: r.name,
    });
    a.push({
      path: s,
      name: r.name,
      size: r.size,
      url: i,
      uploaded_at: new Date().toISOString(),
    });
  }
  return a;
};
function SampleFileRow({ file: e, onRemove: t, removing: a }) {
  return (
    <div className="flex items-center gap-2 px-2.5 py-1.5 bg-gray-800/60 border border-gray-700 rounded-sm text-xs">
      <Paperclip className="w-3.5 h-3.5 text-gray-500 flex-shrink-0" />
      <span className="flex-1 min-w-0 truncate text-gray-200" title={e.name}>
        {e.name}
      </span>
      <span className="text-gray-500 flex-shrink-0">{formatBytes(e.size)}</span>
      {e.url && (
        <a href={e.url} className="text-gray-400 hover:text-gray-100 flex-shrink-0" title="Download">
          <Download className="w-3.5 h-3.5" />
        </a>
      )}
      {t && (
        <button
          onClick={t}
          disabled={a}
          className="text-gray-500 hover:text-red-300 disabled:opacity-40 flex-shrink-0"
          title="Remove"
        >
          <X className="w-3.5 h-3.5" />
        </button>
      )}
    </div>
  );
}
// Wizard: files are only uploaded when the session is created.
function SampleFilesPicker({ files: e, setFiles: t }) {
  let a = useRef(null);
  let [r, o] = useState("");
  let s = (l) => {
    let i = Array.from(l || []);
    let u = i.filter((d) => d.size > SAMPLE_MAX_BYTES);
    o(u.length ? `${u.map((d) => d.name).join(", ")} ${u.length === 1 ? "is" : "are"} over 50 MB and can't be added.` : "");
    t([...e, ...i.filter((d) => d.size <= SAMPLE_MAX_BYTES)]);
  };
  return (
    <div>
      <div className="flex items-baseline justify-between mb-1.5">
        <label className="block text-xs font-semibold uppercase tracking-wider text-gray-400">
          Sample files for testers <span className="normal-case font-normal text-gray-500">(optional)</span>
        </label>
      </div>
      {e.length > 0 && (
        <div className="space-y-1 mb-2">
          {e.map((l, i) => (
            <SampleFileRow file={l} key={i} onRemove={() => t(e.filter((u, d) => d !== i))} />
          ))}
        </div>
      )}
      <button
        type="button"
        onClick={() => a.current?.click()}
        className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-gray-200 bg-gray-800 hover:bg-gray-700 border border-gray-700 rounded-sm transition-colors"
      >
        <Paperclip className="w-3.5 h-3.5" /> {e.length ? "Add more files" : "Attach files"}
      </button>
      <input
        ref={a}
        type="file"
        multiple
        onChange={(l) => {
          s(l.target.files);
          l.target.value = "";
        }}
        className="hidden"
      />
      <div className="text-[11px] text-gray-500 mt-1.5">
        Files testers need, e.g. a zipped sample-files folder. Up to 50 MB each. Testers get a download prompt when
        they open the session.
      </div>
      {r && <div className="text-[11px] text-red-300 mt-1">{r}</div>}
    </div>
  );
}
// Edit session: files upload/remove immediately.
function SampleFilesManager({ session: e }) {
  let [t, a] = useState(null);
  let [r, o] = useState(false);
  let [s, l] = useState("");
  let [i, u] = useState(null);
  let [d, c] = useState(false);
  let p = useRef(null);
  useEffect(() => {
    let b = false;
    (async () => {
      let { data: _, error: S } = await supabase.from("sessions").select("sample_files").eq("id", e.id).single();
      if (b) return;
      if (S) {
        a([]);
        if (isSampleSetupError(S)) c(true);
      } else {
        a(_?.sample_files || []);
      }
    })();
    return () => {
      b = true;
    };
  }, [e.id]);
  let h = async (b) => {
    let { error: _ } = await supabase.from("sessions").update({ sample_files: b }).eq("id", e.id);
    if (_) throw isSampleSetupError(_) ? new Error(SAMPLE_SETUP_MESSAGE) : _;
    a(b);
  };
  let b = async (_) => {
    let S = Array.from(_ || []);
    if (!S.length) return;
    o(true);
    l("");
    try {
      let I = await uploadSampleFiles(e.id, S);
      await h([...(t || []), ...I]);
    } catch (I) {
      l(I.message || "Upload failed");
      if (/one-time Supabase setup/.test(I.message || "")) c(true);
    }
    o(false);
  };
  let _ = async (S) => {
    u(S.path);
    l("");
    try {
      await h((t || []).filter((I) => I.path !== S.path));
      await supabase.storage.from(SAMPLE_BUCKET).remove([S.path]);
    } catch (I) {
      l(I.message || "Couldn't remove the file");
    }
    u(null);
  };
  return (
    <div>
      <div className="text-xs font-semibold uppercase tracking-wider text-gray-400 mb-1.5">
        Sample files for testers{" "}
        {t && t.length > 0 && <span className="normal-case font-normal text-gray-500">({t.length})</span>}
      </div>
      {d && (
        <div className="mb-2 p-3 bg-amber-950/40 border border-amber-700 rounded-sm text-xs text-amber-100 space-y-2">
          <div className="font-semibold">One-time Supabase setup needed</div>
          <div>In Supabase, open the SQL Editor, paste this and click Run. Then close and reopen this window.</div>
          <pre className="bg-gray-950 border border-gray-800 rounded-sm p-2 font-mono text-[10px] text-gray-200 overflow-x-auto whitespace-pre">
            {SAMPLE_SETUP_SQL}
          </pre>
        </div>
      )}
      {t === null ? (
        <div className="text-xs text-gray-500">Loading…</div>
      ) : (
        <>
          {t.length > 0 && (
            <div className="space-y-1 mb-2">
              {t.map((S) => (
                <SampleFileRow file={S} key={S.path} onRemove={() => _(S)} removing={i === S.path} />
              ))}
            </div>
          )}
          <button
            type="button"
            onClick={() => p.current?.click()}
            disabled={r}
            className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-gray-200 bg-gray-800 hover:bg-gray-700 border border-gray-700 rounded-sm transition-colors disabled:opacity-60"
          >
            <Paperclip className="w-3.5 h-3.5" /> {r ? "Uploading…" : t.length ? "Add more files" : "Attach files"}
          </button>
          <input
            ref={p}
            type="file"
            multiple
            onChange={(S) => {
              b(S.target.files);
              S.target.value = "";
            }}
            className="hidden"
          />
          <div className="text-[11px] text-gray-500 mt-1.5">
            Saved straight away. Up to 50 MB each, e.g. a zipped sample-files folder.
          </div>
        </>
      )}
      {s && <div className="text-[11px] text-red-300 mt-1">{s}</div>}
    </div>
  );
}
var linkifyText = (e) =>
  String(e)
    .split(/(https?:\/\/[^\s<>"')]+)/g)
    .map((t, a) =>
      a % 2 === 1 ? (
        <a
          key={a}
          href={t}
          target="_blank"
          rel="noopener noreferrer"
          className="text-sky-300 underline decoration-sky-700 underline-offset-2 break-all hover:text-sky-200"
        >
          {t}
        </a>
      ) : (
        <Fragment key={a}>{t}</Fragment>
      ),
    );
function ReadmeContent({ text: e }) {
  let t = String(e || "").split("\n");
  let a = [];
  let r = [];
  let o = () => {
    if (r.length) {
      a.push(
        <dl className="divide-y divide-gray-800 border border-gray-800 rounded-sm" key={`dl${a.length}`}>
          {r.map((s, l) => (
            <div className="px-3 py-2" key={l}>
              <dt className="text-xs font-semibold text-gray-100">{linkifyText(s[0])}</dt>
              <dd className="text-xs text-gray-400 mt-0.5 leading-relaxed">{linkifyText(s[1])}</dd>
            </div>
          ))}
        </dl>,
      );
      r = [];
    }
  };
  t.forEach((s, l) => {
    let i = s.trim();
    let u = i.indexOf(" :: ");
    if (u > 0) {
      r.push([i.slice(0, u), i.slice(u + 4)]);
      return;
    }
    o();
    if (!i) return;
    if (i.startsWith("## ")) {
      a.push(
        <h3 className="text-xs font-semibold uppercase tracking-wider text-amber-300 pt-3" key={l}>
          {i.slice(3)}
        </h3>,
      );
    } else if (i.startsWith("# ")) {
      a.push(
        <h2 className="text-base font-bold text-gray-100" key={l}>
          {i.slice(2)}
        </h2>,
      );
    } else {
      a.push(
        <p className="text-sm text-gray-300 leading-relaxed" key={l}>
          {linkifyText(i)}
        </p>,
      );
    }
  });
  o();
  return <div className="space-y-2.5">{a}</div>;
}
// Tester portal: full-height side panel prompting testers to download the session's sample files.
var ALREADY_COMPRESSED = /\.(zip|png|jpe?g|gif|webp|pdf|xlsx|xls|docx|pptx|mp4|mov|gz|7z|rar)$/i;
var downloadBlob = (e, t) => {
  let a = URL.createObjectURL(e);
  let r = document.createElement("a");
  r.href = a;
  r.download = t;
  document.body.appendChild(r);
  r.click();
  r.remove();
  setTimeout(() => URL.revokeObjectURL(a), 4000);
};
function TesterSampleFilesBar({ sessionId: e, sessionName: t, myName: n }) {
  let [a, r] = useState([]);
  let o = `uat-pro-samples-ack-${e}-${n}`;
  let [s, l] = useState(() => {
    try {
      return localStorage.getItem(o) === "1";
    } catch {
      return false;
    }
  });
  let [i, u] = useState(false);
  let [d, c] = useState(false);
  let [z, setZip] = useState(null);
  let [zErr, setZErr] = useState("");
  let [readme, setReadme] = useState("");
  let [readmeOpen, setReadmeOpen] = useState(false);
  useEffect(() => {
    let h = false;
    (async () => {
      let { data: b, error: _ } = await supabase.from("sessions").select("readme").eq("id", e).single();
      if (!h && !_ && b?.readme && b.readme.trim()) setReadme(b.readme);
    })();
    return () => {
      h = true;
    };
  }, [e]);
  useEffect(() => {
    if (!readmeOpen) return;
    let h = (b) => {
      if (b.key === "Escape") setReadmeOpen(false);
    };
    window.addEventListener("keydown", h);
    return () => window.removeEventListener("keydown", h);
  }, [readmeOpen]);
  useEffect(() => {
    let h = false;
    (async () => {
      let { data: b, error: _ } = await supabase.from("sessions").select("sample_files").eq("id", e).single();
      if (h || _) return;
      r(b?.sample_files || []);
      if (!s) {
        requestAnimationFrame(() => u(true));
      }
      c(true);
    })();
    return () => {
      h = true;
    };
  }, [e]);
  useEffect(() => {
    if (!i) return;
    let h = (b) => {
      if (b.key === "Escape") u(false);
    };
    window.addEventListener("keydown", h);
    return () => window.removeEventListener("keydown", h);
  }, [i]);
  let hasFiles = d && a.length > 0;
  if (!hasFiles && !readme) return null;
  let p = () => {
    l(true);
    try {
      localStorage.setItem(o, "1");
    } catch {}
  };
  let total = a.reduce((h, b) => h + (b.size || 0), 0);
  let zipAll = async () => {
    if (z !== null) return;
    setZErr("");
    setZip(0);
    try {
      let h = new JSZip();
      let b = new Set();
      for (let _ = 0; _ < a.length; _++) {
        let S = a[_];
        let I = await fetch(S.url);
        if (!I.ok) throw new Error(`Couldn't download ${S.name}`);
        let x = await I.blob();
        let g = S.name;
        for (let v = 2; b.has(g); v++) g = S.name.replace(/(\.[^.]*)?$/, ` (${v})$1`);
        b.add(g);
        h.file(g, x, { compression: ALREADY_COMPRESSED.test(g) ? "STORE" : "DEFLATE" });
        setZip(Math.round(((_ + 1) / a.length) * 80));
      }
      let _ = await h.generateAsync({ type: "blob" }, (S) => setZip(80 + Math.round(S.percent * 0.2)));
      downloadBlob(_, `${(t || "Sample files").replace(/[\\/:*?"<>|]+/g, "-").trim()} - sample files.zip`);
      p();
    } catch (h) {
      setZErr(h.message || "Couldn't build the zip. Download the files one by one instead.");
    }
    setZip(null);
  };
  let close = () => {
    u(false);
    p();
  };
  return (
    <>
      {readme && (
        <Fragment>
          <div
            className={`fixed inset-0 z-30 bg-black/40 sm:hidden transition-opacity duration-300 ${readmeOpen ? "opacity-100" : "opacity-0 pointer-events-none"}`}
            onClick={() => setReadmeOpen(false)}
            aria-hidden="true"
          />
          <aside
            className={`fixed z-40 top-0 bottom-0 right-0 w-full sm:w-[28rem] flex flex-col bg-gray-900 border-l border-gray-700 shadow-2xl transition-transform duration-300 ease-out ${readmeOpen ? "translate-x-0" : "translate-x-full pointer-events-none"}`}
            role="dialog"
            aria-label="Read me"
            aria-hidden={!readmeOpen}
          >
            <div className="flex items-center gap-2.5 px-5 py-4 border-b border-gray-800">
              <BookOpen className="w-4 h-4 text-amber-400 flex-shrink-0" />
              <div className="flex-1 text-base font-semibold text-gray-100">Read me</div>
              <button
                onClick={() => setReadmeOpen(false)}
                className="text-gray-500 hover:text-gray-200 p-0.5"
                title="Close"
              >
                <X className="w-4 h-4" />
              </button>
            </div>
            <div className="flex-1 overflow-y-auto px-5 py-4">
              <ReadmeContent text={readme} />
            </div>
          </aside>
        </Fragment>
      )}
      {hasFiles && (
      <Fragment>
      <div
        className={`fixed inset-0 z-30 bg-black/40 sm:hidden transition-opacity duration-300 ${i ? "opacity-100" : "opacity-0 pointer-events-none"}`}
        onClick={close}
        aria-hidden="true"
      />
      <aside
        className={`fixed z-40 top-0 bottom-0 right-0 w-full sm:w-96 flex flex-col bg-gray-900 border-l border-amber-700/60 shadow-2xl transition-transform duration-300 ease-out ${i ? "translate-x-0" : "translate-x-full pointer-events-none"}`}
        role="dialog"
        aria-label="Sample files"
        aria-hidden={!i}
      >
        <div className="flex items-start gap-2.5 px-5 pt-5 pb-4 border-b border-gray-800">
          <Paperclip className="w-4 h-4 text-amber-400 mt-1 flex-shrink-0" />
          <div className="flex-1 min-w-0">
            <div className="text-base font-semibold text-gray-100">
              {s ? "Sample files" : "Download the sample files first"}
            </div>
            <div className="text-xs text-gray-400 mt-1 leading-relaxed">
              {s
                ? "Files you'll need for these tests."
                : `Some tests need ${a.length === 1 ? "this file" : "these files"}. Download ${a.length === 1 ? "it" : "them"} before you start${a.some((h) => /\.zip$/i.test(h.name)) ? ", then unzip" : ""}.`}
            </div>
          </div>
          <button onClick={close} className="text-gray-500 hover:text-gray-200 p-0.5 flex-shrink-0" title="Close">
            <X className="w-4 h-4" />
          </button>
        </div>
        <div className="flex-1 overflow-y-auto px-5 py-4 space-y-1.5">
          <div className="text-[11px] uppercase tracking-wider text-gray-500 mb-2">
            {a.length} file{a.length === 1 ? "" : "s"} · {formatBytes(total)}
          </div>
          {a.map((h) => (
            <a
              href={h.url}
              onClick={p}
              key={h.path}
              className="flex items-center gap-2.5 px-3 py-2.5 bg-gray-800/70 border border-gray-700 rounded-sm text-sm text-gray-200 transition-colors hover:border-amber-600 hover:bg-gray-800"
            >
              <Download className="w-4 h-4 text-amber-400 flex-shrink-0" />
              <span className="flex-1 min-w-0 truncate" title={h.name}>
                {h.name}
              </span>
              <span className="text-xs text-gray-500 flex-shrink-0">{formatBytes(h.size)}</span>
            </a>
          ))}
        </div>
        <div className="px-5 pt-3 pb-5 border-t border-gray-800 space-y-2.5">
          {a.length > 1 && (
            <button
              onClick={zipAll}
              disabled={z !== null}
              className="relative w-full overflow-hidden flex items-center justify-center gap-2 px-4 py-2.5 text-sm font-medium bg-amber-600 text-gray-950 rounded-sm transition-colors hover:bg-amber-500 disabled:cursor-wait"
            >
              {z !== null && (
                <span className="absolute inset-y-0 left-0 bg-amber-400/60 transition-all" style={{ width: `${z}%` }} />
              )}
              <span className="relative flex items-center gap-2">
                <FileDown className="w-4 h-4" />
                {z !== null ? `Zipping… ${z}%` : `Download all as .zip (${formatBytes(total)})`}
              </span>
            </button>
          )}
          {zErr && <div className="text-xs text-red-300">{zErr}</div>}
          {!s && (
            <button onClick={close} className="w-full text-center text-xs text-gray-400 hover:text-gray-200 py-1">
              I already have them
            </button>
          )}
        </div>
      </aside>
      </Fragment>
      )}
      <div
        className={`fixed z-30 bottom-4 right-4 flex items-center gap-2 transition-all duration-300 ease-out ${i || readmeOpen ? "translate-y-6 opacity-0 pointer-events-none" : "translate-y-0 opacity-100"}`}
      >
        {readme && (
          <button
            onClick={() => {
              u(false);
              setReadmeOpen(true);
            }}
            className="inline-flex items-center gap-1.5 px-3 py-2 text-xs font-medium bg-gray-900 border border-gray-700 text-gray-300 rounded-sm shadow-lg transition-colors hover:text-gray-100 hover:border-gray-600"
            title="Instructions for this test round"
          >
            <BookOpen className="w-3.5 h-3.5" /> Read me
          </button>
        )}
        {hasFiles && (
          <button
            onClick={() => {
              setReadmeOpen(false);
              u(true);
            }}
            className="inline-flex items-center gap-1.5 px-3 py-2 text-xs font-medium bg-gray-900 border border-gray-700 text-gray-300 rounded-sm shadow-lg transition-colors hover:text-gray-100 hover:border-gray-600"
            title="Sample files for these tests"
          >
            <Paperclip className="w-3.5 h-3.5" /> Sample files ({a.length})
          </button>
        )}
      </div>
    </>
  );
}
// Edit session: the tester "Read me" text.
function ReadmeEditor({ session: e }) {
  let [t, a] = useState(null);
  let [r, o] = useState("");
  let [s, l] = useState("");
  let [i, u] = useState(false);
  let [d, c] = useState(false);
  let [p, h] = useState(false);
  let b = useRef(null);
  useEffect(() => {
    let S = false;
    (async () => {
      let { data: I, error: x } = await supabase.from("sessions").select("readme").eq("id", e.id).single();
      if (S) return;
      if (x) {
        if (isReadmeSetupError(x)) c(true);
        a("");
        o("");
      } else {
        a(I?.readme || "");
        o(I?.readme || "");
      }
    })();
    return () => {
      S = true;
    };
  }, [e.id]);
  let _ = async () => {
    u(true);
    l("");
    let { error: S } = await supabase
      .from("sessions")
      .update({ readme: r.trim() ? r : null })
      .eq("id", e.id);
    u(false);
    if (S) {
      if (isReadmeSetupError(S)) c(true);
      l(S.message || "Couldn't save");
    } else {
      a(r);
      l("Saved");
      setTimeout(() => l((I) => (I === "Saved" ? "" : I)), 2000);
    }
  };
  let S = async (I) => {
    let x = Array.from(I || []);
    if (!x.length) return;
    l("");
    let g = await parseImportFiles(x);
    let v = g.find((k) => k.readmeText);
    if (v) {
      o(v.readmeText);
      l(`Loaded from "${v.name}". Check it, then Save.`);
    } else {
      l('No "Start here" / "Read me" style tab found in that file.');
    }
  };
  return (
    <div>
      <div className="flex items-center justify-between mb-1.5">
        <div className="text-xs font-semibold uppercase tracking-wider text-gray-400">Read me for testers</div>
        <div className="flex items-center gap-3 text-[11px]">
          <button type="button" onClick={() => b.current?.click()} className="text-gray-400 hover:text-gray-100">
            Load from sheet…
          </button>
          <button type="button" onClick={() => h(!p)} className="text-gray-400 hover:text-gray-100">
            {p ? "Edit" : "Preview"}
          </button>
        </div>
        <input
          ref={b}
          type="file"
          accept=".xlsx,.xlsm,.xls,.ods"
          onChange={(I) => {
            S(I.target.files);
            I.target.value = "";
          }}
          className="hidden"
        />
      </div>
      {d && (
        <div className="mb-2 p-3 bg-amber-950/40 border border-amber-700 rounded-sm text-xs text-amber-100 space-y-2">
          <div className="font-semibold">One-time Supabase setup needed</div>
          <div>In Supabase, open the SQL Editor, paste this line and click Run. Then reopen this window.</div>
          <pre className="bg-gray-950 border border-gray-800 rounded-sm p-2 font-mono text-[11px] text-gray-200 overflow-x-auto">
            {README_SETUP_SQL}
          </pre>
        </div>
      )}
      {t === null ? (
        <div className="text-xs text-gray-500">Loading…</div>
      ) : p ? (
        <div className="max-h-72 overflow-y-auto p-3 bg-gray-950 border border-gray-800 rounded-sm">
          {r.trim() ? <ReadmeContent text={r} /> : <div className="text-xs text-gray-600 italic">Nothing yet.</div>}
        </div>
      ) : (
        <textarea
          value={r}
          onChange={(I) => o(I.target.value)}
          rows={8}
          placeholder={"# Title\n## Section\nA paragraph of instructions.\nTerm :: what it means"}
          className="w-full px-3 py-2 bg-gray-950 border border-gray-800 text-gray-100 placeholder-gray-600 rounded-sm text-base sm:text-xs font-mono focus:outline-none focus:ring-1 focus:ring-emerald-500 resize-y"
        />
      )}
      <div className="flex items-center justify-between gap-3 mt-1.5">
        <div className="text-[11px] text-gray-500">
          Shown to testers behind a "Read me" button. "# " title, "## " section, "Term :: description".
        </div>
        <div className="flex items-center gap-2 flex-shrink-0">
          {s && <span className={`text-[11px] ${s === "Saved" ? "text-emerald-400" : "text-gray-400"}`}>{s}</span>}
          <button
            type="button"
            onClick={_}
            disabled={i || t === null || r === t}
            className="px-2.5 py-1 text-[11px] font-medium bg-gray-800 border border-gray-700 text-gray-200 rounded-sm hover:bg-gray-700 disabled:opacity-40"
          >
            {i ? "Saving…" : "Save read me"}
          </button>
        </div>
      </div>
    </div>
  );
}
// "Upload updated sheet" flow, opened from Edit session.
function ResyncTestsModal({ session: e, onClose: t, onDone: a }) {
  let [r, o] = useState(true);
  let [s, l] = useState([]);
  let [i, u] = useState({});
  let [d, c] = useState("");
  let [p, h] = useState(false);
  let [b, _] = useState([]);
  let [S, I] = useState(false);
  let [x, g] = useState(false);
  let [v, k] = useState(null);
  let [C, m] = useState(false);
  let [resyncReadme, setResyncReadme] = useState(true);
  let f = useRef(null);
  useEffect(() => {
    let y = false;
    (async () => {
      let [E, D] = await Promise.all([
        supabase.from("test_cases").select("*").eq("session_id", e.id).order("position"),
        supabase.from("findings").select("*").eq("session_id", e.id),
      ]);
      if (y) return;
      if (E.error || D.error) {
        c((E.error || D.error).message);
      } else {
        l(E.data || []);
        let O = {};
        (D.data || []).forEach((B) => {
          O[B.test_case_id] = B;
        });
        u(O);
        if ((E.data || []).length > 0 && !("removed_at" in E.data[0])) h(true);
      }
      o(false);
    })();
    return () => {
      y = true;
    };
  }, [e.id]);
  useEffect(() => {
    let y = (E) => {
      if (E.key === "Escape" && !x) t();
    };
    window.addEventListener("keydown", y);
    return () => window.removeEventListener("keydown", y);
  }, [t, x]);
  let w = async (y) => {
    let E = Array.from(y || []);
    if (E.length === 0) return;
    I(true);
    c("");
    let D = await parseImportFiles(E);
    _(D);
    I(false);
    if (D.every((O) => O.skipped)) {
      c("None of the tabs or files contained a test table.");
    }
  };
  let T = (y) => _(b.map((E) => (E.key === y ? { ...E, include: !E.include } : E)));
  let A = useMemo(() => (b.some((y) => y.include && y.testCases) ? computeResyncPlan(s, b) : null), [s, b]);
  let j = (y) => {
    let E = i[y.id];
    return (
      !!E &&
      ((E.status && E.status !== "Not Tested") ||
        Object.keys(E.step_checks || {}).length > 0 ||
        Object.keys(E.step_notes || {}).length > 0)
    );
  };
  let M = A
    ? A.entries.filter((y) => y.kind === "match" && y.contentChanged && j(y.existing)).length
    : 0;
  let W = A ? A.removed.filter((y) => j(y)).length : 0;
  let V = async () => {
    let rsOn = resyncReadme && b.some((E) => E.readmeText);
    if (!rsOn && (!A || !A.hasChanges)) return;
    g(true);
    c("");
    try {
      let y = A && A.hasChanges ? await applyResyncPlan(e, A, Object.values(i)) : { flagged: 0 };
      let rs = b.find((E) => E.readmeText);
      if (rs && resyncReadme) {
        let { error: E } = await supabase.from("sessions").update({ readme: rs.readmeText }).eq("id", e.id);
        if (E) {
          throw new Error(
            isReadmeSetupError(E)
              ? `Tests updated, but the Read me needs a one-time setup. In the Supabase SQL Editor run: ${README_SETUP_SQL}`
              : E.message,
          );
        }
        y.readmeUpdated = true;
      }
      k(y);
    } catch (y) {
      c(
        /removed_at|content_updated_at/i.test(y.message || "")
          ? "The database needs a quick update first (see the yellow box above)."
          : y.message || "Update failed",
      );
      if (/removed_at|content_updated_at/i.test(y.message || "")) h(true);
    }
    g(false);
  };
  let F = (y) => y.row?.test_id || y.existing?.test_id || y.test_id || "—";
  let se = (y, E, D, O) =>
    E.length > 0 && (
      <details className="group border border-gray-800 rounded-sm">
        <summary className="flex items-center gap-2 px-3 py-2 cursor-pointer select-none text-sm text-gray-200 hover:bg-gray-800/40">
          <ChevronRight className="w-3.5 h-3.5 text-gray-500 transition-transform group-open:rotate-90" />
          <span className={`font-semibold ${D}`}>{E.length}</span> {y}
        </summary>
        <div className="px-3 pb-2.5 pt-1 space-y-1 max-h-48 overflow-y-auto">
          {E.map((B, Y) => (
            <div className="flex items-baseline gap-2 text-xs" key={Y}>
              <span className="font-mono text-emerald-400 flex-shrink-0">{F(B)}</span>
              <span className="text-gray-400 truncate">{O(B)}</span>
            </div>
          ))}
        </div>
      </details>
    );
  let K = (y) => {
    if (y.skipped) return <span className="text-gray-500">{y.skipped}</span>;
    if (!y.include) return <span className="text-gray-500">Not included · left as is</span>;
    let E = A?.perSource.find((O) => O.name === y.name);
    if (!E) return null;
    if (E.isNewTab) return <span className="text-sky-300">New tab · {E.total} tests</span>;
    let D = [
      E.updated && `${E.updated} updated`,
      E.added && `${E.added} new`,
      E.moved && `${E.moved} moved`,
      E.restored && `${E.restored} restored`,
      E.removed && `${E.removed} removed`,
    ].filter(Boolean);
    return <span className={D.length ? "text-amber-300" : "text-gray-500"}>{D.length ? D.join(" · ") : "No changes"}</span>;
  };
  let G = A ? A.added.length + A.updated.length + A.moved.length + A.restored.length + A.removed.length : 0;
  return (
    <div className="fixed inset-0 z-[60] flex items-center justify-center bg-black/70 p-4" onClick={() => !x && t()}>
      <div
        className="bg-gray-900 border border-gray-700 rounded-sm w-full max-w-2xl max-h-[90vh] flex flex-col shadow-2xl"
        onClick={(y) => y.stopPropagation()}
      >
        <div className="flex items-center justify-between px-5 py-3 border-b border-gray-800 flex-shrink-0">
          <div className="min-w-0">
            <h2 className="text-base font-bold text-gray-100">Upload updated sheet</h2>
            <div className="text-xs text-gray-500 truncate">{e.name}</div>
          </div>
          <button onClick={t} disabled={x} className="text-gray-400 hover:text-gray-100 p-1" title="Close">
            <X className="w-5 h-5" />
          </button>
        </div>
        {v ? (
          <div className="flex-1 overflow-y-auto px-5 py-6 text-center">
            <CircleCheck className="w-10 h-10 text-emerald-400 mx-auto mb-3" />
            <div className="text-lg font-semibold text-gray-100 mb-1">Session updated</div>
            <div className="text-sm text-gray-400">
              {!A ? "Read me updated." : null}
              {A && <>{A.updated.length} updated · {A.added.length} new · {A.moved.length} moved
              {A.restored.length ? ` · ${A.restored.length} restored` : ""} · {A.removed.length} removed from sheet</>}
              {A && v.readmeUpdated ? " · Read me updated" : ""}
            </div>
            <div className="text-xs text-gray-500 mt-2">
              All existing results were kept.
              {v.flagged > 0 ? ` ${v.flagged} test${v.flagged === 1 ? " is" : "s are"} flagged "Updated" for testers to re-check.` : ""}
            </div>
          </div>
        ) : (
          <div className="flex-1 overflow-y-auto px-5 py-4 space-y-4">
            {p && (
              <div className="p-3 bg-amber-950/40 border border-amber-700 rounded-sm text-xs text-amber-100 space-y-2">
                <div className="font-semibold">One-time database update needed</div>
                <div>
                  In Supabase, open the SQL Editor, paste these two lines and click Run. Then close and reopen this
                  window.
                </div>
                <pre className="bg-gray-950 border border-gray-800 rounded-sm p-2 font-mono text-[11px] text-gray-200 overflow-x-auto">
                  {`alter table test_cases add column if not exists removed_at timestamptz;
alter table test_cases add column if not exists content_updated_at timestamptz;`}
                </pre>
              </div>
            )}
            {r ? (
              <div className="text-sm text-gray-400">Loading the current tests…</div>
            ) : (
              <>
                <div className="text-xs text-gray-400 leading-relaxed">
                  Tests are matched by tab and test ID. Changed wording is updated, new tests are added, and tests
                  that disappeared from a tab are hidden, never deleted. Status, severity, notes, screenshots and dev
                  notes stay attached. Nothing is saved until you confirm.
                </div>
                <div
                  onClick={() => f.current?.click()}
                  onDragOver={(y) => {
                    y.preventDefault();
                    m(true);
                  }}
                  onDragLeave={() => m(false)}
                  onDrop={(y) => {
                    y.preventDefault();
                    m(false);
                    w(y.dataTransfer.files);
                  }}
                  className={`border-2 border-dashed rounded-sm p-5 text-center cursor-pointer transition-colors ${C ? "border-emerald-500 bg-emerald-950/20" : "border-gray-700 hover:border-gray-600"}`}
                >
                  <Upload className="w-5 h-5 text-gray-500 mx-auto mb-1.5" />
                  <div className="text-sm text-gray-300">
                    {S ? "Reading…" : b.length ? "Drop a different file to replace it" : "Drop the updated sheet here or click to choose"}
                  </div>
                  <div className="text-xs text-gray-500 mt-1">.xlsx (all tabs) or one or more CSVs</div>
                  <input
                    ref={f}
                    type="file"
                    multiple
                    accept=".csv,.xlsx,.xlsm,.xls,.ods"
                    onChange={(y) => {
                      w(y.target.files);
                      y.target.value = "";
                    }}
                    className="hidden"
                  />
                </div>
                <SheetLinkImport disabled={S || x} onFile={(y) => w([y])} />
                {b.length > 0 && (
                  <div className="border border-gray-800 rounded-sm divide-y divide-gray-800 overflow-hidden">
                    {b.map((y) => {
                      let E = !!y.readmeText && y === b.find((O) => O.readmeText);
                      let D = E ? resyncReadme : !!y.include;
                      return (
                        <label
                          key={y.key}
                          className={`flex items-center gap-3 px-3 py-2 text-sm ${y.skipped && !E ? "opacity-50 cursor-not-allowed" : "cursor-pointer hover:bg-gray-800/50"}`}
                        >
                          <input
                            type="checkbox"
                            checked={D}
                            disabled={!!y.skipped && !E}
                            onChange={() => (E ? setResyncReadme(!resyncReadme) : T(y.key))}
                            className="accent-emerald-500 w-4 h-4 flex-shrink-0"
                          />
                          <span className={`flex-1 truncate ${D ? "text-gray-100" : "text-gray-400"}`}>{y.name}</span>
                          <span className="text-xs flex-shrink-0 text-right">
                            {E ? (
                              <span className={D ? "text-sky-300" : "text-gray-500"}>
                                {D ? "Update testers' Read me" : "Read me · left as is"}
                              </span>
                            ) : (
                              K(y)
                            )}
                          </span>
                        </label>
                      );
                    })}
                  </div>
                )}
                {A && (
                  <div className="space-y-2">
                    <div className="text-xs font-semibold uppercase tracking-wider text-gray-400">
                      What will change
                    </div>
                    {G === 0 && !A.hasChanges && (
                      <div className="text-sm text-gray-400">Everything already matches the sheet. Nothing to update.</div>
                    )}
                    {G === 0 && A.hasChanges && (
                      <div className="text-sm text-gray-400">Only the order of tests changes.</div>
                    )}
                    {se("updated", A.updated, "text-amber-300", (y) =>
                      `${y.changedFields.map((E) => RESYNC_FIELD_LABELS[E] || E).join(", ")} changed`,
                    )}
                    {se("new", A.added, "text-sky-300", (y) => `${y.row.title || ""} · ${y.row.module}`)}
                    {se("moved to another tab", A.moved, "text-violet-300", (y) => `${y.existing.module} → ${y.row.module}`)}
                    {se("restored (back in the sheet)", A.restored, "text-emerald-300", (y) => y.row.title || "")}
                    {se("removed from the sheet (hidden, results kept)", A.removed, "text-red-300", (y) => `${y.title || ""} · ${y.module}`)}
                    {A.leftAlone.length > 0 && (
                      <div className="text-xs text-gray-500">
                        Not in this upload, left as is: {A.leftAlone.join(", ")} ({A.leftAloneCount} test
                        {A.leftAloneCount === 1 ? "" : "s"})
                      </div>
                    )}
                    {M > 0 && (
                      <div className="px-3 py-2 bg-orange-950/30 border border-orange-800/70 rounded-sm text-xs text-orange-200">
                        {M} tested test{M === 1 ? " has" : "s have"} reworded steps or expected results. Ticks on
                        unchanged lines are kept; ticks on reworded lines are cleared, and the test is flagged
                        "Updated" for the tester.
                      </div>
                    )}
                    {W > 0 && (
                      <div className="text-xs text-gray-500">
                        {W} of the removed tests already have results. They're kept and come back if the test returns
                        to the sheet.
                      </div>
                    )}
                  </div>
                )}
              </>
            )}
            {d && (
              <div className="px-3 py-2 bg-red-950/40 border border-red-800 rounded-sm text-xs text-red-200">{d}</div>
            )}
          </div>
        )}
        <div className="flex items-center justify-end gap-2 px-5 py-3 border-t border-gray-800 flex-shrink-0">
          {v ? (
            <button
              onClick={a}
              className="px-4 py-1.5 text-sm rounded-sm font-medium bg-emerald-600 text-white hover:bg-emerald-500"
            >
              Done
            </button>
          ) : (
            <>
              <button onClick={t} disabled={x} className="px-3 py-1.5 text-sm text-gray-400 hover:text-gray-200">
                Cancel
              </button>
              <button
                onClick={V}
                disabled={(!(A && A.hasChanges) && !(resyncReadme && b.some((y) => y.readmeText))) || x || p}
                className="px-4 py-1.5 text-sm rounded-sm font-medium bg-emerald-600 text-white hover:bg-emerald-500 disabled:bg-gray-700 disabled:text-gray-500 disabled:cursor-not-allowed"
              >
                {x ? "Updating…" : "Update session"}
              </button>
            </>
          )}
        </div>
      </div>
    </div>
  );
}
function EditSessionModal({ session: e, onClose: t, onSaved: a }) {
  let [resyncOpen, setResyncOpen] = useState(false);
  let [removedTests, setRemovedTests] = useState([]);
  let [restoringId, setRestoringId] = useState(null);
  let loadRemovedTests = async () => {
    let { data: E, error: D } = await supabase
      .from("test_cases")
      .select("id, test_id, title, module, removed_at")
      .eq("session_id", e.id)
      .not("removed_at", "is", null)
      .order("position");
    if (!D) setRemovedTests(E || []);
  };
  useEffect(() => {
    loadRemovedTests();
  }, [e.id]);
  let restoreTest = async (E) => {
    setRestoringId(E);
    let { error: D } = await supabase.from("test_cases").update({ removed_at: null }).eq("id", E);
    setRestoringId(null);
    if (!D) setRemovedTests((O) => O.filter((B) => B.id !== E));
  };
  let [r, o] = useState(e.name || "");
  let [s, l] = useState(e.base_url || "");
  let [i, u] = useState([]);
  let [d, c] = useState([]);
  let [p, h] = useState("");
  let [b, _] = useState("");
  let [S, I] = useState(false);
  let [x, g] = useState("");
  let v = e.testers || [];
  let k = e.developers || [];
  useEffect(() => {
    if (resyncOpen) {
      return;
    }
    let E = (D) => {
      if (D.key === "Escape") {
        t();
      }
    };
    window.addEventListener("keydown", E);
    return () => window.removeEventListener("keydown", E);
  }, [t, resyncOpen]);
  let C = () => {
    let E = p.trim();
    if (!E) {
      return;
    }
    if (![...v, ...i].some((O) => O.toLowerCase() === E.toLowerCase())) {
      u([...i, E]);
    }
    h("");
  };
  let m = () => {
    let E = b.trim();
    if (!E) {
      return;
    }
    if (![...k, ...d].some((O) => O.toLowerCase() === E.toLowerCase())) {
      c([...d, E]);
    }
    _("");
  };
  let f = (E) => u(i.filter((D, O) => O !== E));
  let w = (E) => c(d.filter((D, O) => O !== E));
  let T = r.trim();
  let A = s.trim();
  let j = (e.base_url || "").trim();
  let M = T && T !== e.name;
  let W = A !== j;
  let V = i.length > 0 || d.length > 0;
  let F = !!T && (M || W || V) && !S;
  let se = async () => {
    if (F) {
      I(true);
      g("");
      try {
        let { error: E } = await supabase
          .from("sessions")
          .update({
            name: T,
            testers: [...v, ...i],
            developers: [...k, ...d],
            base_url: A || null,
          })
          .eq("id", e.id);
        if (E) {
          throw E;
        }
        a();
      } catch (E) {
        g(E.message || "Save failed");
        I(false);
      }
    }
  };
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/70 p-4" onClick={t}>
      <div
        className="bg-gray-900 border border-gray-700 rounded-sm w-full max-w-xl max-h-[90vh] flex flex-col shadow-2xl"
        onClick={(E) => E.stopPropagation()}
      >
        <div className="flex items-center justify-between px-5 py-3 border-b border-gray-800 flex-shrink-0">
          <h2 className="text-base font-bold text-gray-100">Edit session</h2>
          <button onClick={t} className="text-gray-400 hover:text-gray-100 p-1" title="Close">
            <X className="w-5 h-5" />
          </button>
        </div>
        <div className="flex-1 overflow-y-auto px-5 py-4 space-y-5">
          <div>
            <label className="block text-xs font-semibold uppercase tracking-wider text-gray-400 mb-1.5">
              Session name
            </label>
            <input
              type="text"
              value={r}
              onChange={(E) => o(E.target.value)}
              className="w-full px-3 py-2 bg-gray-800 border border-gray-700 text-gray-100 rounded-sm text-sm focus:outline-none focus:border-emerald-500"
            />
          </div>
          <div>
            <label className="block text-xs font-semibold uppercase tracking-wider text-gray-400 mb-1.5">
              Path chip prefix <span className="normal-case font-normal text-gray-500">(optional)</span>
            </label>
            <input
              type="text"
              value={s}
              onChange={(E) => l(E.target.value)}
              placeholder="https://localhost:3000"
              autoCapitalize="off"
              autoComplete="off"
              spellCheck="false"
              className={`w-full px-3 py-2 text-sm font-mono bg-gray-800 border text-gray-100 placeholder-gray-500 rounded-sm focus:outline-none focus:border-emerald-500 ${s.trim() ? "border-emerald-800" : "border-gray-700"}`}
            />
            <div className="text-[11px] text-gray-500 mt-1.5 leading-relaxed">
              Default for both portals. Testers and developers can still override this in their own Settings —
              their overrides take precedence.
            </div>
          </div>
          <div>
            <div className="text-xs font-semibold uppercase tracking-wider text-gray-400 mb-1.5">
              Testers <span className="normal-case font-normal text-gray-500">({v.length + i.length})</span>
            </div>
            {v.length > 0 && (
              <div className="flex flex-wrap gap-1.5 mb-2">
                {v.map((E) => (
                  <span
                    className="inline-flex items-center gap-1.5 px-2 py-0.5 bg-gray-800 border border-gray-700 text-gray-300 rounded-sm text-xs"
                    key={E}
                  >
                    {E}
                  </span>
                ))}
              </div>
            )}
            {i.length > 0 && (
              <div className="flex flex-wrap gap-1.5 mb-2">
                {i.map((E, D) => (
                  <span
                    className="inline-flex items-center gap-1 px-2 py-0.5 bg-emerald-900/40 border border-emerald-700 text-emerald-200 rounded-sm text-xs"
                    key={D}
                  >
                    {E}
                    <button
                      onClick={() => f(D)}
                      className="text-emerald-300 hover:text-emerald-100"
                      title="Remove from this edit"
                    >
                      <X className="w-3 h-3" />
                    </button>
                  </span>
                ))}
              </div>
            )}
            <div className="flex gap-2">
              <input
                type="text"
                value={p}
                onChange={(E) => h(E.target.value)}
                onKeyDown={(E) => {
                  if (E.key === "Enter") {
                    E.preventDefault();
                    C();
                  }
                }}
                placeholder="Add a tester name"
                className="flex-1 px-3 py-1.5 bg-gray-800 border border-gray-700 text-gray-100 placeholder-gray-500 rounded-sm text-sm focus:outline-none focus:border-emerald-500"
              />
              <button
                onClick={C}
                disabled={!p.trim()}
                className="px-3 py-1.5 text-xs font-medium bg-emerald-600 text-white rounded-sm hover:bg-emerald-500 disabled:bg-gray-800 disabled:text-gray-500"
              >
                Add
              </button>
            </div>
          </div>
          <div>
            <div className="text-xs font-semibold uppercase tracking-wider text-gray-400 mb-1.5">
              Developers{" "}
              <span className="normal-case font-normal text-gray-500">({k.length + d.length})</span>
            </div>
            {k.length > 0 && (
              <div className="flex flex-wrap gap-1.5 mb-2">
                {k.map((E) => (
                  <span
                    className="inline-flex items-center gap-1.5 px-2 py-0.5 bg-gray-800 border border-gray-700 text-gray-300 rounded-sm text-xs"
                    key={E}
                  >
                    {E}
                  </span>
                ))}
              </div>
            )}
            {d.length > 0 && (
              <div className="flex flex-wrap gap-1.5 mb-2">
                {d.map((E, D) => (
                  <span
                    className="inline-flex items-center gap-1 px-2 py-0.5 bg-emerald-900/40 border border-emerald-700 text-emerald-200 rounded-sm text-xs"
                    key={D}
                  >
                    {E}
                    <button
                      onClick={() => w(D)}
                      className="text-emerald-300 hover:text-emerald-100"
                      title="Remove from this edit"
                    >
                      <X className="w-3 h-3" />
                    </button>
                  </span>
                ))}
              </div>
            )}
            <div className="flex gap-2">
              <input
                type="text"
                value={b}
                onChange={(E) => _(E.target.value)}
                onKeyDown={(E) => {
                  if (E.key === "Enter") {
                    E.preventDefault();
                    m();
                  }
                }}
                placeholder="Add a developer name"
                className="flex-1 px-3 py-1.5 bg-gray-800 border border-gray-700 text-gray-100 placeholder-gray-500 rounded-sm text-sm focus:outline-none focus:border-emerald-500"
              />
              <button
                onClick={m}
                disabled={!b.trim()}
                className="px-3 py-1.5 text-xs font-medium bg-emerald-600 text-white rounded-sm hover:bg-emerald-500 disabled:bg-gray-800 disabled:text-gray-500"
              >
                Add
              </button>
            </div>
          </div>
          <ReadmeEditor session={e} />
          <SampleFilesManager session={e} />
          <div>
            <div className="text-xs font-semibold uppercase tracking-wider text-gray-400 mb-1.5">Test cases</div>
            <button
              onClick={() => setResyncOpen(true)}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-gray-200 bg-gray-800 hover:bg-gray-700 border border-gray-700 rounded-sm transition-colors"
            >
              <RefreshCw className="w-3.5 h-3.5" /> Upload updated sheet
            </button>
            <div className="text-[11px] text-gray-500 mt-1.5 leading-relaxed">
              Updates wording, adds new tests and tabs, and hides tests removed from the sheet. Existing results are
              kept. You'll see a summary before anything changes.
            </div>
            {removedTests.length > 0 && (
              <details className="group mt-2 border border-gray-800 rounded-sm">
                <summary className="flex items-center gap-2 px-3 py-2 cursor-pointer select-none text-xs text-gray-300 hover:bg-gray-800/40">
                  <ChevronRight className="w-3.5 h-3.5 text-gray-500 transition-transform group-open:rotate-90" />
                  Removed from sheet ({removedTests.length}) · hidden from testers, results kept
                </summary>
                <div className="px-3 pb-2 pt-1 space-y-1 max-h-48 overflow-y-auto">
                  {removedTests.map((E) => (
                    <div className="flex items-center gap-2 text-xs" key={E.id}>
                      <span className="font-mono text-emerald-400 flex-shrink-0">{E.test_id || "—"}</span>
                      <span className="text-gray-400 truncate flex-1">
                        {E.title} · {E.module}
                      </span>
                      <button
                        onClick={() => restoreTest(E.id)}
                        disabled={restoringId === E.id}
                        className="flex-shrink-0 px-2 py-0.5 text-[11px] text-gray-300 border border-gray-700 rounded-sm hover:bg-gray-800 disabled:text-gray-600"
                      >
                        {restoringId === E.id ? "Restoring…" : "Restore"}
                      </button>
                    </div>
                  ))}
                </div>
              </details>
            )}
          </div>
          <div className="text-[10px] text-gray-500 pt-2 border-t border-gray-800">
            Existing testers and developers can't be removed here — only added. Their findings stay
            regardless.
          </div>
          {x && (
            <div className="px-3 py-2 bg-red-950/40 border border-red-800 rounded-sm text-xs text-red-200">
              {x}
            </div>
          )}
        </div>
        <div className="flex items-center justify-end gap-2 px-5 py-3 border-t border-gray-800 flex-shrink-0">
          <button onClick={t} className="px-3 py-1.5 text-sm text-gray-400 hover:text-gray-200">
            Cancel
          </button>
          <button
            onClick={se}
            disabled={!F}
            className="px-4 py-1.5 text-sm rounded-sm font-medium bg-emerald-600 text-white hover:bg-emerald-500 disabled:bg-gray-700 disabled:text-gray-500 disabled:cursor-not-allowed"
          >
            {S ? "Saving…" : "Save"}
          </button>
        </div>
      </div>
      {resyncOpen && (
        <div onClick={(E) => E.stopPropagation()}>
          <ResyncTestsModal
            session={e}
            onClose={() => setResyncOpen(false)}
            onDone={() => {
              setResyncOpen(false);
              a();
            }}
          />
        </div>
      )}
    </div>
  );
}
function CreateSessionWizard({ onCancel: e, onCreated: t }) {
  let [a, r] = useState("");
  let [o, s] = useState(null);
  let [l, i] = useState(null);
  let [u, d] = useState("");
  let [c, p] = useState([]);
  let [h, b] = useState([]);
  let [_, S] = useState("");
  let [I, x] = useState("");
  let [g, v] = useState(randomCode());
  let [k, C] = useState(randomCode());
  let [m, f] = useState("");
  let [w, T] = useState(false);
  let [A, j] = useState("");
  let [M, W] = useState(false);
  let [Q, setImportSources] = useState([]);
  let [sampleFiles, setSampleFiles] = useState([]);
  let [useReadme, setUseReadme] = useState(true);
  let V = useRef(null);
  let F = async () => {
    let P = false;
    try {
      if (navigator.clipboard?.writeText) {
        await navigator.clipboard.writeText(PRO_CSV_PROMPT);
        P = true;
      }
    } catch (X) {
      console.warn("Clipboard API failed, falling back:", X);
    }
    if (!P) {
      try {
        let X = document.createElement("textarea");
        X.value = PRO_CSV_PROMPT;
        X.style.position = "fixed";
        X.style.opacity = "0";
        document.body.appendChild(X);
        X.focus();
        X.select();
        P = document.execCommand("copy");
        document.body.removeChild(X);
      } catch (X) {
        console.error("Fallback copy failed:", X);
      }
    }
    if (P) {
      W(true);
      setTimeout(() => W(false), 2000);
    }
  };
  let [se, E] = useState([]);
  let [D, O] = useState([]);
  let [B, ne] = useState(() => {
    try {
      return JSON.parse(localStorage.getItem("uat-pro-hidden-testers") || "[]");
    } catch {
      return [];
    }
  });
  let [he, ie] = useState(() => {
    try {
      return JSON.parse(localStorage.getItem("uat-pro-hidden-devs") || "[]");
    } catch {
      return [];
    }
  });
  useEffect(() => {
    let P = false;
    (async () => {
      let { data: X, error: ve } = await supabase
        .from("sessions")
        .select("testers, developers")
        .order("created_at", {
          ascending: false,
        })
        .limit(50);
      if (P || ve || !X) {
        return;
      }
      let de = new Set();
      let N = new Set();
      for (let G of X) {
        (G.testers || []).forEach((Q) => Q && de.add(Q));
        (G.developers || []).forEach((Q) => Q && N.add(Q));
      }
      E([...de].sort((G, Q) => G.localeCompare(Q)));
      O([...N].sort((G, Q) => G.localeCompare(Q)));
    })();
    return () => {
      P = true;
    };
  }, []);
  let ue = (P, X) => {
    try {
      localStorage.setItem(P, JSON.stringify(X));
    } catch {}
  };
  let pe = (P) => {
    let X = Array.from(new Set([...B, P]));
    ne(X);
    ue("uat-pro-hidden-testers", X);
  };
  let me = (P) => {
    let X = Array.from(new Set([...he, P]));
    ie(X);
    ue("uat-pro-hidden-devs", X);
  };
  let Y = () => {
    ne([]);
    ue("uat-pro-hidden-testers", []);
  };
  let fe = () => {
    ie([]);
    ue("uat-pro-hidden-devs", []);
  };
  let Ce = se.filter((P) => !B.includes(P) && !c.includes(P));
  let be = D.filter((P) => !he.includes(P) && !h.includes(P));
  useEffect(() => {
    if (l) {
      if (!a && o) {
        let P = o.name
          .replace(/\.[a-z0-9]+$/i, "")
          .replace(/[_-]/g, " ")
          .trim();
        r(P);
      }
      p((P) => Array.from(new Set([...P, ...(l.testers || [])])));
    }
  }, [l]);
  // Accepts one or more CSV files and/or spreadsheets (.xlsx). Each CSV file and each
  // spreadsheet tab becomes a "source" the admin can tick on or off.
  let xe = async (P) => {
    let X = Array.from(P || []);
    if (X.length === 0) {
      return;
    }
    s({
      name: X.length === 1 ? X[0].name : `${X.length} files`,
    });
    d("");
    i(null);
    setImportSources([]);
    let ve = await parseImportFiles(X);
    if (X.length === 1) {
      s(X[0]);
    } else {
      s({
        name: X[0].name,
        label: `${X.length} files`,
      });
    }
    setImportSources(ve);
    let de = combineImportSources(ve);
    if (de) {
      i(de);
    } else if (ve.length === 1 && ve[0].skipped) {
      d(ve[0].skipped);
    } else if (ve.every((N) => N.skipped)) {
      d("None of the tabs or files contained a test table. Each needs a header row with at least an ID column and a Title/Test or Steps column.");
    } else {
      d("Tick at least one tab below to import.");
    }
  };
  let toggleImportSource = (P) => {
    let X = Q.map((de) => (de.key === P ? { ...de, include: !de.include } : de));
    setImportSources(X);
    let ve = combineImportSources(X);
    i(ve);
    d(ve ? "" : "Tick at least one tab below to import.");
  };
  let $ = () => {
    let P = _.trim();
    if (P && !c.includes(P)) {
      p([...c, P]);
    }
    S("");
  };
  let H = () => {
    let P = I.trim();
    if (P && !h.includes(P)) {
      b([...h, P]);
    }
    x("");
  };
  let K = a.trim() && l && g.length === 4 && k.length === 4 && !w;
  let te = async () => {
    T(true);
    j("");
    try {
      // Upload sample files first (under a pre-made session id) so a failed upload never leaves a half-made session.
      let sid = crypto.randomUUID ? crypto.randomUUID() : undefined;
      let uploadedSamples = [];
      if (sampleFiles.length > 0) {
        if (!sid) {
          throw new Error("This browser can't attach files here. Create the session, then add files from Edit.");
        }
        uploadedSamples = await uploadSampleFiles(sid, sampleFiles);
      }
      let { data: P, error: X } = await supabase
        .from("sessions")
        .insert({
          ...(sid ? { id: sid } : {}),
          name: a.trim(),
          testers: c,
          developers: h,
          tester_code: g,
          developer_code: k,
          tester_col_header: l.testerColHeader || "Tester",
          base_url: m.trim() || null,
          ...(uploadedSamples.length ? { sample_files: uploadedSamples } : {}),
          ...(useReadme && Q.find((N) => N.readmeText) ? { readme: Q.find((N) => N.readmeText).readmeText } : {}),
        })
        .select()
        .single();
      if (X && isReadmeSetupError(X)) {
        throw new Error(
          `The Read me needs a one-time Supabase setup. In the SQL Editor run: ${README_SETUP_SQL}  (or untick "Use as Read me" and create the session without it).`,
        );
      }
      if (X) {
        if (uploadedSamples.length) {
          await supabase.storage.from(SAMPLE_BUCKET).remove(uploadedSamples.map((N) => N.path));
          if (isSampleSetupError(X)) throw new Error(SAMPLE_SETUP_MESSAGE);
        }
        throw X;
      }
      let ve = l.testCases.map((N) => ({
        session_id: P.id,
        position: N.position,
        row_num: N.rowNum,
        test_id: N.testId,
        module: N.module,
        feature: N.feature,
        title: N.title,
        pre_conditions: N.preConditions,
        test_data: N.testData,
        steps: N.steps,
        expected_result: N.expectedResult,
        downstream_check: N.downstreamCheck,
        edge_case: N.edgeCase,
        dev_branch_name: N.devBranchName,
        initial_tester: N.initialTester,
        focus: N.focus || null,
        priority: N.priority || null,
      }));
      let { error: de } = await supabase.from("test_cases").insert(ve);
      if (de && /focus|priority/i.test(de.message || "")) {
        // Database hasn't had the focus/priority migration yet: save without those two fields.
        console.warn("focus/priority columns missing — run the latest SQL migration. Saving without them.");
        ({ error: de } = await supabase
          .from("test_cases")
          .insert(ve.map(({ focus: N, priority: G, ...Q }) => Q)));
      }
      if (de) {
        throw de;
      }
      t(P);
    } catch (P) {
      console.error("Create session failed:", P);
      j(P.message || String(P));
      T(false);
    }
  };
  return (
    <div className="min-h-screen bg-gradient-to-br from-gray-950 via-gray-900 to-indigo-950">
      <header className="px-8 py-5 border-b border-gray-800 bg-gray-950/60 backdrop-blur sticky top-0 z-10">
        <div className="max-w-3xl mx-auto flex items-center gap-3">
          <button
            onClick={e}
            className="p-1.5 hover:bg-gray-800 rounded-sm text-gray-400 hover:text-gray-200"
          >
            <ArrowLeft className="w-4 h-4" />
          </button>
          <div className="font-bold text-gray-100">New session</div>
        </div>
      </header>
      <main className="max-w-3xl mx-auto px-8 py-10 space-y-7">
        <div>
          <label className="block text-xs font-semibold uppercase tracking-wider text-gray-400 mb-1.5">
            Session name
          </label>
          <input
            type="text"
            value={a}
            onChange={(P) => r(P.target.value)}
            placeholder="e.g. Q2 2026 UAT — Customisation Module"
            className="w-full px-3 py-2 bg-gray-800 border border-gray-700 text-gray-100 placeholder-gray-500 rounded-sm focus:outline-none focus:ring-2 focus:ring-emerald-500 focus:border-transparent"
          />
        </div>
        <div>
          <div className="flex items-center justify-between mb-1.5">
            <label className="block text-xs font-semibold uppercase tracking-wider text-gray-400">
              Test cases
            </label>
            <button
              type="button"
              onClick={F}
              className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-sm text-[11px] font-medium border transition-colors ${M ? "bg-emerald-950/50 border-emerald-700 text-emerald-200" : "bg-gray-800 border-gray-700 text-gray-300 hover:bg-gray-700 hover:text-emerald-200 hover:border-emerald-700"}`}
              title="Copy a prompt you can paste into ChatGPT/Claude to generate a CSV in this format"
            >
              {M ? (
                <Fragment>
                  <Check className="w-3 h-3" /> Prompt copied
                </Fragment>
              ) : (
                <Fragment>
                  <Copy className="w-3 h-3" /> Don't have a CSV? Copy AI prompt
                </Fragment>
              )}
            </button>
          </div>
          <div
            onClick={() => V.current?.click()}
            onDragOver={(P) => {
              P.preventDefault();
              P.currentTarget.classList.add("border-emerald-500", "bg-emerald-950/20");
            }}
            onDragLeave={(P) => {
              P.currentTarget.classList.remove("border-emerald-500", "bg-emerald-950/20");
            }}
            onDrop={(P) => {
              P.preventDefault();
              P.currentTarget.classList.remove("border-emerald-500", "bg-emerald-950/20");
              xe(P.dataTransfer.files);
            }}
            className="border-2 border-dashed border-gray-700 rounded-sm p-6 text-center cursor-pointer hover:border-gray-600 transition-colors"
          >
            <Upload className="w-7 h-7 text-gray-500 mx-auto mb-2" />
            {l ? (
              <div className="text-sm">
                <div className="text-emerald-300 font-medium">{o?.label || o?.name}</div>
                <div className="text-xs text-gray-500 mt-1">
                  {l.testCases.length} test case{l.testCases.length === 1 ? "" : "s"} parsed ·{" "}
                  {l.testers.length} tester{l.testers.length === 1 ? "" : "s"} detected
                </div>
              </div>
            ) : o && Q.length === 0 && !u ? (
              <div className="text-sm text-gray-400">{o.label || o.name} — parsing…</div>
            ) : (
              <div className="text-sm">
                <div className="text-gray-400">Drop files here or click to choose</div>
                <div className="text-xs text-gray-500 mt-1">
                  A CSV, several CSVs, or a whole Google Sheet downloaded as .xlsx (every tab is read)
                </div>
              </div>
            )}
            <input
              ref={V}
              type="file"
              multiple
              accept=".csv,.xlsx,.xlsm,.xls,.ods"
              onChange={(P) => {
                xe(P.target.files);
                P.target.value = "";
              }}
              className="hidden"
            />
          </div>
          <SheetLinkImport onFile={(P) => xe([P])} />
          {(Q.length > 1 || Q.some((P) => P.skipped)) && (
            <div className="mt-2 border border-gray-800 rounded-sm divide-y divide-gray-800 overflow-hidden">
              <div className="flex items-center justify-between px-3 py-2 bg-gray-900/60 text-[11px] text-gray-400">
                <span>
                  {Q.filter((P) => P.include).length} of {Q.filter((P) => !P.skipped).length} tab
                  {Q.filter((P) => !P.skipped).length === 1 ? "" : "s"} included
                </span>
                <span className="text-gray-500">Tab name is used as the Module when there's no Module column</span>
              </div>
              {Q.map((P) => {
                let X = !!P.readmeText && P === Q.find((N) => N.readmeText);
                let ve = X ? useReadme : !!P.include;
                return (
                  <label
                    key={P.key}
                    className={`flex items-center gap-3 px-3 py-2 text-sm ${P.skipped && !X ? "opacity-50 cursor-not-allowed" : "cursor-pointer hover:bg-gray-800/50"}`}
                  >
                    <input
                      type="checkbox"
                      checked={ve}
                      disabled={!!P.skipped && !X}
                      onChange={() => (X ? setUseReadme(!useReadme) : toggleImportSource(P.key))}
                      className="accent-emerald-500 w-4 h-4 flex-shrink-0"
                    />
                    <span className={`flex-1 truncate ${ve ? "text-gray-100" : "text-gray-400"}`}>{P.name}</span>
                    <span className="text-xs text-gray-500 flex-shrink-0 text-right">
                      {X ? (
                        <span className={ve ? "text-sky-300" : "text-gray-500"}>
                          {ve ? "Read me for testers" : "Read me · not used"}
                        </span>
                      ) : P.skipped ? (
                        P.skipped
                      ) : (
                        `${P.testCases.length} test${P.testCases.length === 1 ? "" : "s"}${P.hidden ? " · hidden tab" : ""}`
                      )}
                    </span>
                  </label>
                );
              })}
            </div>
          )}
          {l?.duplicateIds?.length > 0 && (
            <div className="mt-2 p-2.5 bg-amber-950/30 border border-amber-800/60 rounded-sm text-xs text-amber-200">
              Heads up: {l.duplicateIds.length} test ID{l.duplicateIds.length === 1 ? " appears" : "s appear"} more
              than once ({l.duplicateIds.slice(0, 5).join(", ")}
              {l.duplicateIds.length > 5 ? ", …" : ""}). They'll still import, but may be confusing for testers.
            </div>
          )}
          <div className="mt-2 text-[11px] leading-relaxed space-y-1">
            <div className="text-gray-500">
              <span className="font-semibold text-emerald-400">Required: </span>
              <span className="font-mono text-gray-300">ID, Title (or Test), Steps</span>
            </div>
            <div className="text-gray-500">
              <span className="font-semibold text-gray-400">Recommended: </span>
              <span className="font-mono text-gray-400">
                Module (otherwise the tab or file name is used), Feature, Pre-conditions, Test Data, Expected
                Result, Downstream Check, Edge Case, Dev Branch Name, Tester
              </span>
            </div>
            <div className="text-gray-500">
              Tip: name the tester column{" "}
              <span className="font-mono text-gray-400">Tester [1 - Alice, 2 - Bob]</span> to auto-map
              initials to names.
            </div>
          </div>
          {u && (
            <div className="mt-2 p-3 bg-red-950/40 border border-red-800 rounded-sm text-xs text-red-300">
              {u}
            </div>
          )}
        </div>
        <div>
          <label className="block text-xs font-semibold uppercase tracking-wider text-gray-400 mb-1.5">
            Testers{" "}
            {c.length > 0 && <span className="normal-case font-normal text-gray-500">({c.length})</span>}
          </label>
          <div className="text-[11px] text-gray-500 mb-2">
            Auto-filled from the CSV's tester column. Add or remove as needed.
          </div>
          <NameChipList items={c} onRemove={(P) => p(c.filter((X, ve) => ve !== P))} />
          <div className="flex gap-2 mt-2">
            <input
              type="text"
              value={_}
              onChange={(P) => S(P.target.value)}
              onKeyDown={(P) => {
                if (P.key === "Enter") {
                  P.preventDefault();
                  $();
                }
              }}
              placeholder="Tester name"
              className="flex-1 px-3 py-1.5 bg-gray-800 border border-gray-700 text-gray-100 placeholder-gray-500 rounded-sm text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500 focus:border-transparent"
            />
            <button
              onClick={$}
              disabled={!_.trim()}
              className="px-3 py-1.5 bg-gray-800 text-gray-100 text-sm font-medium rounded-sm hover:bg-gray-700 disabled:opacity-50 disabled:cursor-not-allowed"
            >
              Add
            </button>
          </div>
          <NameSuggestions
            suggestions={Ce}
            hiddenCount={B.length}
            onAdd={(P) => {
              if (!c.includes(P)) {
                p([...c, P]);
              }
            }}
            onHide={pe}
            onRestoreHidden={Y}
          />
        </div>
        <div>
          <label className="block text-xs font-semibold uppercase tracking-wider text-gray-400 mb-1.5">
            Developers{" "}
            {h.length > 0 && <span className="normal-case font-normal text-gray-500">({h.length})</span>}
          </label>
          <div className="text-[11px] text-gray-500 mb-2">
            People who'll receive the dev report. Each can mark items done independently.
          </div>
          <NameChipList items={h} onRemove={(P) => b(h.filter((X, ve) => ve !== P))} />
          <div className="flex gap-2 mt-2">
            <input
              type="text"
              value={I}
              onChange={(P) => x(P.target.value)}
              onKeyDown={(P) => {
                if (P.key === "Enter") {
                  P.preventDefault();
                  H();
                }
              }}
              placeholder="Developer name"
              className="flex-1 px-3 py-1.5 bg-gray-800 border border-gray-700 text-gray-100 placeholder-gray-500 rounded-sm text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500 focus:border-transparent"
            />
            <button
              onClick={H}
              disabled={!I.trim()}
              className="px-3 py-1.5 bg-gray-800 text-gray-100 text-sm font-medium rounded-sm hover:bg-gray-700 disabled:opacity-50 disabled:cursor-not-allowed"
            >
              Add
            </button>
          </div>
          <NameSuggestions
            suggestions={be}
            hiddenCount={he.length}
            onAdd={(P) => {
              if (!h.includes(P)) {
                b([...h, P]);
              }
            }}
            onHide={me}
            onRestoreHidden={fe}
          />
        </div>
        <div className="grid grid-cols-2 gap-4">
          <CodeField label="Tester access code" value={g} onChange={v} onRegen={() => v(randomCode())} />
          <CodeField label="Developer access code" value={k} onChange={C} onRegen={() => C(randomCode())} />
        </div>
        <SampleFilesPicker files={sampleFiles} setFiles={setSampleFiles} />
        <div>
          <label className="block text-xs font-semibold uppercase tracking-wider text-gray-400 mb-1.5">
            Path chip prefix <span className="normal-case font-normal text-gray-500">(optional)</span>
          </label>
          <input
            type="text"
            value={m}
            onChange={(P) => f(P.target.value)}
            placeholder="https://localhost:3000"
            autoCapitalize="off"
            autoComplete="off"
            spellCheck="false"
            className={`w-full px-3 py-1.5 text-sm font-mono bg-gray-800 border rounded-sm text-gray-100 placeholder-gray-500 focus:outline-none focus:ring-2 focus:ring-emerald-500 focus:border-transparent ${m ? "border-emerald-800" : "border-gray-700"}`}
          />
          <div className="text-[11px] text-gray-500 mt-1.5 leading-relaxed">
            {m ? (
              <Fragment>
                Path chips like <span className="font-mono text-gray-400">/escalations</span> will become
                clickable links to{" "}
                <span className="text-emerald-400 font-mono break-all">
                  {m.trim().replace(/\/+$/, "")}/escalations
                </span>
                . Testers and developers can override this in their own Settings.
              </Fragment>
            ) : (
              <Fragment>
                Leave blank to fall back to{" "}
                <span className="font-mono text-gray-500">https://localhost:3000</span>. Path chips like{" "}
                <span className="font-mono text-gray-400">/escalations</span> become clickable links to that
                prefix. Testers and developers can override in their own Settings.
              </Fragment>
            )}
          </div>
        </div>
        {A && (
          <div className="p-3 bg-red-950/40 border border-red-800 rounded-sm text-sm text-red-300">
            <strong>Create failed:</strong> {A}
          </div>
        )}
        <div className="flex items-center gap-3 pt-2">
          <button
            onClick={te}
            disabled={!K}
            className="px-5 py-2.5 bg-emerald-600 text-white text-sm font-bold rounded-sm hover:bg-emerald-500 disabled:bg-gray-800 disabled:text-gray-500 disabled:cursor-not-allowed"
          >
            {w ? "Creating session…" : "Create session"}
          </button>
          <button onClick={e} className="px-4 py-2.5 text-sm text-gray-400 hover:text-gray-200">
            Cancel
          </button>
        </div>
      </main>
    </div>
  );
}
function NameChipList({ items: e, onRemove: t }) {
  if (e.length === 0) {
    return null;
  } else {
    return (
      <div className="flex flex-wrap gap-1.5">
        {e.map((a, r) => (
          <span
            className="inline-flex items-center gap-1.5 px-2.5 py-1 bg-gray-800 border border-gray-700 rounded-sm text-xs text-gray-200"
            key={r}
          >
            {a}
            <button
              onClick={() => t(r)}
              className="text-gray-500 hover:text-red-400 leading-none"
              aria-label={`Remove ${a}`}
            >
              ×
            </button>
          </span>
        ))}
      </div>
    );
  }
}
function NameSuggestions({ suggestions: e, hiddenCount: t, onAdd: a, onHide: r, onRestoreHidden: o }) {
  if (e.length === 0 && t === 0) {
    return null;
  } else {
    return (
      <div className="mt-3">
        <div className="flex items-center justify-between mb-1.5">
          <div className="text-[10px] uppercase tracking-wider text-gray-500 font-semibold">
            From previous sessions
          </div>
          {t > 0 && (
            <button
              onClick={o}
              className="text-[10px] text-gray-500 hover:text-gray-300"
              title="Bring back names you previously hid"
            >
              Restore {t} hidden
            </button>
          )}
        </div>
        {e.length > 0 ? (
          <div className="flex flex-wrap gap-1.5">
            {e.map((s) => (
              <div
                className="inline-flex items-stretch bg-gray-900 border border-gray-700 rounded-sm overflow-hidden"
                key={s}
              >
                <button
                  onClick={() => a(s)}
                  className="px-2 py-1 text-xs text-gray-300 hover:text-emerald-300 hover:bg-gray-800 inline-flex items-center gap-1"
                  title={`Add ${s}`}
                >
                  <Plus className="w-3 h-3" />
                  {s}
                </button>
                <button
                  onClick={() => r(s)}
                  className="px-1.5 text-gray-600 hover:text-red-400 hover:bg-gray-800 border-l border-gray-700 inline-flex items-center"
                  title="Hide this suggestion (won't show again on new sessions)"
                  aria-label={`Hide ${s} from suggestions`}
                >
                  <X className="w-3 h-3" />
                </button>
              </div>
            ))}
          </div>
        ) : (
          <div className="text-[11px] text-gray-600 italic">
            All previous names are already added or hidden.
          </div>
        )}
      </div>
    );
  }
}
function CodeField({ label: e, value: t, onChange: a, onRegen: r }) {
  return (
    <div>
      <label className="block text-xs font-semibold uppercase tracking-wider text-gray-400 mb-1.5">{e}</label>
      <div className="flex gap-2">
        <input
          type="text"
          inputMode="numeric"
          pattern={"\\d{4}"}
          maxLength={4}
          value={t}
          onChange={(o) => a(o.target.value.replace(/\D/g, "").slice(0, 4))}
          className="flex-1 px-3 py-2 bg-gray-800 border border-gray-700 text-gray-100 font-mono text-lg tracking-widest text-center rounded-sm focus:outline-none focus:ring-2 focus:ring-emerald-500 focus:border-transparent"
        />
        <button
          onClick={r}
          className="px-2.5 bg-gray-800 text-gray-400 hover:text-gray-100 rounded-sm"
          title="Generate new code"
        >
          <RotateCcw className="w-4 h-4" />
        </button>
      </div>
    </div>
  );
}
function SessionCreatedScreen({ session: e, onDone: t }) {
  let a = sessionUrl(e.id);
  return (
    <div className="min-h-screen bg-gradient-to-br from-gray-950 via-gray-900 to-indigo-950">
      <header className="px-8 py-5 border-b border-gray-800 bg-gray-950/60 backdrop-blur">
        <div className="max-w-3xl mx-auto flex items-center gap-3">
          <button
            onClick={t}
            className="p-1.5 hover:bg-gray-800 rounded-sm text-gray-400 hover:text-gray-200"
          >
            <ArrowLeft className="w-4 h-4" />
          </button>
          <div className="font-bold text-gray-100">{e.name}</div>
        </div>
      </header>
      <main className="max-w-3xl mx-auto px-8 py-10">
        <div className="mb-8 flex items-center gap-3">
          <div className="w-10 h-10 bg-emerald-900/50 border border-emerald-700 rounded-sm flex items-center justify-center">
            <Check className="w-5 h-5 text-emerald-400" />
          </div>
          <div>
            <h2 className="text-lg font-bold text-gray-100">Session ready</h2>
            <p className="text-xs text-gray-500">
              One link covers everyone — recipients pick their role on arrival and enter the matching code
              below.
            </p>
          </div>
        </div>
        <ShareLinksPanel
          url={a}
          testerCode={e.tester_code}
          devCode={e.developer_code}
          testers={e.testers || []}
          developers={e.developers || []}
        />
        <details className="mt-3 group">
          <summary className="text-xs text-gray-500 hover:text-gray-300 cursor-pointer select-none flex items-center gap-1">
            <ChevronRight className="w-3 h-3 transition-transform group-open:rotate-90" />
            Direct role links (skip the role picker)
          </summary>
          <div className="mt-3 space-y-2">
            <CopyUrlRow label="Tester link" url={portalUrl("tester", e.id)} />
            <CopyUrlRow label="Developer link" url={portalUrl("dev", e.id)} />
          </div>
        </details>
        <div className="mt-8">
          <button
            onClick={t}
            className="px-4 py-2 bg-gray-800 text-gray-100 text-sm font-medium rounded-sm hover:bg-gray-700"
          >
            Back to sessions
          </button>
        </div>
      </main>
    </div>
  );
}
function ShareLinksPanel({ url: e, testerCode: t, devCode: a, testers: r, developers: o }) {
  let [s, l] = useState(null);
  let i = async (u, d) => {
    try {
      await navigator.clipboard.writeText(u);
      l(d);
      setTimeout(() => l(null), 1500);
    } catch {}
  };
  return (
    <div className="p-5 bg-gray-900/60 border border-gray-800 rounded-sm">
      <div className="text-xs font-semibold uppercase tracking-wider text-emerald-400 mb-3">
        Shareable link
      </div>
      <div className="flex gap-2 mb-4">
        <input
          readOnly={true}
          value={e}
          className="flex-1 px-2 py-1.5 bg-gray-950 border border-gray-800 text-gray-300 text-xs rounded-sm font-mono"
        />
        <button
          onClick={() => i(e, "url")}
          className="px-2.5 py-1.5 bg-gray-800 text-gray-100 text-xs font-medium rounded-sm hover:bg-gray-700 flex items-center gap-1.5"
        >
          <Copy className="w-3 h-3" /> {s === "url" ? "Copied!" : "Copy"}
        </button>
      </div>
      <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
        <AccessCodeCard
          label="Tester code"
          code={t}
          people={r}
          copied={s === "testerCode"}
          onCopy={() => i(t, "testerCode")}
        />
        <AccessCodeCard
          label="Developer code"
          code={a}
          people={o}
          copied={s === "devCode"}
          onCopy={() => i(a, "devCode")}
        />
      </div>
    </div>
  );
}
function AccessCodeCard({ label: e, code: t, people: a, copied: r, onCopy: o }) {
  return (
    <div className="p-3 bg-gray-950 border border-gray-800 rounded-sm">
      <div className="flex items-center justify-between mb-2">
        <div className="text-[10px] uppercase tracking-wider text-gray-400 font-semibold">{e}</div>
        <div className="text-[10px] text-gray-500">
          {a.length} {a.length === 1 ? "person" : "people"}
        </div>
      </div>
      <div className="flex items-center gap-2 mb-2">
        <div className="px-3 py-1.5 bg-gray-900 border border-gray-800 text-gray-100 font-mono text-lg tracking-widest rounded-sm flex-1 text-center">
          {t}
        </div>
        <button
          onClick={o}
          className="px-2 py-1.5 bg-gray-800 text-gray-100 text-xs font-medium rounded-sm hover:bg-gray-700 flex items-center gap-1"
          title="Copy code"
        >
          <Copy className="w-3 h-3" /> {r ? "Copied!" : "Copy"}
        </button>
      </div>
      {a.length > 0 ? (
        <div className="flex flex-wrap gap-1">
          {a.map((s, l) => (
            <span
              className="text-[10px] px-1.5 py-0.5 bg-gray-900 border border-gray-800 rounded-sm text-gray-400"
              key={l}
            >
              {s}
            </span>
          ))}
        </div>
      ) : (
        <div className="text-[10px] text-gray-600 italic">No names assigned</div>
      )}
    </div>
  );
}
function CopyUrlRow({ label: e, url: t }) {
  let [a, r] = useState(false);
  return (
    <div className="flex items-center gap-2 text-xs">
      <span className="text-[10px] uppercase tracking-wider text-gray-500 font-semibold w-24 flex-shrink-0">
        {e}
      </span>
      <input
        readOnly={true}
        value={t}
        className="flex-1 px-2 py-1 bg-gray-950 border border-gray-800 text-gray-400 rounded-sm font-mono"
      />
      <button
        onClick={async () => {
          try {
            await navigator.clipboard.writeText(t);
            r(true);
            setTimeout(() => r(false), 1500);
          } catch {}
        }}
        className="px-2 py-1 bg-gray-800 text-gray-300 rounded-sm hover:bg-gray-700 flex items-center gap-1"
      >
        <Copy className="w-3 h-3" /> {a ? "Copied" : "Copy"}
      </button>
    </div>
  );
}
function useIsMobile(e = 768) {
  let [t, a] = useState(() => typeof window !== "undefined" && window.innerWidth < e);
  useEffect(() => {
    let r = () => a(window.innerWidth < e);
    window.addEventListener("resize", r);
    return () => window.removeEventListener("resize", r);
  }, [e]);
  return t;
}
function useBaseUrlState(e, t) {
  let a = `uat-pro-baseurl-${e}`;
  let r = "https://localhost:3000";
  let o = () => (t && t.trim()) || r;
  let s = () => {
    try {
      return localStorage.getItem(a);
    } catch {
      return null;
    }
  };
  let [l, i] = useState(s);
  useEffect(() => {
    i(s());
  }, [a]);
  return [
    l !== null ? l : o(),
    (p) => {
      let h = (p || "").trim();
      i(h);
      try {
        localStorage.setItem(a, h);
      } catch {}
    },
    l !== null,
    () => {
      i(null);
      try {
        localStorage.removeItem(a);
      } catch {}
    },
  ];
}
function useSidebarArrowNav(e, t, a) {
  useEffect(() => {
    let r = (o) => {
      if (o.key !== "ArrowUp" && o.key !== "ArrowDown") {
        return;
      }
      let s = o.target?.tagName;
      if (
        s === "INPUT" ||
        s === "TEXTAREA" ||
        o.target?.isContentEditable ||
        document.querySelector(".fixed.inset-0.z-50") ||
        e.length === 0
      ) {
        return;
      }
      o.preventDefault();
      let l = e.findIndex((u) => u.id === t);
      let i;
      if (l === -1) {
        i = o.key === "ArrowDown" ? 0 : e.length - 1;
      } else if (o.key === "ArrowDown") {
        i = Math.min(l + 1, e.length - 1);
      } else {
        i = Math.max(l - 1, 0);
      }
      if (i !== l) {
        a(e[i].id);
      }
    };
    window.addEventListener("keydown", r);
    return () => window.removeEventListener("keydown", r);
  }, [e, t, a]);
  useEffect(() => {
    if (!t) {
      return;
    }
    let r = document.querySelector(`[data-tc-id="${t}"]`);
    if (r && r.scrollIntoView) {
      r.scrollIntoView({
        block: "nearest",
      });
    }
  }, [t]);
}
function DesktopOnlyNotice({ onContinueAnyway: e }) {
  return (
    <div className="min-h-screen flex items-center justify-center bg-gradient-to-br from-gray-950 via-gray-900 to-emerald-950 p-6">
      <div className="max-w-sm w-full bg-gray-900/60 border border-gray-800 rounded-sm p-7 text-center">
        <div className="w-12 h-12 mx-auto mb-3 bg-emerald-900/40 border border-emerald-800 rounded-sm flex items-center justify-center">
          <Monitor className="w-6 h-6 text-emerald-300" />
        </div>
        <div className="text-xs uppercase tracking-widest text-emerald-400 mb-2">Test Case Collaborator</div>
        <h2 className="text-lg font-bold text-gray-100 mb-2">Best on desktop</h2>
        <p className="text-sm text-gray-400 leading-relaxed">
          This portal is built for larger screens — running through steps, reviewing screenshots, and
          switching between cases needs the sidebar visible alongside the detail pane.
        </p>
        <p className="text-xs text-gray-500 mt-4">
          Open this same link on a laptop or desktop browser to continue.
        </p>
        <button
          onClick={e}
          className="mt-5 text-[11px] text-gray-600 hover:text-gray-400 underline-offset-2 hover:underline"
        >
          Continue anyway
        </button>
      </div>
    </div>
  );
}
function PortalPicker({ sessionId: e }) {
  let [t, a] = useState("loading");
  let [r, o] = useState(null);
  let [s, l] = useState(null);
  let i = useIsMobile();
  let [u, d] = useState(false);
  useEffect(() => {
    if (!e) {
      a("notfound");
      return;
    }
    let p = false;
    (async () => {
      let { data: h, error: b } = await supabase
        .from("sessions")
        .select("id, name")
        .eq("id", e)
        .maybeSingle();
      if (!p) {
        if (b || !h) {
          a("notfound");
          return;
        }
        o(h);
        a("choose");
      }
    })();
    return () => {
      p = true;
    };
  }, [e]);
  let c = (p) => {
    window.history.replaceState(null, "", portalUrl(p, e));
    l(p);
  };
  if (s === "tester") {
    return <TesterEntry sessionId={e} />;
  } else if (s === "dev") {
    return <DevEntry sessionId={e} />;
  } else if (i && !u) {
    return <DesktopOnlyNotice onContinueAnyway={() => d(true)} />;
  } else if (t === "loading") {
    return <Notice>Loading session…</Notice>;
  } else if (t === "notfound") {
    return <Notice variant="error">Session not found. Check the link with the admin.</Notice>;
  } else {
    return (
      <div className="min-h-screen flex items-center justify-center bg-gradient-to-br from-gray-950 via-gray-900 to-emerald-950 p-6">
        <div className="w-full max-w-md">
          <div className="text-center mb-6">
            <div className="text-xs uppercase tracking-widest text-emerald-400 mb-1">UAT session</div>
            <h1 className="text-2xl font-bold text-gray-100">{r.name}</h1>
          </div>
          <div className="bg-gray-900/70 border border-gray-800 rounded-sm p-5">
            <div className="text-sm text-gray-300 mb-1">Pick your role to continue</div>
            <div className="text-xs text-gray-500 mb-4">You'll be asked for an access code next.</div>
            <div className="grid grid-cols-2 gap-3">
              <button
                onClick={() => c("tester")}
                className="flex flex-col items-center gap-2 p-5 bg-gray-800/60 border border-gray-700 rounded-sm hover:bg-emerald-950/40 hover:border-emerald-700 transition-colors group"
              >
                <Users className="w-7 h-7 text-gray-400 group-hover:text-emerald-300" />
                <div className="text-sm font-medium text-gray-100 group-hover:text-emerald-200">
                  I'm a tester
                </div>
                <div className="text-[10px] text-gray-500 leading-tight">
                  Run the cases and report findings
                </div>
              </button>
              <button
                onClick={() => c("dev")}
                className="flex flex-col items-center gap-2 p-5 bg-gray-800/60 border border-gray-700 rounded-sm hover:bg-emerald-950/40 hover:border-emerald-700 transition-colors group"
              >
                <GitBranch className="w-7 h-7 text-gray-400 group-hover:text-emerald-300" />
                <div className="text-sm font-medium text-gray-100 group-hover:text-emerald-200">
                  I'm a developer
                </div>
                <div className="text-[10px] text-gray-500 leading-tight">Review findings, mark fixes</div>
              </button>
            </div>
          </div>
          <div className="text-center mt-4 text-[10px] text-gray-600">
            Click <span className="text-gray-400">Switch</span> inside a portal to return here.
          </div>
        </div>
      </div>
    );
  }
}
function TesterEntry({ sessionId: e }) {
  let [t, a] = useState("loading");
  let [r, o] = useState(null);
  let [s, l] = useState(null);
  let [i, u] = useState("");
  let d = useIsMobile();
  let [c, p] = useState(false);
  useEffect(() => {
    if (!e) {
      a("notfound");
      return;
    }
    let I = false;
    (async () => {
      let { data: x, error: g } = await supabase
        .from("sessions")
        .select("id, name, testers, tester_code, tester_col_header, base_url")
        .eq("id", e)
        .single();
      if (!I) {
        if (g || !x) {
          a("notfound");
          return;
        }
        o(x);
        try {
          let v = JSON.parse(localStorage.getItem(`uat-pro-tester-${e}`) || "{}");
          if (v.code === x.tester_code && (x.testers || []).includes(v.name)) {
            l(v.name);
            a("ready");
            return;
          }
        } catch {}
        a("code");
      }
    })();
    return () => {
      I = true;
    };
  }, [e]);
  let h = (I) => {
    if (I === r.tester_code) {
      u("");
      a("name");
    } else {
      u("Wrong code. Double-check what your admin sent.");
    }
  };
  let b = (I) => {
    try {
      localStorage.setItem(
        `uat-pro-tester-${e}`,
        JSON.stringify({
          code: r.tester_code,
          name: I,
        }),
      );
    } catch {}
    l(I);
    a("ready");
  };
  let _ = () => {
    try {
      localStorage.removeItem(`uat-pro-tester-${e}`);
    } catch {}
    navigateTo(sessionUrl(e));
  };
  if (d && !c) {
    return <DesktopOnlyNotice onContinueAnyway={() => p(true)} />;
  }
  if (t === "loading") {
    return <Notice>Loading session…</Notice>;
  }
  if (t === "notfound") {
    return (
      <Notice variant="error">Session not found. The link may be wrong, or the session was deleted.</Notice>
    );
  }
  let S = () => navigateTo(sessionUrl(e));
  if (t === "code") {
    return (
      <AccessCodeGate label="Tester portal" sessionName={r.name} onSubmit={h} error={i} onBackToPicker={S} />
    );
  } else if (t === "name") {
    return (
      <NamePicker
        label="Tester portal"
        sessionName={r.name}
        names={r.testers || []}
        onSelect={b}
        onBackToPicker={S}
      />
    );
  } else {
    return <TesterPortal session={r} myName={s} onLogout={_} />;
  }
}
function Notice({ children: e, variant: t }) {
  let a = t === "error";
  return (
    <div className="min-h-screen flex items-center justify-center bg-gradient-to-br from-gray-950 via-gray-900 to-emerald-950 p-6">
      <div
        className={`max-w-md text-center p-6 rounded-sm border flex flex-col items-center gap-3 ${a ? "border-red-800 bg-red-950/30 text-red-300" : "border-gray-800 bg-gray-900/60 text-gray-300"}`}
      >
        {!a && (
          <div className="w-7 h-7 rounded-full border-2 border-gray-800 border-t-emerald-500 animate-spin" />
        )}
        <div>{e}</div>
      </div>
    </div>
  );
}
function AccessCodeGate({ label: e, sessionName: t, onSubmit: a, error: r, onBackToPicker: o }) {
  let [s, l] = useState("");
  return (
    <div className="min-h-screen flex items-center justify-center bg-gradient-to-br from-gray-950 via-gray-900 to-emerald-950 p-6">
      <div className="max-w-sm w-full bg-gray-900/60 border border-gray-800 rounded-sm p-7">
        {o && (
          <button
            onClick={o}
            className="inline-flex items-center gap-1 text-[11px] text-gray-500 hover:text-gray-300 mb-3"
          >
            <ArrowLeft className="w-3 h-3" /> Pick a different role
          </button>
        )}
        <div className="text-xs uppercase tracking-wider text-emerald-400 font-semibold mb-1">{e}</div>
        <h2 className="text-lg font-bold text-gray-100 mb-1">{t}</h2>
        <p className="text-xs text-gray-500 mb-5">Enter the 4-digit code your admin shared.</p>
        <form
          onSubmit={(i) => {
            i.preventDefault();
            a(s);
          }}
        >
          <input
            autoFocus={true}
            type="text"
            inputMode="numeric"
            pattern={"\\d{4}"}
            maxLength={4}
            value={s}
            onChange={(i) => l(i.target.value.replace(/\D/g, "").slice(0, 4))}
            placeholder="0000"
            className="w-full px-3 py-3 bg-gray-800 border border-gray-700 text-gray-100 font-mono text-2xl tracking-widest text-center rounded-sm focus:outline-none focus:ring-2 focus:ring-emerald-500"
          />
          {r && <div className="mt-2 text-xs text-red-400">{r}</div>}
          <button
            type="submit"
            disabled={s.length !== 4}
            className="mt-4 w-full px-4 py-2.5 bg-emerald-600 text-white text-sm font-bold rounded-sm hover:bg-emerald-500 disabled:bg-gray-800 disabled:text-gray-500"
          >
            Continue
          </button>
        </form>
      </div>
    </div>
  );
}
function NamePicker({ label: e, sessionName: t, names: a, onSelect: r, onBackToPicker: o }) {
  let [s, l] = useState(null);
  useEffect(() => {
    if (!s) {
      return;
    }
    let i = (u) => {
      if (u.key === "Escape") {
        l(null);
      }
    };
    window.addEventListener("keydown", i);
    return () => window.removeEventListener("keydown", i);
  }, [s]);
  return (
    <div className="min-h-screen flex items-center justify-center bg-gradient-to-br from-gray-950 via-gray-900 to-emerald-950 p-6">
      <div className="max-w-sm w-full bg-gray-900/60 border border-gray-800 rounded-sm p-7">
        {!s && o && (
          <button
            onClick={o}
            className="inline-flex items-center gap-1 text-[11px] text-gray-500 hover:text-gray-300 mb-3"
          >
            <ArrowLeft className="w-3 h-3" /> Pick a different role
          </button>
        )}
        <div className="text-xs uppercase tracking-wider text-emerald-400 font-semibold mb-1">{e}</div>
        <h2 className="text-lg font-bold text-gray-100 mb-1">{t}</h2>
        {s ? (
          <Fragment>
            <p className="text-xs text-gray-500 mb-5">Confirm you are</p>
            <div className="px-4 py-3 mb-5 bg-emerald-950/40 border border-emerald-800 rounded-sm text-center">
              <div className="text-lg font-semibold text-emerald-100">{s}</div>
            </div>
            <div className="grid grid-cols-2 gap-2">
              <button
                onClick={() => l(null)}
                className="px-4 py-2 bg-gray-800 hover:bg-gray-700 border border-gray-700 text-gray-200 text-sm font-medium rounded-sm"
              >
                Back
              </button>
              <button
                onClick={() => r(s)}
                autoFocus={true}
                className="px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white text-sm font-medium rounded-sm"
              >
                Yes
              </button>
            </div>
          </Fragment>
        ) : (
          <Fragment>
            <p className="text-xs text-gray-500 mb-5">Who are you?</p>
            <div className="space-y-1.5">
              {a.length === 0 ? (
                <div className="text-sm text-gray-400 text-center py-4">
                  No names on the roster yet. Ask your admin to add you.
                </div>
              ) : (
                a.map((i) => (
                  <button
                    onClick={() => l(i)}
                    className="w-full px-4 py-2.5 bg-gray-800 hover:bg-emerald-900/40 border border-gray-700 hover:border-emerald-700 text-gray-100 text-left rounded-sm transition-colors"
                    key={i}
                  >
                    {i}
                  </button>
                ))
              )}
            </div>
          </Fragment>
        )}
      </div>
    </div>
  );
}
var TESTER_STATUSES = [
  {
    value: "Not Tested",
    label: "Not run",
    dot: "bg-gray-500",
    badge: "bg-gray-700 text-gray-300 border-gray-600",
  },
  {
    value: "Pass",
    label: "Pass",
    dot: "bg-green-500",
    badge: "bg-green-900/40 text-green-300 border-green-700",
  },
  {
    value: "Fail",
    label: "Fail",
    dot: "bg-red-500",
    badge: "bg-red-900/40 text-red-300 border-red-700",
  },
  {
    value: "Blocked",
    label: "Blocked",
    dot: "bg-orange-500",
    badge: "bg-orange-900/40 text-orange-300 border-orange-700",
  },
  {
    value: "Needs Review",
    label: "Needs Review",
    dot: "bg-amber-500",
    badge: "bg-amber-900/40 text-amber-300 border-amber-700",
  },
];
var PRIORITIES = ["P0", "P1", "P2"];
// Bug severity (stored in findings.bug_severity). Older findings saved P0/P1/P2 are read as Blocker/Major/Minor.
var SEVERITIES = [
  {
    value: "Blocker",
    description: "Can't continue or wrong data.",
    badge: "bg-red-900/40 text-red-300 border-red-700",
  },
  {
    value: "Major",
    description: "Works badly or confusingly.",
    badge: "bg-orange-900/40 text-orange-300 border-orange-700",
  },
  {
    value: "Minor",
    description: "Noticeable rough edge.",
    badge: "bg-amber-900/40 text-amber-300 border-amber-700",
  },
  {
    value: "Cosmetic",
    description: "Visual detail only.",
    badge: "bg-sky-900/40 text-sky-300 border-sky-700",
  },
];
var LEGACY_SEVERITY = { P0: "Blocker", P1: "Major", P2: "Minor" };
var getSeverity = (e) => {
  if (!e) return null;
  let t = LEGACY_SEVERITY[e] || e;
  return SEVERITIES.find((a) => a.value.toLowerCase() === String(t).toLowerCase()) || null;
};
var severityRank = (e) => {
  let t = getSeverity(e);
  return t ? SEVERITIES.indexOf(t) : 99;
};
function SeverityBadge({ value: e, size: t = "sm" }) {
  let a = getSeverity(e);
  if (!a) return null;
  return (
    <span
      className={`inline-flex items-center border rounded-sm font-bold flex-shrink-0 ${t === "xs" ? "text-[9px] px-1 tracking-wider" : "text-xs px-2.5 py-1"} ${a.badge}`}
      title={`Severity: ${a.value} — ${a.description}`}
    >
      {a.value}
    </span>
  );
}
var PRO_CSV_PROMPT = `You are generating a User Acceptance Testing (UAT) test plan as a CSV that will be uploaded into a UAT collaboration tool.

Fill these placeholders with your project context first, then proceed:
- Product or feature scope: [describe the product or what's being tested]
- Modules to cover: [list modules, e.g. "Escalations, Leads, Projects"]
- Testers (2\u20135 names): [e.g. "Alice, Bob, Carol"]
- Target test count: [approximate, e.g. "around 50"]

Output a CSV with exactly these 15 columns in this exact order and header text:

1. ID \u2014 short identifier like T1-E01, T1-L02. A project prefix + a short area code + a number that increments within each module.
2. Module \u2014 top-level area being tested (must match one of the modules above).
3. Feature \u2014 sub-area within the module (e.g. "Create ticket", "Status lifecycle").
4. Title \u2014 concise description. Include the URL/route when the test starts from a specific page, e.g. "Create escalation via /escalations/new". Leading-slash paths render as clickable links in the tool.
5. Tester [1 - <name>, 2 - <name>, ...] \u2014 header in this EXACT format using the tester names above. Cell value is a single digit (1, 2, 3...) assigning that tester. Distribute work roughly evenly.
6. Pre-conditions \u2014 what state must exist before this test runs, or "\u2014" if none.
7. Test Data \u2014 concrete values formatted as a key: value list separated by newlines, e.g. "Subject: Foo\\nPriority: Medium". Use "\u2014" if none.
8. Steps \u2014 numbered list, one step per line: "1. Do X.\\n2. Do Y.\\n3. Verify Z."
9. Expected Result \u2014 what should happen when the feature works correctly.
10. Downstream Check \u2014 secondary verification (related page, related record), or "\u2014" if none.
11. Edge Case \u2014 edge case to flag during this test, or "\u2014" if none.
12. Pass/Fail \u2014 leave EMPTY. Testers fill this in.
13. Tester Notes \u2014 leave EMPTY.
14. Bug Severity (P0/ P1 / P2) \u2014 leave EMPTY.
15. Other Observations \u2014 leave EMPTY.

Output format rules:
- One header row, then one row per test.
- Quote any cell that contains a comma, newline, or double quote.
- Inside quoted cells use REAL newlines for line breaks (not the literal text "\\n").
- Use the em dash "\u2014" for cells where no value applies.
- Cover happy paths first, then error paths, then edge cases.
- Aim for atomic checks \u2014 one row = one verification.
- Reference real URL paths/routes from the product where applicable.
- Output the CSV only \u2014 no markdown code fences, no commentary before or after.

Reference row (shows the expected shape \u2014 replace with your own content):
T1-E01,Escalations,Create ticket,Create escalation via /escalations/new,1,On /escalations page,"Subject: Test escalation
Priority: Medium
Category: Timeline","1. /escalations \u2192 'New Ticket'.
2. Fill required fields.
3. Save.","Ticket created with auto-generated ticket_id (T-YYYY-NNNNNN). Toast confirms.","Verify the ticket_id format follows T-2026-NNNNNN.","\u2014",,,,`;
var REGRESSION_STATUSES = ["Fail", "Blocked", "Needs Review"];
var priorityBadgeClass = (e) => {
  switch (e) {
    case "P0":
      return "bg-red-900/40 text-red-300 border-red-700";
    case "P1":
      return "bg-orange-900/40 text-orange-300 border-orange-700";
    case "P2":
      return "bg-yellow-900/40 text-yellow-300 border-yellow-700";
    default:
      return "bg-gray-900 text-gray-300 border-gray-700";
  }
};
var MAX_UPLOAD_BYTES = 10485760;
var getTesterStatus = (e) => TESTER_STATUSES.find((t) => t.value === e) || TESTER_STATUSES[0];
function TesterPortal({ session: e, myName: t, onLogout: a }) {
  let [r, o] = useState([]);
  let [s, l] = useState({});
  let [i, u] = useState({});
  let [d, c] = useState(true);
  let [p, h] = useState("");
  let [b, _, S, I] = useBaseUrlState(e.id, e.base_url);
  let [x, g] = useState(null);
  let [v, k] = useState("");
  let [C, m] = useState(false);
  let [f, w] = useState("All");
  let [T, A] = useState("mine");
  let [j, M] = useState(true);
  let [W, V] = useState(null);
  let [F, se] = useState(false);
  let [filterPriority, setFilterPriority] = useState("All");
  let [filterFocusArea, setFilterFocusArea] = useState("All");
  let E = `uat-pro-tester-focus-${e.id}-${t}`;
  useEffect(() => {
    try {
      let N = localStorage.getItem(E);
      if (N) {
        let G = JSON.parse(N);
        if (Array.isArray(G) && G.length > 0) {
          V(new Set(G));
        }
      }
    } catch {}
  }, [E]);
  useEffect(() => {
    try {
      if (W && W.size > 0) {
        localStorage.setItem(E, JSON.stringify([...W]));
      } else {
        localStorage.removeItem(E);
      }
    } catch {}
  }, [W, E]);
  let D = useRef(new Set());
  let O = useRef({});
  let B = useRef(s);
  useEffect(() => {
    B.current = s;
  }, [s]);
  let [ne, he] = useState(() => new Set());
  let [ie, ue] = useState(() => new Set());
  let [pe, me] = useState({});
  let Y = (N) =>
    me((G) => {
      if (!G[N]) {
        return G;
      }
      let Q = {
        ...G,
      };
      delete Q[N];
      return Q;
    });
  useEffect(() => {
    let N = false;
    (async () => {
      c(true);
      let [G, Q, z] = await Promise.all([
        supabase.from("test_cases").select("*").eq("session_id", e.id).order("position"),
        supabase.from("findings").select("*").eq("session_id", e.id),
        supabase.from("done_marks").select("*").eq("session_id", e.id),
      ]);
      if (N) {
        return;
      }
      if (G.error || Q.error || z.error) {
        h((G.error || Q.error || z.error).message);
        c(false);
        return;
      }
      o((G.data || []).filter((Le) => !Le.removed_at));
      let re = {};
      (Q.data || []).forEach((Le) => {
        re[Le.test_case_id] = Le;
      });
      l(re);
      let ce = {};
      (z.data || []).forEach((Le) => {
        ce[Le.test_case_id] ||= [];
        ce[Le.test_case_id].push({
          developer_name: Le.developer_name,
          done_at: Le.done_at,
          status: Le.status || "fixed",
          note: Le.note || "",
        });
      });
      u(ce);
      c(false);
      if ((G.data || []).length > 0) {
        g(G.data[0].id);
      }
    })();
    return () => {
      N = true;
    };
  }, [e.id]);
  useEffect(() => {
    let N = supabase
      .channel(`findings-${e.id}`)
      .on(
        "postgres_changes",
        {
          event: "*",
          schema: "public",
          table: "findings",
          filter: `session_id=eq.${e.id}`,
        },
        (G) => {
          let Q = G.new || G.old;
          if (!!Q && !!Q.test_case_id) {
            if (G.eventType === "DELETE") {
              l((z) => {
                let re = {
                  ...z,
                };
                delete re[Q.test_case_id];
                return re;
              });
              return;
            }
            if (!D.current.has(Q.test_case_id)) {
              l((z) => ({
                ...z,
                [Q.test_case_id]: G.new,
              }));
            }
          }
        },
      )
      .subscribe();
    return () => {
      supabase.removeChannel(N);
    };
  }, [e.id]);
  useEffect(() => {
    let N = supabase
      .channel(`tester-done-${e.id}`)
      .on(
        "postgres_changes",
        {
          event: "*",
          schema: "public",
          table: "done_marks",
          filter: `session_id=eq.${e.id}`,
        },
        (G) => {
          if (G.eventType === "INSERT" || G.eventType === "UPDATE") {
            let Q = G.new;
            u((z) => {
              let re = (z[Q.test_case_id] || []).filter((ce) => ce.developer_name !== Q.developer_name);
              return {
                ...z,
                [Q.test_case_id]: [
                  ...re,
                  {
                    developer_name: Q.developer_name,
                    done_at: Q.done_at,
                    status: Q.status || "fixed",
                    note: Q.note || "",
                  },
                ],
              };
            });
          } else if (G.eventType === "DELETE") {
            let Q = G.old;
            u((z) => {
              let re = z[Q.test_case_id] || [];
              return {
                ...z,
                [Q.test_case_id]: re.filter((ce) => ce.developer_name !== Q.developer_name),
              };
            });
          }
        },
      )
      .subscribe();
    return () => {
      supabase.removeChannel(N);
    };
  }, [e.id]);
  let fe = (N, G) => {
    if (r.find((z) => z.id === N)) {
      l((z) => {
        let ce = {
          ...(z[N] || {
            test_case_id: N,
            session_id: e.id,
            status: "Not Tested",
            tested_by: t,
            notes: "",
            other_observations: "",
            bug_severity: "",
            step_checks: {},
            step_notes: {},
            recordings: [],
          }),
          ...G,
        };
        ce.tested_by ||= t;
        ce.updated_by = t;
        ce.updated_at = new Date().toISOString();
        return {
          ...z,
          [N]: ce,
        };
      });
      if (
        G.status &&
        REGRESSION_STATUSES.includes(G.status) &&
        (i[N] || []).filter((re) => re.status === "fixed").length > 0
      ) {
        u((re) => {
          let ce = (re[N] || []).filter((We) => We.status !== "fixed");
          let Le = {
            ...re,
          };
          if (ce.length > 0) {
            Le[N] = ce;
          } else {
            delete Le[N];
          }
          return Le;
        });
        (async () => {
          let { error: re } = await supabase
            .from("done_marks")
            .delete()
            .eq("test_case_id", N)
            .eq("status", "fixed");
          if (re) {
            console.error("Failed to clear fixed marks after rejection:", re);
          }
        })();
      }
      D.current.add(N);
      he((z) => {
        if (z.has(N)) {
          return z;
        }
        let re = new Set(z);
        re.add(N);
        return re;
      });
      clearTimeout(O.current[N]);
      O.current[N] = setTimeout(async () => {
        let z = B.current[N];
        if (!z) {
          return;
        }
        let re = {
          test_case_id: N,
          session_id: e.id,
          status: z.status || "Not Tested",
          tested_by: z.tested_by || t,
          notes: z.notes || "",
          other_observations: z.other_observations || "",
          bug_severity: z.bug_severity || "",
          step_checks: z.step_checks || {},
          step_notes: z.step_notes || {},
          recordings: z.recordings || [],
          screenshots: z.screenshots || [],
          updated_by: t,
          updated_at: new Date().toISOString(),
        };
        let { error: ce } = await supabase.from("findings").upsert(re, {
          onConflict: "test_case_id",
        });
        D.current.delete(N);
        he((Le) => {
          if (!Le.has(N)) {
            return Le;
          }
          let We = new Set(Le);
          We.delete(N);
          return We;
        });
        if (ce) {
          console.error("Save failed:", ce);
        }
      }, 600);
    }
  };
  let Ce = async (N, G) => {
    let Q = Array.from(G).filter((q) => q.type && q.type.startsWith("image/"));
    if (Q.length === 0) {
      return;
    }
    let z = Q.filter((q) => q.size > MAX_UPLOAD_BYTES);
    let re = Q.filter((q) => q.size <= MAX_UPLOAD_BYTES);
    if (z.length > 0) {
      let q = z.map(
        (oe) =>
          `${oe.name || "screenshot"} (${(oe.size / 1024 / 1024).toFixed(1)} MB) \u2014 over the ${MAX_UPLOAD_BYTES / 1024 / 1024} MB limit`,
      );
      me((oe) => ({
        ...oe,
        [N]: q,
      }));
      setTimeout(() => Y(N), 8000);
    }
    if (re.length === 0) {
      return;
    }
    ue((q) => {
      if (q.has(N)) {
        return q;
      }
      let oe = new Set(q);
      oe.add(N);
      return oe;
    });
    let ce = await Promise.all(
      re.map(async (q) => {
        let oe = (q.name || "screenshot").replace(/[^a-zA-Z0-9._-]/g, "_").slice(0, 80);
        let $t = `${e.id}/${N}/${crypto.randomUUID()}-${oe}`;
        let { error: Qe } = await supabase.storage.from("screenshots").upload($t, q, {
          contentType: q.type,
          upsert: false,
        });
        if (Qe) {
          console.error("Screenshot upload failed:", Qe);
          return {
            failed: true,
            msg: `${q.name || "screenshot"} \u2014 ${Qe.message || "upload failed"}`,
          };
        }
        let {
          data: { publicUrl: ya },
        } = supabase.storage.from("screenshots").getPublicUrl($t);
        return {
          path: $t,
          url: ya,
          name: q.name || "screenshot",
          size: q.size || 0,
          uploaded_by: t,
          uploaded_at: new Date().toISOString(),
        };
      }),
    );
    let Le = ce.filter((q) => q && q.failed).map((q) => q.msg);
    let We = ce.filter((q) => q && !q.failed);
    ue((q) => {
      if (!q.has(N)) {
        return q;
      }
      let oe = new Set(q);
      oe.delete(N);
      return oe;
    });
    if (Le.length > 0) {
      me((q) => ({
        ...q,
        [N]: [...(q[N] || []), ...Le],
      }));
      setTimeout(() => Y(N), 8000);
    }
    if (We.length === 0) {
      return;
    }
    let St = B.current[N]?.screenshots || [];
    fe(N, {
      screenshots: [...St, ...We],
    });
  };
  let be = async (N, G) => {
    let Q = B.current[N]?.screenshots || [];
    let z = Q[G];
    let fnd = B.current[N] || {};
    let sn = {};
    Object.entries(fnd.step_notes || {}).forEach(([re, ce]) => {
      sn[re] = renumberFigRefs(ce, G + 1);
    });
    if (
      z &&
      (fe(N, {
        screenshots: Q.filter((re, ce) => ce !== G),
        notes: renumberFigRefs(fnd.notes, G + 1),
        other_observations: renumberFigRefs(fnd.other_observations, G + 1),
        step_notes: sn,
      }),
      z.path)
    ) {
      let { error: re } = await supabase.storage.from("screenshots").remove([z.path]);
      if (re) {
        console.error("Screenshot storage delete failed:", re);
      }
    }
  };
  let xe = useRef(x);
  useEffect(() => {
    xe.current = x;
  }, [x]);
  let $ = useRef(Ce);
  $.current = Ce;
  useEffect(() => {
    let N = (G) => {
      let Q = xe.current;
      if (!Q) {
        return;
      }
      let z = G.clipboardData?.items;
      if (!z) {
        return;
      }
      let re = [];
      for (let ce of z) {
        if (ce.type && ce.type.startsWith("image/")) {
          let Le = ce.getAsFile();
          if (Le) {
            re.push(Le);
          }
        }
      }
      if (re.length > 0) {
        G.preventDefault();
        $.current(Q, re);
      }
    };
    window.addEventListener("paste", N);
    return () => window.removeEventListener("paste", N);
  }, []);
  let H = r.filter((N) => {
    if ((W && !W.has(N.id)) || (T === "mine" && N.initial_tester && N.initial_tester !== t)) {
      return false;
    }
    let Q = s[N.id]?.status || "Not Tested";
    if (f !== "All" && Q !== f) {
      return false;
    }
    if (filterPriority !== "All" && (N.priority || "") !== filterPriority) {
      return false;
    }
    if (filterFocusArea !== "All" && (N.focus || "") !== filterFocusArea) {
      return false;
    }
    if (!v) {
      return true;
    }
    let z = v.toLowerCase();
    return (
      (N.test_id || "").toLowerCase().includes(z) ||
      (N.title || "").toLowerCase().includes(z) ||
      (N.module || "").toLowerCase().includes(z) ||
      (N.feature || "").toLowerCase().includes(z) ||
      (N.focus || "").toLowerCase().includes(z) ||
      (N.priority || "").toLowerCase().includes(z)
    );
  });
  let K = new Set();
  let te = {};
  r.forEach((N) => {
    let G = i[N.id] || [];
    let Q = G.filter((re) => re.status === "fixed");
    if (Q.length > 0) {
      let re = Math.max(...Q.map((We) => new Date(We.done_at).getTime()));
      let ce = s[N.id];
      let Le = ce?.updated_at ? new Date(ce.updated_at).getTime() : 0;
      if (re > Le) {
        K.add(N.id);
      }
    }
    let z = G.filter((re) => re.status === "in_progress").map((re) => re.developer_name);
    if (z.length > 0) {
      te[N.id] = z;
    }
  });
  let P = r.find((N) => N.id === x);
  useSidebarArrowNav(H, x, g);
  let X = P ? s[P.id] : null;
  let ve = r.filter((N) => {
    let G = s[N.id]?.status;
    return G && G !== "Not Tested";
  }).length;
  if (d) {
    return <Notice>Loading test cases…</Notice>;
  }
  if (p) {
    return <Notice variant="error">Could not load session: {p}</Notice>;
  }
  let de = async () => {
    let N = r.map((Q) => Q.id);
    if (N.length === 0) {
      return;
    }
    let G = [];
    for (let Q of N) {
      let z = s[Q];
      if (z?.updated_by === t && Array.isArray(z.screenshots)) {
        for (let re of z.screenshots) {
          if (re?.path) {
            G.push(re.path);
          }
        }
      }
    }
    try {
      let { error: Q } = await supabase.from("findings").delete().in("test_case_id", N).eq("updated_by", t);
      if (Q) {
        throw Q;
      }
      if (G.length > 0) {
        try {
          await supabase.storage.from("screenshots").remove(G);
        } catch (z) {
          console.warn("Some screenshot files could not be removed:", z);
        }
      }
      l((z) => {
        let re = {
          ...z,
        };
        for (let ce of Object.keys(re)) {
          if (re[ce]?.updated_by === t) {
            delete re[ce];
          }
        }
        return re;
      });
    } catch (Q) {
      console.error("Reset failed:", Q);
      alert("Could not reset your inputs. Check your connection and try again.");
    }
  };
  return (
    <BaseUrlContext.Provider value={b}>
      <div className="h-screen flex bg-gray-950 overflow-hidden text-gray-100">
        <TesterSidebar
          open={j}
          onClose={() => M(false)}
          sessionName={e.name}
          myName={t}
          onLogout={a}
          onReset={de}
          testCases={r}
          findings={s}
          completed={ve}
          visibleCases={H}
          reviewFixIds={K}
          wipByTest={te}
          selectedId={x}
          onSelect={g}
          search={v}
          setSearch={k}
          searchOpen={C}
          setSearchOpen={m}
          filterStatus={f}
          setFilterStatus={w}
          filterPriority={filterPriority}
          setFilterPriority={setFilterPriority}
          filterFocusArea={filterFocusArea}
          setFilterFocusArea={setFilterFocusArea}
          scope={T}
          setScope={A}
          focusedIds={W}
          onClearFocus={() => V(null)}
          onOpenFocusModal={() => se(true)}
          baseUrl={b}
          setBaseUrl={_}
          hasBaseUrlOverride={S}
          resetBaseUrl={I}
        />
        <div className="flex-1 min-w-0 flex flex-col">
          {!j && (
            <button
              onClick={() => M(true)}
              className="absolute top-3 left-3 z-20 p-2 bg-gray-900 border border-gray-800 rounded-sm hover:bg-gray-800"
            >
              <PanelLeftOpen className="w-4 h-4 text-gray-400" />
            </button>
          )}
          {P ? (
            <TesterTestDetail
              testCase={P}
              finding={X}
              devMarks={i[P.id] || []}
              isSaving={ne.has(P.id)}
              isUploading={ie.has(P.id)}
              screenshotErrors={pe[P.id] || null}
              onChange={(N) => fe(P.id, N)}
              onAddScreenshots={(N) => Ce(P.id, N)}
              onRemoveScreenshot={(N) => be(P.id, N)}
              onDismissScreenshotError={() => Y(P.id)}
            />
          ) : (
            <div className="flex-1 flex items-center justify-center text-gray-500 text-sm">
              No test case selected.
            </div>
          )}
        </div>
        <FocusModeModal
          open={F}
          onClose={() => se(false)}
          testCases={r.map((N) => ({
            __id: N.id,
            __rowNum: N.row_num,
            ID: N.test_id,
            Title: N.title,
            Module: N.module,
            Feature: N.feature,
          }))}
          currentFocusedIds={W}
          onApply={(N) => V(N)}
        />
        <TesterSampleFilesBar sessionId={e.id} sessionName={e.name} myName={t} />
      </div>
    </BaseUrlContext.Provider>
  );
}
function TesterSidebar({
  open: e,
  onClose: t,
  sessionName: a,
  myName: r,
  onLogout: o,
  onReset: s,
  testCases: l,
  findings: i,
  completed: u,
  visibleCases: d,
  reviewFixIds: c,
  wipByTest: p,
  selectedId: h,
  onSelect: b,
  search: _,
  setSearch: S,
  searchOpen: I,
  setSearchOpen: x,
  filterStatus: g,
  setFilterStatus: v,
  filterPriority: fp = "All",
  setFilterPriority: setFp = () => {},
  filterFocusArea: ff = "All",
  setFilterFocusArea: setFf = () => {},
  scope: k,
  setScope: C,
  focusedIds: m,
  onClearFocus: f,
  onOpenFocusModal: w,
  baseUrl: T,
  setBaseUrl: A,
  hasBaseUrlOverride: j,
  resetBaseUrl: M,
}) {
  let [W, V] = useState(false);
  let [F, se] = useState(false);
  let [E, D] = useState(false);
  let priorityOptions = sortPriorities(Array.from(new Set(l.map((O) => O.priority).filter(Boolean))));
  let focusOptions = Array.from(new Set(l.map((O) => O.focus).filter(Boolean))).sort((O, B) => O.localeCompare(B));
  let hasNoPriority = priorityOptions.length > 0 && l.some((O) => !O.priority);
  let hasNoFocus = focusOptions.length > 0 && l.some((O) => !O.focus);
  if (e) {
    return (
      <aside className="w-80 flex-shrink-0 bg-gray-900 border-r border-gray-800 flex flex-col overflow-hidden">
        <div className="px-4 py-3 border-b border-gray-800">
          <div className="flex items-center justify-between gap-2 mb-2">
            <div className="flex items-center gap-2 min-w-0">
              <button onClick={t} className="p-1 hover:bg-gray-800 rounded-sm flex-shrink-0" title="Collapse">
                <PanelLeftClose className="w-4 h-4 text-gray-400" />
              </button>
              <h2 className="font-bold text-gray-100 truncate">{a}</h2>
            </div>
            <button
              onClick={() => D(true)}
              className="p-1 hover:bg-gray-800 rounded-sm flex-shrink-0"
              title="Settings"
            >
              <Settings2 className="w-4 h-4 text-gray-400" />
            </button>
          </div>
          <div className="text-xs text-gray-400 mb-2">
            Signed in as <span className="text-gray-100 font-medium">{r}</span> <span className="text-gray-500">(Tester)</span>
          </div>
          <div className="text-xs text-gray-500 mb-1">
            {u} / {l.length} done
          </div>
          <div className="w-full bg-gray-800 rounded-full h-1 mb-3">
            <div
              className="bg-emerald-500 h-1 rounded-full transition-all"
              style={{
                width: `${l.length ? (u / l.length) * 100 : 0}%`,
              }}
            />
          </div>
          <div className="inline-flex w-full border border-gray-700 rounded-sm overflow-hidden mb-2">
            <button
              onClick={() => C("mine")}
              className={`flex-1 px-2 py-1 text-xs transition-colors ${k === "mine" ? "bg-emerald-600 text-white" : "bg-gray-900 text-gray-400 hover:text-gray-200 hover:bg-gray-800"}`}
            >
              My tests{" "}
              <span className="tabular-nums opacity-60">
                ({l.filter((O) => !O.initial_tester || O.initial_tester === r).length})
              </span>
            </button>
            <button
              onClick={() => C("all")}
              className={`flex-1 px-2 py-1 text-xs transition-colors border-l border-gray-700 ${k === "all" ? "bg-emerald-600 text-white" : "bg-gray-900 text-gray-400 hover:text-gray-200 hover:bg-gray-800"}`}
            >
              All tests <span className="tabular-nums opacity-60">({l.length})</span>
            </button>
          </div>
          {I ? (
            <SidebarSearch value={_} onChange={S} onClose={() => x(false)} className="mb-2" />
          ) : (
          <div className="flex gap-1.5 mb-2">
            <button
              onClick={() => {
                let O = !I;
                x(O);
                if (!O) {
                  S("");
                }
              }}
              className={`flex-shrink-0 flex items-center justify-center border rounded-sm transition-colors w-8 h-8 ${I || _ ? "bg-emerald-600 border-emerald-600 text-white hover:bg-emerald-500" : "bg-gray-800 border-gray-700 text-gray-400 hover:text-emerald-300 hover:border-emerald-700 hover:bg-gray-700"}`}
              title="Search test cases"
            >
              <Search className="w-3.5 h-3.5" />
            </button>
            <button
              onClick={w}
              className={`flex-shrink-0 flex items-center justify-center border rounded-sm transition-colors w-8 h-8 ${m ? "bg-emerald-600 border-emerald-600 text-white hover:bg-emerald-500" : "bg-gray-800 border-gray-700 text-gray-400 hover:text-emerald-300 hover:border-emerald-700 hover:bg-gray-700"}`}
              title={m ? `Focus filter active (${m.size} tests)` : "Focus on a subset of tests"}
            >
              <Filter className="w-3.5 h-3.5" />
            </button>
            <select
              value={g}
              onChange={(O) => v(O.target.value)}
              className="flex-1 min-w-0 px-2 py-1.5 text-xs bg-gray-800 border border-gray-700 text-gray-100 rounded-sm focus:outline-none focus:ring-1 focus:ring-emerald-500 appearance-none"
              style={{
                backgroundImage: `url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 12 12'><path d='M3 4.5l3 3 3-3' stroke='%239ca3af' stroke-width='1.5' fill='none' stroke-linecap='round' stroke-linejoin='round'/></svg>")`,
                backgroundRepeat: "no-repeat",
                backgroundPosition: "right 8px center",
                paddingRight: "28px",
              }}
            >
              <option value="All">All statuses</option>
              {TESTER_STATUSES.map((O) => (
                <option value={O.value} key={O.value}>
                  {O.label}
                </option>
              ))}
            </select>
          </div>
          )}
          {(priorityOptions.length > 0 || focusOptions.length > 0) && (
            <div className="flex gap-1.5 mb-2">
              {priorityOptions.length > 0 && (
                <FilterSelect value={fp} onChange={setFp} active={fp !== "All"}>
                  <option value="All">All priorities</option>
                  {priorityOptions.map((O) => (
                    <option value={O} key={O}>
                      {O} ({l.filter((B) => B.priority === O).length})
                    </option>
                  ))}
                  {hasNoPriority && <option value="">No priority</option>}
                </FilterSelect>
              )}
              {focusOptions.length > 0 && (
                <FilterSelect value={ff} onChange={setFf} active={ff !== "All"}>
                  <option value="All">All focus areas</option>
                  {focusOptions.map((O) => (
                    <option value={O} key={O}>
                      {O} ({l.filter((B) => B.focus === O).length})
                    </option>
                  ))}
                  {hasNoFocus && <option value="">No focus area</option>}
                </FilterSelect>
              )}
            </div>
          )}
          {m && (
            <div className="mt-2 flex items-center gap-2 px-2.5 py-1.5 bg-emerald-950/40 border border-emerald-800/50 rounded-sm text-xs">
              <Filter className="w-3 h-3 text-emerald-400 flex-shrink-0" />
              <span className="text-emerald-200">
                Focused: {m.size} of {l.length}
              </span>
              <button
                onClick={f}
                className="ml-auto text-emerald-400 hover:text-emerald-200 transition-colors"
                title="Clear focus filter"
              >
                <X className="w-3 h-3" />
              </button>
            </div>
          )}
        </div>
        <div className="flex-1 overflow-y-auto">
          {d.length === 0 ? (
            <EmptyState
              hasAnyTests={l.length > 0}
              suggestions={[
                k === "mine" &&
                  l.length > 0 && {
                    label: "Show all tests instead",
                    onClick: () => C("all"),
                  },
                m && {
                  label: "Clear focus",
                  onClick: f,
                },
                g !== "All" && {
                  label: "Clear status filter",
                  onClick: () => v("All"),
                },
                fp !== "All" && {
                  label: "Clear priority filter",
                  onClick: () => setFp("All"),
                },
                ff !== "All" && {
                  label: "Clear focus area filter",
                  onClick: () => setFf("All"),
                },
                _ && {
                  label: "Clear search",
                  onClick: () => S(""),
                },
              ]}
            />
          ) : (
            d.map((O) => {
              let B = i[O.id];
              let ne = B?.status || "Not Tested";
              let he = getTesterStatus(ne);
              let ie = O.id === h;
              let ue = c && c.has(O.id);
              let pe = (p && p[O.id]) || [];
              let me =
                pe.length === 1
                  ? `${pe[0]} Fixing`
                  : pe.length > 1
                    ? `${pe[0]} +${pe.length - 1} Fixing`
                    : null;
              let Y = (O.feature && O.feature !== "—") || ue || me;
              return (
                <button
                  data-tc-id={O.id}
                  onClick={() => b(O.id)}
                  className={`w-full text-left px-3 py-2.5 border-b border-gray-800 border-l-4 flex items-start gap-2 hover:bg-gray-800/60 ${ie ? "bg-emerald-950/40 border-l-emerald-400 shadow-[inset_6px_0_20px_-4px_rgba(16,185,129,0.4)]" : "border-l-transparent"}`}
                  key={O.id}
                >
                  <span className={`w-2 h-2 rounded-full mt-1.5 flex-shrink-0 ${he.dot}`} />
                  <div className="min-w-0 flex-1">
                    <div className="flex items-center gap-1.5 text-[10px] text-gray-500 mb-0.5">
                      <span className="font-mono text-emerald-400">{O.test_id || "—"}</span>
                      {O.priority && <PriorityBadge value={O.priority} size="xs" />}
                      {sheetUpdateUnseen(O, B) && (
                        <span
                          className="text-[9px] px-1 border rounded-sm font-bold tracking-wider bg-orange-950/50 text-orange-300 border-orange-800 flex-shrink-0"
                          title="Steps or expected results changed in the sheet since you last worked on this test"
                        >
                          UPDATED
                        </span>
                      )}
                      {O.module && <span className="truncate">· {O.module}</span>}
                      {O.focus && <span className="truncate text-gray-600">· {O.focus}</span>}
                    </div>
                    <div
                      className={`text-sm leading-tight ${ie ? "text-gray-100 font-semibold" : "text-gray-300"}`}
                    >
                      {O.title || "Untitled"}
                    </div>
                    {Y && (
                      <div className="flex items-center gap-1.5 mt-0.5 flex-wrap">
                        {ue && (
                          <span
                            className="text-[9px] uppercase tracking-wider font-bold px-1.5 py-0.5 rounded-sm bg-emerald-900/40 text-emerald-300 border border-emerald-700"
                            title="A dev marked this fixed after your last edit — re-test when you can."
                          >
                            Review Fix
                          </span>
                        )}
                        {me && (
                          <span
                            className="text-[9px] uppercase tracking-wider font-bold px-1.5 py-0.5 rounded-sm bg-amber-900/40 text-amber-300 border border-amber-700"
                            title={
                              pe.length > 1 ? `In progress: ${pe.join(", ")}` : `${pe[0]} is working on a fix`
                            }
                          >
                            {me}
                          </span>
                        )}
                        {O.feature && O.feature !== "—" && (
                          <span className="text-[10px] uppercase tracking-wider text-gray-500">
                            {O.feature}
                          </span>
                        )}
                      </div>
                    )}
                  </div>
                  {B?.tested_by && B.tested_by !== "" && (
                    <span
                      className="text-[10px] text-gray-500 mt-1 flex-shrink-0"
                      title={`Last edited by ${B.tested_by}`}
                    >
                      {B.tested_by.charAt(0)}
                    </span>
                  )}
                </button>
              );
            })
          )}
        </div>
        {E && (
          <SettingsModal
            baseUrl={T}
            setBaseUrl={A}
            hasBaseUrlOverride={j}
            resetBaseUrl={M}
            myName={r}
            roleLabel="Signed in"
            onSwitchRequested={() => V(true)}
            onResetRequested={s ? () => se(true) : undefined}
            resetDescription="Delete every status, severity, note, observation, step check, recording, and screenshot you've logged in this session. Other testers' work is not affected."
            onClose={() => D(false)}
          />
        )}
        {W && (
          <ConfirmDialog
            title="Switch user?"
            message="You'll be returned to the role picker and signed out. Your saved progress for this session isn't affected."
            confirmLabel="Switch"
            onConfirm={() => {
              V(false);
              o();
            }}
            onCancel={() => V(false)}
          />
        )}
        {F && (
          <ConfirmDialog
            title="Reset all your inputs?"
            message="Every status, severity, note, observation, step check, recording, and screenshot you've logged in this session will be permanently deleted. This action cannot be undone."
            confirmLabel="Reset everything"
            danger={true}
            onConfirm={() => {
              se(false);
              s();
            }}
            onCancel={() => se(false)}
          />
        )}
      </aside>
    );
  } else {
    return null;
  }
}
function TesterTestDetail({
  testCase: e,
  finding: t,
  devMarks: a,
  isSaving: r,
  isUploading: o,
  screenshotErrors: s,
  onChange: l,
  onAddScreenshots: i,
  onRemoveScreenshot: u,
  onDismissScreenshotError: d,
}) {
  let c = e;
  let p = t || {
    status: "Not Tested",
    tested_by: "",
    notes: "",
    other_observations: "",
    bug_severity: "",
    step_checks: {},
    step_notes: {},
    recordings: [],
  };
  let [h, b] = useState("");
  useEffect(() => {
    b("");
  }, [c.id]);
  let [_, S] = useState(null);
  useEffect(() => {
    S(null);
  }, [c.id]);
  let [I, x] = useState(null);
  useEffect(() => {
    x(null);
  }, [c.id]);
  let g = (a || [])
    .filter((f) => (f.note && f.note.trim()) || f.status)
    .sort((f, w) => new Date(f.done_at) - new Date(w.done_at));
  let v = isFailureStatus(p.status);
  let k = (f) => {
    let w = (p.step_checks || {})[f];
    let T = {
      ...(p.step_checks || {}),
    };
    let ee = String(f).startsWith("exp-");
    if (w === "done") {
      T[f] = ee ? "blocked" : "review";
    } else if (w === "review") {
      T[f] = "blocked";
    } else if (w === "blocked" && ee) {
      T[f] = "unsure";
    } else if (w === "blocked" || w === "unsure") {
      delete T[f];
    } else {
      T[f] = "done";
    }
    l({
      step_checks: T,
    });
  };
  let C = (f, w) => {
    let T = {
      ...(p.step_notes || {}),
    };
    if (w === "" || w == null) {
      delete T[f];
    } else {
      T[f] = w;
    }
    l({
      step_notes: T,
    });
  };
  let m = () => {
    let f = h.trim();
    if (f) {
      l({
        recordings: [
          ...(p.recordings || []),
          {
            url: f,
          },
        ],
      });
      b("");
    }
  };
  return (
    <FigureContext.Provider value={{ shots: p.screenshots || [], open: (f) => x(f) }}>
    <div className="flex-1 overflow-y-auto">
      <div className="max-w-6xl mx-auto px-6 py-6">
        <div className="mb-3 pb-4 border-b border-gray-800 flex items-start gap-4">
          <div className="flex-1 min-w-0">
            <Breadcrumb tc={c} />
            <h1 className="text-xl font-bold text-gray-100">{c.title || "Untitled"}</h1>
            {sheetUpdateUnseen(c, p) && (
              <div className="mt-2 inline-flex items-start gap-2 px-2.5 py-1.5 bg-orange-950/40 border border-orange-800 rounded-sm text-xs text-orange-200">
                <AlertTriangle className="w-3.5 h-3.5 mt-px flex-shrink-0 text-orange-400" />
                <span>
                  Steps updated from the sheet on {new Date(c.content_updated_at).toLocaleDateString()}. Ticks on
                  reworded lines were cleared, so please re-check them. This note goes away once you make a change.
                </span>
              </div>
            )}
          </div>
          <div className="pt-1 flex-shrink-0">
            <SaveIndicator saving={!!r} />
          </div>
        </div>
        {p.updated_by && p.updated_at && (
          <div
            className="mb-5 text-[11px] text-gray-500 flex items-center gap-1.5"
            title={new Date(p.updated_at).toLocaleString()}
          >
            <Clock className="w-3 h-3 flex-shrink-0" />
            <span>
              Last edit by <span className="text-gray-300">{p.updated_by}</span> · {timeAgo(p.updated_at)}
            </span>
          </div>
        )}
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          <div className="space-y-5">
            {c.pre_conditions && c.pre_conditions !== "—" && (
              <DetailField label="Pre-conditions" value={c.pre_conditions} />
            )}
            {c.test_data && c.test_data !== "—" && <DetailField label="Test Data" value={c.test_data} />}
            {(c.steps || []).length > 0 && (
              <div>
                <div className="text-xs font-semibold uppercase tracking-wider text-gray-400 mb-2">Steps</div>
                <div className="space-y-1.5">
                  {(c.steps || []).map((f, w) => {
                    let T = (p.step_checks || {})[w];
                    let A = (p.step_notes || {})[w];
                    return (
                      <StepRow
                        idx={w}
                        step={f}
                        state={T}
                        note={A}
                        onToggle={() => k(w)}
                        onNote={(j) => C(w, j)}
                        key={w}
                      />
                    );
                  })}
                </div>
              </div>
            )}
            {c.expected_result && c.expected_result !== "—" && (
              <ExpectedResultChecklist
                value={c.expected_result}
                checks={p.step_checks}
                onToggle={(f) => k(expectedCheckKey(f))}
                notes={p.step_notes}
                onNote={(f, w) => C(f, w)}
              />
            )}
            {c.downstream_check && c.downstream_check !== "—" && (
              <DetailField label="Downstream Check" value={c.downstream_check} />
            )}
            {c.edge_case && c.edge_case !== "—" && <DetailField label="Edge Case" value={c.edge_case} />}
          </div>
          <div className="space-y-5">
            {g.length > 0 && (
              <div>
                <div className="text-xs font-semibold uppercase tracking-wider text-gray-400 mb-1.5">
                  Developer notes <span className="normal-case font-normal text-gray-500">({g.length})</span>
                </div>
                <div className="space-y-2">
                  {g.map((f) => (
                    <DevMarkCard mark={f} key={f.developer_name} />
                  ))}
                </div>
              </div>
            )}
            <div>
              <div className="text-xs font-semibold uppercase tracking-wider text-gray-400 mb-2">Status</div>
              <div className="flex flex-wrap gap-1.5">
                {TESTER_STATUSES.map((f) => {
                  let w = (p.status || "Not Tested") === f.value;
                  return (
                    <button
                      onClick={() =>
                        !w &&
                        l({
                          status: f.value,
                        })
                      }
                      aria-pressed={w}
                      className={`px-2.5 py-1.5 text-xs font-medium border rounded-sm transition-colors ${w ? f.badge : "bg-gray-900 text-gray-400 border-gray-800 hover:text-gray-200 hover:border-gray-700"}`}
                      key={f.value}
                    >
                      {f.label}
                    </button>
                  );
                })}
              </div>
            </div>
            <div>
              <div className="text-xs font-semibold uppercase tracking-wider text-gray-400 mb-2">Severity</div>
              <div className="flex flex-wrap gap-1.5">
                {SEVERITIES.map((f) => {
                  let w = getSeverity(p.bug_severity)?.value === f.value;
                  return (
                    <button
                      onClick={() =>
                        l({
                          bug_severity: w ? "" : f.value,
                        })
                      }
                      aria-pressed={w}
                      title={w ? "Click to clear" : f.description}
                      className={`px-2.5 py-1.5 text-xs font-medium border rounded-sm transition-colors ${w ? f.badge : "bg-gray-900 text-gray-400 border-gray-800 hover:text-gray-200 hover:border-gray-700"}`}
                      key={f.value}
                    >
                      {f.value}
                    </button>
                  );
                })}
              </div>
              {getSeverity(p.bug_severity) && (
                <div className="mt-2 text-xs text-gray-400">
                  <span className="text-gray-200 font-medium">{getSeverity(p.bug_severity).value}</span> ={" "}
                  {getSeverity(p.bug_severity).description.replace(/^./, (f) => f.toLowerCase())}
                </div>
              )}
            </div>
            <div>
              <div className="text-xs font-semibold uppercase tracking-wider text-gray-400 mb-1.5">
                Tester Notes{" "}
                <span className="normal-case tracking-normal font-normal text-gray-500">· type @ to link a screenshot</span>
              </div>
              <MentionTextarea
                value={p.notes || ""}
                onChange={(f) =>
                  l({
                    notes: prettyArrows(f.target.value),
                  })
                }
                rows={4}
                placeholder="Observations, summary of issues found, deviations from expected result…"
                className="w-full px-3 py-2 bg-gray-800 border border-gray-700 text-gray-100 placeholder-gray-500 rounded-sm text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500 resize-y"
              />
            </div>
            <div>
              <div className="text-xs font-semibold uppercase tracking-wider text-gray-400 mb-1.5">
                Other Observations
              </div>
              <MentionTextarea
                value={p.other_observations || ""}
                onChange={(f) =>
                  l({
                    other_observations: prettyArrows(f.target.value),
                  })
                }
                rows={3}
                placeholder="Anything else worth flagging — UX, performance, side effects…"
                className="w-full px-3 py-2 bg-gray-800 border border-gray-700 text-gray-100 placeholder-gray-500 rounded-sm text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500 resize-y"
              />
            </div>
            <div>
              <div className="text-xs font-semibold uppercase tracking-wider text-gray-400 mb-1.5">
                Recordings{" "}
                {(p.recordings || []).length > 0 && (
                  <span className="normal-case font-normal text-gray-500">
                    ({(p.recordings || []).length})
                  </span>
                )}
              </div>
              <div className="flex gap-2">
                <input
                  type="url"
                  value={h}
                  onChange={(f) => b(f.target.value)}
                  onKeyDown={(f) => {
                    if (f.key === "Enter") {
                      f.preventDefault();
                      m();
                    }
                  }}
                  placeholder="Paste a Google Drive share link…"
                  className="flex-1 px-3 py-2 bg-gray-800 border border-gray-700 text-gray-100 placeholder-gray-500 rounded-sm text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500"
                />
                <button
                  onClick={m}
                  disabled={!h.trim()}
                  className="px-3 py-2 bg-emerald-600 text-white text-sm font-medium rounded-sm hover:bg-emerald-500 disabled:bg-gray-800 disabled:text-gray-500"
                >
                  Add
                </button>
              </div>
              <div className="mt-1.5 text-[11px] text-gray-500 leading-relaxed">
                Upload your recording to Google Drive, set sharing to{" "}
                <span className="text-gray-300 font-medium">"Anyone with the link"</span>, then paste the
                share URL here.<span className="text-gray-600"> Loom / YouTube / Vimeo URLs also work.</span>
              </div>
              {(p.recordings || []).length > 0 && (
                <div className="mt-2 space-y-1.5">
                  {(p.recordings || []).map((f, w) => {
                    let A = parseRecordingUrl(f.url)?.label || "Link";
                    return (
                      <div
                        className="flex items-center gap-2 px-2.5 py-1.5 bg-gray-800/60 border border-gray-700 rounded-sm group"
                        key={w}
                      >
                        <span className="text-[10px] font-semibold uppercase tracking-wider px-1.5 py-0.5 bg-gray-900 text-emerald-300 border border-gray-700 rounded-sm flex-shrink-0">
                          {A}
                        </span>
                        <button
                          onClick={() =>
                            S({
                              recordings: p.recordings,
                              startIdx: w,
                            })
                          }
                          className="text-xs text-gray-300 hover:text-emerald-300 truncate flex-1 min-w-0 text-left"
                          title="Play in app"
                        >
                          {f.url}
                        </button>
                        <a
                          href={f.url}
                          target="_blank"
                          rel="noopener noreferrer"
                          className="p-1 text-gray-500 hover:text-emerald-400 flex-shrink-0"
                          title="Open in new tab"
                        >
                          <ExternalLink className="w-3 h-3" />
                        </a>
                        <button
                          onClick={() =>
                            l({
                              recordings: (p.recordings || []).filter((j, M) => M !== w),
                            })
                          }
                          className="opacity-40 group-hover:opacity-100 hover:text-red-400 text-gray-400 text-sm leading-none px-1"
                          title="Remove"
                        >
                          ×
                        </button>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
            <div>
              <div className="text-xs font-semibold uppercase tracking-wider text-gray-400 mb-1.5">
                Screenshots
                {(p.screenshots || []).length > 0 && (
                  <span className="normal-case font-normal text-gray-500">
                    {" "}
                    ({(p.screenshots || []).length})
                  </span>
                )}
              </div>
              <ScreenshotDropzone disabled={!!o} isUploading={!!o} onFiles={(f) => i && i(f)} />
              {s && s.length > 0 && (
                <div className="mt-2 p-2.5 bg-red-950/40 border border-red-800 rounded-sm flex items-start gap-2">
                  <CircleAlert className="w-4 h-4 text-red-400 flex-shrink-0 mt-0.5" />
                  <div className="flex-1 min-w-0 text-xs text-red-200 space-y-0.5">
                    <div className="font-semibold text-red-100">
                      Couldn't upload {s.length === 1 ? "this file" : `${s.length} files`}
                    </div>
                    {s.map((f, w) => (
                      <div className="opacity-90 break-words" key={w}>
                        {f}
                      </div>
                    ))}
                  </div>
                  <button
                    onClick={() => d && d()}
                    className="text-red-400 hover:text-red-200 flex-shrink-0"
                    title="Dismiss"
                  >
                    <X className="w-3.5 h-3.5" />
                  </button>
                </div>
              )}
              {(p.screenshots || []).length > 0 && (
                <div className="grid grid-cols-2 gap-2 mt-2">
                  {(p.screenshots || []).map((f, w) => (
                    <ScreenshotThumb
                      shot={f}
                      idx={w}
                      onRemove={() => u && u(w)}
                      onOpen={(T) => x(T)}
                      key={f.path || f.url || w}
                    />
                  ))}
                </div>
              )}
            </div>
          </div>
        </div>
      </div>
      {_ && <RecordingPlayer recordings={_.recordings} startIdx={_.startIdx} onClose={() => S(null)} />}
      {I !== null && (p.screenshots || []).length > 0 && (
        <ScreenshotLightbox
          screenshots={p.screenshots || []}
          startIdx={I}
          onClose={() => x(null)}
          editable={true}
          onAnnotationsChange={(f, w) => {
            let A = (p.screenshots || []).map((j, M) =>
              M === f
                ? {
                    ...j,
                    annotations: w,
                  }
                : j,
            );
            l({
              screenshots: A,
            });
          }}
        />
      )}
    </div>
    </FigureContext.Provider>
  );
}
// Expected results written as "✓ ..." lines become a checklist. Each item's tick is saved in
// the finding's step_checks under keys "exp-0", "exp-1"… (no database change needed).
var CHECK_LINE = /^\s*(?:✓|✔|☑|☐|\[\s?[xX✓]?\s?\])\s*/;
var parseExpectedChecklist = (e) => {
  if (!e || typeof e != "string") return null;
  let t = e.split(/\n/);
  if (!t.some((o) => CHECK_LINE.test(o))) return null;
  let a = [];
  let r = [];
  for (let o of t) {
    if (CHECK_LINE.test(o)) {
      a.push(o.replace(CHECK_LINE, "").trim());
    } else if (o.trim()) {
      if (a.length > 0) a[a.length - 1] += `\n${o.trim()}`;
      else r.push(o.trim());
    }
  }
  return { intro: r.join("\n"), items: a.filter(Boolean) };
};
var expectedCheckKey = (e) => `exp-${e}`;
// Per-item note under an expected-result checklist item (saved in step_notes under "exp-<n>").
function ExpectedItemNote({ note: e, onNote: t }) {
  let [a, r] = useState(false);
  if (!t) {
    return e ? (
      <div className="ml-[26px] mt-1 mb-0.5 px-2 py-1 bg-gray-950/70 border-l-2 border-amber-600 text-xs text-gray-200 whitespace-pre-wrap rounded-sm">
        {normalizePath(e)}
      </div>
    ) : null;
  }
  return a || e ? (
    <div className="ml-[26px] mt-1 mb-0.5">
      <MentionTextarea
        autoFocus={a}
        value={e || ""}
        onChange={(o) => t(prettyArrows(o.target.value))}
        onBlur={() => r(false)}
        rows={2}
        placeholder="Note for this item… (type @ to link a screenshot)"
        className="w-full px-2 py-1.5 bg-gray-950 border border-gray-700 text-gray-200 placeholder-gray-500 rounded-sm text-base sm:text-xs focus:outline-none focus:ring-1 focus:ring-emerald-500 resize-y"
      />
    </div>
  ) : (
    <button
      type="button"
      onClick={() => r(true)}
      className="ml-[26px] text-[11px] text-gray-500 hover:text-gray-300"
    >
      + Add note
    </button>
  );
}
function ExpectedResultChecklist({ value: e, checks: t, onToggle: a, notes: nt, onNote: on }) {
  let r = parseExpectedChecklist(e);
  if (!r) {
    return <DetailField label="Expected Result" value={e} highlight={true} />;
  }
  let o = t || {};
  let l = r.items.filter((i, u) => o[expectedCheckKey(u)] === "done").length;
  let i = r.items.filter((u, d) => o[expectedCheckKey(d)] === "blocked").length;
  let n = r.items.filter((u, d) => o[expectedCheckKey(d)] === "unsure").length;
  return (
    <div>
      <div className="flex items-center gap-2 text-xs font-semibold uppercase tracking-wider text-gray-400 mb-1">
        Expected Result
        <span className="normal-case tracking-normal font-normal text-gray-500">
          {l} of {r.items.length} met{i > 0 ? ` · ${i} not met` : ""}
          {n > 0 ? ` · ${n} unsure` : ""}
        </span>
      </div>
      <div className="bg-amber-950/30 border-l-4 border-amber-500 rounded-sm py-1.5">
        {r.intro && <div className="px-3 pb-1 text-sm text-amber-100 whitespace-pre-wrap">{normalizePath(r.intro)}</div>}
        {r.items.map((u, d) => {
          let c = o[expectedCheckKey(d)];
          let p = c === "done";
          let h = c === "blocked";
          let q = c === "unsure";
          let b = p ? "Mark as not met" : h ? "Mark as unsure" : q ? "Clear" : "Mark as met";
          let _ = (
            <span
              className={`w-4 h-4 mt-0.5 rounded-sm flex items-center justify-center flex-shrink-0 transition-colors ${p ? "bg-green-600 border-2 border-green-600" : h ? "bg-red-600 border-2 border-red-600" : q ? "bg-orange-500 border-2 border-orange-500" : `border-2 border-amber-700/70 ${a ? "group-hover:border-amber-500" : ""}`}`}
            >
              {p && <Check className="w-2.5 h-2.5 text-white" strokeWidth={3} />}
              {h && <X className="w-2.5 h-2.5 text-white" strokeWidth={3} />}
              {q && <span className="text-[10px] leading-none font-black text-white">?</span>}
            </span>
          );
          let S = (
            <span className={`flex-1 min-w-0 text-sm whitespace-pre-wrap ${p ? "text-green-200" : h ? "text-red-200" : q ? "text-orange-200" : "text-amber-100"}`}>
              {normalizePath(u)}
              {h && (
                <span className="ml-2 text-[10px] uppercase tracking-wider text-red-400 font-bold">Not as expected</span>
              )}
              {q && <span className="ml-2 text-[10px] uppercase tracking-wider text-orange-400 font-bold">Unsure</span>}
            </span>
          );
          let m = (nt || {})[expectedCheckKey(d)];
          return (
            <div className="px-3 py-1" key={d}>
              {a ? (
                <button
                  type="button"
                  onClick={() => a(d)}
                  title={b}
                  className="group w-full flex items-start gap-2.5 text-left rounded-sm transition-colors hover:bg-amber-900/20"
                >
                  {_}
                  {S}
                </button>
              ) : (
                <div className="flex items-start gap-2.5">
                  {_}
                  {S}
                </div>
              )}
              <ExpectedItemNote note={m} onNote={on ? (w) => on(expectedCheckKey(d), w) : null} />
            </div>
          );
        })}
      </div>
    </div>
  );
}
function DetailField({ label: e, value: t, highlight: a }) {
  return (
    <div>
      <div className="text-xs font-semibold uppercase tracking-wider text-gray-400 mb-1">{e}</div>
      <div
        className={`text-sm whitespace-pre-wrap ${a ? "bg-amber-950/30 border-l-4 border-amber-500 px-3 py-2 rounded-sm text-amber-100" : "text-gray-200"}`}
      >
        {normalizePath(t)}
      </div>
    </div>
  );
}
function StepRow({ idx: e, step: t, state: a, note: r, onToggle: o, onNote: s }) {
  let [l, i] = useState(false);
  let u = a === "done";
  let d = a === "blocked";
  let v = a === "review";
  let c = u ? "Mark as works but needs review" : v ? "Mark as unable to proceed" : d ? "Clear check" : "Mark as done";
  return (
    <div
      className={`p-2.5 border rounded-sm ${u ? "bg-green-900/15 border-green-800" : v ? "bg-orange-900/15 border-orange-800" : d ? "bg-red-900/15 border-red-800" : "bg-gray-900 border-gray-800"}`}
    >
      <div className="flex items-start gap-2.5">
        <div className="text-[11px] font-bold text-emerald-400 mt-1 font-mono w-5 flex-shrink-0 text-right">
          {e + 1}.
        </div>
        <div className="flex-1 min-w-0">
          <div className="text-sm text-gray-200">
            {normalizePath(t)}
            {d && (
              <span className="ml-2 text-[10px] uppercase tracking-wider text-red-400 font-bold">
                Unable to proceed
              </span>
            )}
            {v && (
              <span className="ml-2 text-[10px] uppercase tracking-wider text-orange-400 font-bold">
                Works but needs review
              </span>
            )}
          </div>
          {l || r ? (
            <MentionTextarea
              autoFocus={l}
              value={r || ""}
              onChange={(p) => s(prettyArrows(p.target.value))}
              onBlur={() => i(false)}
              rows={2}
              placeholder="Note for this step… (type @ to link a screenshot)"
              className="mt-2 w-full px-2 py-1.5 bg-gray-950 border border-gray-700 text-gray-200 placeholder-gray-500 rounded-sm text-xs focus:outline-none focus:ring-1 focus:ring-emerald-500 resize-y"
            />
          ) : (
            <button onClick={() => i(true)} className="mt-1 text-[11px] text-gray-500 hover:text-gray-300">
              + Add note
            </button>
          )}
        </div>
        <div className="relative group flex-shrink-0">
          <button
            onClick={o}
            className={`w-5 h-5 mt-0.5 rounded-sm border-2 flex items-center justify-center transition-colors ${u ? "bg-green-600 border-green-600" : v ? "bg-orange-500 border-orange-500" : d ? "bg-red-600 border-red-600" : "border-gray-600 hover:border-gray-500"}`}
          >
            {(u || v) && <Check className="w-3 h-3 text-white" strokeWidth={3} />}
            {d && <X className="w-3 h-3 text-white" strokeWidth={3} />}
          </button>
          <span className="hidden group-hover:block absolute z-50 right-full top-1/2 -translate-y-1/2 mr-2 px-2 py-1 bg-gray-950 border border-gray-700 rounded-sm text-[11px] text-gray-100 whitespace-nowrap shadow-lg pointer-events-none">
            {c}
          </span>
        </div>
      </div>
    </div>
  );
}
function DevEntry({ sessionId: e }) {
  let [t, a] = useState("loading");
  let [r, o] = useState(null);
  let [s, l] = useState(null);
  let [i, u] = useState("");
  let d = useIsMobile();
  let [c, p] = useState(false);
  useEffect(() => {
    if (!e) {
      a("notfound");
      return;
    }
    let I = false;
    (async () => {
      let { data: x, error: g } = await supabase
        .from("sessions")
        .select("id, name, developers, developer_code, base_url")
        .eq("id", e)
        .single();
      if (!I) {
        if (g || !x) {
          a("notfound");
          return;
        }
        o(x);
        try {
          let v = JSON.parse(localStorage.getItem(`uat-pro-dev-${e}`) || "{}");
          if (v.code === x.developer_code && (x.developers || []).includes(v.name)) {
            l(v.name);
            a("ready");
            return;
          }
        } catch {}
        a("code");
      }
    })();
    return () => {
      I = true;
    };
  }, [e]);
  let h = (I) => {
    if (I === r.developer_code) {
      u("");
      a("name");
    } else {
      u("Wrong code. Double-check what your admin sent.");
    }
  };
  let b = (I) => {
    try {
      localStorage.setItem(
        `uat-pro-dev-${e}`,
        JSON.stringify({
          code: r.developer_code,
          name: I,
        }),
      );
    } catch {}
    l(I);
    a("ready");
  };
  let _ = () => {
    try {
      localStorage.removeItem(`uat-pro-dev-${e}`);
    } catch {}
    navigateTo(sessionUrl(e));
  };
  if (d && !c) {
    return <DesktopOnlyNotice onContinueAnyway={() => p(true)} />;
  }
  if (t === "loading") {
    return <Notice>Loading session…</Notice>;
  }
  if (t === "notfound") {
    return (
      <Notice variant="error">Session not found. The link may be wrong, or the session was deleted.</Notice>
    );
  }
  let S = () => navigateTo(sessionUrl(e));
  if (t === "code") {
    return (
      <AccessCodeGate
        label="Developer portal"
        sessionName={r.name}
        onSubmit={h}
        error={i}
        onBackToPicker={S}
      />
    );
  } else if (t === "name") {
    return (
      <NamePicker
        label="Developer portal"
        sessionName={r.name}
        names={r.developers || []}
        onSelect={b}
        onBackToPicker={S}
      />
    );
  } else {
    return <DevPortal session={r} myName={s} onLogout={_} />;
  }
}
function DevPortal({ session: e, myName: t, onLogout: a }) {
  let [r, o] = useState([]);
  let [s, l] = useState({});
  let [i, u] = useState({});
  let [d, c] = useState(true);
  let [p, h] = useState("");
  let [b, _, S, I] = useBaseUrlState(e.id, e.base_url);
  let [x, g] = useState(null);
  let [v, k] = useState("");
  let [C, m] = useState(false);
  let [f, w] = useState("all");
  let [T, A] = useState(true);
  let [j, M] = useState(null);
  let [W, V] = useState(false);
  let F = `uat-pro-dev-focus-${e.id}-${t}`;
  useEffect(() => {
    try {
      let $ = localStorage.getItem(F);
      if ($) {
        let H = JSON.parse($);
        if (Array.isArray(H) && H.length > 0) {
          M(new Set(H));
        }
      }
    } catch {}
  }, [F]);
  useEffect(() => {
    try {
      if (j && j.size > 0) {
        localStorage.setItem(F, JSON.stringify([...j]));
      } else {
        localStorage.removeItem(F);
      }
    } catch {}
  }, [j, F]);
  useEffect(() => {
    let $ = false;
    (async () => {
      c(true);
      let [H, K, te] = await Promise.all([
        supabase.from("test_cases").select("*").eq("session_id", e.id).order("position"),
        supabase.from("findings").select("*").eq("session_id", e.id),
        supabase.from("done_marks").select("*").eq("session_id", e.id),
      ]);
      if ($) {
        return;
      }
      if (H.error || K.error || te.error) {
        h((H.error || K.error || te.error).message);
        c(false);
        return;
      }
      o((H.data || []).filter((de) => !de.removed_at));
      let P = {};
      (K.data || []).forEach((de) => {
        P[de.test_case_id] = de;
      });
      l(P);
      let X = {};
      (te.data || []).forEach((de) => {
        X[de.test_case_id] ||= [];
        X[de.test_case_id].push({
          developer_name: de.developer_name,
          done_at: de.done_at,
          status: de.status || "fixed",
          note: de.note || "",
        });
      });
      u(X);
      c(false);
      let ve = (H.data || []).find((de) => isFailureStatus(P[de.id]?.status));
      g((ve || (H.data || [])[0])?.id || null);
    })();
    return () => {
      $ = true;
    };
  }, [e.id]);
  let se = useRef(i);
  useEffect(() => {
    se.current = i;
  }, [i]);
  let E = useRef({});
  let [D, O] = useState(() => new Set());
  useEffect(() => {
    let $ = supabase
      .channel(`dev-findings-${e.id}`)
      .on(
        "postgres_changes",
        {
          event: "*",
          schema: "public",
          table: "findings",
          filter: `session_id=eq.${e.id}`,
        },
        (H) => {
          let K = H.new || H.old;
          if (K?.test_case_id) {
            if (H.eventType === "DELETE") {
              l((te) => {
                let P = {
                  ...te,
                };
                delete P[K.test_case_id];
                return P;
              });
            } else {
              l((te) => ({
                ...te,
                [K.test_case_id]: H.new,
              }));
            }
          }
        },
      )
      .subscribe();
    return () => {
      supabase.removeChannel($);
    };
  }, [e.id]);
  useEffect(() => {
    let $ = supabase
      .channel(`dev-done-${e.id}`)
      .on(
        "postgres_changes",
        {
          event: "*",
          schema: "public",
          table: "done_marks",
          filter: `session_id=eq.${e.id}`,
        },
        (H) => {
          if (H.eventType === "INSERT" || H.eventType === "UPDATE") {
            let K = H.new;
            u((te) => {
              let P = (te[K.test_case_id] || []).filter((X) => X.developer_name !== K.developer_name);
              return {
                ...te,
                [K.test_case_id]: [
                  ...P,
                  {
                    developer_name: K.developer_name,
                    done_at: K.done_at,
                    status: K.status || "fixed",
                    note: K.note || "",
                  },
                ],
              };
            });
          } else if (H.eventType === "DELETE") {
            let K = H.old;
            u((te) => {
              let P = te[K.test_case_id] || [];
              return {
                ...te,
                [K.test_case_id]: P.filter((X) => X.developer_name !== K.developer_name),
              };
            });
          }
        },
      )
      .subscribe();
    return () => {
      supabase.removeChannel($);
    };
  }, [e.id]);
  let B = async ($, H) => {
    let K = (se.current[$] || []).find((ve) => ve.developer_name === t);
    let te = K?.note || "";
    if (!H && !te) {
      if (!K) {
        return;
      }
      u((de) => ({
        ...de,
        [$]: (de[$] || []).filter((N) => N.developer_name !== t),
      }));
      let { error: ve } = await supabase
        .from("done_marks")
        .delete()
        .eq("test_case_id", $)
        .eq("developer_name", t);
      if (ve) {
        console.error("Clear status failed:", ve);
      }
      return;
    }
    let P = new Date().toISOString();
    u((ve) => {
      let de = (ve[$] || []).filter((N) => N.developer_name !== t);
      return {
        ...ve,
        [$]: [
          ...de,
          {
            developer_name: t,
            done_at: P,
            status: H,
            note: te,
          },
        ],
      };
    });
    let { error: X } = await supabase.from("done_marks").upsert(
      {
        test_case_id: $,
        session_id: e.id,
        developer_name: t,
        done_at: P,
        status: H,
        note: te,
      },
      {
        onConflict: "test_case_id,developer_name",
      },
    );
    if (X) {
      console.error("Set status failed:", X);
    }
  };
  let ne = ($, H) => {
    u((K) => {
      let te = K[$] || [];
      let P = te.find((N) => N.developer_name === t);
      let X = P?.status || null;
      let ve = P?.done_at || new Date().toISOString();
      let de = te.filter((N) => N.developer_name !== t);
      if (!H && !X) {
        return {
          ...K,
          [$]: de,
        };
      } else {
        return {
          ...K,
          [$]: [
            ...de,
            {
              developer_name: t,
              done_at: ve,
              status: X,
              note: H,
            },
          ],
        };
      }
    });
    O((K) => {
      if (K.has($)) {
        return K;
      }
      let te = new Set(K);
      te.add($);
      return te;
    });
    clearTimeout(E.current[$]);
    E.current[$] = setTimeout(async () => {
      let K = (se.current[$] || []).find((te) => te.developer_name === t);
      if (!K || (!K.note && !K.status)) {
        let { error: te } = await supabase
          .from("done_marks")
          .delete()
          .eq("test_case_id", $)
          .eq("developer_name", t);
        if (te) {
          console.error("Note delete failed:", te);
        }
      } else {
        let { error: te } = await supabase.from("done_marks").upsert(
          {
            test_case_id: $,
            session_id: e.id,
            developer_name: t,
            done_at: K.done_at,
            status: K.status,
            note: K.note,
          },
          {
            onConflict: "test_case_id,developer_name",
          },
        );
        if (te) {
          console.error("Note save failed:", te);
        }
      }
      O((te) => {
        if (!te.has($)) {
          return te;
        }
        let P = new Set(te);
        P.delete($);
        return P;
      });
    }, 600);
  };
  let he = ($) => severityRank($);
  let [devSeverity, setDevSeverity] = useState("All");
  let [devResult, setDevResult] = useState("All");
  let [devPriority, setDevPriority] = useState("All");
  let ie = (() => {
    let $ = r.filter((H) => {
      if (j && !j.has(H.id)) {
        return false;
      }
      let te = s[H.id]?.status || "Not Tested";
      let ve = (i[H.id] || []).find((N) => N.developer_name === t)?.status || null;
      if (
        (f === "failed" && !isIssueStatus(te)) ||
        (devResult !== "All" && te !== devResult) ||
        (f === "mywip" && ve !== "in_progress") ||
        (f === "myfixed" && ve !== "fixed")
      ) {
        return false;
      }
      if (devSeverity !== "All" && (getSeverity(s[H.id]?.bug_severity)?.value || "") !== devSeverity) {
        return false;
      }
      if (devPriority !== "All" && (H.priority || "") !== devPriority) {
        return false;
      }
      if (!v) {
        return true;
      }
      let de = v.toLowerCase();
      return (
        (H.test_id || "").toLowerCase().includes(de) ||
        (H.title || "").toLowerCase().includes(de) ||
        (H.module || "").toLowerCase().includes(de) ||
        (H.feature || "").toLowerCase().includes(de)
      );
    });
    if (f !== "failed") {
      return $;
    } else {
      return [...$].sort((H, K) => {
        let te = he(s[H.id]?.bug_severity);
        let P = he(s[K.id]?.bug_severity);
        if (te !== P) {
          return te - P;
        } else {
          return (H.position ?? 0) - (K.position ?? 0);
        }
      });
    }
  })();
  let ue = r.find(($) => $.id === x);
  let pe = ue ? s[ue.id] : null;
  let me = ue ? i[ue.id] || [] : [];
  let Y = me.find(($) => $.developer_name === t)?.status || null;
  let fe = r.filter(($) =>
    (i[$.id] || []).some((H) => H.developer_name === t && H.status === "in_progress"),
  ).length;
  let Ce = r.filter(($) =>
    (i[$.id] || []).some((H) => H.developer_name === t && H.status === "fixed"),
  ).length;
  let be = r.filter(($) => isIssueStatus(s[$.id]?.status)).length;
  useSidebarArrowNav(ie, x, g);
  if (d) {
    return <Notice>Loading session…</Notice>;
  }
  if (p) {
    return <Notice variant="error">Could not load session: {p}</Notice>;
  }
  let xe = async () => {
    try {
      let { error: $ } = await supabase
        .from("done_marks")
        .delete()
        .eq("session_id", e.id)
        .eq("developer_name", t);
      if ($) {
        throw $;
      }
      u((H) => {
        let K = {};
        for (let te of Object.keys(H)) {
          let P = (H[te] || []).filter((X) => X.developer_name !== t);
          if (P.length > 0) {
            K[te] = P;
          }
        }
        return K;
      });
    } catch ($) {
      console.error("Reset failed:", $);
      alert("Could not reset your inputs. Check your connection and try again.");
    }
  };
  return (
    <BaseUrlContext.Provider value={b}>
      <div className="h-screen flex bg-gray-950 overflow-hidden text-gray-100">
        <DevSidebar
          open={T}
          onClose={() => A(false)}
          sessionName={e.name}
          myName={t}
          onLogout={a}
          onReset={xe}
          testCases={r}
          findings={s}
          doneMarks={i}
          myWipCount={fe}
          myFixedCount={Ce}
          failedCount={be}
          visibleCases={ie}
          selectedId={x}
          onSelect={g}
          search={v}
          setSearch={k}
          searchOpen={C}
          setSearchOpen={m}
          filter={f}
          setFilter={w}
          severityFilter={devSeverity}
          setSeverityFilter={setDevSeverity}
          resultFilter={devResult}
          setResultFilter={setDevResult}
          priorityFilter={devPriority}
          setPriorityFilter={setDevPriority}
          focusedIds={j}
          onClearFocus={() => M(null)}
          onOpenFocusModal={() => V(true)}
          baseUrl={b}
          setBaseUrl={_}
          hasBaseUrlOverride={S}
          resetBaseUrl={I}
        />
        <div className="flex-1 min-w-0 flex flex-col relative">
          {!T && (
            <button
              onClick={() => A(true)}
              className="absolute top-3 left-3 z-20 p-2 bg-gray-900 border border-gray-800 rounded-sm hover:bg-gray-800"
            >
              <PanelLeftOpen className="w-4 h-4 text-gray-400" />
            </button>
          )}
          {ue ? (
            <DevTestDetail
              testCase={ue}
              finding={pe}
              marks={me}
              myName={t}
              myStatus={Y}
              myNote={me.find(($) => $.developer_name === t)?.note || ""}
              noteIsSaving={D.has(ue.id)}
              onSetStatus={($) => B(ue.id, $)}
              onSetNote={($) => ne(ue.id, $)}
            />
          ) : (
            <div className="flex-1 flex items-center justify-center text-gray-500 text-sm">
              No test case selected.
            </div>
          )}
        </div>
        <FocusModeModal
          open={W}
          onClose={() => V(false)}
          testCases={r.map(($) => ({
            __id: $.id,
            __rowNum: $.row_num,
            ID: $.test_id,
            Title: $.title,
            Module: $.module,
            Feature: $.feature,
          }))}
          currentFocusedIds={j}
          onApply={($) => M($)}
        />
      </div>
    </BaseUrlContext.Provider>
  );
}
function DevSidebar({
  open: e,
  onClose: t,
  sessionName: a,
  myName: r,
  onLogout: o,
  onReset: s,
  testCases: l,
  findings: i,
  doneMarks: u,
  myWipCount: d,
  myFixedCount: c,
  failedCount: p,
  visibleCases: h,
  selectedId: b,
  onSelect: _,
  search: S,
  setSearch: I,
  searchOpen: x,
  setSearchOpen: g,
  filter: v,
  setFilter: k,
  severityFilter: sevF = "All",
  setSeverityFilter: setSevF = () => {},
  resultFilter: resF = "All",
  setResultFilter: setResF = () => {},
  priorityFilter: priF = "All",
  setPriorityFilter: setPriF = () => {},
  focusedIds: C,
  onClearFocus: m,
  onOpenFocusModal: f,
  baseUrl: w,
  setBaseUrl: T,
  hasBaseUrlOverride: A,
  resetBaseUrl: j,
}) {
  let [M, W] = useState(false);
  let [V, F] = useState(false);
  let [se, E] = useState(false);
  let sevOf = (D) => getSeverity(i[D.id]?.bug_severity)?.value || "";
  let sevCounts = {};
  l.forEach((D) => {
    let O = sevOf(D);
    sevCounts[O] = (sevCounts[O] || 0) + 1;
  });
  let resCounts = {};
  l.forEach((D) => {
    let O = i[D.id]?.status || "Not Tested";
    resCounts[O] = (resCounts[O] || 0) + 1;
  });
  let priOptions = sortPriorities(Array.from(new Set(l.map((D) => D.priority).filter(Boolean))));
  let priHasNone = priOptions.length > 0 && l.some((D) => !D.priority);
  let devTabs = [
    { key: "all", label: "All", count: l.length },
    { key: "failed", label: "Issues", count: p, alert: p > 0, hint: "Failed, blocked or needs review" },
    { key: "mywip", label: "Fixing", count: d, hint: "Tests you've marked In Progress" },
    { key: "myfixed", label: "Fixed", count: c, hint: "Tests you've marked Fixed" },
  ];
  if (e) {
    return (
      <aside className="w-80 flex-shrink-0 bg-gray-900 border-r border-gray-800 flex flex-col overflow-hidden">
        <div className="px-4 py-3 border-b border-gray-800">
          <div className="flex items-center justify-between gap-2 mb-2">
            <div className="flex items-center gap-2 min-w-0">
              <button onClick={t} className="p-1 hover:bg-gray-800 rounded-sm flex-shrink-0" title="Collapse">
                <PanelLeftClose className="w-4 h-4 text-gray-400" />
              </button>
              <h2 className="font-bold text-gray-100 truncate">{a}</h2>
            </div>
            <button
              onClick={() => E(true)}
              className="p-1 hover:bg-gray-800 rounded-sm flex-shrink-0"
              title="Settings"
            >
              <Settings2 className="w-4 h-4 text-gray-400" />
            </button>
          </div>
          <div className="text-xs text-gray-400 mb-2">
            Signed in as <span className="text-gray-100 font-medium">{r}</span> <span className="text-gray-500">(Developer)</span>
          </div>
          <div className="flex items-center justify-between text-xs text-gray-500 mb-1">
            <span>
              {c} / {l.length} fixed by you
            </span>
            {p > 0 && <span className="text-red-400">{p} failed</span>}
          </div>
          <div className="w-full bg-gray-800 rounded-full h-1 mb-3 overflow-hidden">
            <div
              className="bg-emerald-500 h-1 rounded-full transition-all"
              style={{
                width: `${l.length ? (c / l.length) * 100 : 0}%`,
              }}
            />
          </div>
          <div className="grid grid-cols-4 border border-gray-700 rounded-sm overflow-hidden mb-2">
            {devTabs.map((D, O) => {
              let B = v === D.key;
              return (
                <button
                  key={D.key}
                  onClick={() => k(D.key)}
                  title={D.hint}
                  aria-pressed={B}
                  className={`h-8 px-1 flex items-center justify-center gap-1 text-[11px] font-medium transition-colors ${O > 0 ? "border-l border-gray-700" : ""} ${B ? "bg-emerald-600 text-white" : "bg-gray-800/60 text-gray-400 hover:bg-gray-800 hover:text-gray-200"}`}
                >
                  <span className="truncate">{D.label}</span>
                  <span className={B ? "text-emerald-100/80" : D.alert ? "text-red-400" : "text-gray-500"}>{D.count}</span>
                </button>
              );
            })}
          </div>
          {x ? (
            <SidebarSearch value={S} onChange={I} onClose={() => g(false)} />
          ) : (
            <div className="flex gap-1.5">
              <button
                onClick={() => g(true)}
                className={`flex-shrink-0 flex items-center justify-center border rounded-sm transition-colors w-8 h-8 ${S ? "bg-emerald-600 border-emerald-600 text-white hover:bg-emerald-500" : "bg-gray-800 border-gray-700 text-gray-400 hover:text-emerald-300 hover:border-emerald-700 hover:bg-gray-700"}`}
                title="Search test cases"
              >
                <Search className="w-3.5 h-3.5" />
              </button>
              <button
                onClick={f}
                className={`relative flex-shrink-0 flex items-center justify-center border rounded-sm transition-colors w-8 h-8 ${C ? "bg-emerald-600 border-emerald-600 text-white hover:bg-emerald-500" : "bg-gray-800 border-gray-700 text-gray-400 hover:text-emerald-300 hover:border-emerald-700 hover:bg-gray-700"}`}
                title={C ? `Focus filter active (${C.size} tests)` : "Focus on a hand-picked set of tests"}
              >
                <Filter className="w-3.5 h-3.5" />
              </button>
              <FilterSelect value={resF} onChange={setResF} active={resF !== "All"}>
                <option value="All">Tester result</option>
                {TESTER_STATUSES.map((D) => (
                  <option value={D.value} key={D.value}>
                    {D.label} ({resCounts[D.value] || 0})
                  </option>
                ))}
              </FilterSelect>
            </div>
          )}
          <div className="flex gap-1.5 mt-1.5">
              <FilterSelect value={sevF} onChange={setSevF} active={sevF !== "All"}>
                <option value="All">Severity</option>
                {SEVERITIES.map((D) => (
                  <option value={D.value} key={D.value}>
                    {D.value} ({sevCounts[D.value] || 0})
                  </option>
                ))}
                <option value="">Not set ({sevCounts[""] || 0})</option>
              </FilterSelect>
              {priOptions.length > 0 && (
                <FilterSelect value={priF} onChange={setPriF} active={priF !== "All"}>
                  <option value="All">Priority</option>
                  {priOptions.map((D) => (
                    <option value={D} key={D}>
                      {D} ({l.filter((O) => O.priority === D).length})
                    </option>
                  ))}
                  {priHasNone && <option value="">No priority</option>}
                </FilterSelect>
              )}
          </div>
          {C && (
            <div className="mt-2 flex items-center gap-2 px-2.5 py-1.5 bg-emerald-950/40 border border-emerald-800/50 rounded-sm text-xs">
              <Filter className="w-3 h-3 text-emerald-400 flex-shrink-0" />
              <span className="text-emerald-200">
                Focused: {C.size} of {l.length}
              </span>
              <button
                onClick={m}
                className="ml-auto text-emerald-400 hover:text-emerald-200 transition-colors"
                title="Clear focus filter"
              >
                <X className="w-3 h-3" />
              </button>
            </div>
          )}
        </div>
        <div className="flex-1 overflow-y-auto">
          {h.length === 0 ? (
            <EmptyState
              hasAnyTests={l.length > 0}
              suggestions={[
                C && {
                  label: "Clear focus",
                  onClick: m,
                },
                v !== "all" && {
                  label: "Show all tests",
                  onClick: () => k("all"),
                },
                resF !== "All" && {
                  label: "Clear tester result filter",
                  onClick: () => setResF("All"),
                },
                sevF !== "All" && {
                  label: "Clear severity filter",
                  onClick: () => setSevF("All"),
                },
                priF !== "All" && {
                  label: "Clear priority filter",
                  onClick: () => setPriF("All"),
                },
                S && {
                  label: "Clear search",
                  onClick: () => I(""),
                },
              ]}
            />
          ) : (
            h.map((D) => {
              let O = i[D.id];
              let B = O?.status || "Not Tested";
              let ne = getTesterStatus(B);
              let he = u[D.id] || [];
              let ue = he.find((be) => be.developer_name === r)?.status || null;
              let pe = he.filter((be) => be.developer_name !== r);
              let me = pe.filter((be) => be.status === "fixed").length;
              let Y = pe.filter((be) => be.status === "in_progress").length;
              let fe = D.id === b;
              let Ce = ue === "fixed" && B === "Pass";
              return (
                <button
                  data-tc-id={D.id}
                  onClick={() => _(D.id)}
                  className={`w-full text-left px-3 py-2.5 border-b border-gray-800 border-l-4 flex items-start gap-2 hover:bg-gray-800/60 ${fe ? "bg-emerald-950/40 border-l-emerald-400 shadow-[inset_6px_0_20px_-4px_rgba(16,185,129,0.4)]" : "border-l-transparent"} ${Ce ? "opacity-50" : ""}`}
                  key={D.id}
                >
                  <span className={`w-2 h-2 rounded-full mt-1.5 flex-shrink-0 ${ne.dot}`} />
                  <div className="min-w-0 flex-1">
                    <div className="flex items-center gap-1.5 text-[10px] text-gray-500 mb-0.5">
                      <span className="font-mono text-emerald-400">{D.test_id || "—"}</span>
                      {D.module && <span className="truncate">· {D.module}</span>}
                      {getSeverity(O?.bug_severity) && (
                        <span className="ml-auto">
                          <SeverityBadge value={O.bug_severity} size="xs" />
                        </span>
                      )}
                    </div>
                    <div
                      className={`text-sm leading-tight ${fe ? "text-gray-100 font-semibold" : "text-gray-300"}`}
                    >
                      {D.title || "Untitled"}
                    </div>
                  </div>
                  <div className="flex flex-col items-end gap-0.5 mt-0.5 flex-shrink-0">
                    {ue === "fixed" && (
                      <span
                        className="w-4 h-4 bg-emerald-600 rounded-sm flex items-center justify-center"
                        title="You marked this fixed"
                      >
                        <Check className="w-3 h-3 text-white" strokeWidth={3} />
                      </span>
                    )}
                    {ue === "in_progress" && (
                      <span
                        className="w-4 h-4 bg-amber-600 rounded-sm flex items-center justify-center"
                        title="You're working on this"
                      >
                        <Clock className="w-2.5 h-2.5 text-white" strokeWidth={3} />
                      </span>
                    )}
                    {(me > 0 || Y > 0) && (
                      <span className="text-[9px] text-gray-500" title={summarizeDevMarks(pe)}>
                        {me > 0 && `+${me}\u2713`}
                        {me > 0 && Y > 0 && " "}
                        {Y > 0 && `+${Y}\u2026`}
                      </span>
                    )}
                  </div>
                </button>
              );
            })
          )}
        </div>
        {se && (
          <SettingsModal
            baseUrl={w}
            setBaseUrl={T}
            hasBaseUrlOverride={A}
            resetBaseUrl={j}
            myName={r}
            roleLabel="Reviewing"
            onSwitchRequested={() => W(true)}
            onResetRequested={s ? () => F(true) : undefined}
            resetDescription="Delete every In Progress / Fixed mark you've made in this session, along with any notes you wrote. Other developers' work is not affected."
            onClose={() => E(false)}
          />
        )}
        {M && (
          <ConfirmDialog
            title="Switch user?"
            message="You'll be returned to the role picker and signed out. Your saved progress for this session isn't affected."
            confirmLabel="Switch"
            onConfirm={() => {
              W(false);
              o();
            }}
            onCancel={() => W(false)}
          />
        )}
        {V && (
          <ConfirmDialog
            title="Reset all your inputs?"
            message="Every In Progress / Fixed mark and note you've made in this session will be permanently deleted. This action cannot be undone."
            confirmLabel="Reset everything"
            danger={true}
            onConfirm={() => {
              F(false);
              s();
            }}
            onCancel={() => F(false)}
          />
        )}
      </aside>
    );
  } else {
    return null;
  }
}
function summarizeDevMarks(e) {
  let t = e.filter((o) => o.status === "fixed").map((o) => o.developer_name);
  let a = e.filter((o) => o.status === "in_progress").map((o) => o.developer_name);
  let r = [];
  if (t.length) {
    r.push(`Fixed by ${t.join(", ")}`);
  }
  if (a.length) {
    r.push(`In progress: ${a.join(", ")}`);
  }
  return r.join(" · ");
}
// Priority / focus-area helpers (values come straight from the imported sheet, e.g. "P0", "Security").
var priorityRank = (e) => {
  let t = String(e || "").match(/\d+/);
  return t ? parseInt(t[0], 10) : 99;
};
var sortPriorities = (e) => [...e].sort((t, a) => priorityRank(t) - priorityRank(a) || t.localeCompare(a));
var priorityBadgeClass = (e) => {
  let t = priorityRank(e);
  if (t === 0) return "bg-red-950/60 text-red-300 border-red-800";
  if (t === 1) return "bg-amber-950/50 text-amber-300 border-amber-800";
  if (t === 2) return "bg-sky-950/50 text-sky-300 border-sky-800";
  return "bg-gray-800 text-gray-300 border-gray-700";
};
function PriorityBadge({ value: e, size: t = "sm" }) {
  if (!e) return null;
  return (
    <span
      className={`inline-flex items-center border rounded-sm font-mono font-bold flex-shrink-0 ${t === "xs" ? "text-[9px] px-1 py-px" : "text-[10px] px-1.5 py-0.5"} ${priorityBadgeClass(e)}`}
      title={`Priority: ${e}`}
    >
      {e}
    </span>
  );
}
function FilterSelect({ value: e, onChange: t, children: a, active: r }) {
  return (
    <select
      value={e}
      onChange={(o) => t(o.target.value)}
      className={`flex-1 min-w-0 px-2 py-1.5 text-xs bg-gray-800 border text-gray-100 rounded-sm focus:outline-none focus:ring-1 focus:ring-emerald-500 appearance-none ${r ? "border-emerald-600" : "border-gray-700"}`}
      style={{
        backgroundImage: `url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 12 12'><path d='M3 4.5l3 3 3-3' stroke='%239ca3af' stroke-width='1.5' fill='none' stroke-linecap='round' stroke-linejoin='round'/></svg>")`,
        backgroundRepeat: "no-repeat",
        backgroundPosition: "right 8px center",
        paddingRight: "28px",
      }}
    >
      {a}
    </select>
  );
}
function Breadcrumb({ tc: e }) {
  return (
    <div className="flex items-center gap-2 text-xs text-gray-400 font-mono mb-1.5">
      {e.row_num && (
        <span className="bg-gray-800 text-gray-500 px-1.5 py-0.5 rounded-sm text-[10px]">
          Row {e.row_num}
        </span>
      )}
      <span className="text-emerald-400">{e.test_id || "—"}</span>
      {e.module && (
        <Fragment>
          <ChevronRight className="w-3 h-3 text-gray-700 flex-shrink-0" />
          <span>{e.module}</span>
        </Fragment>
      )}
      {e.feature && e.feature !== "—" && (
        <Fragment>
          <ChevronRight className="w-3 h-3 text-gray-700 flex-shrink-0" />
          <span>{e.feature}</span>
        </Fragment>
      )}
      {e.priority && <PriorityBadge value={e.priority} />}
      {e.focus && (
        <span
          className="bg-gray-800 text-gray-300 border border-gray-700 px-1.5 py-0.5 rounded-sm text-[10px] font-sans"
          title={`Focus area: ${e.focus}`}
        >
          {e.focus}
        </span>
      )}
      {e.dev_branch_name && (
        <span
          className="inline-flex items-center gap-1 bg-emerald-950/50 text-emerald-300 border border-emerald-900 px-1.5 py-0.5 rounded-sm text-[10px] normal-case"
          title={`Dev branch: ${e.dev_branch_name}`}
        >
          <GitBranch className="w-3 h-3" />
          <span className="font-mono">{e.dev_branch_name}</span>
        </span>
      )}
    </div>
  );
}
function TabButton({ active: e, onClick: t, children: a }) {
  return (
    <button
      onClick={t}
      className={`px-2 py-1 text-[11px] font-medium rounded-sm transition-colors ${e ? "bg-emerald-600 text-white" : "bg-gray-800 text-gray-400 hover:bg-gray-700 hover:text-gray-200"}`}
    >
      {a}
    </button>
  );
}
function SaveIndicator({ saving: e }) {
  return (
    <span className="inline-flex items-center gap-1.5 text-[10px] uppercase tracking-wider font-medium flex-shrink-0">
      {e ? (
        <Fragment>
          <span className="w-1.5 h-1.5 rounded-full bg-amber-400 animate-pulse" />
          <span className="text-amber-400">Saving…</span>
        </Fragment>
      ) : (
        <Fragment>
          <span className="w-1.5 h-1.5 rounded-full bg-emerald-500" />
          <span className="text-gray-500">Saved</span>
        </Fragment>
      )}
    </span>
  );
}
function EmptyState({ hasAnyTests: e, suggestions: t }) {
  let a = (t || []).filter(Boolean);
  return (
    <div className="p-6 text-center text-xs text-gray-500 space-y-3">
      <div>{e ? "No tests match the current view." : "No test cases yet."}</div>
      {a.length > 0 && (
        <div className="flex flex-col items-center gap-1.5">
          {a.map((r, o) => (
            <button
              onClick={r.onClick}
              className="text-emerald-400 hover:text-emerald-300 underline-offset-2 hover:underline"
              key={o}
            >
              {r.label}
            </button>
          ))}
        </div>
      )}
    </div>
  );
}
function ConfirmDialog({
  title: e,
  message: t,
  confirmLabel: a = "Confirm",
  cancelLabel: r = "Cancel",
  danger: o = false,
  onConfirm: s,
  onCancel: l,
}) {
  useEffect(() => {
    let i = (u) => {
      if (u.key === "Escape") {
        l();
      }
    };
    window.addEventListener("keydown", i);
    return () => window.removeEventListener("keydown", i);
  }, [l]);
  return (
    <div className="fixed inset-0 z-50 bg-black/70 flex items-center justify-center p-4" onClick={l}>
      <div
        className="bg-gray-900 border border-gray-700 rounded-sm w-full max-w-sm shadow-2xl"
        onClick={(i) => i.stopPropagation()}
      >
        <div className="px-5 py-4">
          <h3 className="text-sm font-bold text-gray-100 mb-1">{e}</h3>
          {t && <p className="text-xs text-gray-400 leading-relaxed">{t}</p>}
        </div>
        <div className="flex items-center justify-end gap-2 px-5 py-3 border-t border-gray-800">
          <button onClick={l} className="px-3 py-1.5 text-xs text-gray-400 hover:text-gray-200">
            {r}
          </button>
          <button
            onClick={s}
            autoFocus={true}
            className={`px-3 py-1.5 text-xs rounded-sm font-medium text-white ${o ? "bg-red-600 hover:bg-red-500" : "bg-emerald-600 hover:bg-emerald-500"}`}
          >
            {a}
          </button>
        </div>
      </div>
    </div>
  );
}
function SettingsModal({
  baseUrl: e,
  setBaseUrl: t,
  hasBaseUrlOverride: a,
  resetBaseUrl: r,
  myName: o,
  roleLabel: s,
  onSwitchRequested: l,
  onResetRequested: i,
  resetDescription: u,
  onClose: d,
}) {
  useEffect(() => {
    let c = (p) => {
      if (p.key === "Escape") {
        d();
      }
    };
    window.addEventListener("keydown", c);
    return () => window.removeEventListener("keydown", c);
  }, [d]);
  return (
    <div className="fixed inset-0 z-50 bg-black/70 flex items-center justify-center p-4" onClick={d}>
      <div
        className="bg-gray-900 border border-gray-700 rounded-sm w-full max-w-md shadow-2xl"
        onClick={(c) => c.stopPropagation()}
      >
        <div className="flex items-center justify-between px-5 py-3 border-b border-gray-800">
          <h3 className="text-sm font-bold text-gray-100">Settings</h3>
          <button onClick={d} className="p-1 text-gray-400 hover:text-gray-100" title="Close (Esc)">
            <X className="w-4 h-4" />
          </button>
        </div>
        <div className="px-5 py-4 space-y-5">
          <div>
            <div className="flex items-center justify-between mb-1.5">
              <label className="text-[11px] uppercase tracking-wider text-gray-400 font-semibold">
                Path chip prefix
              </label>
              {a && r && (
                <button
                  onClick={r}
                  className="text-[11px] text-gray-500 hover:text-gray-300"
                  title="Discard your override and follow the session's default"
                >
                  Reset to session default
                </button>
              )}
            </div>
            <input
              type="text"
              value={e}
              onChange={(c) => t(c.target.value)}
              placeholder="https://localhost:3000"
              autoCapitalize="off"
              autoComplete="off"
              spellCheck="false"
              className={`w-full px-2.5 py-1.5 text-xs font-mono bg-gray-800 border rounded-sm text-gray-100 placeholder-gray-600 focus:outline-none focus:ring-1 focus:ring-emerald-600 ${e ? "border-emerald-800" : "border-gray-700"}`}
            />
            <div className="text-[11px] text-gray-500 mt-2 leading-relaxed">
              {e ? (
                <Fragment>
                  Path chips like <span className="font-mono text-gray-400">/escalations</span> become
                  clickable links to{" "}
                  <span className="text-emerald-400 font-mono break-all">
                    {e.replace(/\/+$/, "")}/escalations
                  </span>{" "}
                  (opens in a new tab).
                </Fragment>
              ) : (
                <Fragment>
                  Path chips like <span className="font-mono text-gray-400">/escalations</span> are
                  click-to-copy. Add a prefix to make them clickable links.
                </Fragment>
              )}
            </div>
          </div>
          <div className="border-t border-gray-800" />
          <div>
            <div className="text-[11px] uppercase tracking-wider text-gray-400 font-semibold mb-2">
              Session
            </div>
            <div className="text-xs text-gray-400 mb-3">
              {s} as <span className="text-gray-100 font-medium">{o}</span>
            </div>
            <button
              onClick={() => {
                d();
                l();
              }}
              className="px-3 py-1.5 text-xs font-medium bg-gray-800 hover:bg-gray-700 border border-gray-700 text-gray-200 rounded-sm inline-flex items-center gap-1.5"
            >
              <ArrowLeft className="w-3 h-3" /> Switch user
            </button>
          </div>
          {i && (
            <Fragment>
              <div className="border-t border-gray-800" />
              <div>
                <div className="text-[11px] uppercase tracking-wider text-red-400 font-semibold mb-2">
                  Danger zone
                </div>
                {u && <div className="text-[11px] text-gray-500 mb-3 leading-relaxed">{u}</div>}
                <button
                  onClick={() => {
                    d();
                    i();
                  }}
                  className="px-3 py-1.5 text-xs font-medium bg-red-950/40 hover:bg-red-900/40 border border-red-900 text-red-300 hover:text-red-200 rounded-sm inline-flex items-center gap-1.5"
                >
                  <Trash2 className="w-3 h-3" /> Reset all my inputs
                </button>
              </div>
            </Fragment>
          )}
        </div>
      </div>
    </div>
  );
}
// Search bar that takes over the sidebar's search/filter row while open.
// The X button (or Escape) clears the search and brings the row back.
function SidebarSearch({ value: e, onChange: t, onClose: a, className: r = "" }) {
  let o = () => {
    t("");
    a();
  };
  return (
    <div className={`flex gap-1.5 ${r}`}>
      <div className="relative flex-1 min-w-0">
        <Search className="w-3.5 h-3.5 absolute left-2.5 top-1/2 -translate-y-1/2 text-gray-500 pointer-events-none" />
        <input
          autoFocus={true}
          type="text"
          placeholder="Search ID, title, module, feature…"
          value={e}
          onChange={(l) => t(l.target.value)}
          onKeyDown={(l) => {
            if (l.key === "Escape") {
              o();
            }
          }}
          className="w-full h-8 pl-8 pr-2 text-base sm:text-xs bg-gray-800 border border-gray-700 text-gray-100 placeholder-gray-500 rounded-sm focus:outline-none focus:ring-1 focus:ring-emerald-500"
        />
      </div>
      <button
        onClick={o}
        className="flex-shrink-0 flex items-center justify-center w-8 h-8 border border-gray-700 bg-gray-800 text-gray-400 rounded-sm transition-colors hover:text-gray-100 hover:bg-gray-700"
        title="Close search"
        aria-label="Close search"
      >
        <X className="w-3.5 h-3.5" />
      </button>
    </div>
  );
}
function ScreenshotDropzone({ disabled: e, isUploading: t, onFiles: a }) {
  let r = useRef(null);
  let [o, s] = useState(false);
  return (
    <div
      onClick={() => {
        if (!e) {
          r.current?.click();
        }
      }}
      onDragOver={(l) => {
        l.preventDefault();
        if (!e) {
          s(true);
        }
      }}
      onDragLeave={() => s(false)}
      onDrop={(l) => {
        l.preventDefault();
        s(false);
        if (!e && l.dataTransfer.files.length > 0) {
          a(l.dataTransfer.files);
        }
      }}
      className={`border-2 border-dashed rounded-sm p-4 text-center transition-colors ${e ? "border-gray-800 bg-gray-900/40 cursor-wait" : o ? "border-emerald-500 bg-emerald-950/30 cursor-copy" : "border-gray-700 bg-gray-900/30 hover:border-emerald-600 hover:bg-emerald-950/10 cursor-pointer"}`}
    >
      {t ? (
        <div className="flex items-center justify-center gap-2 text-amber-400 text-xs">
          <span className="w-1.5 h-1.5 rounded-full bg-amber-400 animate-pulse" />
          <span>Uploading…</span>
        </div>
      ) : (
        <Fragment>
          <div className="text-sm text-gray-300">Click, drag, or paste screenshots</div>
          <div className="text-[10px] text-gray-500 mt-0.5">
            Tip: paste with Ctrl/Cmd+V anywhere on this page · max 10 MB per file
          </div>
        </Fragment>
      )}
      <input
        ref={r}
        type="file"
        accept="image/*"
        multiple={true}
        className="hidden"
        onChange={(l) => {
          if (l.target.files.length > 0) {
            a(l.target.files);
          }
          l.target.value = "";
        }}
      />
    </div>
  );
}
function ScreenshotThumb({ shot: e, idx: t, onRemove: a, onOpen: r }) {
  let _Component19 = r
    ? (s) => <button onClick={() => r(t)} className="block w-full text-left" title="View full size" {...s} />
    : (s) => <a href={e.url} target="_blank" rel="noopener noreferrer" className="block" {...s} />;
  return (
    <div className="relative group border border-gray-800 rounded-sm overflow-hidden bg-gray-950">
      <div className="absolute top-1.5 left-1.5 z-10 bg-emerald-600 text-white text-[10px] font-bold px-1.5 py-0.5 rounded-sm shadow">
        Fig. {t + 1}
      </div>
      <_Component19>
        <img
          src={e.url}
          alt={e.name || `Screenshot ${t + 1}`}
          className="w-full h-28 object-cover"
          loading="lazy"
        />
      </_Component19>
      {a && (
        <button
          onClick={(s) => {
            s.preventDefault();
            s.stopPropagation();
            a();
          }}
          className="absolute top-1.5 right-1.5 z-10 opacity-0 group-hover:opacity-100 bg-red-600 hover:bg-red-500 text-white p-1 rounded-sm shadow transition-opacity"
          title="Remove screenshot"
        >
          <Trash2 className="w-3.5 h-3.5" />
        </button>
      )}
      <div
        className="px-2 py-1 text-[10px] text-gray-500 truncate bg-gray-900 border-t border-gray-800"
        title={e.name}
      >
        {e.name || `screenshot-${t + 1}`}
      </div>
    </div>
  );
}
function ScreenshotLightbox({
  screenshots: e,
  startIdx: t = 0,
  onClose: a,
  editable: r = false,
  onAnnotationsChange: o,
}) {
  let s = Math.min(Math.max(t, 0), Math.max(0, (e?.length || 1) - 1));
  let [l, i] = useState(s);
  let u = e && e[l];
  let d = u?.annotations || [];
  let c = l > 0;
  let p = e && l < e.length - 1;
  let [h, b] = useState("select");
  let [_, S] = useState(null);
  let [I, x] = useState(null);
  let [editingText, setEditingText] = useState(null);
  useEffect(() => {
    b("select");
    S(null);
    x(null);
    setEditingText(null);
  }, [l]);
  let g = useRef(null);
  let v = useRef(null);
  let [k, C] = useState({
    w: 0,
    h: 0,
  });
  useEffect(() => {
    if (!g.current) {
      return;
    }
    let M = () => {
      let V = g.current.getBoundingClientRect();
      C({
        w: V.width,
        h: V.height,
      });
    };
    M();
    let W;
    if (typeof ResizeObserver !== "undefined") {
      W = new ResizeObserver(M);
      W.observe(g.current);
    } else {
      window.addEventListener("resize", M);
    }
    return () => {
      if (W) {
        W.disconnect();
      } else {
        window.removeEventListener("resize", M);
      }
    };
  }, [u?.url]);
  let m = (M) => {
    if (o) {
      o(l, M);
    }
  };
  let f = (M) => Math.max(0, Math.min(1, M));
  let w = (M, W) => {
    if (!v.current) {
      return {
        x: 0,
        y: 0,
      };
    }
    let V = v.current.getBoundingClientRect();
    return {
      x: f((M - V.left) / V.width),
      y: f((W - V.top) / V.height),
    };
  };
  useEffect(() => {
    let M = (W) => {
      let V = W.target?.tagName;
      if (V !== "INPUT" && V !== "TEXTAREA" && !W.target?.isContentEditable) {
        if (W.key === "Escape") {
          if (I) {
            x(null);
            return;
          }
          if (h !== "select") {
            b("select");
            return;
          }
          if (_) {
            S(null);
            return;
          }
          a();
          return;
        }
        if (W.key === "ArrowLeft" && c) {
          i((F) => F - 1);
          return;
        }
        if (W.key === "ArrowRight" && p) {
          i((F) => F + 1);
          return;
        }
        if ((W.key === "Delete" || W.key === "Backspace") && r && _) {
          W.preventDefault();
          m(d.filter((F) => F.id !== _));
          S(null);
        }
      }
    };
    window.addEventListener("keydown", M);
    return () => window.removeEventListener("keydown", M);
  }, [c, p, a, h, _, I, d, r, l]);
  let T = (M) => {
    if (!r || M.button !== 0) {
      return;
    }
    let { x: W, y: V } = w(M.clientX, M.clientY);
    if (h === "text") {
      M.preventDefault();
      let F = "t_" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
      m([
        ...d,
        {
          id: F,
          type: "text",
          x1: W,
          y1: V,
          x2: W,
          y2: V,
          text: "",
        },
      ]);
      S(F);
      setEditingText(F);
      return;
    }
    if (h === "arrow" || h === "circle") {
      M.preventDefault();
      let F = "a_" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
      m([
        ...d,
        {
          id: F,
          type: h,
          x1: W,
          y1: V,
          x2: W,
          y2: V,
        },
      ]);
      S(F);
      x({
        mode: "create",
        id: F,
      });
      try {
        M.currentTarget.setPointerCapture(M.pointerId);
      } catch {}
      return;
    }
    S(null);
  };
  let A = (M) => {
    if (!I) {
      return;
    }
    let { x: W, y: V } = w(M.clientX, M.clientY);
    if (I.mode === "create" || I.mode === "endpoint-end") {
      m(
        d.map((F) =>
          F.id === I.id
            ? {
                ...F,
                x2: W,
                y2: V,
              }
            : F,
        ),
      );
    } else if (I.mode === "endpoint-start") {
      m(
        d.map((F) =>
          F.id === I.id
            ? {
                ...F,
                x1: W,
                y1: V,
              }
            : F,
        ),
      );
    } else if (I.mode === "move") {
      let F = W - I.origX;
      let se = V - I.origY;
      m(
        d.map((E) =>
          E.id !== I.id
            ? E
            : {
                ...E,
                x1: f(I.origAnn.x1 + F),
                y1: f(I.origAnn.y1 + se),
                x2: f(I.origAnn.x2 + F),
                y2: f(I.origAnn.y2 + se),
              },
        ),
      );
    }
  };
  let j = (M) => {
    if (I) {
      try {
        M.currentTarget.releasePointerCapture(M.pointerId);
      } catch {}
      if (I.mode === "create") {
        let W = d.find((V) => V.id === I.id);
        let tooSmall =
          W &&
          (W.type === "circle"
            ? Math.abs(W.x2 - W.x1) < 0.015 && Math.abs(W.y2 - W.y1) < 0.015
            : Math.hypot(W.x2 - W.x1, W.y2 - W.y1) < 0.02);
        if (tooSmall) {
          m(d.filter((F) => F.id !== I.id));
          S(null);
        }
        // Stay in the drawing tool so several marks can be added in a row.
      }
      x(null);
    }
  };
  if (u) {
    return (
      <div className="fixed inset-0 z-50 bg-black/90 flex items-center justify-center p-4" onClick={a}>
        <div className="w-full max-w-5xl flex flex-col" onClick={(M) => M.stopPropagation()}>
          <div className="flex items-center gap-2 px-4 py-2.5 bg-gray-950 border border-gray-800 border-b-0 rounded-t-sm flex-shrink-0">
            <span className="text-[10px] font-semibold uppercase tracking-wider px-1.5 py-0.5 bg-emerald-600 text-white border border-emerald-700 rounded-sm flex-shrink-0">
              Fig. {l + 1}
            </span>
            <span className="text-xs text-gray-300 truncate min-w-0 flex-1" title={u.name}>
              {u.name || `screenshot-${l + 1}`}
            </span>
            {e.length > 1 && (
              <span className="text-[10px] text-gray-500 flex-shrink-0 tabular-nums">
                {l + 1} of {e.length}
              </span>
            )}
            {r && (
              <Fragment>
                {[
                  { key: "arrow", label: "Add arrow", Icon: Pencil, hint: "Click + drag on the image to draw arrows" },
                  { key: "circle", label: "Add circle", Icon: CircleIcon, hint: "Click + drag on the image to circle an area" },
                  { key: "text", label: "Add text", Icon: Type, hint: "Click on the image to add a text label" },
                ].map((M) => {
                  let W = h === M.key;
                  return (
                    <button
                      key={M.key}
                      onClick={() => {
                        b(W ? "select" : M.key);
                        S(null);
                      }}
                      aria-pressed={W}
                      className={`px-2 py-1 text-[11px] rounded-sm border inline-flex items-center gap-1 flex-shrink-0 ${W ? "bg-red-900/50 border-red-700 text-red-100" : "bg-gray-800 border-gray-700 text-gray-300 hover:bg-gray-700"}`}
                      title={W ? "Click to finish (or press Esc)" : M.hint}
                    >
                      {W ? <Check className="w-3 h-3" /> : <M.Icon className="w-3 h-3" />}
                      {W ? `Done ${M.key === "arrow" ? "arrows" : M.key === "circle" ? "circles" : "text"}` : M.label}
                    </button>
                  );
                })}
                {d.length > 0 && (
                  <button
                    onClick={() => {
                      m([]);
                      S(null);
                    }}
                    className="px-2 py-1 text-[11px] rounded-sm bg-gray-800 border border-gray-700 text-gray-300 hover:bg-gray-700 hover:text-red-300 inline-flex items-center gap-1 flex-shrink-0"
                    title="Remove all arrows, circles and text on this screenshot"
                  >
                    <Trash2 className="w-3 h-3" /> Clear
                  </button>
                )}
              </Fragment>
            )}
            <a
              href={u.url}
              target="_blank"
              rel="noopener noreferrer"
              className="p-1.5 text-gray-400 hover:text-gray-100 flex-shrink-0"
              title="Open in new tab"
            >
              <ExternalLink className="w-4 h-4" />
            </a>
            <button
              onClick={a}
              className="p-1.5 text-gray-400 hover:text-gray-100 flex-shrink-0"
              title="Close (Esc)"
            >
              <X className="w-4 h-4" />
            </button>
          </div>
          <div
            className="relative bg-black border border-gray-800 flex items-center justify-center"
            style={{
              minHeight: "300px",
            }}
          >
            <div ref={g} className="relative inline-block">
              <img
                src={u.url}
                alt={u.name || `Screenshot ${l + 1}`}
                className="max-h-[75vh] max-w-full object-contain block select-none"
                draggable={false}
                key={u.url}
              />
              {k.w > 0 && (
                <svg
                  ref={v}
                  width={k.w}
                  height={k.h}
                  viewBox={`0 0 ${k.w} ${k.h}`}
                  className="absolute inset-0"
                  style={{
                    cursor: h === "arrow" || h === "circle" ? "crosshair" : h === "text" ? "text" : "default",
                    pointerEvents: r || d.length > 0 ? "auto" : "none",
                    touchAction: "none",
                  }}
                  onPointerDown={T}
                  onPointerMove={A}
                  onPointerUp={j}
                  onPointerCancel={j}
                >
                  {d.map((M) => {
                    if (M.type !== "arrow" && M.type !== "circle") {
                      return null;
                    }
                    let isCircle = M.type === "circle";
                    let W = M.x1 * k.w;
                    let V = M.y1 * k.h;
                    let F = M.x2 * k.w;
                    let se = M.y2 * k.h;
                    let E = _ === M.id;
                    let D = E ? "#fb7185" : "#ef4444";
                    return (
                      <g key={M.id}>
                        {isCircle ? (
                          <Fragment>
                            {E && r && (
                              <rect
                                x={Math.min(W, F)}
                                y={Math.min(V, se)}
                                width={Math.abs(F - W)}
                                height={Math.abs(se - V)}
                                fill="none"
                                stroke="#fb7185"
                                strokeWidth={1}
                                strokeDasharray="4 4"
                                pointerEvents="none"
                              />
                            )}
                            <ellipse
                              cx={(W + F) / 2}
                              cy={(V + se) / 2}
                              rx={Math.abs(F - W) / 2}
                              ry={Math.abs(se - V) / 2}
                              stroke={D}
                              strokeWidth={E ? 4 : 3}
                              fill="none"
                              pointerEvents="none"
                            />
                          </Fragment>
                        ) : (
                          <path
                            d={arrowGeometry(W, V, F, se)}
                            stroke={D}
                            strokeWidth={E ? 4 : 3}
                            fill="none"
                            strokeLinecap="round"
                            strokeLinejoin="round"
                            pointerEvents="none"
                          />
                        )}
                        {r && (
                          <HitShape
                            circle={isCircle}
                            x1={W}
                            y1={V}
                            x2={F}
                            y2={se}
                            stroke="rgba(0,0,0,0.001)"
                            strokeWidth={18}
                            strokeLinecap="round"
                            style={{
                              cursor: h === "select" ? "move" : "crosshair",
                            }}
                            onPointerDown={(O) => {
                              if (h !== "select") {
                                return;
                              }
                              O.stopPropagation();
                              O.preventDefault();
                              S(M.id);
                              let { x: B, y: ne } = w(O.clientX, O.clientY);
                              x({
                                mode: "move",
                                id: M.id,
                                origAnn: {
                                  ...M,
                                },
                                origX: B,
                                origY: ne,
                              });
                              try {
                                O.currentTarget.setPointerCapture(O.pointerId);
                              } catch {}
                            }}
                          />
                        )}
                        {E && r && (
                          <Fragment>
                            <circle
                              cx={W}
                              cy={V}
                              r={8}
                              fill="#fff"
                              stroke="#ef4444"
                              strokeWidth={2}
                              style={{
                                cursor: "crosshair",
                              }}
                              onPointerDown={(O) => {
                                if (h === "select") {
                                  O.stopPropagation();
                                  O.preventDefault();
                                  x({
                                    mode: "endpoint-start",
                                    id: M.id,
                                  });
                                  try {
                                    O.currentTarget.setPointerCapture(O.pointerId);
                                  } catch {}
                                }
                              }}
                            />
                            <circle
                              cx={F}
                              cy={se}
                              r={8}
                              fill="#fff"
                              stroke="#ef4444"
                              strokeWidth={2}
                              style={{
                                cursor: "crosshair",
                              }}
                              onPointerDown={(O) => {
                                if (h === "select") {
                                  O.stopPropagation();
                                  O.preventDefault();
                                  x({
                                    mode: "endpoint-end",
                                    id: M.id,
                                  });
                                  try {
                                    O.currentTarget.setPointerCapture(O.pointerId);
                                  } catch {}
                                }
                              }}
                            />
                          </Fragment>
                        )}
                      </g>
                    );
                  })}
                </svg>
              )}
              {k.w > 0 &&
                d
                  .filter((M) => M.type === "text")
                  .map((M) => {
                    let W = _ === M.id;
                    let V = editingText === M.id && r;
                    let F = {
                      position: "absolute",
                      left: `${M.x1 * k.w}px`,
                      top: `${M.y1 * k.h}px`,
                      maxWidth: `${Math.max(140, k.w - M.x1 * k.w)}px`,
                      fontSize: "18px",
                      lineHeight: 1.3,
                      fontWeight: 600,
                      color: "#fff",
                      background: "#dc2626",
                      padding: "4px 8px",
                      borderRadius: "3px",
                      whiteSpace: "pre-wrap",
                      wordBreak: "break-word",
                      boxShadow: W && r ? "0 0 0 2px #0b0f17, 0 0 0 4px #fb7185" : "0 1px 3px rgba(0,0,0,0.4)",
                      pointerEvents: !r || h === "arrow" || h === "circle" ? "none" : "auto",
                      touchAction: "none",
                      cursor: h === "select" ? "move" : "text",
                    };
                    let se = (E) => {
                      let D = E.trim();
                      setEditingText(null);
                      m(D ? d.map((O) => (O.id === M.id ? { ...O, text: E } : O)) : d.filter((O) => O.id !== M.id));
                      if (!D) S(null);
                    };
                    if (V) {
                      return (
                        <AnnotationTextInput
                          key={M.id}
                          initial={M.text || ""}
                          style={{ ...F, cursor: "text", minWidth: "140px", border: "none", outline: "none", resize: "none" }}
                          onCommit={se}
                        />
                      );
                    }
                    return (
                      <div
                        key={M.id}
                        style={F}
                        title={r && h === "select" ? "Drag to move · double-click to edit · Delete to remove" : undefined}
                        onPointerDown={(E) => {
                          if (!r) return;
                          E.stopPropagation();
                          E.preventDefault();
                          S(M.id);
                          if (h === "text") {
                            setEditingText(M.id);
                            return;
                          }
                          let { x: D, y: O } = w(E.clientX, E.clientY);
                          x({ mode: "move", id: M.id, origAnn: { ...M }, origX: D, origY: O });
                          try {
                            E.currentTarget.setPointerCapture(E.pointerId);
                          } catch {}
                        }}
                        onPointerMove={A}
                        onPointerUp={j}
                        onPointerCancel={j}
                        onDoubleClick={(E) => {
                          if (!r) return;
                          E.stopPropagation();
                          setEditingText(M.id);
                        }}
                      >
                        {M.text}
                      </div>
                    );
                  })}
            </div>
          </div>
          {(e.length > 1 || r) && (
            <div className="flex items-center justify-between px-3 py-2 bg-gray-950 border border-gray-800 border-t-0 rounded-b-sm flex-shrink-0 gap-3">
              {e.length > 1 ? (
                <button
                  onClick={() => i((M) => M - 1)}
                  disabled={!c}
                  className="flex items-center gap-1 px-2.5 py-1 text-xs text-gray-300 hover:text-gray-100 disabled:text-gray-700 disabled:cursor-not-allowed"
                  title="Previous (←)"
                >
                  <ChevronLeft className="w-3.5 h-3.5" /> Previous
                </button>
              ) : (
                <span />
              )}
              {r && (
                <span className="text-[10px] text-gray-500 truncate">
                  {h === "arrow"
                    ? "Drag on the image to draw arrows. Add as many as you like, then click Done or press Esc."
                    : h === "circle"
                      ? "Drag across an area to circle it. Add as many as you like, then click Done or press Esc."
                      : h === "text"
                        ? "Click where the label should go and type. Enter to finish, Shift+Enter for a new line."
                      : _
                        ? "Drag the handles to reshape, drag the outline to move, Delete to remove."
                        : d.length > 0
                          ? "Click a mark to select it. Double-click text to edit. Press Esc to close."
                          : 'Pick "Add arrow", "Add circle" or "Add text", then use it on the image.'}
                </span>
              )}
              {e.length > 1 ? (
                <button
                  onClick={() => i((M) => M + 1)}
                  disabled={!p}
                  className="flex items-center gap-1 px-2.5 py-1 text-xs text-gray-300 hover:text-gray-100 disabled:text-gray-700 disabled:cursor-not-allowed"
                  title="Next (→)"
                >
                  Next <ChevronRight className="w-3.5 h-3.5" />
                </button>
              ) : (
                <span />
              )}
            </div>
          )}
        </div>
      </div>
    );
  } else {
    return null;
  }
}
// Inline editor for a text label on a screenshot. Enter (or clicking away) saves; Shift+Enter adds a line.
function AnnotationTextInput({ initial: e, style: t, onCommit: a }) {
  let [r, o] = useState(e);
  let s = useRef(null);
  let l = useRef(false);
  let i = (u) => {
    if (l.current) return;
    l.current = true;
    a(u);
  };
  useLayoutEffect(() => {
    let u = s.current;
    if (!u) return;
    u.style.height = "auto";
    u.style.height = `${u.scrollHeight}px`;
  }, [r]);
  useEffect(() => {
    let u = s.current;
    if (u) {
      u.focus();
      u.setSelectionRange(u.value.length, u.value.length);
    }
  }, []);
  return (
    <textarea
      ref={s}
      rows={1}
      value={r}
      placeholder="Type…"
      onChange={(u) => o(u.target.value)}
      onPointerDown={(u) => u.stopPropagation()}
      onKeyDown={(u) => {
        u.stopPropagation();
        if ((u.key === "Enter" && !u.shiftKey) || u.key === "Escape") {
          u.preventDefault();
          i(r);
        }
      }}
      onBlur={() => i(r)}
      className="placeholder-red-200"
      style={{ ...t, display: "block", overflow: "hidden", fontFamily: "inherit" }}
    />
  );
}
// Invisible, wide outline used to grab an arrow (line) or circle (ellipse) in the lightbox.
function HitShape({ circle: e, x1: t, y1: a, x2: r, y2: o, ...s }) {
  if (e) {
    return (
      <ellipse
        cx={(t + r) / 2}
        cy={(a + o) / 2}
        rx={Math.abs(r - t) / 2}
        ry={Math.abs(o - a) / 2}
        fill="none"
        pointerEvents="stroke"
        {...s}
      />
    );
  }
  return <line x1={t} y1={a} x2={r} y2={o} {...s} />;
}
function arrowGeometry(e, t, a, r) {
  let o = a - e;
  let s = r - t;
  let l = Math.hypot(o, s);
  if (l < 1) {
    return `M ${e} ${t}`;
  }
  let i = Math.min(22, Math.max(8, l * 0.28));
  let u = Math.PI / 7;
  let d = Math.atan2(s, o);
  let c = a - i * Math.cos(d - u);
  let p = r - i * Math.sin(d - u);
  let h = a - i * Math.cos(d + u);
  let b = r - i * Math.sin(d + u);
  return `M ${e} ${t} L ${a} ${r} M ${c} ${p} L ${a} ${r} L ${h} ${b}`;
}
function RecordingPlayer({ recordings: e, startIdx: t = 0, onClose: a }) {
  let r = Math.min(Math.max(t, 0), Math.max(0, (e?.length || 1) - 1));
  let [o, s] = useState(r);
  let l = e && e[o];
  let i = l ? parseRecordingUrl(l.url) : null;
  let u = o > 0;
  let d = e && o < e.length - 1;
  useEffect(() => {
    let c = (p) => {
      if (p.key === "Escape") {
        a();
        return;
      }
      if (p.key === "ArrowLeft" && u) {
        s((h) => h - 1);
      }
      if (p.key === "ArrowRight" && d) {
        s((h) => h + 1);
      }
    };
    window.addEventListener("keydown", c);
    return () => window.removeEventListener("keydown", c);
  }, [u, d, a]);
  if (l) {
    return (
      <div className="fixed inset-0 z-50 bg-black/90 flex items-center justify-center p-4" onClick={a}>
        <div
          className="w-full max-w-4xl bg-gray-950 border border-gray-800 rounded-sm flex flex-col shadow-2xl"
          onClick={(c) => c.stopPropagation()}
        >
          <div className="flex items-center gap-2 px-4 py-2.5 border-b border-gray-800 flex-shrink-0">
            <span className="text-[10px] font-semibold uppercase tracking-wider px-1.5 py-0.5 bg-gray-900 text-emerald-300 border border-gray-700 rounded-sm flex-shrink-0">
              {i?.label || "Link"}
            </span>
            <span className="text-xs text-gray-400 truncate min-w-0 flex-1" title={l.url}>
              {l.url}
            </span>
            {e.length > 1 && (
              <span className="text-[10px] text-gray-500 flex-shrink-0 tabular-nums">
                {o + 1} of {e.length}
              </span>
            )}
            <a
              href={l.url}
              target="_blank"
              rel="noopener noreferrer"
              className="p-1.5 text-gray-400 hover:text-gray-100 flex-shrink-0"
              title="Open in new tab"
            >
              <ExternalLink className="w-4 h-4" />
            </a>
            <button
              onClick={a}
              className="p-1.5 text-gray-400 hover:text-gray-100 flex-shrink-0"
              title="Close (Esc)"
            >
              <X className="w-4 h-4" />
            </button>
          </div>
          <div className="relative aspect-video bg-black">
            {i?.embed ? (
              <iframe
                src={i.embed}
                title={`Recording ${o + 1}`}
                className="absolute inset-0 w-full h-full"
                frameBorder="0"
                allow="autoplay; fullscreen; picture-in-picture; clipboard-write"
                allowFullScreen={true}
                key={i.embed}
              />
            ) : i?.isVideoFile ? (
              <video
                src={l.url}
                controls={true}
                autoPlay={true}
                className="absolute inset-0 w-full h-full"
                key={l.url}
              />
            ) : (
              <div className="absolute inset-0 flex flex-col items-center justify-center gap-3 p-8 text-center">
                <div className="text-sm text-gray-300">This link can't be embedded in the app.</div>
                <div className="text-xs text-gray-500">
                  Loom, YouTube, Google Drive, Vimeo, and direct video files play here. Anything else opens
                  externally.
                </div>
                <a
                  href={l.url}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-emerald-600 text-white text-xs font-medium rounded-sm hover:bg-emerald-500 mt-1"
                >
                  <ExternalLink className="w-3.5 h-3.5" /> Open in new tab
                </a>
              </div>
            )}
          </div>
          {e.length > 1 && (
            <div className="flex items-center justify-between px-3 py-2 border-t border-gray-800 flex-shrink-0">
              <button
                onClick={() => s((c) => c - 1)}
                disabled={!u}
                className="flex items-center gap-1 px-2.5 py-1 text-xs text-gray-300 hover:text-gray-100 disabled:text-gray-700 disabled:cursor-not-allowed"
                title="Previous (←)"
              >
                <ChevronLeft className="w-3.5 h-3.5" /> Previous
              </button>
              <button
                onClick={() => s((c) => c + 1)}
                disabled={!d}
                className="flex items-center gap-1 px-2.5 py-1 text-xs text-gray-300 hover:text-gray-100 disabled:text-gray-700 disabled:cursor-not-allowed"
                title="Next (→)"
              >
                Next <ChevronRight className="w-3.5 h-3.5" />
              </button>
            </div>
          )}
        </div>
      </div>
    );
  } else {
    return null;
  }
}
function DevMarkCard({ mark: e }) {
  let t = e.status === "fixed";
  let a = e.status === "in_progress";
  return (
    <div className="px-3 py-2 bg-gray-900/60 border border-gray-800 rounded-sm">
      <div className="flex items-center gap-2 mb-1">
        <span className="text-sm font-medium text-gray-200">{e.developer_name}</span>
        {t && (
          <span className="text-[10px] uppercase tracking-wider font-bold px-1.5 py-0.5 rounded-sm bg-emerald-900/40 text-emerald-300 border border-emerald-700 whitespace-nowrap">
            Fixed
          </span>
        )}
        {a && (
          <span className="text-[10px] uppercase tracking-wider font-bold px-1.5 py-0.5 rounded-sm bg-amber-900/40 text-amber-300 border border-amber-700 whitespace-nowrap">
            In progress
          </span>
        )}
        <span className="ml-auto text-[10px] text-gray-500" title={new Date(e.done_at).toLocaleString()}>
          {timeAgo(e.done_at)}
        </span>
      </div>
      {e.note ? (
        <div className="text-sm text-gray-300 whitespace-pre-wrap">{e.note}</div>
      ) : (
        <div className="text-xs text-gray-600 italic">(no note left)</div>
      )}
    </div>
  );
}
function DevTestDetail({
  testCase: e,
  finding: t,
  marks: a,
  myName: r,
  myStatus: o,
  myNote: s,
  noteIsSaving: l,
  onSetStatus: i,
  onSetNote: u,
}) {
  let d = e;
  let c = t || {};
  let p = c.status || "Not Tested";
  let h = isFailureStatus(p);
  let b = getTesterStatus(p);
  let [_, S] = useState(s);
  useEffect(() => {
    S(s);
  }, [d.id]);
  let [I, x] = useState(null);
  useEffect(() => {
    x(null);
  }, [d.id]);
  let [g, v] = useState(null);
  useEffect(() => {
    v(null);
  }, [d.id]);
  let k = a.filter((C) => C.developer_name !== r).sort((C, m) => new Date(C.done_at) - new Date(m.done_at));
  return (
    <FigureContext.Provider value={{ shots: c.screenshots || [], open: (f) => v(f) }}>
    <div className="flex-1 overflow-y-auto">
      <div className="max-w-6xl mx-auto px-6 py-6">
        <div className="mb-6 pb-4 border-b border-gray-800 flex items-start gap-4 flex-wrap">
          <div className="flex-1 min-w-0">
            <Breadcrumb tc={d} />
            <h1 className="text-xl font-bold text-gray-100">{d.title || "Untitled"}</h1>
          </div>
          <div className="flex items-center gap-2 flex-wrap">
            <span
              className={`px-2.5 py-1 text-xs font-bold uppercase tracking-wider border rounded-sm ${b.badge}`}
            >
              {p}
            </span>
            {getSeverity(c.bug_severity) && <SeverityBadge value={c.bug_severity} />}
            <span className="self-stretch w-px bg-gray-800 mx-1" aria-hidden="true" />
            <button
              onClick={() => i(o === "in_progress" ? null : "in_progress")}
              className={`px-3 py-1.5 text-xs font-bold rounded-sm border flex items-center gap-1.5 transition-colors ${o === "in_progress" ? "bg-amber-600 text-white border-amber-600 hover:bg-amber-500" : "bg-gray-800 text-gray-300 border-gray-700 hover:bg-gray-700"}`}
              title={o === "in_progress" ? "Click to clear status" : "Mark that you're working on this"}
            >
              <Clock className="w-3.5 h-3.5" strokeWidth={2.5} />
              In Progress
            </button>
            <button
              onClick={() => i(o === "fixed" ? null : "fixed")}
              className={`px-3 py-1.5 text-xs font-bold rounded-sm border flex items-center gap-1.5 transition-colors ${o === "fixed" ? "bg-emerald-600 text-white border-emerald-600 hover:bg-emerald-500" : "bg-gray-800 text-gray-300 border-gray-700 hover:bg-gray-700"}`}
              title={o === "fixed" ? "Click to clear status" : "Mark this as fixed"}
            >
              <Check className="w-3.5 h-3.5" strokeWidth={3} />
              {o === "fixed" ? "Fixed" : "Mark as Fixed"}
            </button>
          </div>
        </div>
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          <div className="space-y-5">
            {d.pre_conditions && d.pre_conditions !== "—" && (
              <DetailField label="Pre-conditions" value={d.pre_conditions} />
            )}
            {d.test_data && d.test_data !== "—" && <DetailField label="Test Data" value={d.test_data} />}
            {(d.steps || []).length > 0 && (
              <div>
                <div className="text-xs font-semibold uppercase tracking-wider text-gray-400 mb-2">Steps</div>
                <div className="space-y-1.5">
                  {(d.steps || []).map((C, m) => {
                    let f = (c.step_checks || {})[m];
                    let w = (c.step_notes || {})[m];
                    let T = f === "done";
                    let A = f === "blocked";
                    let R = f === "review";
                    return (
                      <div
                        className={`p-2.5 border rounded-sm ${T ? "bg-green-900/15 border-green-800" : R ? "bg-orange-900/15 border-orange-800" : A ? "bg-red-900/15 border-red-800" : "bg-gray-900 border-gray-800"}`}
                        key={m}
                      >
                        <div className="flex items-start gap-2.5">
                          <div
                            className={`w-5 h-5 mt-0.5 rounded-sm flex items-center justify-center flex-shrink-0 ${T ? "bg-green-600" : R ? "bg-orange-500" : A ? "bg-red-600" : "border-2 border-gray-700"}`}
                          >
                            {(T || R) && <Check className="w-3 h-3 text-white" strokeWidth={3} />}
                            {A && <X className="w-3 h-3 text-white" strokeWidth={3} />}
                          </div>
                          <div className="text-[11px] font-bold text-emerald-400 mt-1 font-mono w-5 text-right flex-shrink-0">
                            {m + 1}.
                          </div>
                          <div className="flex-1 min-w-0">
                            <div className="text-sm text-gray-200">
                              {normalizePath(C)}
                              {A && (
                                <span className="ml-2 text-[10px] uppercase tracking-wider text-red-400 font-bold">
                                  Unable to proceed
                                </span>
                              )}
                              {R && (
                                <span className="ml-2 text-[10px] uppercase tracking-wider text-orange-400 font-bold">
                                  Works but needs review
                                </span>
                              )}
                            </div>
                            {w && (
                              <div className="mt-1.5 bg-amber-950/30 border-l-2 border-amber-500 px-2 py-1 text-xs text-amber-100 rounded-sm whitespace-pre-wrap">
                                {normalizePath(w)}
                              </div>
                            )}
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}
            {d.expected_result && d.expected_result !== "—" && (
              <ExpectedResultChecklist value={d.expected_result} checks={c.step_checks} notes={c.step_notes} />
            )}
            {d.downstream_check && d.downstream_check !== "—" && (
              <DetailField label="Downstream Check" value={d.downstream_check} />
            )}
            {d.edge_case && d.edge_case !== "—" && <DetailField label="Edge Case" value={d.edge_case} />}
          </div>
          <div className="space-y-5">
            <div className="bg-gray-900/40 border border-gray-800 rounded-sm p-4">
              <div className="text-xs uppercase tracking-wider text-emerald-400 font-semibold mb-3">
                Tester findings
              </div>
              {c.tested_by && (
                <div className="mb-3">
                  <div className="text-[10px] uppercase tracking-wider text-gray-500 mb-0.5">Tester</div>
                  <div className="text-sm text-gray-200">{c.tested_by}</div>
                </div>
              )}

              <div className="mb-3">
                <div className="text-[10px] uppercase tracking-wider text-gray-500 mb-0.5">Notes</div>
                <div className="text-sm text-gray-200 whitespace-pre-wrap bg-gray-950 border border-gray-800 px-3 py-2 rounded-sm">
                  {c.notes ? normalizePath(c.notes) : <span className="text-gray-600 italic">No notes</span>}
                </div>
              </div>
              {c.other_observations && (
                <div className="mb-3">
                  <div className="text-[10px] uppercase tracking-wider text-gray-500 mb-0.5">
                    Other observations
                  </div>
                  <div className="text-sm text-gray-200 whitespace-pre-wrap bg-gray-950 border border-gray-800 px-3 py-2 rounded-sm">
                    {normalizePath(c.other_observations)}
                  </div>
                </div>
              )}
              {(c.recordings || []).length > 0 && (
                <div className="mb-3">
                  <div className="text-[10px] uppercase tracking-wider text-gray-500 mb-1">
                    Recordings ({(c.recordings || []).length})
                  </div>
                  <div className="space-y-1.5">
                    {(c.recordings || []).map((C, m) => {
                      let w = parseRecordingUrl(C.url)?.label || "Link";
                      return (
                        <div
                          className="flex items-center gap-2 px-2.5 py-1.5 bg-gray-950 border border-gray-800 rounded-sm hover:border-emerald-700 hover:bg-gray-900 group"
                          key={m}
                        >
                          <span className="text-[10px] font-semibold uppercase tracking-wider px-1.5 py-0.5 bg-gray-900 text-emerald-300 border border-gray-700 rounded-sm flex-shrink-0">
                            {w}
                          </span>
                          <button
                            onClick={() =>
                              x({
                                recordings: c.recordings,
                                startIdx: m,
                              })
                            }
                            className="text-xs text-gray-300 group-hover:text-emerald-300 truncate flex-1 min-w-0 text-left"
                            title="Play in app"
                          >
                            {C.url}
                          </button>
                          <a
                            href={C.url}
                            target="_blank"
                            rel="noopener noreferrer"
                            className="p-1 text-gray-500 hover:text-emerald-400 flex-shrink-0"
                            title="Open in new tab"
                          >
                            <ExternalLink className="w-3 h-3" />
                          </a>
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}
              {(c.screenshots || []).length > 0 && (
                <div className="mb-3">
                  <div className="text-[10px] uppercase tracking-wider text-gray-500 mb-1">
                    Screenshots ({(c.screenshots || []).length})
                  </div>
                  <div className="grid grid-cols-2 gap-2">
                    {(c.screenshots || []).map((C, m) => (
                      <ScreenshotThumb shot={C} idx={m} onOpen={(f) => v(f)} key={C.path || C.url || m} />
                    ))}
                  </div>
                </div>
              )}
              {c.updated_by && c.updated_at && (
                <div className="text-[10px] text-gray-500 pt-2 mt-2 border-t border-gray-800">
                  Last tester edit: {c.updated_by} · {new Date(c.updated_at).toLocaleString()}
                </div>
              )}
            </div>
            <div className="bg-gray-900/40 border border-gray-800 rounded-sm p-4">
              <div className="text-xs uppercase tracking-wider text-emerald-400 font-semibold mb-3">
                Dev notes
              </div>
              <div className="mb-4">
                <div className="flex items-baseline justify-between mb-1">
                  <div className="text-[10px] uppercase tracking-wider text-gray-500">
                    Your note{" "}
                    <span className="normal-case text-gray-600">— visible to testers and other devs</span>
                  </div>
                  <SaveIndicator saving={!!l} />
                </div>
                <textarea
                  value={_}
                  onChange={(C) => {
                    let m = prettyArrows(C.target.value);
                    S(m);
                    u(m);
                  }}
                  rows={3}
                  className="w-full px-3 py-2 bg-gray-950 border border-gray-800 text-gray-100 placeholder-gray-600 rounded-sm text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500 resize-y"
                />
              </div>
              {k.length > 0 && (
                <div>
                  <div className="text-[10px] uppercase tracking-wider text-gray-500 mb-2">
                    Other dev activity ({k.length})
                  </div>
                  <div className="space-y-2">
                    {k.map((C) => (
                      <DevMarkCard mark={C} key={C.developer_name} />
                    ))}
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>
      </div>
      {I && <RecordingPlayer recordings={I.recordings} startIdx={I.startIdx} onClose={() => x(null)} />}
      {g !== null && (c.screenshots || []).length > 0 && (
        <ScreenshotLightbox screenshots={c.screenshots || []} startIdx={g} onClose={() => v(null)} />
      )}
    </div>
    </FigureContext.Provider>
  );
}
var root = ReactDOM.createRoot(document.getElementById("root"));
root.render(<App />);
