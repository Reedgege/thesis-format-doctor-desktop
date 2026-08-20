"use strict";

const state = {
  thesis: null,    // {name, data(base64)}
  template: null,  // {name, data(base64)}
  profile: null,   // {name, data(base64)}  抽取后的模板画像
};

const $ = (id) => document.getElementById(id);

function readFile(file) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => {
      const b64 = (r.result.split(",")[1] || "");
      resolve({ name: file.name, data: b64 });
    };
    r.onerror = reject;
    r.readAsDataURL(file);
  });
}

// ---- 文件选择 ----
$("fileThesis").addEventListener("change", async (e) => {
  const f = e.target.files[0];
  if (!f) return;
  state.thesis = await readFile(f);
  $("thesisName").textContent = f.name;
});
$("fileTemplate").addEventListener("change", async (e) => {
  const f = e.target.files[0];
  if (!f) return;
  state.template = await readFile(f);
  $("templateName").textContent = f.name;
});

// 拖拽
["dropThesis", "dropTemplate"].forEach((id) => {
  const el = $(id);
  const input = id === "dropThesis" ? $("fileThesis") : $("fileTemplate");
  el.addEventListener("dragover", (e) => { e.preventDefault(); el.classList.add("over"); });
  el.addEventListener("dragleave", () => el.classList.remove("over"));
  el.addEventListener("drop", async (e) => {
    e.preventDefault(); el.classList.remove("over");
    const f = e.dataTransfer.files[0];
    if (!f) return;
    const data = await readFile(f);
    if (id === "dropThesis") { state.thesis = data; $("thesisName").textContent = f.name; }
    else { state.template = data; $("templateName").textContent = f.name; }
  });
});

// 画像清除
$("clearProfile").addEventListener("click", () => {
  state.profile = null;
  $("profileBox").hidden = true;
});

// ---- 状态显示 ----
function setStatus(kind, text) {
  const el = $("status");
  el.className = "status " + kind;
  el.textContent = text;
}
function escapeHtml(s) {
  return (s || "").replace(/[&<>]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;" }[c]));
}

// ---- 通用请求 ----
async function postJSON(route, body) {
  const resp = await fetch(route, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  return resp.json();
}

function lockButtons(locked) {
  ["btnCheck", "btnFix", "btnRef", "btnProfile"].forEach((b) => {
    $(b).disabled = locked;
  });
}

// ---- 结果渲染 ----
function renderCheck(r) {
  const res = r.result || {};
  const sev = res.severity_counts || {};
  const high = sev["高"] || 0, mid = sev["中"] || 0, low = sev["低"] || 0;
  const noneCls = high + mid + low === 0 ? "none" : "";
  let html = `<div class="sev-row">
    <div class="sev-card high ${high ? "" : noneCls}"><b>${high}</b><span>高危</span></div>
    <div class="sev-card mid ${mid ? "" : noneCls}"><b>${mid}</b><span>中危</span></div>
    <div class="sev-card low ${low ? "" : noneCls}"><b>${low}</b><span>低危</span></div>
  </div>`;
  html += `<div class="meta">
    <span>正文段落：${res.para_count ?? "-"}</span>
    <span>表格：${res.table_count ?? "-"}</span>
    <span>图片：${res.image_count ?? "-"}</span>
    <span>标题样式：内置 ${res.heading_builtin ?? 0} / 自定义 ${res.heading_custom ?? 0}</span>
    ${res.template_driven ? '<span>📌 模板驱动诊断</span>' : '<span>📋 通用规范体检</span>'}
  </div>`;

  const issues = res.issues || [];
  if (issues.length === 0) {
    html += `<p class="meta">✅ 未发现明显格式红线问题。</p>`;
  } else {
    html += `<div class="section-h">一、问题清单（按严重度）</div>`;
    for (const it of issues) {
      const cls = it.severity === "高" ? "high" : it.severity === "中" ? "mid" : "low";
      html += `<div class="issue ${cls}"><span class="tag">${escapeHtml(it.severity)}</span>${escapeHtml(it.msg)}</div>`;
    }
  }
  const struct = res.structural_pages || [];
  if (struct.length) {
    html += `<div class="section-h">二、结构页检查（只读，建议对照模板手动确认）</div>`;
    for (const it of struct) {
      const cls = it.severity === "高" ? "high" : it.severity === "中" ? "mid" : "low";
      html += `<div class="issue ${cls}"><span class="tag">${escapeHtml(it.severity)}</span>${escapeHtml(it.msg)}</div>`;
    }
  }
  if (r.report) {
    html += `<div class="dl"><a href="${r.report}" download="格式体检报告.md">⬇ 下载 Markdown 报告</a></div>`;
  }
  $("output").innerHTML = html;
}

function renderSummary(r, title) {
  let html = `<div class="section-h">${title}</div>`;
  if (r.summary) html += `<pre class="summary">${escapeHtml(r.summary)}</pre>`;
  if (r.files && Object.keys(r.files).length) {
    html += `<div class="dl">`;
    const labels = { fixed: "⬇ 下载处理后的论文.docx", report: "⬇ 下载修改明细.docx", profile: "⬇ 下载格式画像.json" };
    for (const [k, url] of Object.entries(r.files)) {
      html += `<a href="${url}" download>${labels[k] || "⬇ 下载 " + k}</a>`;
    }
    html += `</div>`;
  }
  $("output").innerHTML = html;
}

function renderProfile(r) {
  let html = `<div class="section-h">模板格式画像</div>`;
  if (r.summary) html += `<pre class="summary">${escapeHtml(r.summary)}</pre>`;
  if (r.profile) {
    html += `<div class="section-h">画像 JSON（已自动留存，将用于后续诊断/套用）</div>`;
    html += `<pre class="code">${escapeHtml(JSON.stringify(r.profile, null, 2))}</pre>`;
  }
  if (r.files && r.files.profile) {
    html += `<div class="dl"><a href="${r.files.profile}" download>⬇ 下载 profile.json</a></div>`;
  }
  $("output").innerHTML = html;
}

// ---- 操作 ----
$("btnCheck").addEventListener("click", async () => {
  if (!state.thesis) { setStatus("err", "请先选择待处理论文 .docx"); return; }
  lockButtons(true);
  setStatus("run", "正在体检…");
  try {
    const body = { file: state.thesis, profile: state.profile || state.template };
    const r = await postJSON("/api/check", body);
    if (!r.ok) { setStatus("err", "体检失败：" + (r.error || "")); $("output").innerHTML = `<pre class="summary">${escapeHtml(r.detail || "")}</pre>`; }
    else { setStatus("ok", "体检完成"); renderCheck(r); }
  } catch (e) {
    setStatus("err", "请求异常：" + e.message);
  } finally { lockButtons(false); }
});

$("btnFix").addEventListener("click", async () => {
  if (!state.thesis) { setStatus("err", "请先选择待处理论文 .docx"); return; }
  lockButtons(true);
  setStatus("run", "正在套标题 / 出目录…（大文档稍慢）");
  try {
    const body = { file: state.thesis, profile: state.profile || state.template, make_report: true, no_comments: false };
    const r = await postJSON("/api/fix", body);
    if (!r.ok) { setStatus("err", "处理失败：" + (r.error || "")); $("output").innerHTML = `<pre class="summary">${escapeHtml(r.detail || "")}</pre>`; }
    else { setStatus("ok", "处理完成，可下载结果"); renderSummary(r, "套标题 & 出目录 · 结果"); }
  } catch (e) {
    setStatus("err", "请求异常：" + e.message);
  } finally { lockButtons(false); }
});

$("btnRef").addEventListener("click", async () => {
  if (!state.thesis) { setStatus("err", "请先选择待处理论文 .docx"); return; }
  lockButtons(true);
  setStatus("run", "正在重排参考文献…");
  try {
    const body = { file: state.thesis };
    const r = await postJSON("/api/reformat", body);
    if (!r.ok) { setStatus("err", "处理失败：" + (r.error || "")); $("output").innerHTML = `<pre class="summary">${escapeHtml(r.detail || "")}</pre>`; }
    else { setStatus("ok", "重排完成，可下载结果"); renderSummary(r, "参考文献重排 · 结果"); }
  } catch (e) {
    setStatus("err", "请求异常：" + e.message);
  } finally { lockButtons(false); }
});

$("btnProfile").addEventListener("click", async () => {
  const src = state.template || state.thesis;
  if (!src) { setStatus("err", "请先选择学校模板 .docx"); return; }
  lockButtons(true);
  setStatus("run", "正在抽取模板格式画像…");
  try {
    const body = { file: src };
    const r = await postJSON("/api/profile", body);
    if (!r.ok) { setStatus("err", "抽取失败：" + (r.error || "")); $("output").innerHTML = `<pre class="summary">${escapeHtml(r.detail || "")}</pre>`; }
    else {
      state.profile = src;  // 自动留存画像供后续使用
      $("profileBox").hidden = false;
      setStatus("ok", "画像已抽取并留存"); renderProfile(r);
    }
  } catch (e) {
    setStatus("err", "请求异常：" + e.message);
  } finally { lockButtons(false); }
});

// 引擎加载自检
window.addEventListener("DOMContentLoaded", () => {
  setStatus("idle", "等待操作…（文件全程在本地处理，不会上传）");
  $("engineBadge").textContent = "就绪 · 本地引擎";
});
