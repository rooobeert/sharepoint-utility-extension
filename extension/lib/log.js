// Plain-text and JSON forms of a run log. The run object only holds results
// (no request bodies, cookies or digests), so both are safe to share.

const LABEL = { success: 'OK  ', warning: 'WARN', failed: 'FAIL', throttled: 'FAIL' };

/** Show the part of an operation URL after the site URL. */
export function shortUrl(url, site) {
  if (!url) return '(no URL)';
  const s = (site || '').toLowerCase();
  return s && url.toLowerCase().startsWith(s) ? url.slice(s.length) : url;
}

export function stageSummaryText(stage) {
  if (stage.status === 'skipped') return 'skipped';
  if (!stage.summary) return 'not sent';
  const { succeeded, warnings, failed } = stage.summary;
  return `${succeeded} succeeded, ${failed} failed` + (warnings ? `, ${warnings} already existed` : '');
}

export function formatRunText(run) {
  const lines = [
    'SharePoint batch run',
    `Site:     ${run.site}`,
    `Started:  ${run.startedAt}`,
    `Finished: ${run.finishedAt || '(running)'}`,
  ];
  run.stages.forEach((stage, i) => {
    lines.push('', `== Stage ${i + 1}: ${stage.name} [${stage.status}] ${stageSummaryText(stage)}`);
    if (stage.error) lines.push(`ERROR: ${stage.error}`);
    for (const w of stage.warnings || []) lines.push(`Warning: ${w}`);
    for (const n of stage.notes || []) lines.push(`Note: ${n}`);
    for (const p of stage.placeholders || []) {
      lines.push(`Placeholder ${p.label} = ${p.id || 'UNRESOLVED (' + p.error + ')'}`);
    }
    for (const w of stage.waits || []) {
      lines.push(`Waited ${Math.round(w.ms / 1000)} s: ${w.reason} (retry ${w.retry} of ${w.of})`);
    }
    for (const op of stage.operations || []) {
      const status = op.status == null ? '---' : op.status;
      const at = op.line ? `line ${op.line}`.padEnd(9) + ' ' : '';
      let line = `[${LABEL[op.outcome] || op.outcome}] ${String(op.index + 1).padStart(3)} ${at}${status} ${op.method} ${shortUrl(op.url, run.site)}`;
      if (op.attempts > 1) line += ` (attempt ${op.attempts})`;
      if (op.outcome !== 'success' && op.errorMessage) line += `\n        ${op.errorMessage}`;
      lines.push(line);
    }
  });
  return lines.join('\n') + '\n';
}

export function formatRunJson(run) {
  return JSON.stringify(run, null, 2);
}
