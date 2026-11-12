import { AlertTriangle, Check, CircleDashed, Loader2, XCircle } from 'lucide-react';
import type { RedlineTab, TabState } from '../tabs/model.js';

/**
 * The Claude-Cowork-style task checklist.
 *
 * Its whole job is making "comparing document 37 of 100" legible: one row per
 * comparison, with its own state, so a long batch is readable at a glance rather
 * than hidden behind a single global spinner.
 */
export function TaskPlan({ tabs }: { tabs: TabState }) {
  const jobs = tabs.tabs.filter((t): t is RedlineTab => t.kind === 'redline');

  if (jobs.length === 0) {
    return (
      <div className="flex items-center gap-2 px-3 pb-3 text-[11px] text-text-faint">
        <CircleDashed size={12} aria-hidden />
        <span>Nothing running.</span>
      </div>
    );
  }

  const done = jobs.filter((j) => j.status.state === 'ready').length;
  const failed = jobs.filter((j) => j.status.state === 'failed').length;
  const running = jobs.filter((j) => j.status.state === 'running').length;

  return (
    <div className="px-3 pb-3">
      <p className="pb-1.5 text-[11px] text-text-muted">
        {running > 0
          ? `Comparing ${done + 1} of ${jobs.length}`
          : `${done} of ${jobs.length} complete`}
        {failed > 0 ? ` · ${failed} failed` : ''}
      </p>

      <ul className="flex flex-col gap-1">
        {jobs.map((job) => (
          <li key={job.id} className="flex items-start gap-1.5 text-[11px] leading-snug">
            <span className="mt-0.5 shrink-0">
              <StatusIcon job={job} />
            </span>
            <span className="min-w-0 flex-1">
              <span
                className={
                  'block truncate ' +
                  (job.status.state === 'ready' ? 'text-text-muted' : 'text-text')
                }
                title={job.title}
              >
                {job.title}
              </span>
              {job.status.state === 'running' ? (
                <span className="block text-text-faint">{job.status.step}</span>
              ) : null}
              {job.status.state === 'failed' ? (
                <span className="block text-danger">{job.status.message}</span>
              ) : null}
            </span>
          </li>
        ))}
      </ul>
    </div>
  );
}

function StatusIcon({ job }: { job: RedlineTab }) {
  switch (job.status.state) {
    case 'running':
      return <Loader2 size={12} className="animate-spin text-accent" aria-hidden />;
    case 'ready':
      return <Check size={12} className="text-redline-ins" aria-hidden />;
    case 'failed':
      return <AlertTriangle size={12} className="text-danger" aria-hidden />;
    case 'cancelled':
      return <XCircle size={12} className="text-text-faint" aria-hidden />;
  }
}
