# _sandbox.sh — concurrency-robust git sync for the chrono sbatch jobs,
# adapted from v_1/src/stress_tests/sbatch/_common.sh but targeting the
# SANDBOX branch: chrono cluster jobs sync and push yarin-sandbox, NEVER
# main (SLA section 8). Same machinery: all git mutations serialized
# under an flock, rebase always onto the single ref FETCH_HEAD after an
# explicit fetch, --autostash so a dirty NFS tree never aborts, stale
# rebase state / index.lock from killed jobs cleared first.
#
# artifacts is gitignored, so jobs must copy the small results
# (results.parquet, gate reports) into reports/ — which IS
# tracked — before commit_push_sandbox.
#
# Source it AFTER cd-ing into the repo:
#     source sbatch/_sandbox.sh
#     sync_sandbox
#     ... produce results ...
#     commit_push_sandbox "C?: message (job ${SLURM_JOB_ID})" chrono/reports

GIT_LOCK="${GIT_LOCK:-$HOME/.tabfm_git.lock}"
SANDBOX_BRANCH="${SANDBOX_BRANCH:-claude/tabular-model-agent-exp-wvni5o}"

_git_abort_inflight() {
    git rebase --abort 2>/dev/null || true
    git merge  --abort 2>/dev/null || true
    # a corrupt rebase dir from a killed job survives a failed --abort
    # and blocks every later rebase; we hold the lock, so any state
    # still present here is stale — remove it outright.
    rm -rf .git/rebase-merge .git/rebase-apply 2>/dev/null || true
    find .git -maxdepth 1 -name index.lock -mmin +2 -delete 2>/dev/null || true
}

# Bring the working tree up to date with origin/yarin-sandbox.
#
# REVIEW FIX (wave B1): this used to fetch + rebase WITHOUT ever checking
# the branch out. On a cluster clone sitting on main that rebased main
# onto the sandbox and left HEAD on main, so the later
# `git push origin HEAD:yarin-sandbox` would have pushed MAIN's history
# into the sandbox branch. It now checks out the branch explicitly and
# refuses to continue unless HEAD really is on it; failures return 1
# instead of a silent no-op, so callers can abort the job.
sync_sandbox() {
    local ok=1
    ( flock -w 300 9 || true
      _git_abort_inflight
      for _i in 1 2 3 4 5; do
          # CLUSTER FIX (jobs 32706/32723): `checkout -B ... FETCH_HEAD`
          # RESET the branch to the remote, silently discarding commits a
          # previous job had made but could not push (token expired) --
          # C1's extraction-meta commits vanished this way. Check the
          # branch out as it is (create it only if absent), then rebase
          # whatever is local onto the fetched tip; nothing is dropped.
          if git fetch origin "${SANDBOX_BRANCH}" \
             && { git checkout -q "${SANDBOX_BRANCH}" 2>/dev/null \
                  || git checkout -q -b "${SANDBOX_BRANCH}" FETCH_HEAD; } \
             && git rebase --autostash FETCH_HEAD; then
              exit 0
          fi
          _git_abort_inflight
          sleep $((RANDOM % 8 + 5))
      done
      exit 1
    ) 9>"$GIT_LOCK"
    ok=$?
    local head
    head="$(git rev-parse --abbrev-ref HEAD)"
    if [ "$ok" -ne 0 ] || [ "$head" != "${SANDBOX_BRANCH}" ]; then
        echo "[sync_sandbox] FAILED: HEAD is '${head}', wanted "\
             "'${SANDBOX_BRANCH}' (sync rc=${ok}) — refusing to run" >&2
        return 1
    fi
    return 0
}

# Stage + commit the given paths and push to yarin-sandbox — all inside
# one lock so overlapping array tasks never collide on the index.
# Usage: commit_push_sandbox "msg" path [path...]
# REVIEW FIX (wave B1): returns non-zero when the push never lands, so a
# job cannot report success having pushed nothing.
commit_push_sandbox() {
    local msg="$1"; shift
    ( flock -w 600 9 || true
      _git_abort_inflight
      git add "$@" 2>/dev/null || true
      # BIG-FILE GUARD (2026-09-03). Twice in one day a job staged a
      # multi-hundred-MB checkpoint that had been copied under
      # chrono/reports; pack-objects was then OOM-killed on the login
      # node and EVERY push from the cluster failed for hours, with the
      # results of unrelated jobs stuck behind it. Nothing this repo
      # tracks is legitimately larger than a few MB, so unstage anything
      # over the limit and say so, rather than poisoning the branch.
      local _big
      _big="$(git diff --cached --name-only --diff-filter=AM | while read -r _f; do
                  [ -f "$_f" ] && [ "$(stat -c%s "$_f" 2>/dev/null || echo 0)" -gt "${MAX_COMMIT_BYTES:-20000000}" ] \
                      && printf '%s\n' "$_f"; done)"
      if [ -n "$_big" ]; then
          echo "[commit_push_sandbox] REFUSING to commit oversized file(s):" >&2
          printf '  %s\n' $_big >&2
          git reset -q -- $_big
      fi
      if git diff --cached --quiet; then
          echo "[commit_push_sandbox] nothing to commit"
      else
          git commit -m "$msg" || echo "[commit_push_sandbox] commit failed"
      fi
      for _i in 1 2 3 4 5; do
          if git fetch origin "${SANDBOX_BRANCH}" \
             && git rebase --autostash FETCH_HEAD \
             && git push origin "HEAD:${SANDBOX_BRANCH}"; then
              echo "[commit_push_sandbox] pushed to ${SANDBOX_BRANCH}"
              exit 0
          fi
          _git_abort_inflight
          sleep $((RANDOM % 8 + 5))
      done
      echo "[commit_push_sandbox] FAILED after 5 attempts" >&2
      exit 1
    ) 9>"$GIT_LOCK"
}

# Push this job's slurm log to the sandbox branch on EVERY exit, success
# or failure. Without it the only jobs whose output reaches the repo are
# the ones that finished, i.e. exactly the ones nobody needs to read; a
# failure like job 32500 could only be diagnosed by hand-copying the log
# out of a browser terminal.
#
# Usage, right after `sync_sandbox` (which must already have run, so the
# checkout is on the sandbox branch):
#     enable_log_push
LOG_PUSH_LINES="${LOG_PUSH_LINES:-4000}"

push_job_log() {
    local rc=$?
    local out dest
    out="$(scontrol show job "${SLURM_JOB_ID:-0}" 2>/dev/null \
           | tr ' ' '\n' | sed -n 's/^StdOut=//p' | head -1)"
    [ -n "$out" ] && [ -f "$out" ] || return 0
    dest="reports/logs/${SLURM_JOB_NAME:-job}_${SLURM_JOB_ID:-0}.log"
    mkdir -p reports/logs
    {
        echo "# job ${SLURM_JOB_ID:-?} ${SLURM_JOB_NAME:-?} exit=${rc} $(date -u)"
        echo "# last ${LOG_PUSH_LINES} lines of ${out}"
        tail -n "${LOG_PUSH_LINES}" "$out"
    } > "$dest" 2>/dev/null || return 0
    commit_push_sandbox \
        "logs: ${SLURM_JOB_NAME:-job} ${SLURM_JOB_ID:-0} exit=${rc}" \
        reports/logs || true
}

enable_log_push() { trap push_job_log EXIT; }
