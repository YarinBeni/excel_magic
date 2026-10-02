#!/bin/bash
# runner.sh — git-driven remote control for the cluster.
#
# WHY. The cluster is reachable only through a browser terminal, so every
# command had to be copy-pasted by hand and every log screenshotted back.
# This loop turns the sandbox branch into a command channel instead:
#
#   agent/inbox/NNN_name.sh   a script pushed from the outside
#   agent/outbox/NNN_name.out its stdout+stderr, pushed back
#   agent/done/NNN_name.sh    the script after it ran (moved)
#
# Every ~60 s: fetch + rebase the sandbox branch, run any inbox script
# not yet in done/ (sorted, one at a time, under a timeout), commit the
# output, push. A heartbeat file is pushed once an hour so the outside
# can tell "no work" from "runner dead".
#
# SECURITY. Anything committed to inbox/ on this branch runs as the
# cluster user. Only people with push access to the repo can do that;
# keep it that way. Scripts run with `bash -e`, working dir = repo root,
# conda env `thesis` active.
#
# Start once (from the repo root) as a long sbatch job:
#     sbatch sbatch/AGENT_runner.sbatch
# or on the login node:      nohup bash agent/runner.sh &
# Stop: touch agent/STOP (from anywhere) and push, or scancel.

set -uo pipefail
cd "$(dirname "$0")/.." || exit 1
REPO="$PWD"
AGENT="agent"
POLL="${AGENT_POLL:-60}"
HEARTBEAT_EVERY="${AGENT_HEARTBEAT:-3600}"
DEFAULT_TIMEOUT="${AGENT_TIMEOUT:-1800}"

if [ -f ~/miniconda3/etc/profile.d/conda.sh ]; then
    source ~/miniconda3/etc/profile.d/conda.sh && conda activate thesis
fi

log() { echo "[runner $(date -u +%H:%M:%S)] $*"; }

# One line saying WHY a git fetch/push failed (expired token? network?
# rejected?), pulled from git's own chatter. Pushes fail silently otherwise:
# on 2026-09-03 an expired token stopped every push from the cluster for
# hours while jobs kept finishing, and nothing on the branch said so.
git_reason() {
    printf '%s\n' "$1" | grep -iE 'fatal|error|remote:|denied|auth|could not|timed out|unable|reject' \
        | tail -2 | tr '\n' ' ' | cut -c1-300
}

self_sha() { sha1sum "$AGENT/runner.sh" | cut -c1-12; }
START_SHA="$(self_sha)"
last_beat=0

while true; do
    # re-source each loop so fixes to the git helpers take effect live
    source sbatch/_sandbox.sh
    if ! synclog="$(sync_sandbox 2>&1)"; then
        log "sync failed; retry next tick -- $(git_reason "$synclog")"; sleep "$POLL"; continue
    fi

    if [ -f "$AGENT/STOP" ]; then
        log "STOP file present; exiting"
        git rm -q --cached "$AGENT/STOP" 2>/dev/null; rm -f "$AGENT/STOP"
        commit_push_sandbox "agent: stopped" "$AGENT" >/dev/null 2>&1
        exit 0
    fi

    # runner.sh updated by a push -> restart into the new version
    if [ "$(self_sha)" != "$START_SHA" ]; then
        log "runner.sh changed; re-exec"
        exec bash "$AGENT/runner.sh"
    fi

    ran=0
    for job in $(ls "$AGENT"/inbox/*.sh 2>/dev/null | sort); do
        name="$(basename "$job" .sh)"
        [ -e "$AGENT/done/$name.sh" ] && continue
        out="$AGENT/outbox/$name.out"
        # per-script timeout override: a line `# TIMEOUT=7200`
        t="$(sed -n 's/^# *TIMEOUT=\([0-9]*\).*/\1/p' "$job" | head -1)"
        t="${t:-$DEFAULT_TIMEOUT}"
        log "run $name (timeout ${t}s)"
        {
            echo "# $name  host=$(hostname)  start=$(date -u)  timeout=${t}s"
            echo "# HEAD=$(git rev-parse --short HEAD)"
            echo "# ---------------------------------------------------------"
            timeout "$t" bash -e "$job" 2>&1
            rc=$?
            echo "# ---------------------------------------------------------"
            echo "# exit=$rc  end=$(date -u)"
        } > "$out" 2>&1
        git mv -f "$job" "$AGENT/done/$name.sh" 2>/dev/null \
            || mv -f "$job" "$AGENT/done/$name.sh"
        msg="agent: $name (exit $(tail -1 "$out" | sed 's/.*exit=\([0-9]*\).*/\1/'))"
        if ! pushlog="$(commit_push_sandbox "$msg" "$AGENT" 2>&1)"; then
            # the commit is kept locally and rides along with the next push
            # that succeeds; say WHY it failed so `tail` of this log is enough
            log "push failed for $name -- $(git_reason "$pushlog")"
        fi
        ran=1
    done

    now=$(date +%s)
    if [ "$ran" -eq 0 ] && [ $((now - last_beat)) -ge "$HEARTBEAT_EVERY" ]; then
        { echo "alive $(date -u)"; echo "host $(hostname)";
          echo "job ${SLURM_JOB_ID:-login}"; squeue -u "$USER" -h 2>/dev/null; } \
            > "$AGENT/heartbeat"
        if ! pushlog="$(commit_push_sandbox "agent: heartbeat" "$AGENT/heartbeat" 2>&1)"; then
            log "heartbeat push failed -- $(git_reason "$pushlog")"
        fi
        last_beat=$now
    fi
    sleep "$POLL"
done
