#!/usr/bin/env python3
"""
LODHA GCR — ROUTINE BRIDGE CLIENT (no service account, no Google Cloud
project, no Drive/Sheets sharing at all)

Your Workspace admin blocks sharing items with external accounts —
including a service account's iam.gserviceaccount.com identity — and that
policy cannot be relaxed. This script routes around it entirely: it never
talks to Google's APIs directly. It talks to your own already-deployed
Apps Script web app over plain HTTPS, authenticated by a secret token you
generate yourself. Nothing is shared with anyone; a token is not a Drive
permission, so the policy that blocks external sharing does not apply.

The web app already runs inside your domain (deployed as Execute as: Me),
so every read and write it does — reading the queue, computing the rating,
building the evidence workbook — happens with the deploying user's own
access. This script only carries JSON back and forth.

CLI (each subcommand prints one JSON object to stdout, mirroring the
previous gates_engine.py contract so ROUTINE_PROMPT.md needs no changes):

  read_queue
  claim --row N
  evaluate --sub '<json>' --anchor '<json>' --ledger '<json>'
  workbook --sub '<json>' --anchor '<json>' --ledger '<json>' --notes '<json>'
  writeback --row N --verdict V --confidence C --summary S --reasoning R
             --gaps '["...","..."]' --escalations E --evidence_url U
  fail --row N --reason "..."

ENVIRONMENT
  GCR_APP_URL     the web app's /exec URL (Deploy > Manage deployments)
  GCR_TOKEN       must match the ROUTINE_TOKEN Script Property in Code.gs

  pip install requests
  (no google-api-python-client, no google-auth, no openpyxl — the bridge
  does all of that work on the Apps Script side)
"""

import argparse, json, os, sys
import requests

APP_URL = os.environ.get("GCR_APP_URL", "")
TOKEN   = os.environ.get("GCR_TOKEN", "")

def call(action, **fields):
    if not APP_URL:
        print("GCR_APP_URL is not set", file=sys.stderr); sys.exit(2)
    if not TOKEN:
        print("GCR_TOKEN is not set", file=sys.stderr); sys.exit(2)
    payload = {"action": action, "token": TOKEN}
    payload.update(fields)
    resp = requests.post(APP_URL, json=payload, timeout=90)
    resp.raise_for_status()
    out = resp.json()
    if not out.get("ok", True) and "error" in out:
        print(json.dumps(out), file=sys.stderr)
        sys.exit(1)
    return out

def jarg(s, default=None):
    if not s:
        return default
    return json.loads(s)

def cmd_read_queue(a):
    print(json.dumps(call("read_queue")))

def cmd_claim(a):
    print(json.dumps(call("claim", row=a.row)))

def cmd_evaluate(a):
    out = call("evaluate", sub=jarg(a.sub, {}), anchor=jarg(a.anchor, None), ledger=jarg(a.ledger, {}))
    print(json.dumps(out))

def cmd_workbook(a):
    out = call("workbook", sub=jarg(a.sub, {}), anchor=jarg(a.anchor, {}),
               ledger=jarg(a.ledger, {}), notes=jarg(a.notes, {}))
    print(json.dumps(out))

def cmd_writeback(a):
    out = call("writeback", row=a.row, verdict=a.verdict, confidence=a.confidence,
               summary=a.summary, reasoning=a.reasoning, gaps=jarg(a.gaps, []),
               escalations=a.escalations, evidence_url=a.evidence_url,
               reference=a.reference or "")
    print(json.dumps(out))

def cmd_fail(a):
    print(json.dumps(call("fail", row=a.row, reason=a.reason)))

def main():
    p = argparse.ArgumentParser()
    sub = p.add_subparsers(dest="cmd", required=True)
    sub.add_parser("read_queue").set_defaults(fn=cmd_read_queue)
    c = sub.add_parser("claim"); c.add_argument("--row", type=int, required=True); c.set_defaults(fn=cmd_claim)
    c = sub.add_parser("evaluate")
    c.add_argument("--sub", required=True); c.add_argument("--anchor", default="")
    c.add_argument("--ledger", required=True); c.set_defaults(fn=cmd_evaluate)
    c = sub.add_parser("workbook")
    c.add_argument("--sub", required=True); c.add_argument("--anchor", default="")
    c.add_argument("--ledger", required=True); c.add_argument("--notes", default="")
    c.set_defaults(fn=cmd_workbook)
    c = sub.add_parser("writeback")
    c.add_argument("--row", type=int, required=True); c.add_argument("--verdict", required=True)
    c.add_argument("--confidence", default=""); c.add_argument("--summary", required=True)
    c.add_argument("--reasoning", required=True); c.add_argument("--gaps", default="[]")
    c.add_argument("--escalations", default=""); c.add_argument("--evidence_url", default="")
    c.add_argument("--reference", default=""); c.set_defaults(fn=cmd_writeback)
    c = sub.add_parser("fail"); c.add_argument("--row", type=int, required=True)
    c.add_argument("--reason", required=True); c.set_defaults(fn=cmd_fail)
    a = p.parse_args(); a.fn(a)

if __name__ == "__main__":
    main()
