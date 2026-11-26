import copy, json, os
from datetime import datetime, timezone
from google import genai
from google.genai import types

def get_wb():
    return {"program": "Meridian Retail Bank Enterprise AI Transformation", "approved_records": {"milestones": [{"id": "MS-01", "name": "Core Banking Cloud Data Lake Ingestion", "status": "COMPLETED", "owner": "Sarah Jenkins", "evidence_id": "EVD-2026-089"}], "change_requests": []}, "incoming_updates": [{"record_type": "milestones", "id": "MS-02", "name": "Credit Scoring Model Validation", "status": "ON_TRACK", "owner": "Dr. Aris Thorne", "owner_confirmed": True, "evidence_id": "EVD-2026-114"}, {"record_type": "change_requests", "id": "CR-04", "title": "Decommission Legacy Mainframe Fraud Engine", "status": "PROPOSED", "owner": None, "owner_confirmed": False, "evidence_id": None}], "audit_log": []}

def agent_0(wb):
    ts = datetime.now(timezone.utc).isoformat()
    q = []
    for u in wb["incoming_updates"]:
        if u.get("owner_confirmed") is True and u.get("evidence_id"):
            wb["approved_records"][u["record_type"]].append({k:v for k,v in u.items() if k!="record_type"})
            wb["audit_log"].append({"ts": ts, "actor": "AGENT_0", "action": "ACCEPTED", "id": u["id"], "evidence_id": u["evidence_id"]})
        else:
            q.append({"id": u["id"], "status": "QUARANTINED_UNCONFIRMED"})
            wb["audit_log"].append({"ts": ts, "actor": "AGENT_0", "action": "REJECTED", "id": u["id"], "evidence_id": "UNKNOWN"})
    wb["incoming_updates"] = []
    wb["quarantined"] = q
    return wb

def ask_agent(c, role, task, wb):
    for m in [x.name.replace("models/","") for x in c.models.list() if "gemini" in x.name and "embed" not in x.name]:
        try:
            r = c.models.generate_content(model=m, contents=f"Role: {role}. Task: {task}. STRICT READ-ONLY. Mark missing data as UNKNOWN. Data: {json.dumps(copy.deepcopy(wb))}", config=types.GenerateContentConfig(temperature=0.1, response_mime_type="application/json"))
            return json.loads(r.text)
        except Exception:
            continue
    return {"error": "No available model"}

if __name__ == "__main__":
    c = genai.Client(api_key=os.environ.get("GEMINI_API_KEY"))
    wb = agent_0(get_wb())
    out = {"PMO": ask_agent(c, "PMO Analyst", "Analyze schedule and milestones", wb), "BA": ask_agent(c, "Business Analyst", "Analyze scope and CRs", wb), "Governance": ask_agent(c, "Governance Officer", "Verify audit_log and SoD", wb)}
    signoffs = {"PMO": True, "BA": False, "Governance": True}
    pack = {"title": "Meridian Steering Committee Pack", "sections": {k: (v if signoffs[k] else "PENDING OWNER VALIDATION") for k,v in out.items()}, "audit_log": wb["audit_log"]}
    print(json.dumps(pack, indent=2))
