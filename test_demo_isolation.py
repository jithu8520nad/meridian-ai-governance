import streamlit as st
import record_review as rr

def test_demo_visitors_are_isolated(monkeypatch, tmp_path):
    import tempfile
    monkeypatch.setenv("MERIDIAN_PUBLIC_DEMO", "1")
    monkeypatch.setattr(tempfile, "gettempdir", lambda: str(tmp_path))
    first, second = {}, {}
    monkeypatch.setattr(st, "session_state", first)
    path_a = rr.database_path()
    item = {"record_hash": "demo-record", "outcome": rr.REJECT}
    rr.save_decision(item)
    assert rr.load_history() == [item]
    assert rr.database_path() == path_a
    monkeypatch.setattr(st, "session_state", second)
    assert rr.database_path() != path_a
    assert rr.load_history() == []
    monkeypatch.setattr(st, "session_state", first)
    assert rr.load_history() == [item]
