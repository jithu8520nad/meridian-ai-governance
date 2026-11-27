from datetime import date
import pytest
from record_review import decision, open_items, ACCEPT, RETURN, REJECT

def entry(outcome, **kwargs):
    args=dict(sheet="Changes",record={"ID":"CHG-005"},outcome=outcome,reviewer="Reviewer",
              reason="Check cost basis",owner="Finance",due=date(2026,12,4),
              workbook_hash="hash",context="current")
    return decision(**(args | kwargs))

def test_queue_holds_only_affected_context_and_keeps_history():
    history=[entry(RETURN),entry(REJECT, context="other")]
    assert len(open_items(history,"current"))==1
    assert "Finance" in history[0]["affected_sections"]
    history.append(entry(ACCEPT))
    assert open_items(history,"current")==[]
    assert len(history)==3

def test_rejection_needs_a_correction_owner_and_due_date():
    with pytest.raises(ValueError):
        entry(REJECT, owner="")
    with pytest.raises(ValueError):
        entry(RETURN, due=None)

def test_changed_record_needs_separate_disposition():
    assert len(open_items([entry(REJECT),entry(ACCEPT,record={"ID":"CHG-005","Cost":2})],"current"))==1

@pytest.mark.parametrize('tab_index', [2, 3, 4, 5])
def test_streamlit_reject_then_accept(tmp_path, monkeypatch, tab_index):
    monkeypatch.setenv('MERIDIAN_REVIEW_DB', str(tmp_path / 'review.sqlite3'))
    from streamlit.testing.v1 import AppTest
    app=AppTest.from_file("app.py",default_timeout=30).run()
    assert not app.exception
    assert len(app.get("download_button")) > 1
    def choose(outcome):
        panel=app.tabs[tab_index]
        next(w for w in panel.selectbox if w.label=="Decision").select(outcome)
        next(w for w in panel.text_input if w.label=="Simulated reviewer").input("Reviewer")
        next(w for w in panel.text_area if w.label=="Reason / evidence reference").input("Reconciled against ledger")
        next(w for w in panel.text_input if w.label=="Correction owner").input("Finance")
        next(w for w in panel.date_input if w.label=="Correction due").set_value(date(2026,12,4))
        next(w for w in panel.button if w.label=="Record decision").click()
        app.run()
        assert not app.exception
    choose(REJECT)
    assert any("exports are held" in w.value for w in app.warning)
    assert len(app.get("download_button")) == 1
    panel=app.tabs[5]
    next(w for w in panel.text_input if w.label=="Closure reviewer (simulated)").input("Reviewer")
    next(w for w in panel.text_area if w.label=="Correction evidence or documented override reference").input("Evidence-001")
    next(w for w in panel.checkbox if w.label=="I confirm the correction or authorised override has been verified").check()
    next(w for w in panel.button if w.label=="Close exception").click()
    app.run()
    assert not app.exception
    assert not any("exports are held" in w.value for w in app.warning)
    assert len(app.get("download_button")) > 1
    assert len(app.session_state["record_review_history"]) == 2


def test_persistent_exception_and_explicit_closure(tmp_path, monkeypatch):
    from record_review import save_decision, load_history, unresolved_items
    monkeypatch.setenv("MERIDIAN_REVIEW_DB", str(tmp_path / "review.sqlite3"))
    rejected = entry(REJECT)
    save_decision(rejected)
    assert unresolved_items(load_history()) == [rejected]
    with pytest.raises(ValueError):
        save_decision(entry(ACCEPT, context="new", workbook_hash="new"))
    save_decision(entry(ACCEPT, context="new", workbook_hash="new"), closure=rejected["record_hash"])
    assert not unresolved_items(load_history())
    assert len(load_history()) == 2
