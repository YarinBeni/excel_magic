import pytest

from tabfm_auto.models import MODELS, is_available, parse_model_spec
from tabfm_auto.models.manifest import find
from tabfm_auto.models.weights import status


def test_parse_spec():
    assert parse_model_spec("kumo-tabular-s:n_estimators=4,device=cpu") == ("kumo-tabular-s", {"n_estimators": 4, "device": "cpu"})


def test_manifest_and_status_cover_every_model():
    for name in MODELS:
        s = status(name)
        assert s["state"]
    assert find("tabiclv2").name == "tabicl"
    with pytest.raises(KeyError):
        find("no-such-model")


def test_classical_always_available():
    assert is_available("hgb") and is_available("rf")


def test_split_model_specs_keeps_commas_inside_specs():
    from tabfm_auto.models import split_model_specs

    assert split_model_specs("tabpfn:n_estimators=4,device=cuda,tabicl:device=cuda,hgb") == [
        "tabpfn:n_estimators=4,device=cuda", "tabicl:device=cuda", "hgb"]
    assert split_model_specs("a;b:x=1,y=2;c") == ["a", "b:x=1,y=2", "c"]
