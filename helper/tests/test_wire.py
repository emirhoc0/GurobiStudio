import math

import pytest

from gurobistudio.wire import NotSerializable, decode, encode


def test_decode_shapes():
    assert decode(3) == 3 and decode(2.5) == 2.5 and decode("a") == "a" and decode(True) is True
    assert decode(["a", "b"]) == ["a", "b"]
    assert decode([["a", 1], ["b", 2]]) == [("a", 1), ("b", 2)]
    assert decode({"__dict__": [["a", 1]]}) == {"a": 1}
    assert decode({"__dict__": [[["a", "x"], 1]]}) == {("a", "x"): 1}
    assert decode({"__dict__": [["a", [1, 2]]]})["a"] == [1, 2]


def test_round_trip_with_tuple_keys():
    value = {("a", 1): 2.5, ("b", 2): "text"}
    assert decode(encode(value)) == value
    assert decode(encode([("a", "b")])) == [("a", "b")]


def test_encode_edge_values():
    assert encode(math.inf) == "inf"
    assert encode({"a": math.nan}) == {"__dict__": [["a", "nan"]]}
    assert encode(None) is None


def test_encode_rejects_objects():
    with pytest.raises(NotSerializable):
        encode(object())
    with pytest.raises(NotSerializable):
        encode({"a": object()})
