"""JSON form of data items. JSON can't hold tuple keys, so dicts travel as {"__dict__": [[k, v]]}."""

import math

DICT_TAG = "__dict__"


class NotSerializable(TypeError):
    pass


def _tuple_if_list(v):
    return tuple(_tuple_if_list(x) for x in v) if isinstance(v, list) else v


def decode(value):
    if isinstance(value, dict):
        return {_tuple_if_list(k): decode(v) for k, v in value[DICT_TAG]}
    if isinstance(value, list):
        return [_tuple_if_list(v) for v in value]
    return value


def _scalar(v):
    if hasattr(v, "item") and callable(v.item):  # numpy scalars
        v = v.item()
    if isinstance(v, float) and not math.isfinite(v):
        return str(v)  # inf and nan aren't valid JSON, so they reach the sheet as text
    if v is None or isinstance(v, bool | int | float | str):
        return v
    raise NotSerializable(type(v).__name__)


def _entry(v):
    return [_entry(x) for x in v] if isinstance(v, tuple | list) else _scalar(v)


def encode(value):
    if isinstance(value, dict):  # includes gurobipy.tupledict
        return {DICT_TAG: [[_entry(k), _entry(v)] for k, v in value.items()]}
    if isinstance(value, list | tuple):
        return [_entry(v) for v in value]
    return _scalar(value)
