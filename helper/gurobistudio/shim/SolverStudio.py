"""What `from SolverStudio import *` provides. The runner fills this in before the model runs."""


class _Unsupported:
    def __getattr__(self, name):
        raise NotImplementedError(
            "The Excel Application object is not supported. "
            "Read and write the sheet through data items instead."
        )


Application = _Unsupported()
__all__ = ["Application"]


def _install(items):
    globals().update(items)
    __all__.extend(items)
