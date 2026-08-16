"""MkDocs hook that filters griffe docstring-annotation warnings.

Strict builds (``mkdocs build -s``) fail on any warning. Griffe warns when a
documented parameter or return value has no type annotation; these are style
nits in the source docstrings, not documentation errors, so they are dropped
here. All other griffe warnings still fail the strict build.
"""

import logging

_DOCSTRING_NITS = ("No type or annotation for",)


class _DropDocstringAnnotationWarnings(logging.Filter):
    def filter(self, record: logging.LogRecord) -> bool:
        if record.levelno == logging.WARNING:
            message = record.getMessage()
            return not any(nit in message for nit in _DOCSTRING_NITS)
        return True


logging.getLogger("mkdocs.plugins.griffe").addFilter(_DropDocstringAnnotationWarnings())
