"""Test options: ``--run-slow`` also runs the tests marked ``slow`` (as docx-agent's).  The
Excel oracle's tests (marked ``oracle``) are left out unless asked for (``pytest -m
oracle``, pyproject's ``addopts``), and never run under xdist's parallel workers: there is
one Excel (tests/oracle.py)."""

import os

import pytest


def pytest_addoption(parser):
    parser.addoption("--run-slow", action="store_true", default=False,
                     help="also run the tests marked slow (the 200,000-row memory measurement)")


def pytest_collection_modifyitems(config, items):
    if os.environ.get("PYTEST_XDIST_WORKER"):
        serial = pytest.mark.skip(reason="the Excel oracle runs serially: pytest -m oracle, "
                                  "without -n")
        for item in items:
            if item.get_closest_marker("oracle") is not None:
                item.add_marker(serial)
    if config.getoption("--run-slow"):
        return
    later = pytest.mark.skip(reason="slow: run with --run-slow")
    for item in items:
        if item.get_closest_marker("slow") is not None:
            item.add_marker(later)
