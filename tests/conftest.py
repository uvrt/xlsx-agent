"""Test options: ``--run-slow`` also runs the tests marked ``slow`` (as docx-agent's)."""

import pytest


def pytest_addoption(parser):
    parser.addoption("--run-slow", action="store_true", default=False,
                     help="also run the tests marked slow (the 200,000-row memory measurement)")


def pytest_collection_modifyitems(config, items):
    if config.getoption("--run-slow"):
        return
    later = pytest.mark.skip(reason="slow: run with --run-slow")
    for item in items:
        if item.get_closest_marker("slow") is not None:
            item.add_marker(later)
