"""The workbook's document properties (production R62): ``wb.properties()`` and
``wb.set_properties(...)``.

::

    wb.properties().core["title"]                   # "Stock list"
    wb.properties().custom["Project"].value         # "Alpha" (and .type: "text")
    wb.set_properties(title="Stock list", subject=None, custom={"Run": 42, "Old": None})

The core and custom properties are ooxml-edit's (:mod:`ooxml_edit.properties`), which every
library reads and writes the same way: core properties by name (``title``, ``subject``,
``creator`` -- ``author`` too --, ``keywords``, ``description``, ``last_modified_by``,
``revision``, ``created``, ``modified``, ``category``, ``content_status``, ``language``,
``identifier``, ``version``, ``last_printed``), custom ones as name -> value with their
type (text, integer, number, bool, date).  ``None`` removes one.  A change is one undoable
step (or part of the open batch), checked before anything changes, every problem listed
(``bad_property``).

**The modified date.**  Nothing here or in a save changes it on its own: a save writes
``docProps/core.xml`` as it is (its bytes, when no property was set), so ``modified`` and
``last_modified_by`` stay what they were until they are set.  Excel does otherwise
(measured, Excel for Mac 16.106, 2026-11-09): every save set ``modified`` to the time of
the save and ``lastModifiedBy`` to Excel's user name, and kept ``created``.  A host that
wants the same passes ``modified=datetime.now(timezone.utc)`` (and its user's name).

**Excel's own names** (measured the same day, through its object model): Title, Subject,
Author, Keywords, Comments and Category are these core properties (Comments is
``description``, Author is ``creator``); Manager and Company are extended properties
(``docProps/app.xml``), which are not set here.

Property values are content: a result names properties, never their values, and a
redacted workbook's properties read redacted.
"""

from __future__ import annotations

from dataclasses import dataclass, field, replace
from datetime import datetime
from typing import Any, Mapping

from ooxml_edit.properties import (CustomProperty, PropertyError, check_core, check_custom,
                                   core_changes, custom_changes, read_core, read_custom,
                                   w3cdtf, write_core, write_custom)

from . import EditError


@dataclass(frozen=True)
class Properties:
    """A workbook's properties: ``core`` by name (text; dates as aware ``datetime``;
    ``revision`` an ``int``) and ``custom`` by name (:class:`~ooxml_edit.properties.
    CustomProperty`: ``value``, ``type``, ``linked``)."""

    core: dict[str, Any] = field(default_factory=dict, repr=False)
    custom: dict[str, CustomProperty] = field(default_factory=dict, repr=False)

    def as_dict(self) -> dict[str, Any]:
        return {"core": {name: w3cdtf(value) if isinstance(value, datetime) else value
                         for name, value in self.core.items()},
                "custom": {name: prop.as_dict() for name, prop in self.custom.items()}}


@dataclass(frozen=True)
class PropertiesChange:
    """What :func:`set_properties` changed: the core and custom properties by name (never
    their values), and whether anything did (``note`` says why not)."""

    core: tuple[str, ...] = ()
    custom: tuple[str, ...] = ()
    note: str | None = None

    @property
    def changed(self) -> bool:
        return bool(self.core or self.custom)

    def as_dict(self) -> dict[str, Any]:
        out: dict[str, Any] = {"changed": self.changed, "core": list(self.core),
                               "custom": list(self.custom)}
        if self.note:
            out["note"] = self.note
        return out


def properties(workbook: Any) -> Properties:
    """The workbook's core and custom properties (see the module)."""
    core = read_core(workbook.package)
    custom = read_custom(workbook.package)
    if workbook._redaction is not None:
        core = {name: workbook._redacted(value) if isinstance(value, str) else value
                for name, value in core.items()}
        custom = {name: replace(prop, value=workbook._redacted(prop.value))
                  if isinstance(prop.value, str) else prop for name, prop in custom.items()}
    return Properties(core, custom)


def set_properties(workbook: Any, core: Mapping[str, Any],
                   custom: Mapping[str, Any] | None = None) -> PropertiesChange:
    """Set (a value) or remove (``None``) core properties by name and custom ones (see the
    module); one step.  Already so: ``changed`` false."""
    if workbook.strict:
        raise EditError("strict")
    problems: list[tuple[str, str]] = []
    checked_core: dict[str, Any] = {}
    checked_custom: dict[str, Any] = {}
    try:
        checked_core = check_core(core)
    except PropertyError as error:
        problems += error.problems
    if custom is not None and not isinstance(custom, Mapping):
        problems.append(("custom", "a mapping of name -> value (None removes one)"))
    elif custom:
        try:
            checked_custom = check_custom(workbook.package, custom)
        except PropertyError as error:
            problems += [(f"custom:{name}", why) for name, why in error.problems]
    if problems:
        raise EditError("bad_property", "; ".join(f"{name}: {why}" for name, why in problems),
                        addresses=[name for name, _ in problems],
                        problems=[("bad_property", name) for name, _ in problems])
    core_names = core_changes(workbook.package, checked_core)
    custom_names = custom_changes(workbook.package, checked_custom)
    if not core_names and not custom_names:
        return PropertiesChange(note="the properties are already so")
    with workbook.editor.batch():
        write_core(workbook.package, {name: checked_core[name] for name in core_names})
        write_custom(workbook.package, {name: checked_custom[name] for name in custom_names})
    return PropertiesChange(tuple(core_names), tuple(custom_names))


__all__ = ["Properties", "PropertiesChange", "properties", "set_properties"]
