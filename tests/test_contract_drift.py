"""Cross-layer drift checks: the contract, the handlers and the handbook must agree.

Why this file exists
--------------------
A handler that reads an argument the contract does not declare is a real defect and an easy one
to introduce: the argument passes neither validation nor the generated schema, so the caller's
intent is silently dropped *or* the call is refused for a reason that contradicts the
documentation. That is exactly what happened with ``verify`` on ``outlook.create_draft`` -- the
handler honoured it, the contract had never heard of it, and the documented opt-out was refused.

The guard is structural rather than a one-off assertion: the handler modules are read as source
and every argument they consult is required to be declared by the operation that consults it.
That catches the whole class, including the next argument someone adds.
"""
from __future__ import annotations

import inspect
import re

import pytest

from microsoft365 import handlers
from microsoft365.handlers import calendar, files, outlook, planner, teams, todo
from microsoft365.validation import ARGUMENT_CONTRACTS

HANDLER_MODULES = (calendar, files, outlook, planner, teams, todo)

#: ``arguments.get("name")`` / ``arguments["name"]`` -- how a handler reads one input.
ARGUMENT_READ = re.compile(r"""arguments(?:\.get\(|\[)\s*["']([a-z_][a-z0-9_]*)["']""")

#: Names a handler may read that are not caller arguments: the dispatch key itself, and any
#: local bookkeeping the handler keeps in the same mapping.
NOT_ARGUMENTS = frozenset({"action"})

#: Arguments a handler reads *in order to refuse them*, where the specific refusal is better
#: than the contract's generic "the argument X is not accepted". Each entry is a reviewed
#: exception, not a blanket exemption: an undeclared read anywhere else still fails this file.
DEFENSIVE_READS = {
    # A draft is composed from its own body and recipients; accepting a message_id here would
    # mean silently drafting from a different message than the one the body describes.
    "outlook.create_draft": frozenset({"message_id"}),
}


def handler_operation_keys(module) -> dict:
    """Map each ``@handler(service, operation)``-decorated function to its operation key."""
    found = {}
    for name, function in vars(module).items():
        if not callable(function) or not hasattr(function, "__wrapped__"):
            continue
        source = inspect.getsource(function) if _has_source(function) else ""
        for service, operation in re.findall(
            r"""@handler\(\s*["']([a-z0-9_]+)["']\s*,\s*["']([a-z0-9_]+)["']\s*\)""", source
        ):
            found[function] = f"{service}.{operation}"
    return found


def _has_source(function) -> bool:
    try:
        inspect.getsource(function)
        return True
    except (OSError, TypeError):  # pragma: no cover - sources are present in this checkout
        return False


def registered_handlers():
    """Every handler in the plugin, as ``(operation_key, function)`` pairs."""
    pairs = []
    for module in HANDLER_MODULES:
        source = inspect.getsource(module)
        for match in re.finditer(
            r"""@handler\(\s*["']([a-z0-9_]+)["']\s*,\s*["']([a-z0-9_]+)["']\s*\)\s*\ndef\s+([a-z_][a-z0-9_]*)""",
            source,
        ):
            service, operation, function_name = match.groups()
            function = getattr(module, function_name, None)
            if function is not None:
                pairs.append((f"{service}.{operation}", function))
    return pairs


def assert_declared_arguments(operation: str, function) -> list[str]:
    """Every argument the handler reads must be declared by its operation's contract."""
    contract = ARGUMENT_CONTRACTS.get(operation)
    assert contract is not None, f"{operation} has a handler but no argument contract"
    declared = {spec.name for spec in contract.properties}
    exempt = DEFENSIVE_READS.get(operation, frozenset())
    read = set(ARGUMENT_READ.findall(inspect.getsource(function))) - NOT_ARGUMENTS - exempt
    undeclared = sorted(read - declared)
    assert not undeclared, (
        f"{operation} reads {undeclared} but its contract declares only {sorted(declared)}"
    )
    return sorted(read)


def test_every_handler_argument_is_declared_by_its_contract():
    """The whole class of defect, not one instance: a handler cannot read what is undeclared."""
    handlers_found = registered_handlers()
    assert len(handlers_found) >= 25, "the handler scan found implausibly few handlers"
    for operation, function in handlers_found:
        assert_declared_arguments(operation, function)


#: The operations whose write is confirmed by a re-read. Reviewed list: the confirmation lives
#: in a per-handler or shared helper, so it is asserted both ways below.
CONFIRMABLE_WRITES = {
    "outlook.create_draft": outlook,
    "calendar.create_events": calendar,
    "sharepoint.upload_files": files,
    "onedrive.upload_files": files,
}


def test_every_confirmable_write_declares_the_verify_opt_out():
    """A handler that honours ``verify`` must publish it, or the documented opt-out is refused."""
    for operation, module in CONFIRMABLE_WRITES.items():
        declared = {spec.name for spec in ARGUMENT_CONTRACTS[operation].properties}
        assert "verify" in declared, f"{operation} honours verify but does not declare it"
        assert "verify" in inspect.getsource(module), (
            f"{operation} declares verify but nothing in its module reads it"
        )


def test_the_confirmable_write_list_still_covers_every_module_that_reads_verify():
    """A module that reads ``verify`` must have all its writes on the reviewed list."""
    reading_modules = {
        module for module in HANDLER_MODULES if "verify" in inspect.getsource(module)
    }
    listed_modules = set(CONFIRMABLE_WRITES.values())
    assert reading_modules <= listed_modules, (
        f"these modules read verify but no write of theirs is listed: "
        f"{sorted(m.__name__ for m in reading_modules - listed_modules)}"
    )


def test_the_deprecated_overwrite_alias_is_declared_and_cannot_contradict_the_new_name():
    """The handbook promises ``overwrite`` still works; the contract must therefore accept it."""
    for operation in ("sharepoint.upload_files", "onedrive.upload_files"):
        contract = ARGUMENT_CONTRACTS[operation]
        names = {spec.name for spec in contract.properties}
        assert {"overwrite", "conflict_behavior"} <= names, f"{operation} lost the alias"
        spec = next(spec for spec in contract.properties if spec.name == "overwrite")
        assert spec.kind == "boolean" and spec.required is False
        # Both at once is refused rather than one silently outranking the other.
        forms = [exclusion.forms for exclusion in contract.exclusions]
        assert (("overwrite",), ("conflict_behavior",)) in forms


def test_the_verify_opt_out_is_a_boolean_in_every_contract_that_has_it():
    declaring = [
        key for key, contract in ARGUMENT_CONTRACTS.items()
        if any(spec.name == "verify" for spec in contract.properties)
    ]
    assert declaring, "no contract declares verify at all"
    for key in declaring:
        spec = next(spec for spec in ARGUMENT_CONTRACTS[key].properties if spec.name == "verify")
        assert spec.kind == "boolean", f"{key}.verify is {spec.kind}, not a boolean"
        assert spec.required is False, f"{key}.verify is required; the opt-out must be optional"


def test_no_contract_declares_an_argument_no_handler_reads_for_a_write():
    """A declared write argument nobody reads is either dead or a handler bug -- surface both."""
    write_keys = [key for key in ARGUMENT_CONTRACTS if key in _EXECUTABLE_WRITES]
    for key in write_keys:
        contract = ARGUMENT_CONTRACTS[key]
        module_source = "".join(
            inspect.getsource(module) for module in HANDLER_MODULES
        )
        for spec in contract.properties:
            # ``action`` is the dispatch key and ``user_id``/``drive_id``/... are routing; only
            # the behavioural arguments are worth asserting a reader for.
            if spec.name in {"action", "user_id", "drive_id", "item_id", "item_path", "chat_id"}:
                continue
            assert spec.name in module_source, (
                f"{key} declares {spec.name}, which no handler mentions"
            )


_EXECUTABLE_WRITES = frozenset(
    {
        "outlook.create_draft",
        "outlook.send",
        "sharepoint.upload_files",
        "onedrive.upload_files",
    }
)


@pytest.mark.parametrize("operation", sorted(_EXECUTABLE_WRITES))
def test_an_executable_write_still_has_a_handler(operation):
    keys = [key for key, _ in registered_handlers()]
    assert operation in keys, f"{operation} is executable but has no registered handler"


def test_the_handbook_does_not_promise_an_operation_that_does_not_exist():
    """Every ``service.operation`` named in the reference docs must exist in the registry."""
    from pathlib import Path

    from microsoft365.contract import OPERATION_REGISTRY

    root = Path(handlers.__file__).resolve().parents[2]
    documented = set()
    for path in (root / "docs").rglob("*.md"):
        documented.update(
            re.findall(r"\b((?:outlook|calendar|sharepoint|onedrive|teams|planner|todo)\.[a-z_]+)\b",
                       path.read_text(encoding="utf-8"))
        )
    assert documented, "no documented operation names were found"
    unknown = sorted(name for name in documented if name not in OPERATION_REGISTRY)
    assert not unknown, f"the docs name operations that do not exist: {unknown}"