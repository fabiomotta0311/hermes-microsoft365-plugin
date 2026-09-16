"""Standalone, opt-in Microsoft 365 native Hermes plugin."""
from .registration import register_plugin


def register(ctx) -> None:
    """Register preflight, policy hook, and executable capability tools."""
    register_plugin(ctx)


__all__ = ["register"]
