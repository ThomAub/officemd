"""Pipeline definitions for the OfficeMD local provider.

To register these pipelines inside a ParseBench checkout, import this module
from `parse_bench.inference.pipelines.parse.register_parse_pipelines` (or call
`register_officemd_pipelines(register_fn)` directly), passing the same
`register_fn` used for built-in pipelines.

Example patch applied to `src/parse_bench/inference/pipelines/parse.py`:

    from parse_bench.inference.pipelines.officemd_pipelines import (
        register_officemd_pipelines,
    )

    def register_parse_pipelines(register_fn):
        # ...existing registrations...
        register_officemd_pipelines(register_fn)
"""

from __future__ import annotations

from collections.abc import Callable

from parse_bench.schemas.pipeline import PipelineSpec
from parse_bench.schemas.product import ProductType

# Importing the provider module ensures `@register_provider("officemd_local")`
# fires before any pipeline referencing it is instantiated.
import parse_bench.inference.providers.parse.officemd_local  # noqa: F401


def register_officemd_pipelines(register_fn: Callable[[PipelineSpec], None]) -> None:
    """Register all OfficeMD-backed pipelines with ParseBench.

    The default pipeline invokes the CLI from a local OfficeMD checkout via
    `cargo run`. Set `OFFICEMD_REPO_ROOT` in the environment, or override
    `repo_root` on the pipeline config, to point at the checkout.
    """
    register_fn(
        PipelineSpec(
            pipeline_name="officemd_local",
            provider_name="officemd_local",
            product_type=ProductType.PARSE,
            config={
                "cargo_run": True,
                "cargo_profile": "release",
                "extra_args": [],
            },
        )
    )
