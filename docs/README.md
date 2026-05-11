# Documentation

This directory keeps product and engineering intent close to the code without mixing it into implementation files.

## Structure

- `specs/` - active and historical feature specifications, implementation plans, and acceptance criteria.
- `adr/` - architecture decision records for decisions that should survive beyond a single PR.
- `templates/` - reusable document templates for specs and ADRs.

## Specification-Driven Workflow

1. Start with a spec before changing behavior.
   - Define the problem, goals, non-goals, constraints, and acceptance criteria.
   - Prefer observable acceptance criteria over implementation wishes.

2. Link decisions to implementation.
   - Use an ADR when the work introduces a meaningful technical decision or tradeoff.
   - Keep tactical task lists in the spec, not the ADR.

3. Keep verification explicit.
   - Record the expected checks before implementation.
   - Update the spec with the actual verification result before merging.

4. Keep documents living but scoped.
   - Update active specs as the work changes.
   - Avoid retrofitting unrelated decisions into old specs.

## Naming

- Specs: `short-topic.md`
- ADRs: `short-decision.md`
