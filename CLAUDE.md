# docgen-json-to-word — Repository Guidance

## Role

`docgen-json-to-word` owns final Word and Excel document rendering for DocGen.

It consumes prepared document data from upstream DocGen processing and turns it into rendered Office output.

For cross-repository DocGen architecture and ownership, consult:

`../docs/docgen/architecture.md`

Load that document only when the task genuinely requires cross-repository context.

## Stack

- .NET / C#
- Word rendering
- Excel rendering
- dependency injection through the existing application startup/composition root

Use the current solution/project configuration as the source of truth for exact frameworks, package versions, scripts, and build behavior.

## Architecture Boundaries

This repository owns:

- Word-specific rendering behavior
- Excel-specific rendering behavior
- document-format-specific layout and formatting logic
- renderer-side interpretation of established DocGen contracts

Do not move upstream business/domain processing into the renderer.

Do not duplicate data retrieval or Azure DevOps logic here.

Do not patch malformed upstream data silently when the actual defect belongs to Content Control, Data Provider, or another upstream producer.

When a rendering problem is caused by an upstream contract defect, identify the owning repository before broadening scope.

## Important Entry Points

Likely routing anchors include:

- `WordService.cs`
- `ExcelService.cs`
- `Startup.cs`

Use targeted discovery from the relevant renderer/service rather than broad solution scans.

Before adding a new rendering abstraction, check whether an existing Word/Excel service, helper, builder, formatter, or extension point already owns the behavior.

## Renderer Ownership

Keep Word-specific and Excel-specific behavior separated where the document formats genuinely differ.

Share behavior only when the semantics are actually common.

Do not generalize formatting abstractions solely because two implementations currently look similar.

Keep orchestration thin around renderer services; formatting/layout behavior should remain in the owning rendering layer rather than leaking into controllers or transport-facing code.

## Input Contract Safety

The renderer consumes data produced by upstream DocGen repositories.

Treat changes to assumptions about:

- required fields
- nullability
- collection shape
- ordering
- style metadata
- table structure
- document/section hierarchy

as potential cross-repository contract changes.

Before changing renderer input expectations:

1. identify the upstream producer,
2. verify the existing contract,
3. preserve backwards compatibility where practical,
4. avoid silently accepting malformed input when it would hide an upstream defect,
5. add focused regression coverage.

## Word and Excel Output Safety

Rendering changes can affect document appearance even when the code change is small.

When modifying formatting/layout behavior:

- identify whether the change is Word-only, Excel-only, or shared,
- preserve unrelated formatting,
- avoid broad style normalization,
- verify representative output when practical,
- treat visual/output regressions as behavior regressions.

For bugs affecting a specific document structure, prefer a focused regression case over broad renderer rewrites.

## Dependency Injection

Use the existing dependency-injection/composition model.

New or materially modified service/infrastructure dependencies should be registered through the existing DI mechanism.

Do not introduce service locator patterns or instantiate infrastructure dependencies directly inside domain/rendering logic without a clear reason.

Avoid broad DI refactors outside the requested scope.

## Async and Cancellation

For asynchronous code:

- propagate `CancellationToken` through async call chains where supported,
- do not block on async operations using `.Result`, `.Wait()`, or equivalent sync-over-async patterns in normal application code,
- preserve asynchronous I/O behavior in request or rendering hot paths.

Use `async void` only for legitimate framework event-handler scenarios.

## Resource Management

Dispose streams, documents, writers, and other disposable resources correctly.

Prefer existing ownership/disposal patterns.

Do not leave file handles, streams, or document objects open across rendering operations.

For async-disposable resources, use the appropriate async disposal pattern when supported.

## Error Handling

Do not silently swallow rendering failures.

Preserve useful context such as the affected document operation or rendering stage without logging sensitive document content unnecessarily.

Do not convert failures into partially valid-looking output unless that is an explicit existing behavior.

Follow current repository error-handling conventions before introducing new wrappers.

## Performance

Document generation may process large documents, tables, or repeated elements.

Watch for:

- repeated expensive enumeration,
- unnecessary full materialization,
- avoidable document-tree traversal,
- synchronous I/O in hot paths,
- repeated object creation inside large rendering loops.

Do not perform speculative performance rewrites without evidence.

Preserve correctness and output fidelity over micro-optimization.

## Commands

Use the current solution/project configuration as authority.

Typical commands:

- Build: `dotnet build JsonToWord.sln`
- Tests: `dotnet test JsonToWord.sln`

Run the narrowest relevant test/project first when possible.

For rendering changes, supplement automated tests with representative output verification when practical and when the affected behavior is visual/format-sensitive.

## Source of Truth

Use current source/configuration as authority for:

- rendering ownership
- DI registration
- input contracts
- Word/Excel behavior
- scripts and build behavior
- test configuration

Do not encode volatile test counts, package versions, file sizes, or temporary implementation status in this file.
