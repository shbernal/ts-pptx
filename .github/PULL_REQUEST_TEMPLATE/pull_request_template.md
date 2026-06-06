# Submission Guidelines

- Keep source changes in `src/` and tests in `test/`.
- Treat `dist/`, `src/bld`, and `types/` as generated artifacts unless the change explicitly refreshes package outputs.
- New and updated public properties must be reflected in source interfaces and generated declarations.
- New and updated features should include focused tests and, when useful, a demo update.
- Review `docs/development.md`, `docs/testing.md`, and `docs/runtime-and-package-support.md` before changing package behavior.

## Change Summary
<!--- Required: Provide a general summary of your changes -->

## Change Description
<!--- Optional: Describe your changes in detail if complex or summary is not sufficiently detailed -->
<!--- Optional: Describe any new npm libraries needed -->

## Change Type

- [ ] Bug fix
- [ ] New feature
- [ ] Documentation update

## Related Issue
<!--- Optional: If this change is targeted towards an existing Issue -->

## Motivation and Context
<!--- Required: Why is this change required? What does it add or what problem does it solve? -->

## Checklist before requesting a review

- [ ] My code follows the style guidelines of this project.
- [ ] I have performed a self-review of my code.
- [ ] I have included code or tests that prove my fix is effective or that my feature works.
- [ ] I ran the relevant checks from `docs/testing.md`.
- [ ] For emitted OOXML changes, I added or updated a schema fixture and ran `pnpm run test:schema`.
- [ ] For package-boundary changes, I ran `pnpm run test:package`.
- [ ] I did not reintroduce CommonJS or IIFE/global browser bundle support.

## Screenshots / Sample Code (if appropriate)

Thanks for your contribution!
