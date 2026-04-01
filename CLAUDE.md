# Project Name — Developer Guide

<!-- AUTO-DEPLOYED SKELETON — Customize every section below for your project -->
<!-- Delete this comment block when done customizing -->

## Project Overview
<!-- What this project does, who it's for, key tech stack -->

## Quick Start
```bash
# Backend
# cd backend && ...

# Frontend
# cd frontend && ...
```

## Architecture
<!-- Key technologies, patterns, directory structure -->

## Documentation Lookup Table

<!-- Map features to their docs and key files. Example:
| Feature | Documentation | Key Backend Files | Key Frontend Files |
|---------|--------------|-------------------|-------------------|
| Auth    | docs/auth.md | auth/views.py     | context/Auth.tsx   |
-->

## Key Patterns

### Backend
<!-- Naming conventions, model patterns, permission classes, etc. -->

### Frontend
<!-- Component patterns, state management, API client conventions -->

### Testing
```bash
# How to run tests
```

## Code Quality Standards

### Code Review Checkpoints
<!-- Replacing the code-review plugin with embedded behavior -->
After completing any feature or significant change, perform a self-review:
- [ ] All tests pass
- [ ] No hardcoded secrets, credentials, or PII
- [ ] API endpoints have proper authentication/authorization
- [ ] Error handling covers edge cases
- [ ] No unused imports or dead code introduced
- [ ] Documentation updated if behavior changed

### Security
- Never commit secrets, API keys, or credentials
- Validate all user input at system boundaries
- Use parameterized queries (no raw SQL string interpolation)
- Sanitize output to prevent XSS

## Data Sanitization
<!-- For open-source projects -->
- Use fake data only (Faker, 555-xxxx phones, generic names)
- Seed commands and local data are gitignored
- Pre-commit hook scans for PII patterns

## Important Gotchas
<!-- Project-specific foot-guns and non-obvious behavior -->
