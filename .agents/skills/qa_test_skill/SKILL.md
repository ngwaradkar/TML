---
name: qa-test-skill
description: Run comprehensive unit and integration tests for the TCF PPC Dashboard and report results.
---

# TCF PPC Dashboard QA Testing Skill

Use this skill when running test cases, validating parsing logic, or checking mathematical constraints in the dashboard calculations.

## Running Tests
Run the comprehensive test suite using `pytest`:
```bash
pytest test_comprehensive_suite.py -v
```

## Adding Test Cases
When adding new features or fixing bugs:
1. Identify target modules (Parsers, Calculators, Utilities, Reset rules).
2. Append parameterized cases to the relevant test function in `test_comprehensive_suite.py`.
3. Verify that all 100+ assertions continue to pass successfully.
