# Benchmarks

These benchmarks are for **human document work**, not just feature coverage.

The goal is to answer:

> Can an agent use `deword` to complete realistic Word-document tasks that a human would otherwise do manually?

## Folder layout

```text
benchmarks/
  README.md
  <challenge>/
    README.md           # scenario + scoring
    prompt.txt          # prompt to give the agent
    fixture-spec.md     # exactly how to build the input doc in Word
    judge.json          # semantic + visual checks
    input/
    gold/
    candidate/
    renders/
```

## Evaluation philosophy

Each benchmark should be judged in 3 layers:

1. **Semantic correctness**
   - required updates are present
   - old values are gone
   - protected text stays unchanged

2. **Visual fidelity**
   - candidate doc still looks correct in Microsoft Word
   - compare exported PDFs or screenshots against a gold render

3. **File health**
   - doc opens in Word
   - page count is sane
   - no visible corruption

## First challenge

Start with:

- `board-memo-refresh`

It is intentionally small enough to build by hand in Word today, while still feeling like real executive-assistant / operations work.
