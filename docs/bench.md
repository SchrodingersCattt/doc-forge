# Controlled conversion bench

Synthetic fixtures only. Nothing here is a manuscript, a consumer template, or a machine path. `python -m docforge.bench --iterations 3` builds small, medium, and complex Markdown in a temporary directory and times three operations, each in its own process:

| Kind | Contents |
| --- | --- |
| small | Two paragraphs. |
| medium | One heading, emphasis, a 2x2 table, and 40 paragraphs. |
| complex | Headings, emphasis, a 2x2 table, one 8x8 PNG, and 120 sections. |

| Operation | What is timed |
| --- | --- |
| `md-docx` | `convert` from Markdown to DOCX. |
| `exact-reuse` | `convert` from that DOCX to a second DOCX. The build of the source DOCX is not included. |
| `restyle` | `restyle` of that DOCX with a Normal 12 pt profile. The build of the source DOCX is not included. |

Peak RSS is `ru_maxrss` for that process, in kilobytes, as reported by Linux. Package size is the output DOCX in bytes.

## This run

Runner: Python 3.12.3 on Linux 6.12.94+ (x86_64). Median of 3 iterations. Before is `main` at `3b9db9f`, measured on this same machine with the same fixture text before the branch changes. After is this branch.

| Kind | Operation | Before s | After s | Wall | Before RSS KiB | After RSS KiB | RSS | Before bytes | After bytes |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| small | md-docx | 0.1191 | 0.1129 | -5.2% | 53056 | 53584 | +1.0% | 36857 | 36857 |
| medium | md-docx | 0.1391 | 0.1363 | -2.0% | 53500 | 53956 | +0.9% | 37578 | 37578 |
| complex | md-docx | 0.2940 | 0.2724 | -7.3% | 56028 | 56220 | +0.3% | 39236 | 39236 |
| small | exact-reuse | 0.0023 | 0.0019 | -16.5% | 53088 | 53576 | +0.9% | 36857 | 36857 |
| medium | exact-reuse | 0.0034 | 0.0032 | -3.5% | 53524 | 54008 | +0.9% | 37578 | 37578 |
| complex | exact-reuse | 0.0085 | 0.0084 | -1.1% | 56140 | 56272 | +0.2% | 39236 | 39236 |
| small | restyle | 0.0223 | 0.0211 | -5.1% | 53076 | 53580 | +0.9% | 36850 | 36850 |
| medium | restyle | 0.0291 | 0.0264 | -9.4% | 53496 | 53996 | +0.9% | 37558 | 37558 |
| complex | restyle | 0.0470 | 0.0457 | -2.9% | 56204 | 56196 | -0.0% | 39227 | 39227 |

No operation rose by 25% in wall time or peak RSS. Package sizes were unchanged. Exact-reuse times are a few milliseconds, so the percentage there is run noise, not a speedup claim. `pytest tests/test_bench.py` checks that one small case returns a positive time, RSS, and size. It does not replace the three-iteration command above.
