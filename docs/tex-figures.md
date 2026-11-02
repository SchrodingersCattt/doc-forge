# TeX figure text

`tex2docx` reads a `figure` or `figure*` environment on its own. This is
separate from the [review bundle](tex-review-roundtrip.md), which does not
recompile TeX or rebuild figures.

Literal text beside `\includegraphics` is kept in source order and written
at the body font size. The caption paragraph, including every `\label` on
that float, stays at the caption size. A `figure` with no
`\includegraphics` is refused.

These tokens carry no visible figure text, so they are ignored. Each
`\includegraphics` is still written:

- layout and size switches: `\centering`, `\hfill`, `\hfil`, `\small`,
  `\footnotesize`
- spacing: `\quad`, `\qquad`, `\xspace`, `\,`, `\:`, `\;`, `\!`
- line breaks: `\\`, `\newline`, and `\linebreak`. A star or a bracketed
  argument is ignored with the command, as in `\\*`, `\\[0.5em]`,
  `\newline[0.5em]`, and `\linebreak[4]`

Any other command, math, or special character (`\ { } $ # % & _ ^ ~`) in
the figure body stops conversion. The error is `Unsupported visible content
in figure environment`, and the output file is not written. Refused
examples include `\raisebox`, `\hspace`, `minipage`, and `\textbf`.

An optional argument on `\includegraphics`, such as
`[width=\linewidth]`, belongs to that graphic. It is not text between
graphics.
