
# LeanCTX usage

- Prefer ctx_search for locating code and ctx_tree for project structure.
- Prefer ctx_read for reading files and relevant line ranges.
- Prefer ctx_shell for commands with potentially large output.
- Never dump entire files or directories when a focused excerpt is sufficient.
- For builds and tests, return failures and relevant surrounding lines,
  not the complete successful output.
- Use native tools when LeanCTX cannot perform the required operation.


