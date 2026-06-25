# Contributing

Thanks for your interest in contributing to PDF to Excel Converter! This guide
explains how to set up the project and submit changes.

## Getting Started

1. Fork the repository and clone your fork:
   ```bash
   git clone https://github.com/TsvetanG2/PDF-To-Excel-Converter.git
   cd PDF-To-Excel-Converter
   ```
2. (Recommended) Create and activate a virtual environment:
   ```bash
   python -m venv .venv
   source .venv/bin/activate      # Windows: .venv\Scripts\activate
   ```
3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
4. Run the app locally:
   ```bash
   python pdftoexcel.py
   ```
   Then open <http://localhost:5000>.

## How to Contribute

1. Create a branch for your work:
   ```bash
   git checkout -b feature/your-feature
   ```
2. Make your changes.
3. Test that the app still runs and converts a sample PDF correctly (both
   "All Text + Tables" and "Tables Only" modes).
4. Add an entry to `CHANGELOG.md` under `[Unreleased]`.
5. Commit and push, then open a Pull Request describing what you changed and why.

Small, focused PRs are easier to review than large ones.

## Code Style

- Follow [PEP 8](https://peps.python.org/pep-0008/) for Python code.
- Keep functions small and focused; prefer clear names over comments.
- Run a linter before submitting:
  ```bash
  pip install flake8
  flake8 .
  ```

## What to Work On

Good areas for contribution:

- Improving table-detection accuracy (pdfplumber / tabula-py tuning)
- Handling edge cases: scanned PDFs, multi-column layouts, merged cells
- Better error messages when a PDF can't be parsed
- File-upload hardening (size limits, type validation, temp-file cleanup)
- Tests for the conversion logic
- UI/UX improvements to the upload page

Check the [open issues](https://github.com/TsvetanG2/PDF-To-Excel-Converter/issues)
for ideas, or start a thread in
[Discussions](https://github.com/TsvetanG2/PDF-To-Excel-Converter/discussions).

## Reporting Bugs & Security Issues

- **Bugs / feature requests:** open a
  [GitHub issue](https://github.com/TsvetanG2/PDF-To-Excel-Converter/issues)
  with steps to reproduce and a sample PDF if possible.
- **Security vulnerabilities:** do **not** open a public issue — follow the
  process in [SECURITY.md](SECURITY.md).

## Code of Conduct

By participating, you agree to abide by the
[Code of Conduct](CODE_OF_CONDUCT.md).

## License

By contributing, you agree that your contributions will be licensed under the
[MIT License](LICENSE.md) that covers this project.
