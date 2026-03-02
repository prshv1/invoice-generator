# Todo list for Python Module & Features

## Python Module Improvements
- Add `__main__.py` to `src/invoice_generator/` to support clean execution via `python -m invoice_generator` (doing `python -c "from invoice_generator.cli import main_generator"` is too verbose for users).
- Explicitly expose the `extract_from_image` function in `__init__.py` and abstract it so power users can run OCR/LLMs individually, beyond just running `images_to_invoice`.
- Add mapping to `pyproject.toml` or `setup.py` so the module installs a native CLI bin like `invoice-gen`.

## CI
- Add GitHub Actions workflows for automated builds on pushes to `main`.

## Upcoming Features (from README Roadmap)
- [ ] Web interface — browser-based upload & download
- [ ] Email integration — auto-send invoices to buyers
- [ ] IGST support — inter-state invoice detection
