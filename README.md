## 🎯 Business Objective
The core purpose of this automation is to run every single day, grab 4 distinct large daily data exports (**Issuance, Inventory, Receipt, and R&R**), and automatically compile them into 4 separate sheets of a Master Inventory Dashboard. It dynamically calculates reporting formulas across millions of cells, replacing hours of manual, memory-heavy Excel copy-pasting.

## 🚀 Key Features

*   **Hybrid Architecture:** Uses Robot Framework for orchestration, logging, and error tracking, paired with Python data models for raw calculation speed.
*   **Massive Data Handling:** Implemented a custom, low-memory String-Split HTML parser to bypass pandas `MemoryError` limitations when ingesting 1GB+ `.xls` disguised HTML files securely in under 40 seconds.
*   **Format Interoperability:** Implemented a continuous `try-except` fallback architecture `(pyxlsb -> openpyxl -> xlrd -> strict Custom HTML parser)` that auto-adapts to daily changes in source-file exports seamlessly.
*   **Dynamic Placement:** Automatically auto-detects dynamically shifting input data widths (`len(df.columns)`) to append calculated formula columns to the exact edge reliably every run.
*   **Zero-Zombie Safe Process:** Integrated Python COM `App.Quit()` protocols inside `try...finally` teardowns to guarantee no background `Excel.exe` tasks hang on the server after completion.

## 📁 Repository Structure
*   `Process 1/` - Contains the Python extraction classes and the Robot Framework execution suites for Issuance, Inventory, RnR, and Receipt modules.
*   `INV_Dashboard.robot` - The unified production orchestrator combining all 4 flows into a single memory-safe Excel session.

*(Note: Sensitive company templates and raw data sets have been intentionally omitted from this public repository.)*
