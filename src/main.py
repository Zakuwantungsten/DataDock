import argparse
from tkinter import Tk, filedialog
from pathlib import Path
from typing import Iterable, List

try:
    from .config import get_column_aliases
    from .merger import ExcelMerger
except ImportError:  # pragma: no cover - supports running as a script
    from config import get_column_aliases
    from merger import ExcelMerger


ALLOWED_SUFFIXES = [".xlsx", ".xlsm"]


def collect_excel_files(paths: Iterable[str], recursive: bool) -> List[Path]:
    files: List[Path] = []
    for raw in paths:
        path = Path(raw)
        if path.is_dir():
            pattern = "**/*" if recursive else "*"
            for suffix in ALLOWED_SUFFIXES:
                files.extend(path.glob(f"{pattern}{suffix}"))
            continue

        if path.is_file() and path.suffix.lower() in ALLOWED_SUFFIXES:
            files.append(path)

    unique = {file.resolve(): file for file in files}
    return sorted(unique.values())


def select_files_with_dialog() -> List[Path]:
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    file_paths = filedialog.askopenfilenames(
        title="Select Excel files",
        filetypes=[("Excel files", "*.xlsx *.xlsm")],
    )
    root.destroy()
    return [Path(path) for path in file_paths]


def select_output_with_dialog(default_name: str) -> Path | None:
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    file_path = filedialog.asksaveasfilename(
        title="Save merged report",
        defaultextension=".xlsx",
        initialfile=default_name,
        filetypes=[("Excel files", "*.xlsx")],
    )
    root.destroy()
    if not file_path:
        return None
    return Path(file_path)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Merge multiple Excel files into a single report sheet."
    )
    parser.add_argument(
        "paths",
        nargs="*",
        help="Excel file paths or folders containing Excel files",
    )
    parser.add_argument(
        "--output",
        default=None,
        help="Output Excel file path",
    )
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="Scan folders recursively",
    )
    args = parser.parse_args()

    if args.paths:
        file_paths = collect_excel_files(args.paths, args.recursive)
    else:
        file_paths = select_files_with_dialog()
    if not file_paths:
        print("No Excel files were selected.")
        return

    merger = ExcelMerger()
    merger.set_files(file_paths)

    aliases_map = get_column_aliases()
    tables = merger.merge_files(aliases_map=aliases_map)

    if merger.errors:
        print("Some files were skipped:")
        for message in merger.errors:
            print(f"- {message}")

    if not tables:
        print("No tables were generated. Check aliases and input files.")
        return

    if args.output:
        output_path = Path(args.output)
    else:
        output_path = select_output_with_dialog("merged_report.xlsx")
        if output_path is None:
            print("No output file selected.")
            return

    output_path.parent.mkdir(parents=True, exist_ok=True)
    merger.export_excel_tables(tables, output_path)
    print(f"Saved: {output_path.resolve()}")


if __name__ == "__main__":
    main()
