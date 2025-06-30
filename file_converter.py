#!/usr/bin/env python3
"""
Office File to PDF Converter
A robust tool for converting Office documents (Word, PowerPoint) to PDF format
with rich progress visualization and comprehensive logging.
"""

import os
import sys
import argparse
import comtypes.client
import shutil
import logging
import time
from pathlib import Path
from typing import List, Tuple, Dict, Optional
from dataclasses import dataclass
from contextlib import contextmanager

from rich.progress import Progress, BarColumn, TextColumn, TimeRemainingColumn, MofNCompleteColumn
from rich.console import Console
from rich.panel import Panel
from rich.table import Table


@dataclass
class Config:
    """Configuration constants and settings."""
    # File format constants for COM objects
    WD_FORMAT_PDF: int = 17
    PPT_FORMAT_PDF: int = 32

    # Supported file extensions
    SUPPORTED_EXTENSIONS: set = frozenset({'.ppt', '.pptx', '.doc', '.docx'})
    PDF_EXTENSION: str = '.pdf'

    # Default directories
    DEFAULT_OUTPUT_SUBDIR: str = 'converted_pdf'
    DEFAULT_LOG_FILE: str = 'conversion_log.txt'

    # UI settings
    PROGRESS_DELAY: float = 0.1
    MAX_FILENAME_DISPLAY: int = 40


class FileConverter:
    """Main file converter class handling the conversion process."""

    def __init__(self, config: Config = None, console: Console = None):
        self.config = config or Config()
        self.console = console or Console()
        self.apps: Dict[str, Optional[object]] = {}
        self.logger = self._setup_logging()

    def _setup_logging(self, log_file: str = None) -> logging.Logger:
        """Setup logging configuration."""
        log_file = log_file or self.config.DEFAULT_LOG_FILE

        # Remove previous log file if it exists
        if os.path.exists(log_file):
            os.remove(log_file)

        # Create logger
        logger = logging.getLogger(__name__)
        logger.setLevel(logging.INFO)

        # Create file handler
        file_handler = logging.FileHandler(log_file)
        file_handler.setLevel(logging.INFO)

        # Create formatter
        formatter = logging.Formatter(
            '%(asctime)s - %(levelname)s - %(message)s'
        )
        file_handler.setFormatter(formatter)

        # Add handler to logger
        if not logger.handlers:
            logger.addHandler(file_handler)

        return logger

    def _identify_tasks(self, input_dir: Path, output_dir: Path) -> List[Tuple[str, str]]:
        """Identify all files that need to be processed."""
        tasks = []

        for file_path in input_dir.iterdir():
            if not file_path.is_file():
                continue

            file_ext = file_path.suffix.lower()
            filename = file_path.name

            if file_ext == self.config.PDF_EXTENSION:
                if input_dir != output_dir:
                    tasks.append(('copy', filename))

            elif file_ext in self.config.SUPPORTED_EXTENSIONS:
                pdf_filename = file_path.stem + self.config.PDF_EXTENSION
                output_path = output_dir / pdf_filename

                if not output_path.exists():
                    tasks.append(('convert', filename))
                else:
                    self.logger.info(
                        f"Skipping conversion for '{filename}', PDF already exists.")
                    self.console.log(
                        f"[yellow]Skipping conversion for '{filename}', PDF already exists.[/yellow]")

        return tasks

    @contextmanager
    def _com_applications(self, tasks: List[Tuple[str, str]]):
        """Context manager for COM applications lifecycle."""
        try:
            needs_ppt = any(
                task[1].lower().endswith(('.ppt', '.pptx'))
                for task in tasks if task[0] == 'convert'
            )
            needs_word = any(
                task[1].lower().endswith(('.doc', '.docx'))
                for task in tasks if task[0] == 'convert'
            )

            if needs_ppt:
                self.console.log(
                    "[bold blue]Initializing PowerPoint...[/bold blue]")
                self.apps['PowerPoint.Application'] = comtypes.client.CreateObject(
                    'PowerPoint.Application')
                try:
                    # Try to set the window state to minimized (not always supported)
                    ppt_app = self.apps['PowerPoint.Application']
                    if hasattr(ppt_app, 'WindowState'):
                        ppt_app.WindowState = 2  # ppWindowMinimized
                except (AttributeError, comtypes.COMError):
                    # If minimizing fails, continue without it
                    pass

            if needs_word:
                self.console.log("[bold blue]Initializing Word...[/bold blue]")
                self.apps['Word.Application'] = comtypes.client.CreateObject(
                    'Word.Application')
                self.apps['Word.Application'].Visible = False

            yield self.apps

        finally:
            self._cleanup_com_applications()

    def _cleanup_com_applications(self):
        """Clean up COM applications."""
        self.console.log("[bold blue]Closing COM applications...[/bold blue]")
        for app_name, app in self.apps.items():
            if app:
                try:
                    app.Quit()
                except comtypes.COMError as e:
                    self.logger.error(f"Error while quitting {app_name}: {e}")
        self.apps.clear()

    def _copy_file(self, input_path: Path, output_path: Path) -> bool:
        """Copy a single file."""
        try:
            shutil.copy2(input_path, output_path)
            self.logger.info(
                f"Copied '{input_path.name}' to output directory.")
            return True
        except (shutil.Error, IOError) as e:
            self.logger.error(f"Could not copy '{input_path.name}': {e}")
            return False

    def _convert_file(self, input_path: Path, output_path: Path) -> bool:
        """Convert a single file to PDF."""
        file_ext = input_path.suffix.lower()

        try:
            if file_ext in {'.ppt', '.pptx'}:
                app = self.apps.get('PowerPoint.Application')
                if not app:
                    raise RuntimeError(
                        "PowerPoint application not initialized")

                presentation = app.Presentations.Open(
                    str(input_path), ReadOnly=True, WithWindow=False)
                presentation.SaveAs(
                    str(output_path), self.config.PPT_FORMAT_PDF)
                presentation.Close()

            elif file_ext in {'.doc', '.docx'}:
                app = self.apps.get('Word.Application')
                if not app:
                    raise RuntimeError("Word application not initialized")

                doc = app.Documents.Open(str(input_path), ReadOnly=True)
                doc.SaveAs(str(output_path), self.config.WD_FORMAT_PDF)
                doc.Close()

            self.logger.info(
                f"Successfully converted '{input_path.name}' to PDF.")
            return True

        except (OSError, comtypes.COMError, RuntimeError) as e:
            self.logger.error(f"Failed to convert '{input_path.name}': {e}")
            return False

    def process_files(self, input_dir: str, output_dir: str = None) -> None:
        """Main processing function."""
        input_path = Path(input_dir).resolve()

        if not input_path.exists() or not input_path.is_dir():
            raise ValueError(f"Input directory does not exist: {input_dir}")

        if output_dir:
            output_path = Path(output_dir).resolve()
        else:
            output_path = input_path / self.config.DEFAULT_OUTPUT_SUBDIR

        output_path.mkdir(exist_ok=True)

        # Display initial information
        self._display_info(input_path, output_path)

        # Identify tasks
        tasks = self._identify_tasks(input_path, output_path)

        if not tasks:
            self.logger.info("No files to copy or convert.")
            self.console.log(
                "[bold green]No new files to copy or convert.[/bold green]")
            return

        # Process tasks
        with self._com_applications(tasks):
            self._process_tasks_with_progress(tasks, input_path, output_path)

    def _display_info(self, input_path: Path, output_path: Path):
        """Display initial processing information."""
        info_table = Table(show_header=False, box=None)
        info_table.add_column("Label", style="bold yellow")
        info_table.add_column("Path", style="green")

        info_table.add_row("Input Directory:", str(input_path))
        info_table.add_row("Output Directory:", str(output_path))
        info_table.add_row("Log File:", self.config.DEFAULT_LOG_FILE)

        self.console.print(
            Panel(info_table, title="Processing Information", border_style="blue"))

    def _process_tasks_with_progress(self, tasks: List[Tuple[str, str]],
                                     input_path: Path, output_path: Path):
        """Process all tasks with a rich progress bar."""
        with Progress(
            TextColumn("[bold blue]{task.description}", justify="right"),
            BarColumn(bar_width=None),
            "[progress.percentage]{task.percentage:>3.1f}%",
            "•",
            MofNCompleteColumn(),
            "•",
            TimeRemainingColumn(),
            console=self.console
        ) as progress:

            task_id = progress.add_task(
                "Processing files...", total=len(tasks))
            success_count = 0

            for task_type, filename in tasks:
                # Update progress description
                display_name = filename[:self.config.MAX_FILENAME_DISPLAY]
                if len(filename) > self.config.MAX_FILENAME_DISPLAY:
                    display_name += "..."

                progress.update(
                    task_id, description=f"[green]Processing[/] {display_name}")

                input_file_path = input_path / filename

                if task_type == 'copy':
                    output_file_path = output_path / filename
                    if self._copy_file(input_file_path, output_file_path):
                        success_count += 1

                elif task_type == 'convert':
                    pdf_filename = Path(filename).stem + \
                        self.config.PDF_EXTENSION
                    output_file_path = output_path / pdf_filename
                    if self._convert_file(input_file_path, output_file_path):
                        success_count += 1

                progress.update(task_id, advance=1)
                time.sleep(self.config.PROGRESS_DELAY)

            # Display final results
            self._display_results(len(tasks), success_count)

    def _display_results(self, total_tasks: int, success_count: int):
        """Display processing results."""
        failed_count = total_tasks - success_count

        results_table = Table(show_header=False, box=None)
        results_table.add_column("Metric", style="bold")
        results_table.add_column("Count", style="bold")

        results_table.add_row("Total Files:", str(total_tasks))
        results_table.add_row("Successfully Processed:",
                              f"[green]{success_count}[/green]")

        if failed_count > 0:
            results_table.add_row("Failed:", f"[red]{failed_count}[/red]")

        self.console.print(
            Panel(results_table, title="Processing Results", border_style="green"))


def create_argument_parser() -> argparse.ArgumentParser:
    """Create and configure argument parser."""
    parser = argparse.ArgumentParser(
        description="Convert Office documents to PDF format",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s /path/to/documents
  %(prog)s /path/to/documents -o /path/to/output
  %(prog)s /path/to/documents --log-file custom_log.txt
        """
    )

    parser.add_argument(
        'input_directory',
        help='Input directory containing files to convert'
    )

    parser.add_argument(
        '-o', '--output',
        dest='output_directory',
        help='Output directory (default: input_directory/converted_pdf)'
    )

    parser.add_argument(
        '--log-file',
        dest='log_file',
        default=Config.DEFAULT_LOG_FILE,
        help=f'Log file path (default: {Config.DEFAULT_LOG_FILE})'
    )

    parser.add_argument(
        '--version',
        action='version',
        version='%(prog)s 2.0.0'
    )

    return parser


def main():
    """Main entry point."""
    parser = create_argument_parser()

    # If no arguments provided, use interactive mode
    if len(sys.argv) == 1:
        console = Console()
        console.print(
            Panel.fit("🚀 Office File to PDF Converter", style="bold blue"))

        while True:
            input_directory = console.input(
                "[bold yellow]Enter the input directory path: [/bold yellow]")
            if os.path.isdir(input_directory):
                break
            console.print(
                "❌ [bold red]Invalid directory. Please enter a valid path.[/bold red]")

        converter = FileConverter(console=console)
        converter.process_files(input_directory)

    else:
        # Command line mode
        args = parser.parse_args()

        try:
            converter = FileConverter()
            converter._setup_logging(args.log_file)
            converter.process_files(
                args.input_directory, args.output_directory)

        except ValueError as e:
            print(f"Error: {e}", file=sys.stderr)
            sys.exit(1)
        except KeyboardInterrupt:
            print("\nOperation cancelled by user.", file=sys.stderr)
            sys.exit(1)
        except Exception as e:
            print(f"Unexpected error: {e}", file=sys.stderr)
            sys.exit(1)

    print("\n🎉 Processing complete.")


if __name__ == "__main__":
    main()
