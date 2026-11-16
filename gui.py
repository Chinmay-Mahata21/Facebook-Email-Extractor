
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import threading
import queue
import time

from excel_handler import (
    read_facebook_urls,
    initialize_output,
    get_processed_urls,
    save_result
)

from scraper import (
    create_browser,
    scrape_facebook_page
)


# -----------------------------
# Application settings
# -----------------------------

DEFAULT_INPUT = "input/facebook_urls.xlsx"
DEFAULT_OUTPUT = "output/extracted_emails.xlsx"

message_queue = queue.Queue()
stop_event = threading.Event()


# -----------------------------
# Main application
# -----------------------------

class FacebookExtractorGUI:

    def __init__(self, root):

        self.root = root
        self.root.title("Facebook Business Email Extractor")
        self.root.geometry("750x650")
        self.root.minsize(650, 580)

        self.input_path = tk.StringVar(value=DEFAULT_INPUT)
        self.output_path = tk.StringVar(value=DEFAULT_OUTPUT)
        self.url_limit = tk.StringVar()

        self.total_urls = 0
        self.found_count = 0
        self.not_found_count = 0
        self.error_count = 0

        self.is_running = False

        self.create_widgets()

        self.root.after(100, self.process_queue)

        # Automatically count URLs on startup
        self.count_urls()

    # -------------------------
    # GUI layout
    # -------------------------

    def create_widgets(self):

        main_frame = ttk.Frame(self.root, padding=20)
        main_frame.pack(fill="both", expand=True)

        # Heading
        ttk.Label(
            main_frame,
            text="Facebook Business Email Extractor",
            font=("Arial", 18, "bold")
        ).pack(pady=(0, 5))

        ttk.Label(
            main_frame,
            text="Extract publicly visible business emails from Facebook pages."
        ).pack(pady=(0, 20))

        # Input file
        ttk.Label(
            main_frame,
            text="Input Excel File:"
        ).pack(anchor="w")

        input_frame = ttk.Frame(main_frame)
        input_frame.pack(fill="x", pady=(5, 15))

        ttk.Entry(
            input_frame,
            textvariable=self.input_path
        ).pack(side="left", fill="x", expand=True)

        ttk.Button(
            input_frame,
            text="Browse",
            command=self.browse_input
        ).pack(side="left", padx=(8, 0))

        # Output file
        ttk.Label(
            main_frame,
            text="Output Excel File:"
        ).pack(anchor="w")

        output_frame = ttk.Frame(main_frame)
        output_frame.pack(fill="x", pady=(5, 15))

        ttk.Entry(
            output_frame,
            textvariable=self.output_path
        ).pack(side="left", fill="x", expand=True)

        ttk.Button(
            output_frame,
            text="Browse",
            command=self.browse_output
        ).pack(side="left", padx=(8, 0))

        # URL limit
        ttk.Label(
            main_frame,
            text="Maximum URLs to Process:"
        ).pack(anchor="w")

        limit_frame = ttk.Frame(main_frame)
        limit_frame.pack(fill="x", pady=(5, 5))

        ttk.Entry(
            limit_frame,
            textvariable=self.url_limit,
            width=15
        ).pack(side="left")

        ttk.Label(
            limit_frame,
            text="Leave blank to process all URLs."
        ).pack(side="left", padx=10)

        # URL count
        self.url_count_label = ttk.Label(
            main_frame,
            text="Total URLs: 0"
        )
        self.url_count_label.pack(anchor="w", pady=(5, 15))

        ttk.Button(
            main_frame,
            text="Refresh URL Count",
            command=self.count_urls
        ).pack(anchor="w", pady=(0, 15))

        ttk.Separator(main_frame).pack(fill="x", pady=10)

        # Start and Stop buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x", pady=10)

        self.start_button = ttk.Button(
            button_frame,
            text="Start Extraction",
            command=self.start_extraction
        )
        self.start_button.pack(side="left", expand=True, fill="x", padx=(0, 5))

        self.stop_button = ttk.Button(
            button_frame,
            text="Stop",
            command=self.stop_extraction,
            state="disabled"
        )
        self.stop_button.pack(side="left", expand=True, fill="x", padx=(5, 0))

        # Progress
        ttk.Label(
            main_frame,
            text="Progress:"
        ).pack(anchor="w", pady=(10, 5))

        self.progress = ttk.Progressbar(
            main_frame,
            mode="determinate",
            maximum=100
        )
        self.progress.pack(fill="x")

        self.progress_label = ttk.Label(
            main_frame,
            text="Waiting to start..."
        )
        self.progress_label.pack(anchor="w", pady=5)

        # Counters
        stats_frame = ttk.Frame(main_frame)
        stats_frame.pack(fill="x", pady=15)

        self.found_label = ttk.Label(
            stats_frame,
            text="Found: 0",
            foreground="green"
        )
        self.found_label.pack(side="left", expand=True)

        self.not_found_label = ttk.Label(
            stats_frame,
            text="Not Found: 0"
        )
        self.not_found_label.pack(side="left", expand=True)

        self.error_label = ttk.Label(
            stats_frame,
            text="Errors: 0",
            foreground="red"
        )
        self.error_label.pack(side="left", expand=True)

        # Log
        ttk.Label(
            main_frame,
            text="Extraction Log:"
        ).pack(anchor="w")

        log_frame = ttk.Frame(main_frame)
        log_frame.pack(fill="both", expand=True, pady=5)

        self.log_box = tk.Text(
            log_frame,
            height=12,
            state="disabled",
            wrap="word",
            font=("Consolas", 9)
        )
        self.log_box.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(
            log_frame,
            command=self.log_box.yview
        )
        scrollbar.pack(side="right", fill="y")

        self.log_box.config(
            yscrollcommand=scrollbar.set
        )

    # -------------------------
    # Browse files
    # -------------------------

    def browse_input(self):

        file_path = filedialog.askopenfilename(
            title="Select Input Excel File",
            filetypes=[("Excel Files", "*.xlsx")]
        )

        if file_path:
            self.input_path.set(file_path)
            self.count_urls()

    def browse_output(self):

        file_path = filedialog.asksaveasfilename(
            title="Select Output Excel File",
            defaultextension=".xlsx",
            initialfile="extracted_emails.xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )

        if file_path:
            self.output_path.set(file_path)

    # -------------------------
    # Count URLs
    # -------------------------

    def count_urls(self):

        try:
            urls = read_facebook_urls(
                self.input_path.get()
            )

            self.total_urls = len(urls)

            self.url_count_label.config(
                text=f"Total valid unique URLs: {self.total_urls}"
            )

        except Exception as error:

            self.total_urls = 0

            self.url_count_label.config(
                text="Unable to count URLs."
            )

            messagebox.showerror(
                "Input Error",
                str(error)
            )

    # -------------------------
    # Logging
    # -------------------------

    def log(self, message):

        self.log_box.config(state="normal")

        self.log_box.insert(
            "end",
            message + "\n"
        )

        self.log_box.see("end")

        self.log_box.config(state="disabled")

    # -------------------------
    # Start extraction
    # -------------------------

    def start_extraction(self):

        if self.is_running:
            return

        input_file = Path(self.input_path.get())
        output_file = Path(self.output_path.get())

        if not input_file.is_file():
            messagebox.showerror(
                "Input Error",
                "Please select a valid input Excel file."
            )
            return

        if output_file.suffix.lower() != ".xlsx":
            messagebox.showerror(
                "Output Error",
                "Please select an output file with .xlsx extension."
            )
            return

        limit_text = self.url_limit.get().strip()

        if limit_text:

            try:
                limit = int(limit_text)

                if limit < 1:
                    raise ValueError

            except ValueError:

                messagebox.showerror(
                    "Invalid Limit",
                    "Please enter a positive whole number or leave it blank."
                )
                return

        else:
            limit = None

        self.found_count = 0
        self.not_found_count = 0
        self.error_count = 0

        self.update_counters()

        self.progress["value"] = 0
        self.progress["maximum"] = 100

        self.log_box.config(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.config(state="disabled")

        stop_event.clear()

        self.is_running = True

        self.start_button.config(state="disabled")
        self.stop_button.config(state="normal")

        self.progress_label.config(
            text="Preparing extraction..."
        )

        worker = threading.Thread(
            target=self.extraction_worker,
            args=(str(input_file), str(output_file), limit),
            daemon=True
        )

        worker.start()

    # -------------------------
    # Extraction worker
    # -------------------------

    def extraction_worker(self, input_file, output_file, limit):

        playwright = None
        browser = None
        context = None

        try:

            # Read all URLs
            urls = read_facebook_urls(input_file)

            # Prepare output
            initialize_output(output_file)

            # Skip previously processed URLs
            processed_urls = get_processed_urls(output_file)

            remaining_urls = [
                url for url in urls
                if url not in processed_urls
            ]

            # Apply GUI limit AFTER skipping processed URLs
            if limit is not None:
                remaining_urls = remaining_urls[:limit]

            total = len(remaining_urls)

            message_queue.put((
                "prepared",
                len(urls),
                len(processed_urls),
                total
            ))

            if not total:
                message_queue.put((
                    "log",
                    "No new URLs available to process."
                ))
                return

            # Start browser
            playwright, browser, context, page = create_browser()

            for index, url in enumerate(remaining_urls, start=1):

                if stop_event.is_set():
                    break

                message_queue.put((
                    "log",
                    f"\nProcessing {index}/{total}\n{url}"
                ))

                email, status = scrape_facebook_page(page, url)

                # Save each result immediately
                save_result(
                    url,
                    email,
                    status,
                    output_file
                )

                if status == "Found":
                    self.found_count += 1

                elif status == "Not Found":
                    self.not_found_count += 1

                else:
                    self.error_count += 1

                message_queue.put((
                    "progress",
                    index,
                    total,
                    status,
                    email,
                    url,
                    self.found_count,
                    self.not_found_count,
                    self.error_count
                ))

                # Wait 3 seconds between pages
                if index < total:
                    if stop_event.wait(3):
                        break

            if stop_event.is_set():
                message_queue.put((
                    "log",
                    "Extraction stopped by user."
                ))

        except Exception as error:

            message_queue.put((
                "log",
                f"Unexpected error: {error}"
            ))

        finally:

            if context:
                context.close()

            if browser:
                browser.close()

            if playwright:
                playwright.stop()

            message_queue.put(("finished",))

    # -------------------------
    # Stop extraction
    # -------------------------

    def stop_extraction(self):

        if self.is_running:

            stop_event.set()

            self.stop_button.config(state="disabled")

            self.progress_label.config(
                text="Stopping after the current page..."
            )

            self.log("Stop requested. Finishing current page...")

    # -------------------------
    # Update counters
    # -------------------------

    def update_counters(self):

        self.found_label.config(
            text=f"Found: {self.found_count}"
        )

        self.not_found_label.config(
            text=f"Not Found: {self.not_found_count}"
        )

        self.error_label.config(
            text=f"Errors: {self.error_count}"
        )

    # -------------------------
    # Process GUI messages
    # -------------------------

    def process_queue(self):

        try:

            while True:

                message = message_queue.get_nowait()

                message_type = message[0]

                if message_type == "log":

                    self.log(message[1])

                elif message_type == "prepared":

                    _, total_urls, processed, selected = message

                    self.log(f"Total URLs in Excel: {total_urls}")
                    self.log(f"Previously processed: {processed}")
                    self.log(f"Selected for processing: {selected}")

                    self.progress["maximum"] = max(selected, 1)

                    if selected == 0:
                        self.progress_label.config(
                            text="Nothing to process."
                        )

                elif message_type == "progress":

                    (
                        _,
                        index,
                        total,
                        status,
                        email,
                        url,
                        found,
                        not_found,
                        errors
                    ) = message

                    self.progress["maximum"] = total
                    self.progress["value"] = index

                    self.progress_label.config(
                        text=f"Completed: {index}/{total}"
                    )

                    self.found_count = found
                    self.not_found_count = not_found
                    self.error_count = errors

                    self.update_counters()

                    self.log(f"Status: {status}")

                    if email:
                        self.log(f"Email: {email}")

                elif message_type == "finished":

                    self.is_running = False

                    self.start_button.config(state="normal")
                    self.stop_button.config(state="disabled")

                    self.log("\nExtraction finished.")
                    self.log(f"Found: {self.found_count}")
                    self.log(f"Not Found: {self.not_found_count}")
                    self.log(f"Errors: {self.error_count}")

                    self.progress_label.config(
                        text="Extraction completed or stopped."
                    )

                    messagebox.showinfo(
                        "Extraction Finished",
                        "Extraction process has finished.\n\n"
                        f"Found: {self.found_count}\n"
                        f"Not Found: {self.not_found_count}\n"
                        f"Errors: {self.error_count}\n\n"
                        f"Output:\n{self.output_path.get()}"
                    )

        except queue.Empty:
            pass

        self.root.after(100, self.process_queue)


# -----------------------------
# Run application
# -----------------------------

if __name__ == "__main__":

    root = tk.Tk()

    app = FacebookExtractorGUI(root)

    root.mainloop()
