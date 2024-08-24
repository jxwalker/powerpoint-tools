import asyncio
import argparse
import logging
import sys
import os
import json
from tqdm import tqdm
from typing import List, Tuple
from asyncio_throttle import Throttler

from config import load_config, configure_logging, Config
from extractors import extract_notes_from_presentation
from summarizers import summarize_with_openai, summarize_with_watson, summarize_with_claude, SummarizationError
from writers import write_summary_to_word, write_summary_to_md, write_summary_to_pdf, write_summary_to_txt

async def summarize_note(note: str, slide_number: int, ai: str, config: Config, summarization_level: int, throttler: Throttler) -> str:
    for attempt in range(config.max_retries):
        try:
            if ai == "openai":
                return await summarize_with_openai(note, config, summarization_level, throttler)
            elif ai == "watson":
                return await summarize_with_watson(note, config, summarization_level, throttler)
            elif ai == "claude":
                return await summarize_with_claude(note, config, summarization_level, throttler)
            else:
                return note
        except SummarizationError as e:
            if attempt == config.max_retries - 1:
                logging.error(f"Failed to summarize slide {slide_number} after {config.max_retries} attempts: {str(e)}")
                return f"Error in summarization: {str(e)}"
            else:
                logging.warning(f"Attempt {attempt + 1} failed for slide {slide_number}: {str(e)}. Retrying...")
                await asyncio.sleep(2 ** attempt)  # Exponential backoff

async def main() -> None:
    parser = argparse.ArgumentParser(description="Extract and summarize notes from a PowerPoint presentation.")
    parser.add_argument("presentation", type=str, help="Path to the PowerPoint presentation file.")
    parser.add_argument("output", type=str, help="Path to the output summary file.")
    parser.add_argument("--ai", type=str, choices=["openai", "watson", "claude"], default="watson", help="Specify the AI service to use for summarization.")    
    parser.add_argument("--config", type=str, default="config.json", help="Path to the configuration file.")
    parser.add_argument("--format", type=str, choices=["docx", "md", "pdf", "txt"], default="docx", help="Specify the output format.")
    parser.add_argument("--extract-only", action="store_true", help="Only extract notes without summarizing.")
    parser.add_argument("--summary-only", action="store_true", help="Produce only summaries without notes.")
    parser.add_argument("--verbose", action="store_true", help="Enable verbose logging.")
    parser.add_argument("--summarization-level", type=int, help="Specify the number of bullet points for summarization.")
    
    # Check if any arguments were provided
    if len(sys.argv) == 1:
        parser.print_help(sys.stderr)
        sys.exit(1)

    args = parser.parse_args()
    try:
        await configure_logging(args.verbose)
        try:
            config = await load_config(args.config)
        except FileNotFoundError as e:
            logging.error(f"Configuration file not found: {args.config}")
            logging.error(f"Current working directory: {os.getcwd()}")
            logging.error(f"Full path to config file: {os.path.abspath(args.config)}")
            sys.exit(1)
        except json.JSONDecodeError as e:
            logging.error(f"Error parsing configuration file: {str(e)}")
            sys.exit(1)

        if args.summarization_level is None:
            args.summarization_level = config.default_summarization_level

        notes = await extract_notes_from_presentation(args.presentation)
        summaries = []

        throttler = Throttler(rate_limit=config.rate_limit)

        if not args.extract_only:
            for slide_number, note in tqdm(notes, desc="Summarizing notes"):
                if note == "No notes found." or note == "No notes slide.":
                    summary = note
                elif len(note) < config.min_characters:
                    summary = note
                    logging.info(f"Slide {slide_number}: Note too short for summarization (characters: {len(note)})")
                else:
                    summary = await summarize_note(note, slide_number, args.ai, config, args.summarization_level, throttler)
                
                if args.summary_only:
                    summaries.append((slide_number, summary, ""))
                else:
                    summaries.append((slide_number, summary, note))
        else:
            summaries = [(slide_number, "", note) for slide_number, note in notes]

        if args.format == "docx":
            logging.info(f"Preparing to write {len(summaries)} summaries to {args.output}")
            await write_summary_to_word(summaries, args.output, args.summary_only)
            logging.info(f"Finished writing summaries to {args.output}")
        elif args.format == "md":
            await write_summary_to_md(summaries, args.output, args.summary_only)
        elif args.format == "pdf":
            await write_summary_to_pdf(summaries, args.output, args.summary_only)
        elif args.format == "txt":
            await write_summary_to_txt(summaries, args.output, args.summary_only)
        else:
            logging.error(f"Unsupported output format: {args.format}")
            sys.exit(1)

    except FileNotFoundError as e:
        logging.error(f"File or directory not found: {str(e)}")
        sys.exit(1)
    except PermissionError as e:
        logging.error(f"Permission denied: {str(e)}")
        sys.exit(1)
    except IOError as e:
        logging.error(f"IO Error occurred: {str(e)}")
        sys.exit(1)
    except Exception as e:
        logging.error(f"An unexpected error occurred: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        logging.info("Program interrupted by user. Exiting...")
        sys.exit(0)
    except Exception as e:
        logging.error(f"An unexpected error occurred: {str(e)}")
        sys.exit(1)