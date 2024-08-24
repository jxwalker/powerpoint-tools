# PowerPoint Note Summarizer

This application extracts and summarizes notes from PowerPoint presentations using various AI services. It provides a flexible and efficient way to process presentation notes and generate concise summaries.

## Features

- Extract notes from PowerPoint presentations
- Summarize notes using different AI services (OpenAI, IBM Watson, Anthropic Claude)
- Multiple output formats (DOCX, Markdown, PDF, TXT)
- Configurable summarization levels
- Throttling and retry mechanisms for API rate limiting
- Verbose logging option for debugging

## Requirements

- Python 3.7+
- Required Python packages (install via `pip install -r requirements.txt`):
  - python-pptx
  - python-docx
  - fpdf
  - aiofiles
  - asyncio_throttle
  - ibm_watson
  - openai
  - anthropic
  - tqdm

## Setup

1. Clone the repository:
   ```
   git clone https://github.com/yourusername/powerpoint-note-summarizer.git
   cd powerpoint-note-summarizer
   ```

2. Install the required packages:
   ```
   pip install -r requirements.txt
   ```

3. Create a `config.json` file in the root directory with your API keys and settings:
   ```json
   {
     "watson_api_key": "your_watson_api_key",
     "watson_service_url": "your_watson_service_url",
     "openai_api_key": "your_openai_api_key",
     "openai_model": "gpt-3.5-turbo",
     "openai_max_tokens": 150,
     "openai_temperature": 0.7,
     "default_summarization_level": 3,
     "anthropic_api_key": "your_anthropic_api_key",
     "anthropic_model": "claude-3-haiku-20240307",
     "anthropic_max_tokens": 150,
     "min_characters": 50,
     "rate_limit": 1,
     "max_retries": 3,
     "perplexity_api_key": "your_perplexity_api_key"
   }
   ```

## Usage

Run the application using the following command:

```
python main.py <presentation_path> <output_path> [options]
```

### Arguments:

- `presentation`: Path to the PowerPoint presentation file.
- `output`: Path to the output summary file.

### Options:

- `--ai`: Specify the AI service to use for summarization (openai, watson, claude). Default: watson
- `--config`: Path to the configuration file. Default: config.json
- `--format`: Specify the output format (docx, md, pdf, txt). Default: docx
- `--extract-only`: Only extract notes without summarizing.
- `--summary-only`: Produce only summaries without notes.
- `--verbose`: Enable verbose logging.
- `--summarization-level`: Specify the number of bullet points for summarization.

### Example:

```
python main.py presentations/my_presentation.pptx summaries/output_summary.docx --ai openai --format docx --verbose
```

This command will process `my_presentation.pptx`, summarize its notes using OpenAI, and save the output as a Word document with verbose logging enabled.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
