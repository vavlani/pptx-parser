# Presentation Parser

This project provides tools to process presentation files (PPTX or PDF), extract metadata, slide images, and text content, and generate summaries or descriptions of each slide using Google's Generative AI models.

## Table of Contents

- [Features](#features)
- [Dependencies](#dependencies)
- [Installation](#installation)
- [Setup](#setup)
- [Usage](#usage)
  - [Processing a PDF Presentation](#processing-a-pdf-presentation)
  - [Processing a PPTX Presentation](#processing-a-pptx-presentation)
- [Functions](#functions)
- [Directory Structure](#directory-structure)
- [Notes](#notes)
- [Troubleshooting](#troubleshooting)

## Features

- **Supports PPTX and PDF files**: Process presentations in both formats.
- **Metadata Extraction**: Extracts detailed metadata from presentations.
- **Slide Image Generation**: Generates images for each slide.
- **Text Content Extraction**: Extracts text content from slides.
- **AI-Powered Summaries**: Generates slide summaries using Google's Generative AI models (Gemini).
- **Configurable Options**: Customize slide selection, image resizing, and metadata extraction.

## Dependencies

Ensure you have the following Python libraries installed:

- `os`
- `hashlib`
- `json`
- `tempfile`
- `time`
- `subprocess`
- `python-pptx`
- `pdf2image`
- `PyPDF2`
- `Pillow`
- `google-generativeai`
- `concurrent.futures`

External Dependencies:

- **LibreOffice**: Required for converting PPTX files to PDF.
- **GhostScript**: Required by `pdf2image` to convert PDF pages to images.

## Installation

### Clone the Repository

```bash
git clone <repository_url>
cd <repository_directory>
```

### Install Python Dependencies

Use `pip` to install the required Python libraries:

```bash
pip install -r requirements.txt
```

Alternatively, install them individually:

```bash
pip install python-pptx pdf2image PyPDF2 Pillow google-generativeai
```

### Install External Dependencies

#### LibreOffice

- **Ubuntu/Debian**:

  ```bash
  sudo apt-get install libreoffice
  ```

- **macOS (Homebrew)**:

  ```bash
  brew install --cask libreoffice
  ```

#### GhostScript

- **Ubuntu/Debian**:

  ```bash
  sudo apt-get install ghostscript
  ```

- **macOS (Homebrew)**:

  ```bash
  brew install ghostscript
  ```

## Setup

### Google Generative AI API Key

Set your Google API key as an environment variable:

```bash
export GOOGLE_API_KEY='your_google_api_key_here'
```

Or within your Python script:

```python
import os
os.environ['GOOGLE_API_KEY'] = 'your_google_api_key_here'
```

## Usage

### Processing a PDF Presentation

```python
# Import necessary functions
from your_module import process_presentation, generate_texts_from_images

# Define file paths and options
presentation_file = './source_pptx/Boston Consulting Group Report.pdf'
image_output_dir = 'slide_images'
output_json = 'presentation_info.json'
slides_to_process = [27, 28, 29]  # Specific slides to process
resize_option = (1024, 768)  # Image resizing

# Metadata extraction configuration
pdf_config = {
    'file_name': True,
    'file_format': True,
    'file_size_bytes': True,
    'file_path': True,
    'creation_date': True,
    'last_modified_date': True,
    'number_of_slides': True,
    'title': True,
    'author': True,
    'subject': True,
    'keywords': True,
    'creator': True,
    'producer': True,
    'pdf_version': True,
    'encrypted': True,
    'mod_date': True,
}

# Process the presentation
presentation_data = process_presentation(
    presentation_file_path=presentation_file,
    image_output_path=image_output_dir,
    slides_to_process=slides_to_process,
    image_format='PNG',
    output_json_path=output_json,
    resize_option=resize_option,
    config=pdf_config
)

# Prepare image and prompt list
image_prompt_list = []
prompt = "Provide a summary of this slide."
for slide in presentation_data['slides']:
    image_path = slide['image_path']
    image_prompt_list.append((image_path, prompt))

# Generate descriptions
descriptions = generate_texts_from_images(
    image_prompt_list=image_prompt_list,
    resize_option=None,
    use_file_api=False,
    model_name='gemini-1.5-flash',
    parallel=False,
    api_call_sleep_seconds=5
)

# Add descriptions to presentation data
for slide, description in zip(presentation_data['slides'], descriptions):
    slide['generated_description'] = description

# Save updated data to JSON
with open(output_json, 'w', encoding='utf-8') as f:
    json.dump(presentation_data, f, ensure_ascii=False, indent=4)
```

### Processing a PPTX Presentation

```python
# Metadata extraction configuration
pptx_config = {
    'file_name': True,
    'file_format': True,
    'file_size_bytes': True,
    'file_path': True,
    'creation_date': True,
    'last_modified_date': True,
    'number_of_slides': True,
    'title': True,
    'subject': True,
    'author': True,
    'last_modified_by': True,
    'description': True,
    'keywords': True,
    'created': True,
    'modified': True,
    'last_printed': True,
    'language': True,
}

presentation_file = './source_pptx/samplepptx.pptx'
image_output_dir = 'slide_images'
output_json = 'presentation_info.json'
slides_to_process = None  # Process all slides
resize_option = (1024, 768)

# Process the presentation
presentation_data = process_presentation(
    presentation_file_path=presentation_file,
    image_output_path=image_output_dir,
    slides_to_process=slides_to_process,
    image_format='PNG',
    output_json_path=output_json,
    resize_option=resize_option,
    config=pptx_config
)

# Prepare image and prompt list
image_prompt_list = []
prompt = "Provide a summary of this slide."
for slide in presentation_data['slides']:
    image_path = slide['image_path']
    image_prompt_list.append((image_path, prompt))

# Generate descriptions
descriptions = generate_texts_from_images(
    image_prompt_list=image_prompt_list,
    resize_option=None,
    use_file_api=False,
    model_name='gemini-1.5-flash',
    parallel=False,
    api_call_sleep_seconds=5
)

# Add descriptions to presentation data
for slide, description in zip(presentation_data['slides'], descriptions):
    slide['generated_description'] = description

# Save updated data to JSON
with open(output_json, 'w', encoding='utf-8') as f:
    json.dump(presentation_data, f, ensure_ascii=False, indent=4)
```

## Functions

### `compute_presentation_hash(pptx_file_path)`

Computes a SHA-256 hash of the presentation file for unique identification.

### `extract_presentation_info(presentation_file_path, is_pdf=False, config=None)`

Extracts metadata and high-level information from the presentation.

- **Parameters**:
  - `presentation_file_path`: Path to the presentation file.
  - `is_pdf`: Boolean indicating if the file is a PDF.
  - `config`: Dictionary specifying which metadata properties to extract.
- **Returns**: Dictionary containing the requested metadata.

### `extract_slide_texts(presentation_source, slides_to_process=None, is_pdf=False)`

Extracts text content from each slide.

- **Parameters**:
  - `presentation_source`: Presentation object or file path.
  - `slides_to_process`: List of slide numbers to process.
  - `is_pdf`: Boolean indicating if the file is a PDF.
- **Returns**: List of texts extracted from each slide.

### `generate_slide_images(...)`

Generates image versions of the slides.

- **Parameters**:
  - `presentation_file_path`: Path to the presentation file.
  - `image_output_path`: Directory to save the slide images.
  - `slides_to_process`: List of slide numbers to process.
  - `image_format`: Format for the slide images.
  - `resize_option`: Tuple `(width, height)` to resize images.
  - `is_pdf`: Boolean indicating if the file is a PDF.
- **Returns**: List of paths to the slide images.

### `process_presentation(...)`

Processes the presentation and returns a dictionary with the extracted information.

- **Parameters**:
  - `presentation_file_path`: Path to the presentation file.
  - `image_output_path`: Directory to save the slide images.
  - `slides_to_process`: List of slide numbers to process.
  - `image_format`: Format for the slide images.
  - `output_json_path`: Path to save the output JSON file.
  - `resize_option`: Tuple `(width, height)` to resize images.
  - `config`: Dictionary specifying which metadata properties to extract.
- **Returns**: Dictionary containing presentation info and slides info.

### `generate_text_from_image(...)`

Generates a text description based on the image and prompt using the Gemini API.

- **Parameters**:
  - `image_path`: Path to the image file.
  - `prompt`: Text prompt to guide the model.
  - `resize_option`: Tuple `(width, height)` to resize the image.
  - `use_file_api`: Boolean indicating whether to upload the image using the File API.
  - `model_name`: Name of the Gemini model to use.
  - `api_call_sleep_seconds`: Seconds to wait before making the API call.
- **Returns**: Dictionary with `response_text`, `error_flag`, and `error_message`.

### `generate_texts_from_images(...)`

Generates text descriptions based on images and prompts using the Gemini API.

- **Parameters**:
  - `image_prompt_list`: List of tuples `(image_path, prompt)`.
  - `resize_option`: Tuple `(width, height)` to resize images.
  - `use_file_api`: Boolean indicating whether to upload images using the File API.
  - `model_name`: Name of the Gemini model to use.
  - `parallel`: Boolean indicating whether to parallelize calls.
  - `api_call_sleep_seconds`: Seconds to wait before each API call.
- **Returns**: List of response texts.

## Directory Structure

```
├── source_pptx/
│   ├── samplepptx.pptx
│   ├── Boston Consulting Group Report.pdf
├── slide_images/
│   ├── slide_1.png
│   ├── slide_2.png
│   ├── ...
├── presentation_info.json
├── your_notebook.ipynb
├── README.md
├── requirements.txt
```

- `source_pptx/`: Contains sample presentation files.
- `slide_images/`: Contains generated slide images.
- `presentation_info.json`: Output JSON with presentation data and slide descriptions.
- `your_notebook.ipynb`: The Jupyter notebook with the code.
- `README.md`: Project documentation.
- `requirements.txt`: List of required Python packages.

## Notes

- **LibreOffice**: Used to convert PPTX files to PDF. Ensure it's installed and accessible via command line.
- **GhostScript**: Required by `pdf2image`. Install it to avoid conversion errors.
- **Generative AI API**: Requires a valid API key and access permissions.
- **Rate Limits**: Be mindful of API rate limits and potential costs.
- **Error Handling**: Functions are designed to handle exceptions and provide meaningful error messages.

## Troubleshooting

- **LibreOffice Conversion Errors**: Confirm LibreOffice installation and system PATH configuration.
- **GhostScript Errors**: Ensure GhostScript is installed and accessible.
- **API Errors**: Verify API key validity and model access permissions.
- **File Path Issues**: Use absolute paths or ensure correct relative paths based on the working directory.
- **Module Import Errors**: Verify all dependencies are installed and correctly imported.

## License

This project is licensed under the MIT License.
