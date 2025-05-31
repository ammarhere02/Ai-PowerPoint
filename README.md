![Python](https://img.shields.io/badge/Python-3.x-blue.svg)
![OpenAI](https://img.shields.io/badge/OpenAI-API-green.svg)
![License](https://img.shields.io/badge/License-MIT-yellow.svg)

# AI PowerPoint Slides Generator

## Description

The AI PowerPoint Slides Generator is a Python-based tool designed to automate the creation of professional PowerPoint
presentations using artificial intelligence. This project leverages AI to generate visually appealing and content-rich
slides based on user inputs, streamlining the process of creating presentations for various purposes such as
educational, business, or personal use. The tool supports generating a trimmed version of the slides for quick previews
and a final polished output for professional use.

## Requirements

- Python 3.x
- OpenAI API key
- UNSPLASH API key
- Required Python packages (see requirements.txt)

## Installation and Usage

1. Clone the repository:
   ```bash
   git clone https://github.com/ammarhere02/Ai-PowerPoint.git
   cd Ai-PowerPoint
   ```

2. Install required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Set up environment variables:
    - Create a `.env` file in the project root
    - Add your OpenAI API key: `OPEN_AI=your_api_key_here`
    - Add your Unsplash API key: `UNSPLASH_API_KEY=your_unsplash_key_here`

4. Usage:
    - For trimmed presentation output:
      ```bash
      python main.py
      ```
    - For enhanced presentation output:
      ```bash
      python main2.py
      ```

Note: Ensure you have a PowerPoint file named "orignal.pptx" in the same directory before running the scripts.

## Features

- Automated slide generation using OpenAI GPT models
- Smart slide content prioritization and scoring
- Support for both trimmed and enhanced presentations
- Integration with Unsplash for image content
- Professional formatting and layout optimization

## Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

