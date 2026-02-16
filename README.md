# Using-LLMs-to-simulate-real-life-participants-answered-questionnaires.
# EasyPsych
An intelligent questionnaire simulation system based on large language models (LLMs). It supports multiple mainstream questionnaire file formats, automatically parses questionnaire structures, and generates realistic survey responses tailored to participants' demographic background data.

> [!NOTE]
> A personal note from the developer: The **GUI interface code** and **concurrent API processing module** were developed with the assistance of AI tools. Currently, only Alibaba Cloud Tongyi Qianwen (Qwen) series models are fully tested and supported in the GUI configuration. More detailed model selection options will be added to the GUI in subsequent updates.

---

## Key Features
- 🎯 **Smart Parsing**: Accurately identifies questionnaire dimensions, questions, scoring rules, and reverse-scoring items
- 📊 **Multi-Format Support**: Full compatibility with Excel, CSV, and Word documents
- 🤖 **LLM Integration**: Built-in Alibaba Cloud Tongyi Qianwen (Qwen) LLM API
- 🎲 **Randomization Options**: Supports flexible question order randomization
- 🔄 **Concurrent Processing**: Multi-threaded API calls to significantly improve simulation efficiency
- 💾 **Smart Saving**: Automatically distinguishes between fully completed and interrupted result files
- 🎨 **Intuitive GUI**: Clean, beginner-friendly graphical operation interface

---

## Core Capabilities

### Supported Questionnaire Formats
| Format | Extensions | Description |
|:-------|:-----------|:------------|
| Excel | .xlsx, .xls | Ideal for structured questionnaire data |
| CSV | .csv | Comma-separated values file with wide compatibility |
| Word | .docx | Flexible natural language format for questionnaire design |

### Detailed Functions
1. **Automatic Questionnaire Parsing**
   - Accurately identifies dimension headers
   - Efficiently extracts question content
   - Intelligently parses scoring rules
   - Automatically detects reverse-scoring items

2. **Intelligent Response Simulation**
   - Generates personalized responses based on participant demographic backgrounds
   - Adapts to multiple psychological scale types
   - Concurrent processing for high-volume simulation tasks

3. **Flexible Configuration**
   - Customizable API key settings
   - Model selection (currently Qwen-only, see [Future Plans](#future-plans))
   - Toggle for question randomization
   - Customizable output format settings

4. **Robust GUI & Error Handling**
   - Organized multi-tab settings interface
   - Real-time progress bar for transparent task tracking
   - Detailed status prompts with traceable error logs
   - Complete error handling and interruption recovery mechanism

---

## 🚀 Quick Start

### System Requirements
- Python 3.8 or higher
- Optimized for Windows operating system

### Install Dependencies
Run the following command in your terminal:
```bash
pip install pandas openpyxl python-docx openai tenacity tkinter
```

### Configure API Key
Open the `config.py` file in the project root directory, and fill in your Alibaba Cloud API key:
```python
DASHSCOPE_API_KEY = "your-api-key-here"
BASE_URL = "https://dashscope.aliyuncs.com/compatible-mode/v1"
MODEL_NAME = "qwen-plus"
```

### Run the Program
Execute the following command in your terminal to launch the GUI:
```bash
python EasyPsych_source_code.py
```

---

## 📝 Usage Guide

### Step 1: Prepare Your Files
1. **Questionnaire File**
   - Excel/CSV format: Must include `Question ID`, `Dimension`, `Question Content`, and `Scoring Standard` fields
   - Word format: Use natural language layout; dimension headers **must end with a colon**, and scoring rules **must start with "Coding:"** for accurate parsing

2. **Participant Demographic File**
   - Only Excel format is supported; must contain basic demographic information of the virtual participants

### Step 2: GUI Configuration
1. **API Settings**
   - Enter your API key, or load it directly from the pre-configured `config.py`
   - Select the supported Qwen model variant
   - Set the maximum token limit for single API responses

2. **Questionnaire Settings**
   - Select your prepared questionnaire file
   - Check the box to enable question order randomization
   - Set the maximum number of consecutive questions from the same dimension

3. **File Path Settings**
   - Select your participant demographic file
   - Specify the output directory for the final result files

### Step 3: Start Processing
Click the **"Start Processing"** button, and the program will automatically complete the following workflow:
1. Parse and validate the questionnaire file structure
2. Display a real-time progress bar for the simulation task
3. Call the LLM API concurrently to generate simulated responses
4. Automatically save the final results to the specified directory

---

## ⚙️ Full Configuration Details

### `config.py` Core Settings
| Configuration Item | Description | Default Value |
|:-------------------|:------------|:--------------|
| `DASHSCOPE_API_KEY` | Alibaba Cloud API key (required) | Empty (user must fill in) |
| `BASE_URL` | API base request URL | `https://dashscope.aliyuncs.com/compatible-mode/v1` |
| `MODEL_NAME` | LLM model name | `qwen-plus` |

### Runtime Parameters
| Parameter | Description | Default Value |
|:----------|:------------|:--------------|
| `MAX_TOKENS` | Maximum token count per API response | 512 |
| `TEMPERATURE` | Response diversity (0-1; higher = more diverse) | 0.7 |
| `API_RETRY_TIMES` | Max retry times for failed API calls | 3 |
| `API_RETRY_DELAY` | Initial retry delay (in seconds, exponential backoff) | 2 |

---

## 📦 Packaging & Distribution

### Use the Built-in Automated Script
The project includes a one-click packaging script `new_build_app.py`. Simply run this command in your terminal:
```bash
python new_build_app.py
```

### Script Features
- ✅ Automatically checks and installs PyInstaller if missing
- ✅ Cleans up old build and temporary files automatically
- ✅ Supports embedding custom application icons
- ✅ Automatically includes the `config.py` configuration file
- ✅ Automatically packages the `icons` resource folder
- ✅ Generates a single standalone executable file for easy distribution

### Packaged Output File
After the packaging is complete, the executable file will be generated in the following path:
```
dist/EasyPsych.exe
```

### Distribution Notes
1. **API Key**: Double-check that the `config.py` file contains a valid API key before packaging
2. **Custom Icon**: To use a custom icon, place your `.ico` file at `icons/EasyPsych.ico`
3. **Pre-Distribution Test**: Always test the packaged executable locally before distribution to ensure normal operation

---

## 📂 Project Structure
```text
EasyPsych/
├── EasyPsych_source_code.py    # Main program entry and core logic
├── config.py                    # API and global configuration file
├── new_build_app.py             # One-click automated packaging script
├── CHANGELOG.md                 # Version update changelog
├── README.md                    # Project documentation (this file)
├── icons/                       # Icon resource folder
│   ├── EasyPsych.ico
│   ├── EasyPsych.jpg
│   └── app_icon.png
├── dist/                        # Packaging output directory
│   └── EasyPsych.exe
└── build/                       # Build temporary file directory
```

---

## 🐛 Frequently Asked Questions

### Q: What should I do if the API call fails?
A: First, verify that your API key in `config.py` is correct and that your Alibaba Cloud account has sufficient balance and available quota. You can also check your network connection stability, and refer to the error log in the GUI for detailed troubleshooting.

### Q: Why is my Word document failing to parse?
A: Ensure that all dimension headers in your Word document end with a colon, and all scoring rules start with "Coding:". You can also adjust the document format to match the standard template to improve parsing accuracy.

### Q: How can I modify the output file name?
A: For successfully completed tasks, the result is automatically saved as `EasyPsych_Results.xlsx`. If the task is interrupted, a timestamp will be automatically added to the file name to avoid overwriting existing files.

### Q: The packaged executable won't run, how to fix it?
A: First, confirm that the `config.py` file is correctly included in the packaging process. You can also check for missing dependencies, try reinstalling all required packages locally, and re-run the packaging script.

---

## 🚧 Future Plans
- Add more detailed model selection options in the GUI, including support for other mainstream open-source and commercial LLMs
- Add support for more questionnaire and scale file formats
- Optimize error handling and interruption recovery capabilities
- Add built-in sample questionnaire and participant demographic files for quick testing
- Add cross-platform support for macOS and Linux

---

## 🤝 Contributing
Issues and Pull Requests are warmly welcomed to help improve this project!

---

## 📄 License
This project is for educational and research purposes only.

---

## 🙏 Acknowledgments
- Alibaba Cloud Tongyi Qianwen (Qwen) Large Language Model
- The excellent open-source libraries from the Python community
- All contributors to this project

---

### Format Compatibility Note
This document strictly follows GitHub Flavored Markdown (GFM) specifications. You can directly copy all the content above into your `README.md` file, and it will render perfectly on GitHub without any format errors.
