# Using-LLMs-to-simulate-real-life-participants-answered-questionnaires.
# EasyPsych 

An intelligent questionnaire simulation system based on large language models (LLMs). It supports multiple mainstream questionnaire file formats, automatically parses questionnaire structures, and generates realistic survey responses tailored to participants' demographic background data. 

> [!NOTE] 
> A personal note from the developer: The **GUI interface code** and **concurrent API processing module** were developed with the assistance of AI tools. Currently, only Alibaba Cloud Tongyi Qianwen (Qwen) series models are fully tested and supported in the GUI configuration. More detailed model selection options will be added to the GUI in subsequent updates. 

--- 

## 🎯 Key Features 

### Core Capabilities
- **Smart Parsing**: Accurately identifies questionnaire dimensions, questions, scoring rules, and reverse-scoring items 
- **Multi-Format Support**: Full compatibility with Excel, CSV, and Word documents 
- **LLM Integration**: Built-in Alibaba Cloud Tongyi Qianwen (Qwen) LLM API 
- **Randomization Options**: Supports flexible question order randomization 
- **Concurrent Processing**: Multi-threaded API calls to significantly improve simulation efficiency 
- **Smart Saving**: Automatically distinguishes between fully completed and interrupted result files 
- **Intuitive GUI**: Clean, beginner-friendly graphical operation interface

### Enhanced Features (Latest Updates)
- **Memory Function**: Automatically remembers user settings, language preferences, and file selections
- **Smart Interface Flow**: Language selection and welcome screens only appear on first launch
- **Improved Cancellation**: Enhanced cancel functionality with confirmation dialogs
- **Progress Optimization**: Optimized progress bar with better UI responsiveness

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
1. **Language Selection** (First Launch Only)
   - Choose between Chinese and English interface
   - This selection is remembered for future launches

2. **API Settings** 
   - Enter your API key, or load it directly from the pre-configured `config.py` 
   - Select the supported Qwen model variant 
   - Set the maximum token limit for single API responses 

3. **Questionnaire Settings** 
   - Select your prepared questionnaire file 
   - Check the box to enable question order randomization 
   - Set the maximum number of consecutive questions from the same dimension 

4. **File Requirements** 
   - View detailed file format requirements and specifications
   - Understand supported question formats and scoring rules

5. **File Path Settings** 
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

### Memory Function Settings
| Setting | Description | Storage Location |
|:--------|:------------|:-----------------|
| Language Preference | User's language choice | `user_settings.json` |
| Welcome Screen Status | Whether welcome screen has been shown | `user_settings.json` |
| API Configuration | API key, base URL, model settings | `user_settings.json` |
| Questionnaire Settings | Randomization, token limits, age ranges | `user_settings.json` |
| File Paths | Last used questionnaire and background files | `user_settings.json` |
| Output Settings | Output format, filename, directory | `user_settings.json` |

--- 

## 🔄 Enhanced Features (Latest Updates)

### Smart Memory Function
- **Automatic Settings Persistence**: All user settings are automatically saved and restored
- **Language Preference Memory**: Your language choice is remembered across sessions
- **File Path Memory**: Last used file paths are preserved for convenience
- **One-Time Welcome**: Welcome and language selection screens only appear on first launch

### Improved User Interface
- **Streamlined Workflow**: Reduced repetitive steps for experienced users
- **File Requirements Tab**: Dedicated section for file format specifications
- **Enhanced Progress Tracking**: Better progress bar with improved responsiveness
- **Smart Cancellation**: Confirmation dialog for cancel operations

### Advanced Processing Features
- **Concurrent API Calls**: Multi-threaded processing for faster simulations
- **Error Handling**: Robust error handling with detailed error messages
- **Progress Monitoring**: Real-time progress tracking with status updates
- **Flexible Output**: Multiple output formats (Excel, CSV) with customizable filenames

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
- ✅ Includes all necessary configuration files and dependencies
- ✅ Generates standalone executable for easy distribution

### Packaging Requirements
- Python 3.8+ with PyInstaller installed
- All required dependencies (see Install Dependencies section)
- Sufficient disk space for build process

### Generated Files
After successful packaging, you'll find:
- `dist/EasyPsych.exe` - Main executable file
- `user_settings.json` - User preferences (created on first run)
- All necessary configuration files embedded in the executable

--- 

## 🐛 Troubleshooting 

### Common Issues
1. **API Connection Errors**
   - Verify your API key is correct and has sufficient quota
   - Check internet connectivity
   - Ensure the API endpoint URL is accessible

2. **File Parsing Errors**
   - Verify file formats match supported specifications
   - Check for required columns in Excel/CSV files
   - Ensure proper formatting in Word documents

3. **Memory Function Issues**
   - If settings are not persisting, check write permissions
   - Verify `user_settings.json` file is not corrupted
   - Try deleting `user_settings.json` to reset preferences

### Performance Tips
- Use smaller batch sizes for better memory management
- Enable question randomization for more realistic simulations
- Monitor API usage to avoid rate limiting
- Use appropriate token limits based on questionnaire complexity

--- 

## 🔮 Future Plans 

### Planned Enhancements
- **Multi-Model Support**: Integration with additional LLM providers
- **Advanced Analytics**: Built-in statistical analysis features
- **Batch Processing**: Support for processing multiple questionnaires simultaneously
- **Custom Templates**: User-defined questionnaire templates
- **Export Options**: Additional output formats and customization

### Community Contributions
We welcome contributions from the community! Feel free to:
- Report bugs and suggest improvements
- Submit pull requests for new features
- Share your use cases and success stories
- Help improve documentation and translations

--- 

## 📄 License 

This project is provided as-is for educational and research purposes. Please ensure compliance with the terms of service of any third-party APIs used.

--- 

## 🤝 Contributing 

For questions, bug reports, or feature requests, please open an issue on the project repository.

---

*Last Updated: February 2025*
