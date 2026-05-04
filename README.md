# IT3040 Assignment 1 - Option 1
## Singlish Transliteration Accuracy Testing

This project tests the accuracy of the Chat Sinhala transliteration function 
at https://www.pixelssuite.com/chat-translator using Playwright automation.

## Prerequisites
- Python 3.11 or 3.12
- Google Chrome (recommended)

## Installation

Run the following commands one by one:

```bash
pip install -U pip
pip install playwright openpyxl
playwright install
```

## How to Run

1. Open Command Prompt
2. Navigate to the project folder:
```bash
cd /d D:\test_automation
```
3. Run the automation script:
```bash
python test_automation.py --excel "Assignment 1 - Test cases.xlsx" --url "https://www.pixelssuite.com/chat-translator" --wait-ms 5000 --type-delay-ms 80 --slow-mo-ms 200 --save-every 1 --keep-open
```

## Files
- `test_automation.py` - Playwright automation script
- `Assignment 1 - Test cases.xlsx` - Test cases with 50 negative test cases
