<!-- ABOUT THE PROJECT -->
## About WordSearchText

WordSearchText is a PowerShell script that performs text searches in Microsoft Word documents, including both visible text and text formatted as hidden. This tool is useful for quickly identifying occurrences of specific terms in multiple Word documents within a given directory.

### Built With

* [![PShell][PShell-img]][PShell-url]
* [![Microsoft Word][Word-img]][Word-url]

<!-- GETTING STARTED -->
## Getting Started

Before running this PowerShell script, ensure that Microsoft Word is installed on your system, as the script uses the Word COM object to interact with Word documents.

### Prerequisites

To use this script, you need:
* A Windows system with PowerShell installed.
* Microsoft Word installed and configured.

### Installation

1. Save the script file as `WordSearchText.ps1` on your system.
2. Open PowerShell and navigate to the directory containing the script.

### Usage

You can use the script to search for specific text in Word documents within a given directory. 

#### Common usage:
```sh
Search-WordTextInFiles -Path "C:\Path\To\Documents" -SearchText "hidden text"
```

### Examples:

1. Search for the term "confidential" in all Word documents in a folder:
    ```powershell
    Search-WordTextInFiles -Path "C:\Reports" -SearchText "confidential"
    ```

2. Use the pipeline to specify individual files:
    ```powershell
    "C:\Documents\File1.docx", "C:\Documents\File2.docx" | Search-WordTextInFiles -SearchText "specific term"
    ```

3. Search within the current directory:
    ```powershell
    Search-WordTextInFiles -SearchText "important note"
    ```

The results include a list of files containing the search term and the number of occurrences in each file.

### Notes:

- Only `.docx`, `.docm`, and `.doc` file formats are supported.
- The script operates in read-only mode, ensuring no changes are made to the original documents.

### Cleanup:

The script automatically releases the Word COM object after the search process. However, if the script is interrupted, ensure that the Word application is closed manually to free up system resources.


<!-- MARKDOWN LINKS & IMAGES -->
<!-- https://www.markdownguide.org/basic-syntax/#reference-style-links -->
[PShell-img]: https://img.shields.io/badge/PowerShell-5391FE?style=plastic&logo=powershell&logoColor=5391FEf&labelColor=ffffff
[PShell-url]: https://learn.microsoft.com/en-us/training/modules/introduction-to-powershell/
[Word-img]: https://img.shields.io/badge/Microsoft_Word-2B579A?style=plastic&logo=microsoft-word&logoColor=2B579A&labelColor=ffffff
[Word-url]: https://support.microsoft.com/en-us/word