# OpenXmlBrowser
Generates a .html file (and displays it in the default browser) with the contents of any Open XML file. .docx, .xlsx, .pptx, etc.

# Overview

I created this as a tool to help me work I'm doing with Excel spreadsheets. In the middle of working on it, I realized it can also be used to look at any Open XML format file.

# Features

- Runs the XML through a formatter (that's why all the XML files aren't 2 lines long).
- Files with the extension .PNG, .JPG, and .JPEG are displayed as images 
- Generates a single, self-contained HTML file (embedded CSS and images).


# Generated HTML and CSS

The HTML and CSS came from some pages in the site, https://renenyffenegger.ch/ that I found while Googling. He turned out to be a prolific github contributor. https://github.com/ReneNyffenegger. René has kindly given me permission to use this in my code.

# Example

Command: OpenXmlBrowser D:\Temp\PRs.xlsx

Output: 

![OpenXmlSample](https://user-images.githubusercontent.com/2781885/167685074-d0d313ef-813b-4ea0-8926-7ff149f7b9a8.PNG)
