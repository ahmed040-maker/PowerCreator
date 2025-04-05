# PowerCreator 🎓📊

**Turn studied content into beautiful PowerPoint presentations—automatically.**

## 🚀 What is PowerCreator?

PowerCreator is an LLM-powered tool that takes in URLs (articles, video transcripts, etc.), summarizes and combines their content, and transforms it into a well-designed PowerPoint presentation using Microsoft Designer.

Whether you're a student who’s studied hard and now needs slides fast, or just someone who hates building PowerPoints manually—this tool is for you.

---

## 🔧 Features

- ✅ Input any list of URLs (articles, videos, blogs, etc.)
- ✅ Scrapes and summarizes content using powerful LLMs (OpenAI/LLama/etc.)
- ✅ Automatically generates presentation outline and speaker notes
- ✅ Uses Microsoft Designer to produce professional slide decks
- ✅ Bundled executable version for non-technical users
- ✅ Sample demo video included

---

## 🛠 Tech Stack

- **Python**
- **LLM API (OpenAI GPT-4 / Claude / etc.)**
- **BeautifulSoup (soon to be replaced by Selenium)**
- **PPTX / Microsoft Designer**
- **Text and content summarization logic**
- **LangChain for source handling (planned)**

---

## 📹 Demo

Watch the video demo here: 

---

## ✅ Usage

### Option 1: As a Python script (for developers)
1. Clone the repo
2. Install requirements  
   ```bash
   pip install -r requirements.txt


3. Obtain an api key from a provider (Open AI, DeepSeek) or run ollama locally with an open source model
4. Setup The configs for model name, url, api key, language ..etc
5. Run the script
6. Provide a list of source URLs
7. Get your .pptx file ready to use 🎉

### Option 2: Use the prebuilt executable (Soon)
No need to install anything.

Just download and run the .exe from the release/ folder (when it is available).

Paste your links and let the tool work its magic.

## Next Steps / Roadmap
🔁 Replace BeautifulSoup with Selenium for dynamic content support
🖼 Add support for more slide types (e.g., title+image, bullet list, quotes)
🗒 Auto-generate summarized outline notes for presenters
🎨 Build a custom presentation designer (remove reliance on MS Designer)
🌀 Add support for Prezi and other non-PPT formats
📂 Support for file inputs: PDFs, Word Docs, Excel sheets, plain text, etc.

## Why This Project?
This is part of my LLM Engineering mini-project series—where I explore how LLMs can solve practical problems and speed up real-life workflows.

I plan to build more tools like this and share them with the community. If you find this useful or have suggestions, feel free to contribute or reach out!

## Contact
Connect with me on LinkedIn
Star the repo if you like it ⭐

## License
MIT License. Free to use and modify.
