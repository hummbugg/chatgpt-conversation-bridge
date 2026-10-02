# ChatGPT Conversation Bridge Documentation

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

## Document Navigator

### Overview & Setup

- [Purpose](#purpose)
- [Requirements](#requirements)
- [Installing Python](#installing-python)
- [Download and Extract ChatGPT Conversation Bridge](#download-and-extract-chatgpt-conversation-bridge)

### Archiving a Conversation

- [Normal Workflow](#normal-workflow)
- [Create the ChatGPT Shared Link](#create-the-chatgpt-shared-link)

### Running Conversation Bridge

- [Run ChatGPT Conversation Bridge](#run-chatgpt-conversation-bridge)
- [Supported Input Types](#supported-input-types)
- [What Happens on the First Run](#what-happens-on-the-first-run)
- [Console Report and Status: SUCCESS](#console-report-and-status-success)
- [Output Locations](#output-locations)

### Archive Management

- [Permanent Archive and DOCX Regeneration](#permanent-archive-and-docx-regeneration)
- [Archive Collision Protection](#archive-collision-protection)
- [Safety and File-Handling Rules](#safety-and-file-handling-rules)

### The Generated DOCX

- [DOCX Content and Formatting](#docx-content-and-formatting)
- [Images and Attachments](#images-and-attachments)
- [Reference URL Handling](#reference-url-handling)
- [Opening the DOCX in Microsoft Word](#opening-the-docx-in-microsoft-word)
- [Opening the DOCX in LibreOffice Writer](#opening-the-docx-in-libreoffice-writer)
- [Read-Only / Open-View-Only Behavior](#read-only--open-view-only-behavior)
- [Intentionally Editing the DOCX](#intentionally-editing-the-docx)

### Testing & Reference

- [Tested Platforms and Python Versions](#tested-platforms-and-python-versions)
- [Production Testing and Validation](#production-testing-and-validation)
- [Known Limitations](#known-limitations)
- [Troubleshooting](#troubleshooting)

---

## Purpose

The primary purpose of ChatGPT Conversation Bridge is to make it practical to continue a valuable ChatGPT conversation after that conversation has reached its maximum length. Instead of losing the working context or manually reconstructing it, the completed conversation can be converted into a readable DOCX and carried forward into a new ChatGPT conversation.

For chat continuation, create a ChatGPT shared link for the completed or near-limit conversation and pass that shared URL directly to ChatGPT Conversation Bridge. The program retrieves the public shared-conversation data and supported uploaded images, creates a permanent ZIP archive, and generates a readable DOCX. The DOCX gives the new chat a detailed record of the earlier discussion, including user and assistant messages, timestamps when available, formatting, public reference URLs, and preserved supported images.

The DOCX also serves an important archival purpose. It creates a portable, human-readable record of the conversation that can be opened independently of ChatGPT in Microsoft Word, LibreOffice Writer, or another compatible DOCX reader.

The permanent ZIP created by ChatGPT Conversation Bridge preserves the conversation data and supported uploaded images needed to regenerate the DOCX without retrieving the shared conversation again. The DOCX is the practical document for reading, reviewing, archiving, and supplying prior conversation context to a continuation chat.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Requirements

ChatGPT Conversation Bridge requires:

- A desktop computer running Windows, Linux, or macOS.
- `Python 3.10` or later.
- The Python `curl_cffi` and `tzdata` packages. Version `0.16.3` of `curl_cffi` was used during final Version 2.0.0 cross-platform validation.
- Internet access is required when creating a new archive from a ChatGPT shared URL. Internet access is not required when regenerating a DOCX from an existing permanent archive ZIP.
- A ChatGPT shared-conversation URL for the conversation being archived.
- Microsoft Word, LibreOffice Writer, or another compatible DOCX reader to view the generated document.

The exact operating systems and Python versions that were tested are documented separately under [Tested Platforms and Python Versions](#tested-platforms-and-python-versions).

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Installing Python

ChatGPT Conversation Bridge requires `Python 3.10` or later and the third-party Python `curl_cffi` and `tzdata` packages.

Version `0.16.3` of `curl_cffi` was used during final Version 2.0.0 cross-platform validation.

### Windows

Open **Command Prompt** and check the installed Python version:

```text
python --version
```

If `Python 3.10` or later is reported, Python does not need to be installed again.

If Python is not installed or the version is older than `Python 3.10`, download and install the current version of Python from the official Python website:

```text
https://www.python.org/downloads/windows/
```

After installation, open a new **Command Prompt** and verify the version:

```text
python --version
```

Confirm that `Python 3.10` or later is reported.

#### cURL Installation for Windows

Install the required Python `curl_cffi` and `tzdata` packages:

```text
python -m pip install curl_cffi tzdata
```

Then verify the installed `curl_cffi` package version:

```text
python -c "import curl_cffi; print(curl_cffi.__version__)"
```

Version `0.16.3` was used during final Version 2.0.0 validation.

### Ubuntu / Linux

Open **Terminal** and install Python:

```text
sudo apt install -y python3-full
```

Then verify the installed version:

```text
python3 --version
```

Confirm that `Python 3.10` or later is reported.

The installation command above is for Ubuntu and other Linux distributions that use the `apt` package manager. Other Linux distributions may use a different package manager.

#### cURL Installation for Ubuntu

ChatGPT Conversation Bridge requires the Python `curl_cffi` and `tzdata` packages. The `curl_cffi` package is separate from the ordinary `curl` command-line program.

If you want to verify that the ordinary `curl` command is available, run:

```text
curl --version
```

If it is not installed on Ubuntu, it can normally be installed with:

```text
sudo apt install -y curl
```

Install the required Python `curl_cffi` and `tzdata` packages with:

```text
sudo python3 -m pip install --break-system-packages curl_cffi tzdata
```

The `--break-system-packages` option deliberately permits `pip` to install the packages into an externally managed Python installation. Use it here only if you intend to install `curl_cffi` and `tzdata` into that system Python environment.

Then verify the installed `curl_cffi` package version:

```text
python3 -c "import curl_cffi; print(curl_cffi.__version__)"
```

Version `0.16.3` was used during final Version 2.0.0 validation.

### macOS

Open **Terminal** and check the installed Python version:

```text
python3 --version
```

If `Python 3.10` or later is reported, Python does not need to be installed again.

If Python is not installed or the version is older than `Python 3.10`, download the current macOS installer from the official Python website:

```text
https://www.python.org/downloads/macos/
```

Open the downloaded installer and follow the installation prompts.

After installation, open a new **Terminal** and verify Python and `pip`:

```text
python3 --version
python3 -m pip --version
```

Confirm that `Python 3.10` or later is reported.

#### Python Package Installation for macOS

Install the required Python `curl_cffi` and `tzdata` packages:

```text
python3 -m pip install curl_cffi tzdata
```

Then verify the installed `curl_cffi` package version:

```text
python3 -c "import curl_cffi; print(curl_cffi.__version__)"
```

Version `0.16.3` was used during final Version 2.0.0 validation.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Download and Extract ChatGPT Conversation Bridge

ChatGPT Conversation Bridge is distributed as a ZIP file under **Assets** on the project's GitHub release page.

Download the ZIP file for the version of ChatGPT Conversation Bridge that you want to install.

### Extract the ZIP into Your Home Directory

Extract the downloaded ZIP file directly into your user **home directory**.

Do **not** create the `chatgpt-conversation-bridge` directory yourself before extracting the ZIP. The ZIP file already contains the `chatgpt-conversation-bridge` directory, and extracting the ZIP into your home directory creates it automatically.

Your home directory depends on your operating system.

**Windows**

```text
C:\Users\<username>
```

**Ubuntu/Linux**

```text
/home/<username>
```

**macOS**

```text
/Users/<username>
```

Replace `<username>` with your operating-system username.

On Ubuntu/Linux and macOS, `~` is commonly used as shorthand for the current user's home directory. For example, on Ubuntu/Linux:

```text
~
```

refers to:

```text
/home/<username>
```

On Ubuntu, the graphical **Files** application normally displays `/home/<username>` simply as **Home**.

### Result After Extraction

After the ZIP has been extracted into your home directory, it creates the `chatgpt-conversation-bridge` directory.

The resulting location is:

**Windows**

```text
C:\Users\<username>\chatgpt-conversation-bridge
```

**Ubuntu/Linux**

```text
/home/<username>/chatgpt-conversation-bridge
```

**macOS**

```text
/Users/<username>/chatgpt-conversation-bridge
```

Throughout this documentation, this extracted `chatgpt-conversation-bridge` directory is referred to as the **ChatGPT Conversation Bridge working folder**.

Do not extract the ZIP into a `chatgpt-conversation-bridge` directory that you created yourself. Doing so may create an unnecessary nested directory such as:

```text
chatgpt-conversation-bridge/
└── chatgpt-conversation-bridge/
```

### Extracted Files and Folders

Immediately after the Version 2.0.0 release ZIP has been extracted, and before ChatGPT Conversation Bridge has been run, the working folder contains:

```text
chatgpt-conversation-bridge/
├── documentation/
│   ├── chatgpt-continuing-a-chat.md
│   ├── chatgpt-conversation-bridge-documentation.md
│   └── chatgpt-conversation-continuation-instructions.md
├── chatgpt_conversation_bridge.py
├── LICENSE
└── README.md
```

The production Python script is:

```text
chatgpt_conversation_bridge.py
```

The project documentation is stored under:

```text
documentation/
```

The `archive` and `docx` directories are not included in the extracted release package and do not need to be created manually. ChatGPT Conversation Bridge creates them automatically when they are needed.

After the program has been run successfully with a valid ChatGPT shared-conversation URL, the working folder has the following structure:

```text
chatgpt-conversation-bridge/
├── archive/
├── documentation/
│   ├── chatgpt-continuing-a-chat.md
│   ├── chatgpt-conversation-bridge-documentation.md
│   └── chatgpt-conversation-continuation-instructions.md
├── docx/
├── chatgpt_conversation_bridge.py
├── LICENSE
└── README.md
```

The `archive` directory contains permanent conversation archives created by ChatGPT Conversation Bridge.

The `docx` directory contains the generated DOCX conversation documents.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Normal Workflow

ChatGPT Conversation Bridge uses a ChatGPT shared-conversation URL as its normal input. The program retrieves the public shared-conversation data directly, so saving the conversation through a web browser is no longer required.

To obtain the shared-conversation URL, open the conversation you want to preserve and click **Share** in the upper-right corner of the ChatGPT conversation page. ChatGPT automatically copies the shared-conversation URL to the clipboard and displays a confirmation that the link was copied.

A shared URL has the following general form:

```text
https://chatgpt.com/share/<UUID>
```

Because the program works from the shared conversation rather than a browser-saved copy of the page, the extraction workflow does not depend on Google Chrome, Microsoft Edge, Opera, Firefox, or another particular web browser.

Before running ChatGPT Conversation Bridge, make sure the shared version reflects the conversation state you intend to preserve. The data available through a shared link can sometimes lag behind the current live conversation, so recently added messages may not immediately appear in the shared version.

When the shared version reflects the conversation state you want to preserve, use the shared-conversation URL that ChatGPT copied to the clipboard as the input to ChatGPT Conversation Bridge.


ChatGPT Conversation Bridge uses a ChatGPT shared-conversation URL directly. There is no browser Save As step, no HTML/HTM input file, no `_files` companion folder, and no requirement to load or scroll through the shared conversation in a web browser before running the program.

The overall workflow is:

1. Click **Share** in the ChatGPT conversation you want to preserve. ChatGPT automatically copies the shared-conversation URL to the clipboard.
2. Run ChatGPT Conversation Bridge and supply the shared-conversation URL as its input.
3. ChatGPT Conversation Bridge retrieves the public shared-conversation data.
4. The program retrieves supported uploaded images that are available through the shared conversation.
5. The program creates a permanent ZIP archive under `archive/`.
6. The program creates the readable DOCX under `docx/`.
7. Confirm that the program finishes with `Status: SUCCESS`.
8. Review the generated DOCX.
9. Preserve the permanent ZIP under `archive/` for future DOCX regeneration.

After a permanent archive has been created, ChatGPT Conversation Bridge can regenerate its DOCX directly from that ZIP without retrieving the shared conversation again.

The following sections describe each part of this workflow in detail.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Create the ChatGPT Shared Link

Begin in the original ChatGPT conversation that you want to preserve.

1. Click **Share** in the upper-right corner of the ChatGPT conversation page.
2. The shared-conversation URL is copied to the clipboard automatically.
3. Confirm that ChatGPT displays the message indicating that the link was copied.

The shared URL has the following general form:

```text
https://chatgpt.com/share/<UUID>
```

ChatGPT Conversation Bridge uses this URL directly. You do not need to open the shared conversation in another browser tab, scroll through it, use Developer Tools, or save the page to your computer.

Before running ChatGPT Conversation Bridge, make sure the shared version of the conversation reflects the conversation state you intend to preserve. The data available through a shared link can sometimes lag behind the current live conversation, so recently added messages may not immediately appear in the shared version.

When the conversation has reached the state you want to preserve, click **Share** in the upper-right corner of the ChatGPT conversation page. The shared-conversation URL is copied to the clipboard automatically, and ChatGPT displays a confirmation that the link was copied. Proceed to [Run ChatGPT Conversation Bridge](#run-chatgpt-conversation-bridge).

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Run ChatGPT Conversation Bridge

Open a terminal or command prompt in the ChatGPT Conversation Bridge working folder.

When creating a new archive, run the program with the ChatGPT shared-conversation URL as its single command-line argument.

**Windows**

```text
python chatgpt_conversation_bridge.py "https://chatgpt.com/share/<UUID>"
```

**Linux/macOS**

```text
python3 chatgpt_conversation_bridge.py "https://chatgpt.com/share/<UUID>"
```

Replace `https://chatgpt.com/share/<UUID>` with the shared-conversation URL that was copied to the clipboard when you clicked **Share** in ChatGPT.

ChatGPT Conversation Bridge retrieves the public shared-conversation data directly. It then retrieves supported uploaded images that are available through the shared conversation, creates the permanent archive, and generates the DOCX.

Allow the program to finish. A successful run ends with:

```text
Status: SUCCESS
```

Do not delete or replace any output files until the program has completed successfully.

To regenerate a DOCX later from an existing permanent archive, supply the archive ZIP as the command-line argument instead of a shared-conversation URL. This process is described under [Permanent Archive and DOCX Regeneration](#permanent-archive-and-docx-regeneration).

The following sections explain input handling, first-run processing, the console report, and output locations in detail.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Supported Input Types

ChatGPT Conversation Bridge accepts one command-line argument identifying the source to process. Version 2.0.0 supports two source types:

1. A ChatGPT shared-conversation URL for creating a new permanent archive and DOCX.
2. An existing ChatGPT Conversation Bridge archive ZIP for regenerating a DOCX without retrieving the shared conversation again.

### ChatGPT Shared-Conversation URL

To create a new archive and DOCX, supply the shared-conversation URL copied to the clipboard when you click **Share** in ChatGPT.

**Windows**

```text
python chatgpt_conversation_bridge.py "https://chatgpt.com/share/<UUID>"
```

**Linux/macOS**

```text
python3 chatgpt_conversation_bridge.py "https://chatgpt.com/share/<UUID>"
```

ChatGPT Conversation Bridge retrieves the public shared-conversation data and available supported uploaded images, then creates the permanent archive and DOCX.

### Existing Archive ZIP

To regenerate a DOCX from an existing permanent archive, supply the ZIP archive name instead of a shared-conversation URL.

Several archive-path forms are supported.

#### Archive Filename Only

When only the ZIP filename is supplied, ChatGPT Conversation Bridge looks for that file in its `archive` folder.

**Windows**

```text
python chatgpt_conversation_bridge.py "Ear Crevice in Dog.zip"
```

**Linux/macOS**

```text
python3 chatgpt_conversation_bridge.py "Ear Crevice in Dog.zip"
```

#### Archive Folder Path

The archive folder can also be included explicitly.

**Windows**

```text
python chatgpt_conversation_bridge.py "archive\Ear Crevice in Dog.zip"
```

or:

```text
python chatgpt_conversation_bridge.py ".\archive\Ear Crevice in Dog.zip"
```

**Linux/macOS**

```text
python3 chatgpt_conversation_bridge.py "archive/Ear Crevice in Dog.zip"
```

or:

```text
python3 chatgpt_conversation_bridge.py "./archive/Ear Crevice in Dog.zip"
```

#### Explicit Current-Relative ZIP Path

A ZIP can also be specified explicitly relative to the current working folder.

**Windows**

```text
python chatgpt_conversation_bridge.py ".\Ear Crevice in Dog.zip"
```

**Linux/macOS**

```text
python3 chatgpt_conversation_bridge.py "./Ear Crevice in Dog.zip"
```

When this form is used with only a ZIP filename, ChatGPT Conversation Bridge resolves the archive through its `archive` folder.

#### Absolute ZIP Path

An absolute path to an existing archive ZIP is also supported.

**Windows example**

```text
python chatgpt_conversation_bridge.py "D:\chatgpt-conversation-bridge\archive\Ear Crevice in Dog.zip"
```

**Linux example**

```text
python3 chatgpt_conversation_bridge.py "/home/<username>/chatgpt-conversation-bridge/archive/Ear Crevice in Dog.zip"
```

**macOS example**

```text
python3 chatgpt_conversation_bridge.py "/Users/<username>/chatgpt-conversation-bridge/archive/Ear Crevice in Dog.zip"
```

Quoted arguments are recommended, especially when a conversation or path contains spaces.

The browser-capture inputs used by earlier versions are no longer supported. Version 2.0.0 does not process saved `.html` or `.htm` files or their `_files` companion folders.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## What Happens on the First Run

When ChatGPT Conversation Bridge is run with a ChatGPT shared-conversation URL, it retrieves the public shared-conversation data directly rather than processing a browser-saved copy of the conversation.

The program:

1. Retrieves the public shared-conversation data from the supplied ChatGPT shared URL.
2. Validates the conversation structure and message chain.
3. Identifies uploads referenced by the conversation.
4. Retrieves supported uploaded images that are available through the shared conversation.
5. Creates a permanent ZIP archive containing the conversation data and the supported uploaded images that were successfully retrieved.
6. Creates the DOCX from the retrieved conversation data and images.
7. Reports the processing results and finishes with `Status: SUCCESS` when the conversion completes successfully.

The permanent archive contains the source data needed by ChatGPT Conversation Bridge to regenerate the DOCX later without retrieving the shared conversation again.

The archive has the following general structure:

```text
CONVERSATION_NAME.zip
├── conversation.json
└── uploads/
    ├── file_<stable-id>.<extension>
    └── ...
```

`conversation.json` contains the archived shared-conversation data. The `uploads` folder contains the supported uploaded images that were successfully retrieved for that conversation.

If processing fails, do not treat a partial result as a successfully completed archive. Investigate the reported error and run ChatGPT Conversation Bridge again after the problem has been corrected.

After a successful first run, the permanent ZIP under `archive/` becomes the retained source from which the DOCX can later be regenerated.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Console Report and Status: SUCCESS

While ChatGPT Conversation Bridge runs, it prints a console report describing the conversation it processed and the results of the conversion.

The report provides processing information such as the source being used, conversation and message statistics, upload and image results, recovered references and timestamps when applicable, output information, and the final processing status.

A successful run ends with:

```text
Status: SUCCESS
```

Do not treat the conversion as successfully completed until the program reaches `Status: SUCCESS`.

If the program reports an error or does not reach `Status: SUCCESS`, investigate the reported problem before deleting, moving, or replacing an existing permanent archive or other source files.

The console report is also useful when comparing repeated runs or validating the same conversation on another supported platform because it provides a consistent summary of what the program processed.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Output Locations

ChatGPT Conversation Bridge keeps the program, permanent source archives, generated DOCX files, and documentation separated within the project folder.

After the program has been run successfully with a valid ChatGPT shared-conversation URL, the project layout is:

```text
chatgpt-conversation-bridge/
├── archive/
│   └── CONVERSATION_NAME.zip
├── documentation/
│   ├── chatgpt-continuing-a-chat.md
│   ├── chatgpt-conversation-bridge-documentation.md
│   └── chatgpt-conversation-continuation-instructions.md
├── docx/
│   └── CONVERSATION_NAME.docx
├── chatgpt_conversation_bridge.py
├── LICENSE
└── README.md
```

The generated DOCX is stored under:

```text
docx/CONVERSATION_NAME.docx
```

The permanent archive is stored under:

```text
archive/CONVERSATION_NAME.zip
```

The `archive` and `docx` folders are created automatically when needed.

The DOCX is the practical document for reading, reviewing, archiving, and supplying prior conversation context to a continuation chat.

The permanent ZIP preserves the archived shared-conversation data and supported uploaded images that were successfully retrieved. It can later be used directly by ChatGPT Conversation Bridge to regenerate the DOCX without retrieving the shared conversation again.

Keep the permanent ZIP if you want to retain the ability to regenerate the DOCX from the archived conversation data.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Permanent Archive and DOCX Regeneration

After a successful run from a ChatGPT shared-conversation URL, ChatGPT Conversation Bridge retains the conversation data and successfully retrieved supported uploaded images in a permanent ZIP under the `archive` folder.

The permanent archive allows the DOCX to be regenerated later without retrieving the shared conversation again. Internet access is not required for DOCX regeneration from an existing archive.

To regenerate the DOCX, run ChatGPT Conversation Bridge with the existing archive ZIP as its command-line argument.

For example, when the archive is already in the program's `archive` folder, you can supply only its filename.

**Windows**

```text
python chatgpt_conversation_bridge.py "Ear Crevice in Dog.zip"
```

**Linux/macOS**

```text
python3 chatgpt_conversation_bridge.py "Ear Crevice in Dog.zip"
```

ChatGPT Conversation Bridge resolves the filename through its `archive` folder, reads `conversation.json` and the archived supported uploaded images directly from the ZIP, and generates the DOCX under `docx/`.

You can also specify the archive folder explicitly.

**Windows**

```text
python chatgpt_conversation_bridge.py "archive\Ear Crevice in Dog.zip"
```

or:

```text
python chatgpt_conversation_bridge.py ".\archive\Ear Crevice in Dog.zip"
```

**Linux/macOS**

```text
python3 chatgpt_conversation_bridge.py "archive/Ear Crevice in Dog.zip"
```

or:

```text
python3 chatgpt_conversation_bridge.py "./archive/Ear Crevice in Dog.zip"
```

An explicit current-relative ZIP path and an absolute ZIP path are also supported, as described under [Supported Input Types](#supported-input-types).

The permanent archive does not need to be extracted manually. Leave the ZIP intact and allow ChatGPT Conversation Bridge to read it directly.

Regenerating the DOCX does not modify the permanent archive. The ZIP remains the retained source, while the DOCX can be recreated when needed.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Archive Collision Protection

ChatGPT Conversation Bridge protects an existing permanent archive from being silently overwritten when processing a ChatGPT shared-conversation URL.

When a shared conversation is processed, the program determines the conversation name and the corresponding permanent archive filename under `archive/`.

If that permanent archive already exists, ChatGPT Conversation Bridge stops rather than replacing it automatically.

This protection prevents an established archive from being unintentionally replaced by another retrieval of the shared conversation, including a retrieval made after the shared conversation has changed.

If an archive collision is reported, preserve the existing ZIP until you have deliberately determined how you want to handle the existing archive and the newly requested conversation state.

Do not delete or overwrite an existing permanent archive merely to bypass the collision check. If you intentionally want to create a new archive for an updated conversation, first decide how the existing archive should be preserved or renamed.

Archive collision protection applies when creating a permanent archive from a shared-conversation URL. Regenerating a DOCX from an existing archive ZIP reads that archive as the source and does not replace it.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Safety and File-Handling Rules

ChatGPT Conversation Bridge is designed to keep the permanent archive separate from the generated DOCX and to avoid silently replacing an existing archive.

Follow these file-handling rules:

- Do not treat a partial or failed run as a successfully completed archive.
- Do not treat a conversion as successfully completed until the program reaches `Status: SUCCESS`.
- Preserve permanent ZIP archives under `archive/` if you may need to regenerate their DOCX files later.
- Do not manually extract a permanent ZIP merely to regenerate its DOCX. ChatGPT Conversation Bridge reads the archive directly.
- Do not manually modify `conversation.json` or files under `uploads/` inside a permanent archive unless you deliberately intend to alter the archived source.
- If an archive collision is reported, preserve the existing archive until you have deliberately determined how the existing archive and the newly requested conversation state should be handled.
- Do not delete or overwrite an existing permanent archive merely to bypass archive collision protection.
- If a run fails or does not reach `Status: SUCCESS`, investigate the reported problem before deleting, moving, or replacing an existing permanent archive or other source files.
- A generated DOCX can be recreated from its permanent archive, but the permanent archive contains the retained source data used for regeneration. Do not assume that keeping only the DOCX provides the same recovery capability.

These rules help protect the retained conversation data and supported uploaded images while keeping permanent source archives separate from generated DOCX files.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## DOCX Content and Formatting

ChatGPT Conversation Bridge converts the ChatGPT conversation into a portable DOCX designed to preserve both the readable conversation and important structural information from the original chat.

The generated DOCX preserves, when available:

- User and assistant messages in chronological order.
- **YOU** and **CHATGPT** message identification.
- Message timestamps.
- Paragraphs, headings, lists, tables, code blocks, inline code, bold text, italic text, and other supported conversation formatting.
- Public reference URLs recovered from the conversation data.
- Supported uploaded images that were successfully retrieved and preserved.
- Clear placeholders for supported image or attachment records that cannot be recovered.

User-authored soft line breaks are preserved so that intentional line structure in **YOU** messages is not unnecessarily collapsed.

For visual separation, message bodies use subtle background shading:

- **YOU** message body: `#FFFBF2`
- **CHATGPT** message body: `#F7FFF5`

Code blocks retain monospaced formatting and gray shading so that they remain visually distinct from ordinary conversation text.

The generated DOCX is intended primarily as a faithful readable record and continuation document rather than as an editable recreation of the ChatGPT web interface.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Images and Attachments

ChatGPT Conversation Bridge attempts to preserve uploaded images and attachment information identified in the shared-conversation data.

For supported uploaded raster images, the program uses the image's stable file identifier to request access to the uploaded file and retrieve its original bytes. Successfully retrieved supported images are preserved in the permanent archive and embedded in the generated DOCX.

The retrieved file bytes are treated as authoritative when determining the raster image type. This avoids relying only on a filename or metadata label when the file itself identifies the actual image format.

When an expected image cannot be recovered, ChatGPT Conversation Bridge writes a filename-specific placeholder:

```text
[Image unavailable: filename]
```

When an attachment record cannot be recovered as an embedded image or other supported content, the program writes:

```text
[Attachment unavailable: filename]
```

These placeholders intentionally identify the unavailable item without inserting a giant data URL, a `None` value, or an ambiguous generic upload message.

Attachment and image records are reconciled so that a failed image does not incorrectly consume the position of a later recoverable image, and metadata-only attachments do not take a rendered image slot.

Successfully retrieved supported uploaded images are stored under `uploads/` inside the permanent archive. When a DOCX is later regenerated from that archive, ChatGPT Conversation Bridge uses those archived image bytes directly rather than retrieving the images again.

An upload that is represented in the conversation data is not necessarily available as an embeddable image. Unsupported image formats and non-image attachments can therefore be represented by unavailable-item placeholders rather than embedded content.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Reference URL Handling

ChatGPT Conversation Bridge preserves public reference URLs that can be recovered from the conversation data.

Reference URLs are written into the DOCX as readable plain-text URLs rather than being converted into active hyperlinks.

This keeps the visible destination available in the archival document and avoids changing the displayed URL through generated hyperlink behavior.

Only reference information that is available in the retrieved or archived conversation data can be preserved. ChatGPT Conversation Bridge does not independently reconstruct missing external references that are not present in the conversation data.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Opening the DOCX in Microsoft Word

The generated DOCX can be opened directly in Microsoft Word.

The production document uses Word's recommended write-protection setting to encourage open-view-only behavior while still allowing intentional editing.

When the document opens in Word, review it normally as a conversation record. If Word presents the document in a protected or view-oriented state, this is expected behavior for the generated DOCX.

During final Windows validation, the document could still be intentionally placed into an editable state using Word's available editing control.

The read-only/open-view-only behavior is intended to reduce accidental changes to the archival conversation. It is not intended to prevent a user who deliberately chooses to edit the document from doing so.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Opening the DOCX in LibreOffice Writer

The generated DOCX can also be opened in LibreOffice Writer.

During final validation on `Ubuntu 24.04 LTS` and `macOS`, LibreOffice Writer opened the generated document in a read-only state as intended.

The document remains fully available for reading and review. If intentional editing is required, LibreOffice Writer's **Edit Mode** can be used to switch the document into an editable state.

The exact appearance of application controls can vary by LibreOffice version and operating system, but the validated behavior is that the document opens for safe viewing while still allowing deliberate editing.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Read-Only / Open-View-Only Behavior

ChatGPT Conversation Bridge marks the generated DOCX with recommended write protection.

The DOCX contains the Word setting:

```xml
<w:writeProtection w:recommended="true"/>
```

This is intentionally advisory rather than enforced document protection. The generated DOCX is not password-protected, and ChatGPT Conversation Bridge does not use an enforced editing restriction to prevent the user from changing the document.

The purpose is to make accidental editing less likely while preserving the user's ability to intentionally edit the document when necessary.

Final native-application validation confirmed the intended behavior:

- Microsoft Word on `Windows 11` opens the document in a view-oriented state while still providing a way to enable editing.
- LibreOffice Writer on `Ubuntu 24.04 LTS` opens the document read-only and allows intentional editing through **Edit Mode**.
- LibreOffice Writer on `macOS` opens the document read-only and allows intentional editing through **Edit Mode**.

This behavior was validated with the final production output rather than inferred only from the DOCX package setting.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Intentionally Editing the DOCX

The generated DOCX is intended to open in a view-oriented or read-only state to reduce accidental modification of the archived conversation, but intentional editing remains available.

**Microsoft Word on Windows**

Use Word's available editing control to switch from the view-oriented state into editing when you deliberately want to modify the document.

**LibreOffice Writer on Ubuntu or macOS**

Use **Edit Mode** to switch the document from read-only viewing into an editable state.

Once editing has been enabled, changes can be made and saved like changes to another DOCX.

If the goal is to preserve an authoritative archival copy of the conversation, keep the original generated DOCX unchanged and make intentional edits to a separate copy.

The permanent ZIP under `archive/` remains the retained source archive and can be used by ChatGPT Conversation Bridge to regenerate the DOCX later.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Tested Platforms and Python Versions

The final production version of ChatGPT Conversation Bridge Version 2.0.0 was originally runtime-tested on the following operating systems and Python versions:

| Operating System | Python Version | Runtime Result |
| --- | --- | --- |
| `Windows 11` | `Python 3.12.3` | PASS |
| `Ubuntu 24.04 LTS` | `Python 3.12.3` | PASS |

Additional fresh-install testing performed after the Version 2.0.0 release validated the current installation procedure on the following systems:

| Operating System | Python Version | Runtime Result |
| --- | --- | --- |
| `Windows 11` | `Python 3.14` | PASS |
| `macOS` (Intel) | `Python 3.14.7` | PASS |
| `macOS` (Apple Silicon) | `Python 3.14.7` | PASS |

The program requires `Python 3.10` or later.

ChatGPT Conversation Bridge Version 2.0.0 requires the third-party Python `curl_cffi` and `tzdata` packages. Version `0.16.3` of `curl_cffi` was used during final Version 2.0.0 cross-platform validation.

Internet access is required when creating a new archive from a ChatGPT shared-conversation URL. Internet access is not required when regenerating a DOCX from an existing permanent archive ZIP.

Version 2.0.0 does not require a particular web browser for conversation extraction because the program retrieves the public shared-conversation data directly from the supplied ChatGPT shared URL.

Microsoft Word was used to validate the generated DOCX on `Windows 11`. LibreOffice Writer was used for native DOCX validation on `Ubuntu 24.04 LTS` and `macOS`.

These results document the environments actually tested for this release. They should not be interpreted as a guarantee for every operating-system version, Python version, dependency version, office-suite version, hardware configuration, or future ChatGPT shared-conversation implementation.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Production Testing and Validation

ChatGPT Conversation Bridge Version 2.0.0 was tested with multiple ChatGPT conversations representing a wide range of conversation sizes, message counts, uploads, references, and document complexity.

The Windows validation corpus included large and small conversations, conversations containing many uploaded images, conversations containing non-image attachments, unsupported image-like uploads, and conversations containing no recoverable images.

Across the Windows test corpus, 181 uploaded images were identified. Of these, 178 were supported raster images that ChatGPT Conversation Bridge handled and preserved successfully. The remaining three were SVG uploads, which are not supported for embedding by Version 2.0.0.

Testing also covered permanent-archive creation and subsequent DOCX regeneration directly from the archive ZIP without retrieving the shared conversation again.

Cross-platform runtime validation was performed on:

- `Windows 11`
- `Ubuntu 24.04 LTS`
- `macOS`

The final production source used for cross-platform validation was identical on all three operating systems. Its SHA-256 fingerprint was:

```text
3BA0F9C0ACE6CFC6251A6ED8F1A2938950709CC53CD9611DC81DB96AA94314E6
```

The final production source compiled successfully on Windows and completed the required runtime validation. The same source fingerprint was verified on Ubuntu and macOS, and the cross-platform runtime tests completed successfully.

Validation included both the normal shared-conversation workflow and regeneration from an existing permanent archive. Archive-path handling was tested using filename-only and explicit relative archive paths, with platform-appropriate path separators.

Final validation also covered:

- Conversation-chain validation.
- Creation of the permanent archive containing `conversation.json` and retrieved supported uploaded images.
- DOCX generation from a ChatGPT shared-conversation URL.
- DOCX regeneration directly from a permanent archive ZIP.
- Preservation of supported uploaded raster images.
- Filename-specific unavailable-image and unavailable-attachment placeholders.
- Preservation of intentional soft line breaks in **YOU** messages.
- Literal-asterisk handling without damaging intended Markdown emphasis.
- Distinct subtle background shading for **YOU** and **CHATGPT** message bodies.
- Public reference URL preservation.
- Message timestamp preservation when available.
- Recommended DOCX write-protection behavior.
- Native DOCX viewing and intentional editing behavior in Microsoft Word and LibreOffice Writer.

The final production source completed the validated cross-platform tests without requiring platform-specific changes to the ChatGPT Conversation Bridge source itself.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Known Limitations

ChatGPT Conversation Bridge works from the information and files made available through the ChatGPT shared-conversation interface or from information and files already preserved in a permanent ChatGPT Conversation Bridge archive.

Known limitations include:

- The shared-conversation data can lag behind the current live ChatGPT conversation. Recently added messages may therefore not immediately appear in the data available through the shared link. Before creating the permanent archive, make sure the shared version reflects the conversation state you intend to preserve.
- ChatGPT Conversation Bridge can preserve only information represented in the shared-conversation data. Content that is not present in that data cannot be independently reconstructed by the program.
- An upload represented in the conversation data is not necessarily available for retrieval through the shared conversation.
- Version 2.0.0 embeds supported raster image formats. Unsupported image formats, including SVG, are not embedded as supported raster images.
- Non-image attachments are not embedded in the DOCX as supported raster images. When appropriate, unavailable image or attachment records are represented by `[Image unavailable: filename]` or `[Attachment unavailable: filename]`.
- If an uploaded image cannot be retrieved when the permanent archive is created, the program cannot later recover that image from the archive unless its bytes were successfully preserved there during archive creation.
- Public reference URLs can be preserved only when the necessary reference information is present in the conversation data.
- The normal shared-URL workflow depends on the ChatGPT public shared-conversation interface used by Version 2.0.0. A future change to that interface or its returned data could require a corresponding update to ChatGPT Conversation Bridge.
- The normal shared-URL workflow requires Internet access. DOCX regeneration from an existing permanent archive does not.
- The DOCX preserves supported conversation structure and formatting, but it is not intended to reproduce the ChatGPT web interface pixel-for-pixel.
- Recommended write protection is advisory. It reduces accidental editing but does not prevent deliberate modification of the DOCX.
- Validation documents the tested operating systems, Python versions, dependency versions, and office applications. Untested environments or future dependency versions may behave differently.

Preserve the permanent ZIP under `archive/` so that the DOCX can be regenerated later from the archived conversation data and supported uploaded images without retrieving the shared conversation again.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Troubleshooting

If ChatGPT Conversation Bridge does not produce the expected result, start with the command-line source argument and the final console status.

**The shared-conversation URL is not accepted**

Confirm that the command contains the complete ChatGPT shared-conversation URL and that it has the expected general form:

```text
https://chatgpt.com/share/<UUID>
```

Make sure the URL was not accidentally truncated or altered when it was pasted into the command.

**The shared conversation does not contain the latest messages**

The data available through a ChatGPT shared link can sometimes lag behind the current live conversation.

Verify that the shared version reflects the conversation state you intend to preserve. If recently added messages are not yet represented in the shared data, do not treat the resulting archive as a complete copy of the current live conversation.

**Expected uploaded images are unavailable**

An upload represented in the conversation data is not necessarily available for retrieval through the shared conversation.

ChatGPT Conversation Bridge attempts to retrieve supported uploaded images when the permanent archive is created. If an expected image cannot be retrieved, the DOCX can contain:

```text
[Image unavailable: filename]
```

Unsupported image formats and non-image attachments are not embedded as supported raster images. An unavailable attachment can be represented as:

```text
[Attachment unavailable: filename]
```

Review the console report to compare the number of uploads, embedded images, and unavailable upload references.

**The program does not reach `Status: SUCCESS`**

Do not treat the conversion as successfully completed.

Review the reported error before deleting, moving, renaming, or replacing an existing permanent archive or other source files. Do not treat a partial or failed result as a valid permanent archive.

**An archive collision is reported**

A permanent archive with the conversation's archive name already exists under `archive/`. ChatGPT Conversation Bridge stops instead of silently replacing that archive.

Preserve the existing ZIP until you have deliberately determined how the existing archive and the newly requested conversation state should be handled.

Do not simply delete or overwrite the existing archive to bypass the collision protection.

**The program cannot find an archive ZIP**

Check the archive argument supplied on the command line.

If only the ZIP filename is supplied, ChatGPT Conversation Bridge looks for that ZIP under its `archive` folder.

For example:

**Windows**

```text
python chatgpt_conversation_bridge.py "Ear Crevice in Dog.zip"
```

**Linux/macOS**

```text
python3 chatgpt_conversation_bridge.py "Ear Crevice in Dog.zip"
```

Explicit relative and absolute archive paths are also supported as described under [Supported Input Types](#supported-input-types).

**The DOCX opens read-only or in a view-oriented state**

This is expected. The generated DOCX uses recommended write protection to reduce accidental editing.

In Microsoft Word, use the available editing control when intentional editing is required. In LibreOffice Writer, use **Edit Mode**.

**The DOCX needs to be regenerated**

Run ChatGPT Conversation Bridge with the existing permanent archive ZIP as the source. Internet access is not required for archive-based regeneration.

For example:

**Windows**

```text
python chatgpt_conversation_bridge.py "Ear Crevice in Dog.zip"
```

**Linux/macOS**

```text
python3 chatgpt_conversation_bridge.py "Ear Crevice in Dog.zip"
```

The permanent ZIP does not need to be extracted manually. ChatGPT Conversation Bridge reads the archive directly and creates the regenerated DOCX under `docx/`.

**A required Python package is not installed**

ChatGPT Conversation Bridge Version 2.0.0 requires the Python `curl_cffi` and `tzdata` packages.

If Python reports that `curl_cffi` or `tzdata` cannot be imported, install the required packages using the platform-specific installation instructions in this documentation, then run ChatGPT Conversation Bridge again.

**A Homebrew Python installation on macOS reports an installation or version problem**

Homebrew is not required to run ChatGPT Conversation Bridge. The Python.org installation described under [Installing Python](#installing-python) is the normal macOS installation path used by this documentation.

If a Homebrew-installed Python reports an `externally-managed-environment` error when installing the required packages, the following command was successfully tested:

```text
python3 -m pip install --break-system-packages curl_cffi tzdata
```

If `python3 --version` still reports an older Python version immediately after installing Python with Homebrew, refresh the shell's command lookup and check the version again:

```text
rehash
python3 --version
```

**A problem remains after checking the items above**

Preserve the permanent archive and the console output from the failed run. The console report contains useful processing counts and error information that can help identify whether the problem occurred while retrieving the shared conversation, retrieving an upload, creating the archive, or generating the DOCX.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)