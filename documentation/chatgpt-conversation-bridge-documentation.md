# ChatGPT Conversation Bridge Documentation

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

## Document Navigator

### Overview & Setup

- [Purpose](#purpose)
- [Requirements](#requirements)
- [Installing Python](#installing-python)
- [Browser Compatibility](#browser-compatibility)
- [Initial Folder Setup](#initial-folder-setup)

### Archiving a Conversation

- [Normal Workflow](#normal-workflow)
- [Create the ChatGPT Shared Link](#create-the-chatgpt-shared-link)
- [Standard Loading Method](#standard-loading-method)
- [Verified Loading Method](#verified-loading-method)
- [Save the Loaded Conversation with Chrome](#save-the-loaded-conversation-with-chrome)
- [Verify the Saved Browser Capture](#verify-the-saved-browser-capture)

### Running Conversation Bridge

- [Run ChatGPT Conversation Bridge](#run-chatgpt-conversation-bridge)
- [Source Lookup Order](#source-lookup-order)
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

For chat continuation, save the completed or near-limit conversation in Google Chrome using **Save As** with the file type explicitly set to **Web Page, Complete**, run ChatGPT Conversation Bridge, and upload the generated DOCX into a new ChatGPT conversation. The DOCX gives the new chat a detailed record of the earlier discussion, including user and assistant messages, timestamps when available, formatting, public reference URLs, and locally preserved supported images.

The DOCX also serves an important archival purpose. It creates a portable, human-readable record of the conversation that can be opened independently of ChatGPT in Microsoft Word, LibreOffice Writer, or another compatible DOCX reader.

The permanent ZIP created by ChatGPT Conversation Bridge preserves the browser-saved source needed to regenerate the DOCX, while the DOCX is the practical document for reading, reviewing, archiving, and supplying prior conversation context to a continuation chat.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Browser Compatibility

Google Chrome is the browser validated for the ChatGPT Conversation Bridge workflow. In testing on `Windows 11`, `Ubuntu 24.04 LTS`, and `macOS Monterey`, Chrome successfully preserved the uploaded-image resources needed by the program when a ChatGPT shared conversation was saved using **Web Page, Complete**.

For this workflow, use Google Chrome, right-click the shared-conversation page, select **Save As**, and explicitly set the file-type selection to **Web Page, Complete**.

Do not rely on Microsoft Edge, Opera, or Firefox when uploaded images must be preserved. During testing, Edge and Opera did not save the tested uploaded-image resources, and earlier Firefox testing likewise failed to preserve the tested uploaded PNG resources. Conversation text may still be processable, but missing saved-image resources prevent those images from being embedded in the generated DOCX.

These results describe the browser versions, operating systems, and manual Save As workflow tested during development of ChatGPT Conversation Bridge. They are not a claim about every browser version, operating-system release, Linux distribution, or future ChatGPT page format.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Requirements

ChatGPT Conversation Bridge requires:

- A desktop computer running Windows, Linux, or macOS. Runtime validation for this release was performed specifically on `Windows 11`, `Ubuntu 24.04 LTS`, and `macOS Monterey`.
- `Python 3.10` or later. Runtime validation used `Python 3.12.3` on `Windows 11` and `Ubuntu 24.04 LTS`, and `Python 3.14.7` on `macOS Monterey`.
- No third-party Python packages. ChatGPT Conversation Bridge uses the Python standard library.
- Google Chrome to save the ChatGPT shared-conversation page. For the tested workflow, explicitly select **Web Page, Complete** in Chrome's Save As dialog.
- Microsoft Word, LibreOffice Writer, or another compatible DOCX reader to view the generated document.

The exact operating systems and Python versions that were tested are documented separately under [Tested Platforms and Python Versions](#tested-platforms-and-python-versions).

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Installing Python

ChatGPT Conversation Bridge requires `Python 3.10` or later. No third-party Python packages are required.

If `Python 3.10` or later is already installed, you do not need to install Python again.

### Windows

Open **Command Prompt** and check the installed Python version:

    python --version

If `Python 3.10` or later is reported, no further action is required.

If Python is not installed or the version is older than `Python 3.10`, download and install the current version of Python from the official Python website:

    https://www.python.org/downloads/windows/

After installation, open a new **Command Prompt** and verify the version:

    python --version

Confirm that `Python 3.10` or later is reported.

### Ubuntu / Linux

Open **Terminal** and install Python:

    sudo apt install -y python3-full

Then verify the installed version:

    python3 --version

Confirm that `Python 3.10` or later is reported.

The installation command above is for Ubuntu and other Linux distributions that use the `apt` package manager. Other Linux distributions may use a different package manager.

### macOS

Open **Terminal** and check the installed Python version:

    python3 --version

If `Python 3.10` or later is reported, no further action is required.

If Python is not installed or the version is older than `Python 3.10`, download the current macOS installer from the official Python website:

    https://www.python.org/downloads/macos/

Open the downloaded installer and follow the installation prompts.

After installation, open a new **Terminal** and verify the version:

    python3 --version

Confirm that `Python 3.10` or later is reported.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Initial Folder Setup

Create one working folder for ChatGPT Conversation Bridge. The tested locations are:

**Windows**

    D:\chatgpt-conversation-bridge

**Linux**

    /home/<username>/chatgpt-conversation-bridge

**macOS**

    /Users/<username>/chatgpt-conversation-bridge

Place the ChatGPT Conversation Bridge script directly in the working folder.

**Windows**

    D:\chatgpt-conversation-bridge\chatgpt_conversation_bridge.py

**Linux/macOS**

    ~/chatgpt-conversation-bridge/chatgpt_conversation_bridge.py

The project uses the following folder structure:

    chatgpt-conversation-bridge/
    ├── archive/
    ├── documentation/
    ├── docx/
    └── chatgpt_conversation_bridge.py

The `archive` and `docx` folders do not need to be created manually. ChatGPT Conversation Bridge creates them automatically when needed.

The `documentation` folder contains the project documentation.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Normal Workflow

ChatGPT Conversation Bridge provides two ways to prepare a shared ChatGPT conversation before saving it:

- **Standard Loading Method** — does not require Chrome Developer Tools and provides the simpler procedure.
- **Verified Loading Method** — uses Chrome Developer Tools to observe uploaded-image requests and provides additional confirmation that available image resources have been retrieved.

Both methods lead to the same browser-save procedure using **Web Page, Complete**.

The overall workflow is:

1. Create the ChatGPT shared link for the conversation you want to preserve.
2. Open the shared conversation in Google Chrome using either the Standard Loading Method or the Verified Loading Method.
3. Allow the conversation and its available resources to finish loading.
4. Save the loaded shared conversation using **Web Page, Complete**.
5. Verify that Chrome created both the conversation HTML/HTM file and its matching `_files` companion folder.
6. Run ChatGPT Conversation Bridge.
7. Confirm that the program finishes with `Status: SUCCESS`.
8. Review the generated DOCX.
9. Preserve the permanent ZIP under `archive/` for future DOCX regeneration.

The following sections describe each part of this workflow in detail.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Create the ChatGPT Shared Link

Begin in the original ChatGPT conversation that you want to preserve.

1. Click **Share** in ChatGPT.
2. Create or update the shared link for the conversation.
3. Copy the shared-conversation URL.
4. Open a new blank Google Chrome tab.
5. Paste the shared-conversation URL into the address bar.

At this point, choose either the **Standard Loading Method** or the **Verified Loading Method** described in the following sections.

Do not save the conversation from the original ChatGPT conversation page. ChatGPT Conversation Bridge is designed to process the browser capture made from the loaded shared-conversation page.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Standard Loading Method

The Standard Loading Method is the simpler way to prepare the shared conversation for saving. It does not require Chrome Developer Tools.

1. Open the shared-conversation URL in Google Chrome.
2. Allow the conversation to load.
3. Scroll through the conversation from top to bottom so that the conversation content and available resources have an opportunity to load.
4. For a large conversation, use **Ctrl+Home** to return to the top and **Ctrl+End** to move to the bottom as needed while checking that the conversation has loaded.
5. After the conversation appears to have finished loading, wait at least 30 additional seconds before saving it.
6. Save the loaded conversation using the procedure in [Save the Loaded Conversation with Chrome](#save-the-loaded-conversation-with-chrome).

If the saved browser capture appears incomplete or expected images are missing, repeat the loading and saving process. You can also use the [Verified Loading Method](#verified-loading-method) for additional confirmation that available image resources have been retrieved before saving.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Verified Loading Method

The Verified Loading Method uses Chrome Developer Tools to provide additional confirmation that available uploaded-image resources have been requested while the shared conversation is loading.

1. Open a new blank Google Chrome tab.
2. Open Chrome Developer Tools by pressing **F12**.
3. Select the **Network** panel.
4. Select **Fetch/XHR** in the Network panel.
5. With Developer Tools still open, paste the shared-conversation URL into the Chrome address bar and load the page.
6. Allow the conversation to load completely. For a large conversation, work through the conversation from top to bottom so that its content and available resources have an opportunity to load.
7. Watch the Network panel while the conversation loads. Uploaded-image requests can appear with names beginning with `file_`.
8. Allow network activity and conversation loading to settle before saving the page.
9. Save the loaded conversation using the procedure in [Save the Loaded Conversation with Chrome](#save-the-loaded-conversation-with-chrome).

Use **Fetch/XHR** rather than relying on the **Img** filter when checking uploaded-image activity. During testing, the Img filter was not a reliable way to identify the uploaded-image requests used by this workflow.

The Network panel is a verification aid. Its purpose is to provide additional evidence that available resources were requested before the shared conversation is saved; it does not change the required Chrome **Web Page, Complete** save procedure.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Save the Loaded Conversation with Chrome

After the shared conversation and its available resources have finished loading, save the page from Google Chrome.

1. Right-click the loaded shared-conversation page.
2. Select **Save As**.
3. In Chrome's Save As dialog, explicitly select **Web Page, Complete** as the file type. Do not rely on the default selection.
4. Save the page directly in the ChatGPT Conversation Bridge working folder.
5. Wait for Chrome to finish saving the page before closing the tab.

Chrome should create two items in the working folder:

- The conversation `.html` or `.htm` file.
- A matching companion folder whose name ends in `_files`.

The HTML/HTM file and its matching `_files` folder together form the browser capture used by ChatGPT Conversation Bridge. Do not move, rename, or separate one from the other before running the program.

The Save As default can vary. During Ubuntu testing, Chrome was observed to default to different save formats depending on the destination folder. Always explicitly select **Web Page, Complete** rather than assuming Chrome has already selected it.

After the save is complete, close the shared-conversation tab. During testing, continuing to work from the shared-conversation page after saving created a risk of accidentally starting or creating a duplicate conversation.

Before running ChatGPT Conversation Bridge, verify the saved browser capture as described in [Verify the Saved Browser Capture](#verify-the-saved-browser-capture).

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Verify the Saved Browser Capture

Before running ChatGPT Conversation Bridge, verify that Chrome created a complete browser capture in the working folder.

Confirm that both of these items exist:

- The conversation `.html` or `.htm` file.
- The matching companion folder whose name ends in `_files`.

The HTML/HTM file and matching `_files` folder must remain together in the working folder. ChatGPT Conversation Bridge uses the saved page and its companion resources when creating the DOCX and permanent archive.

If the `_files` folder is missing, do not run ChatGPT Conversation Bridge. Return to the shared conversation in Chrome and repeat the loading and **Web Page, Complete** save procedure.

If the `_files` folder exists but expected uploaded images are missing from the saved browser capture, return to the loading procedure before running the program. Repeat the [Standard Loading Method](#standard-loading-method), or use the [Verified Loading Method](#verified-loading-method) for additional confirmation that available image resources have been retrieved.

Once the HTML/HTM file and matching `_files` folder have been verified, the browser capture is ready for ChatGPT Conversation Bridge.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Run ChatGPT Conversation Bridge

Open a terminal or command prompt in the ChatGPT Conversation Bridge working folder.

Run the program with the ChatGPT conversation name as its single command-line argument.

**Windows**

    python chatgpt_conversation_bridge.py "Find Image Upload Links"

**Linux/macOS**

    python3 chatgpt_conversation_bridge.py "Find Image Upload Links"

Replace `Find Image Upload Links` with the exact conversation name you want ChatGPT Conversation Bridge to process.

The program accepts exactly one conversation name. You normally do not need to include `.html` or `.htm`; ChatGPT Conversation Bridge looks for the supported source automatically. The extension can be supplied as an optional source hint when needed.

On a normal first run, the browser-saved HTML/HTM file and matching `_files` folder must already be present in the working folder. On a later regeneration run, ChatGPT Conversation Bridge can instead use the permanent archive when the live browser capture is no longer present.

Allow the program to finish. A successful run ends with:

    Status: SUCCESS

The following sections explain the source lookup order, first-run processing, console report, and output locations in detail.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Source Lookup Order

When ChatGPT Conversation Bridge is run with a conversation name, it looks for a usable source in this order:

1. `CONVERSATION_NAME.html`
2. `CONVERSATION_NAME.htm`
3. `archive/CONVERSATION_NAME.zip`

This lookup order allows the same command to be used for both a new browser capture and a later DOCX regeneration from the permanent archive.

If a matching live HTML/HTM browser capture is present in the working folder, ChatGPT Conversation Bridge uses that source before looking for the permanent ZIP.

If neither a matching live browser capture nor a matching permanent archive can be found, the program cannot process that conversation.

The optional `.html` or `.htm` extension supplied with the conversation name acts as a source hint; it does not change the overall purpose of the lookup process.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## What Happens on the First Run

On the first successful run for a newly saved conversation, ChatGPT Conversation Bridge processes the live browser capture from the working folder.

The program:

1. Reads and validates the saved HTML/HTM conversation and its companion resources.
2. Extracts the conversation content and creates the DOCX.
3. Creates a permanent ZIP archive containing the browser-saved source needed for future DOCX regeneration.
4. Reopens and tests the permanent archive to verify that it contains the expected saved browser capture.
5. Removes the temporary live browser capture only after the permanent archive has been successfully created and verified.

This ordering protects the original browser capture during processing. If the run fails before the permanent archive has been successfully created and verified, the original live capture is intentionally left in place so that it is not lost.

A partial or failed archive is not treated as a valid replacement for the original browser capture.

After a successful first run, the permanent ZIP becomes the retained source from which the DOCX can later be regenerated.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Console Report and Status: SUCCESS

While ChatGPT Conversation Bridge runs, it prints a console report describing the conversation it processed and the results of the conversion.

The report provides processing information such as the source being used, conversation/message statistics, recovered uploads and references when applicable, output information, and the final processing status.

A successful run ends with:

    Status: SUCCESS

Do not treat the conversion as successfully completed until the program reaches `Status: SUCCESS`.

If the program reports an error or does not reach `Status: SUCCESS`, preserve the existing browser capture and investigate the reported problem before deleting, moving, or replacing source files.

The console report is also useful when comparing repeated runs or validating the same conversation on another supported platform because it provides a consistent summary of what the program processed.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Output Locations

ChatGPT Conversation Bridge keeps the working program, permanent source archive, generated DOCX, and documentation separated within the project folder.

The normal project layout is:

    chatgpt-conversation-bridge/
    ├── archive/
    │   └── CONVERSATION_NAME.zip
    ├── documentation/
    ├── docx/
    │   └── CONVERSATION_NAME.docx
    └── chatgpt_conversation_bridge.py

The generated DOCX is stored under:

    docx/CONVERSATION_NAME.docx

The permanent browser-capture archive is stored under:

    archive/CONVERSATION_NAME.zip

The `archive` and `docx` folders are created automatically when needed.

The DOCX is the practical document for reading, reviewing, archiving, and supplying prior conversation context to a continuation chat. The permanent ZIP retains the browser-saved source so that the DOCX can be regenerated later without repeating the Chrome Save As procedure.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Permanent Archive and DOCX Regeneration

After a successful first run, ChatGPT Conversation Bridge retains the browser-saved source as a permanent ZIP under the `archive` folder.

The permanent archive allows the DOCX to be regenerated later without saving the ChatGPT shared conversation from Chrome again.

To regenerate the DOCX, run the same command used for the original conversion.

**Windows**

    python chatgpt_conversation_bridge.py "Find Image Upload Links"

**Linux/macOS**

    python3 chatgpt_conversation_bridge.py "Find Image Upload Links"

Replace `Find Image Upload Links` with the conversation name you want to regenerate.

When the live HTML/HTM browser capture is no longer present, ChatGPT Conversation Bridge finds the matching permanent ZIP under `archive/`, reads the saved browser capture from that archive, and creates the DOCX again under `docx/`.

Regeneration does not require extracting the permanent ZIP manually. Leave the archive intact and allow ChatGPT Conversation Bridge to use it directly.

The permanent archive is therefore the retained source for future regeneration, while the DOCX can be recreated when needed.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Archive Collision Protection

ChatGPT Conversation Bridge protects an existing permanent archive from being silently replaced by a different live browser capture with the same conversation name.

If a permanent archive already exists under `archive/` and a new live HTML/HTM browser capture with the same base conversation name is also present, ChatGPT Conversation Bridge stops instead of overwriting the existing archive.

The newly saved live browser capture is left unchanged so that the conflict can be investigated safely.

This protection prevents an established permanent archive from being unintentionally replaced merely because another browser capture uses the same conversation name.

Resolve the naming or source conflict before running ChatGPT Conversation Bridge again. Do not delete or overwrite the existing permanent archive unless you have deliberately determined which source should be retained.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Safety and File-Handling Rules

ChatGPT Conversation Bridge is designed to preserve the source browser capture until a safe permanent archive has been created and verified.

Follow these file-handling rules:

- Keep the saved HTML/HTM file and its matching `_files` folder together before the first run.
- Do not manually delete the live browser capture before ChatGPT Conversation Bridge reports a successful conversion.
- Do not treat a partial or failed ZIP as a valid permanent archive.
- Allow ChatGPT Conversation Bridge to remove the temporary live browser capture only after the permanent archive has been successfully created and verified.
- Preserve the permanent ZIP under `archive/` if you may need to regenerate the DOCX later.
- Do not manually extract the permanent ZIP merely to regenerate the DOCX; ChatGPT Conversation Bridge can read the archive directly.
- Do not overwrite an existing permanent archive when an archive collision is reported. Investigate and resolve the conflicting sources first.
- If a run fails or does not reach `Status: SUCCESS`, preserve the available source files and investigate the reported problem before making file changes.

These rules keep the permanent browser-saved source separate from the generated DOCX and reduce the risk of losing the only usable source for a conversation.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## DOCX Content and Formatting

ChatGPT Conversation Bridge converts the saved ChatGPT conversation into a portable DOCX designed to preserve both the readable conversation and important structural information from the original chat.

The generated DOCX preserves, when available:

- User and assistant messages in chronological order.
- **YOU** and **CHATGPT** message identification.
- Message timestamps.
- Paragraphs, headings, lists, tables, code blocks, inline code, bold text, italic text, and other supported conversation formatting.
- Public reference URLs recovered from the saved conversation.
- Supported uploaded images that are available in the browser-saved resources.
- Clear placeholders for supported image or attachment records that cannot be recovered from the saved browser capture.

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

ChatGPT Conversation Bridge attempts to preserve uploaded images and attachment information that can be recovered from the saved shared-conversation page and its companion resources.

For supported uploaded raster images whose saved bytes are available, the program embeds the recovered image in the DOCX.

The browser-saved file bytes are treated as authoritative when determining the recovered raster image type. This avoids relying only on a filename or metadata label when the saved file itself identifies the actual image format.

When an expected image cannot be recovered, ChatGPT Conversation Bridge writes a filename-specific placeholder:

    [Image unavailable: filename]

When an attachment record cannot be recovered as an embedded image or other supported content, the program writes:

    [Attachment unavailable: filename]

These placeholders intentionally identify the unavailable item without inserting a giant data URL, a `None` value, or an ambiguous generic upload message.

Attachment and image records are reconciled so that a failed image does not incorrectly consume the position of a later recoverable image, and metadata-only attachments do not take a rendered image slot.

Whether an uploaded image can be embedded ultimately depends on the resources preserved by Chrome when the shared conversation was saved. This is why the Chrome loading and **Web Page, Complete** procedure is important when image preservation matters.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Reference URL Handling

ChatGPT Conversation Bridge preserves public reference URLs that can be recovered from the saved conversation.

Reference URLs are written into the DOCX as readable plain-text URLs rather than being converted into active hyperlinks.

This keeps the visible destination available in the archival document and avoids changing the displayed URL through generated hyperlink behavior.

Only reference information that is available in the saved conversation can be preserved. ChatGPT Conversation Bridge does not independently reconstruct missing external references that were not present in the saved source.

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

During final validation on `Ubuntu 24.04 LTS` and `macOS Monterey`, LibreOffice Writer opened the generated document in a read-only state as intended.

The document remains fully available for reading and review. If intentional editing is required, LibreOffice Writer's **Edit Mode** can be used to switch the document into an editable state.

The exact appearance of application controls can vary by LibreOffice version and operating system, but the validated behavior is that the document opens for safe viewing while still allowing deliberate editing.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Read-Only / Open-View-Only Behavior

ChatGPT Conversation Bridge marks the generated DOCX with recommended write protection.

The DOCX contains the Word setting:

    <w:writeProtection w:recommended="true"/>

This is intentionally advisory rather than enforced document protection. The generated DOCX is not password-protected, and ChatGPT Conversation Bridge does not use an enforced editing restriction to prevent the user from changing the document.

The purpose is to make accidental editing less likely while preserving the user's ability to intentionally edit the document when necessary.

Final native-application validation confirmed the intended behavior:

- Microsoft Word on `Windows 11` opens the document in a view-oriented state while still providing a way to enable editing.
- LibreOffice Writer on `Ubuntu 24.04 LTS` opens the document read-only and allows intentional editing through **Edit Mode**.
- LibreOffice Writer on `macOS Monterey` opens the document read-only and allows intentional editing through **Edit Mode**.

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

The permanent ZIP under `archive/` remains the retained browser-source archive and can be used by ChatGPT Conversation Bridge to regenerate the DOCX later.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Tested Platforms and Python Versions

The final production version of ChatGPT Conversation Bridge was runtime-tested on the following operating systems and Python versions:

| Operating System | Python Version | Runtime Result |
| --- | --- | --- |
| `Windows 11` | `Python 3.12.3` | PASS |
| `Ubuntu 24.04 LTS` | `Python 3.12.3` | PASS |
| `macOS Monterey` | `Python 3.14.7` | PASS |

The program requires `Python 3.10` or later and uses only the Python standard library. No third-party Python packages are required.

Google Chrome was used for the validated browser-save workflow on all three operating systems.

Microsoft Word was used to validate the generated DOCX on `Windows 11`. LibreOffice Writer was used for native DOCX validation on `Ubuntu 24.04 LTS` and `macOS Monterey`.

These results document the environments actually tested for this release. They should not be interpreted as a guarantee for every operating-system version, Python version, browser version, office-suite version, Linux distribution, or future ChatGPT page format.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Production Testing and Validation

The final production script was tested with eight saved ChatGPT conversations representing a range of conversation sizes, message counts, uploads, references, and document complexity.

Each of the eight conversations was converted on:

- `Windows 11`
- `Ubuntu 24.04 LTS`
- `macOS Monterey`

This produced 24 final production conversions. All 24 completed successfully.

For every conversation, the reported processing metrics matched across the three operating systems. No cross-platform metric mismatches were found.

The generated DOCX files were also compared across platforms using a standard-library DOCX comparison process that checked meaningful package content including paragraph text, tables, embedded media hashes, relationships, and package members. All eight Windows-to-Ubuntu comparisons and all eight Ubuntu-to-macOS comparisons passed for meaningful document content.

Raw DOCX file hashes were not required to match because package-level differences can exist without changing the meaningful document content.

All 24 generated DOCX files were also visually inspected, and the final documents passed that review.

Final validation additionally covered:

- Preservation of intentional soft line breaks in **YOU** messages.
- Literal-asterisk handling without damaging intended Markdown emphasis.
- Distinct subtle background shading for **YOU** and **CHATGPT** message bodies.
- Final image and attachment recovery behavior and filename-specific unavailable-item placeholders.
- Cross-platform DOCX content equivalence.
- Recommended write-protection behavior.
- Native open-view-only/read-only behavior and intentional editing in Microsoft Word and LibreOffice Writer.

The final production script completed this validation with no required platform-specific code changes.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Known Limitations

ChatGPT Conversation Bridge works from the information and resources preserved in the browser-saved shared-conversation page. It cannot recover content that was never included in that saved source.

Known limitations include:

- Uploaded images or attachments that Chrome did not preserve in the saved browser capture cannot be reconstructed by ChatGPT Conversation Bridge.
- An unavailable image or attachment may therefore appear in the DOCX as `[Image unavailable: filename]` or `[Attachment unavailable: filename]`.
- Browser behavior can change. Google Chrome is the browser validated for this workflow; the tested Edge, Opera, and Firefox workflows did not reliably preserve the uploaded-image resources required by the tested conversations.
- The saved ChatGPT page structure may change in the future. A significant ChatGPT page-format change could require corresponding changes to ChatGPT Conversation Bridge.
- Only information present in the saved conversation can be converted. Missing external references, missing browser resources, or content omitted from the shared page cannot be independently reconstructed by the program.
- The DOCX preserves supported conversation structure and formatting, but it is not intended to reproduce the ChatGPT web interface pixel-for-pixel.
- Recommended write protection is advisory. It reduces accidental editing but does not prevent deliberate modification of the DOCX.
- Validation documents the tested operating systems, Python versions, Chrome workflow, and office applications; untested environments may behave differently.

Preserve the permanent ZIP under `archive/` so that the DOCX can be regenerated from the saved source if needed.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

---

## Troubleshooting

If ChatGPT Conversation Bridge does not produce the expected result, start with the source browser capture and the final console status.

**The program cannot find the conversation source**

Confirm that the working folder contains either:

- The conversation `.html` or `.htm` file with its matching `_files` folder, or
- The matching permanent ZIP under `archive/`.

Also confirm that the conversation name supplied on the command line matches the source name.

**The `_files` companion folder is missing**

Do not run the first conversion from an incomplete browser capture. Return to the shared conversation in Chrome, allow it to load, and save it again with **Web Page, Complete** explicitly selected.

**Expected uploaded images are missing**

Return to the shared conversation and repeat the loading procedure before saving again. For additional confirmation of uploaded-image activity, use the [Verified Loading Method](#verified-loading-method).

If Chrome did not preserve a required image resource in the saved browser capture, ChatGPT Conversation Bridge cannot reconstruct that missing file afterward.

**The program does not reach `Status: SUCCESS`**

Preserve the existing browser capture and permanent archive, if present. Review the reported error before deleting, moving, renaming, or replacing source files.

Do not treat a failed or partial ZIP as a valid permanent archive.

**An archive collision is reported**

An existing permanent archive and a new live browser capture have the same conversation name. ChatGPT Conversation Bridge stops to protect the existing archive.

Determine which source should be retained and resolve the naming or source conflict before running the program again. Do not simply overwrite the existing archive.

**The DOCX opens read-only or in a view-oriented state**

This is expected. The generated DOCX uses recommended write protection to reduce accidental editing.

In Microsoft Word, use the available editing control when intentional editing is required. In LibreOffice Writer, use **Edit Mode**.

**The DOCX needs to be regenerated**

Run ChatGPT Conversation Bridge again with the same conversation name. If the live browser capture is gone, the program can use the permanent ZIP under `archive/` directly. Manual ZIP extraction is not required.

[↑ Back to Document Navigator](#chatgpt-conversation-bridge-documentation)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)