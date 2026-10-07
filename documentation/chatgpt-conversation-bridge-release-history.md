# Project Release, Version History and Compatibility Information

- [Release Information](#release-information)
- [Version History](#version-history)
  - [Version 2.0.1](#version-201)
  - [Version 2.0.0](#version-200)
  - [Version 1.0.0](#version-100)
- [Future Compatibility Risk: ChatGPT Shared-Link Storage Changes](#future-compatibility-risk-chatgpt-shared-link-storage-changes)
  - [Existing Version 2 Archives](#existing-version-2-archives)

## Release Information

Use the latest published release of ChatGPT Conversation Bridge unless you have a specific reason to use an earlier version.

Published releases are available from the project's GitHub Releases page:

https://github.com/hummbugg/chatgpt-conversation-bridge/releases

Each GitHub release is a fixed snapshot of the repository source and documentation at the time that version was published. The release notes provide a concise summary of that release.

The Version History below provides a more detailed record of the significant features, behavior changes, bug fixes, safety and reliability improvements, dependency changes, and compatibility changes introduced by each public release.

[↑ Back to Document Navigator](#project-release-version-history-and-compatibility-information)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

## Version History

This history covers the significant features, behavior changes, fixes, safety and reliability improvements, dependency changes, and compatibility changes introduced by each public release of ChatGPT Conversation Bridge. Releases are listed in reverse chronological order.

### Version 2.0.1

Version 2.0.1 focused on portability, safety, reliability, and refinement of the Share URL and permanent-archive architecture introduced in Version 2.0.0.

**Behavior Change — Simplified offline archive regeneration**

Offline DOCX regeneration now accepts the conversation/archive name rather than a ZIP filename or path. The script locates the corresponding ZIP in its `archive` directory automatically.

For example:

Windows:

```text
python chatgpt_conversation_bridge.py "Ear Crevice in Dog"
```

Linux/macOS:

```text
python3 chatgpt_conversation_bridge.py "Ear Crevice in Dog"
```

Passing a `.zip` filename or an archive path is no longer supported. Share URL usage is unchanged.

**Portable filename and path safety**

Filename handling was expanded for safe operation across Windows, Linux, and macOS. Conversation titles are normalized and sanitized for filesystem use while preserving Unicode characters when possible. The script also protects against Windows reserved filenames and enforces filename character, UTF-8 byte, and complete output-path limits.

If the project is located too deeply in the filesystem for a safe output path, the script reports a clear error so the project can be moved to a shallower location.

**Computer-local message timestamps**

Message timestamps are now formatted using the computer's local timezone, including the timezone's historical daylight-saving-time rules, instead of always being formatted in U.S. Eastern Time.

The `tzdata` and `tzlocal` packages were added to the documented Python dependencies to support this behavior consistently across supported platforms.

**Safer DOCX replacement**

When regenerating a DOCX that already exists, the complete replacement document is now generated as a temporary candidate before the existing DOCX is replaced. The completed candidate is committed with an atomic filesystem replacement operation only after generation succeeds.

This prevents the existing DOCX from being deliberately deleted before its replacement has been successfully created.

**Improved archive and DOCX commit ordering**

During Share URL processing, the completed DOCX remains temporary while the permanent conversation archive is created and verified. The DOCX is committed to its final location only after the required archive operation succeeds.

This improves failure handling and reduces the possibility of leaving a partially completed set of permanent output files.

**Clearer completion reporting**

Successful Share URL and offline archive operations now report explicit `SUCCESS` status messages that identify the completed workflow.

**Documentation and cross-platform validation**

Installation, dependency, offline-regeneration, filename/path safety, timezone, output-safety, and troubleshooting documentation was expanded.

The final Version 2.0.1 implementation was validated on Windows, Ubuntu Linux, and macOS, including offline archive regeneration, timezone formatting, and long Unicode-capable conversation filenames. Additional fresh-install testing was performed with Python 3.14 on Windows and both Intel and Apple Silicon macOS systems.

[↑ Back to Document Navigator](#project-release-version-history-and-compatibility-information)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

### Version 2.0.0

Version 2.0.0 introduced a major architectural change in how ChatGPT Conversation Bridge acquires and permanently archives conversations. The established conversation-validation and DOCX-conversion capabilities from Version 1.0.0 were retained while the browser-saved HTML acquisition and archive system was replaced.

**Major Architecture Change — Direct ChatGPT Share URL retrieval**

The script now accepts a public ChatGPT Share URL and retrieves the shared conversation directly from ChatGPT. Users no longer need to use a browser's Save Page feature to create an HTML/HTM file and companion `_files` directory before running the script.

This changed the acquisition layer while retaining the existing conversion capabilities for producing the formatted DOCX.

**JSON-based permanent conversation archives**

The permanent archive format was redesigned around the data retrieved from the Share URL. Instead of preserving the browser-saved HTML/HTM page and its companion `_files` directory, Version 2.0.0 archives the retrieved conversation JSON together with successfully retrieved uploaded images.

These archives provide the source material needed to regenerate the conversation DOCX later without retrieving that conversation from its ChatGPT Share URL again.

**Direct retrieval of available uploaded images**

Uploaded images associated with the shared conversation can now be retrieved directly through the ChatGPT Share workflow and stored in the permanent archive for use in the generated DOCX and later offline regeneration.

When an uploaded image or file cannot be retrieved, the established placeholder behavior preserves a visible indication of the unavailable upload rather than silently omitting it.

**New offline archive regeneration workflow**

Offline DOCX regeneration was redesigned to operate on the new Version 2 archive format containing conversation JSON and retrieved images rather than on the browser-saved HTML archives used by Version 1.0.0.

This allows an already captured Version 2 conversation to be regenerated without contacting ChatGPT again.

**New Share-retrieval dependency**

The `curl_cffi` package was introduced as a third-party Python dependency to support retrieval of public ChatGPT Share conversations.

This was a significant change from Version 1.0.0, which required no third-party Python packages and operated entirely on files that had already been saved locally.

**Operational reporting**

Version 2.0.0 added reporting appropriate to the new Share URL workflow, including retrieval progress and expanded completion information about the conversation, generated DOCX, permanent archive, messages, uploads, timestamps, and references.

The permanent archive is verified as part of the Share workflow, and an existing archive is protected from being silently overwritten.

[↑ Back to Document Navigator](#project-release-version-history-and-compatibility-information)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

### Version 1.0.0

Version 1.0.0 was the initial public release of ChatGPT Conversation Bridge. It established the conversation-validation and DOCX-conversion foundation that continued into later versions.

**Browser-saved ChatGPT conversation capture**

The original workflow operated entirely on local files. Users saved a ChatGPT shared conversation from a browser as an HTML/HTM page and its companion `_files` directory, then ran ChatGPT Conversation Bridge against that saved capture.

The script searched for a live HTML/HTM capture first and then for an existing archived capture. It did not contact ChatGPT or any other network service and required no third-party Python packages.

**Conversation validation and DOCX conversion**

The script extracted the conversation data from the saved page, validated the conversation chain, selected the visible user and assistant messages, and generated a formatted Word DOCX containing the conversation.

The conversion engine preserved substantial Markdown formatting, including headings, lists, tables, blockquotes, fenced code blocks, inline code, bold and italic text, Markdown escapes, horizontal rules, and line-break behavior.

**Images, uploads, references, and timestamps**

Locally available uploaded images could be embedded in the DOCX while preserving their aspect ratio and limiting their displayed size. PNG, JPEG, GIF, and WebP image data was recognized from the actual file contents.

When an uploaded image or file was referenced by the conversation but was not available in the saved browser capture, the DOCX included a visible placeholder instead of silently omitting the upload.

Public reference URLs associated with assistant responses were preserved as plain text, and available message creation times were included in the DOCX. Version 1.0.0 formatted those timestamps in U.S. Eastern Time.

**Verified permanent archives**

After successfully generating a DOCX from a new browser capture, the script created a permanent ZIP archive containing the saved HTML/HTM page and its companion `_files` directory. The archive was reopened and verified before the original browser capture was removed.

If an archive for that conversation already existed, the newly saved browser capture was left unchanged rather than silently replacing the existing permanent archive.

**Offline DOCX regeneration**

An existing permanent archive could be used to regenerate the DOCX directly from the ZIP without extracting or modifying the archive and without contacting ChatGPT.

This local archive-and-regeneration workflow provided the original preservation model for ChatGPT Conversation Bridge before the Share URL and JSON-based archive architecture introduced in Version 2.0.0.

[↑ Back to Document Navigator](#project-release-version-history-and-compatibility-information)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

## Future Compatibility Risk: ChatGPT Shared-Link Storage Changes

ChatGPT Conversation Bridge depends on the structure and behavior of public ChatGPT shared-conversation data. That representation and the mechanisms used to retrieve it are controlled by OpenAI and may change independently of this project.

A future change to ChatGPT's public Share system could therefore require corresponding changes to ChatGPT Conversation Bridge even when the Python script itself has not otherwise changed. Such a change could potentially affect Share URL retrieval, the structure of returned conversation data, conversation traversal, upload and image retrieval, timestamps, or other information used to construct the permanent archive and DOCX.

If a previously working version of ChatGPT Conversation Bridge begins failing with newly created Share URLs, an upstream change to ChatGPT's public Share representation or retrieval behavior should be considered as a possible cause. This does not necessarily indicate that the local installation, Python environment, or an existing archive is damaged.

### Existing Version 2 Archives

Existing Version 2 archives are importantly different from new live Share captures.

After a Share conversation has been successfully captured and archived, offline DOCX regeneration uses the conversation data and available uploaded images already stored in that permanent archive. It does not retrieve that conversation from its ChatGPT Share URL again.

As a result, a future upstream ChatGPT change that prevents or disrupts new Share URL retrieval would not necessarily prevent DOCX regeneration from Version 2 archives that were successfully created before that change.

Permanent archives should therefore be retained. In addition to supporting offline DOCX regeneration, they preserve the successfully captured conversation data independently of later changes to the live ChatGPT Share representation.

This should not be interpreted as a guarantee that every future version of ChatGPT Conversation Bridge will remain compatible with every archive created by every earlier version. If the project's archive format changes in the future, corresponding compatibility or migration handling may be required.

[↑ Back to Document Navigator](#project-release-version-history-and-compatibility-information)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)
