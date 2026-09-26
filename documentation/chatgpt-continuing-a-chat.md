# Continuing a Chat

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)

## Document Navigator


**Have you already completed the [Normal Workflow](chatgpt-conversation-bridge-documentation.md#normal-workflow) steps that finish by running ChatGPT Conversation Bridge to create the DOCX (Microsoft Word / LibreOffice Writer) document and the Archive ZIP? If not, you must complete the [Normal Workflow](chatgpt-conversation-bridge-documentation.md#normal-workflow) before proceeding with Continuing a Chat.**


-   [Purpose](#purpose)
-   [Standard Initial Chat Continuation Message](#standard-initial-chat-continuation-message)
-   [How to Submit the Standard ChatGPT Conversation Continuation Instructions](#how-to-submit-the-standard-chatgpt-conversation-continuation-instructions)
-   [Standard ChatGPT Conversation Continuation Instructions to be Copied](#standard-chatgpt-conversation-continuation-instructions-to-be-copied)

------------------------------------------------------------------------

Ready-to-copy standard text for continuing a ChatGPT conversation

## Purpose

This document provides the standard copy-and-paste text used when continuing a ChatGPT conversation from a continuation DOCX created by ChatGPT Conversation Bridge.

The companion Markdown version on GitHub is intended to provide convenient online access to the same templates, including copy-to-clipboard access where practical.

## Standard Initial Chat Continuation Message

Drag and drop the continuation DOCX which is the document previously created by ChatGPT Conversation Bridge from now on referred to as continuation DOCX, into ChatGPT's message composer but do not submit the message yet.

Copy and paste the following standard message into ChatGPT's message composer. Submit the message and wait for ChatGPT to acknowledge that the continuation document has been received.

``` text
CHAT CONTINUATION

The attached DOCX contains the previous ChatGPT conversation that this new chat
will continue.

Do not read, analyze, summarize, or begin using the attached continuation
document yet.

I will provide the ChatGPT Conversation Continuation Instructions next. Those
instructions tell you how to use this continuation document and how to
initialize this new conversation.

For now, only acknowledge that the continuation document has been received and
wait for the instruction document.
```

Once you receive the reply from ChatGPT proceed to the next section "**How to Submit the Standard ChatGPT Conversation Continuation Instructions**"

## How to Submit the Standard ChatGPT Conversation Continuation Instructions

Since ChatGPT has already confirmed that it has received the continuation DOCX, we now must give ChatGPT instructions stating what to do with the continuation document.

The complete standard continuation-instructions template appears in the next section. Copy the complete template following the instructions in the section **Standard ChatGPT Conversation Continuation Instructions to be Copied** below and do not submit yet.

## Standard ChatGPT Conversation Continuation Instructions to be Copied

Copy the complete **ChatGPT Conversation Continuation Instructions** below and paste them into ChatGPT's message composer.

After pasting the instructions, scroll to the top of ChatGPT's message composer so that you can see `My Continuation Document.docx` on the `CONTINUATION DOCUMENT:` line.

`My Continuation Document.docx` is a placeholder for the actual continuation DOCX. Replace `My Continuation Document.docx` with the exact filename of the continuation DOCX that was uploaded in the section **Standard Initial Chat Continuation Message** above.

Verify that the filename is correct, then submit the message. The remaining standard continuation instructions are intended to stay unchanged.

``` text
ChatGPT Conversation Continuation Instructions

Generic instructions for continuing a previous ChatGPT conversation from a generated DOCX

Continuation Document

CONTINUATION DOCUMENT: My Continuation Document.docx

For each new continuation, change only the filename above so that it exactly matches the
uploaded DOCX containing the previous conversation.

The remaining instructions are intended to stay the same for every conversation.

Purpose

These instructions tell a new ChatGPT conversation how to continue from the named continuation
DOCX with as little loss of context as practical.

The continuation DOCX is the chronological record of the previous conversation. The objective is
continuation, not redesign, reinterpretation, or replacement of the earlier discussion.

Continuation Declaration

This chat is a direct continuation of the ChatGPT conversation preserved in the CONTINUATION
DOCUMENT named above.

Treat the continuation DOCX and the current chat as one chronological conversation when
historical checking is requested.

The continuation DOCX contains the older history. The current chat contains the newer history. A
later explicit decision supersedes an earlier conflicting decision unless the later record
explicitly says otherwise.

Do not silently restore an older requirement, conclusion, preference, or decision merely because
it appears in the continuation DOCX.

Initial Synchronization

When these instructions and the named continuation DOCX are first provided to a new chat,
perform a one-time initial synchronization before resuming normal work.

During initial synchronization, review the continuation DOCX sufficiently to understand:

- the subject of the conversation;
- the important established context;
- the exact point where the previous conversation stopped.

Do not attempt to redesign the subject matter while performing synchronization.

When initial synchronization is complete, begin the continuation with:

FULL CHECK: ON

Keep FULL CHECK: ON until the user changes it.

What Information Should Be Recovered?

During initial synchronization, recover, when applicable:

- important decisions, conclusions, requirements, preferences, and agreements;
- completed work and results actually reported;
- tests, checks, or validations actually performed;
- important limitations, warnings, or known problems;
- deferred work and unresolved questions;
- the working method or procedure established by the user;
- the exact point where the previous conversation stopped and the likely next step.

Interpret the Historical Record Carefully

Distinguish carefully among discussion, suggestion, approval, implementation, testing,
verification, and final acceptance.

Do not claim that something reached a later state unless the record actually establishes it.

- A suggestion or proposed idea is not automatically an approved decision.
- Text, code, instructions, or a solution produced by ChatGPT is not automatically something the
  user implemented.
- Implemented work is not automatically tested or verified.
- A successful test does not automatically establish that the result became the final design.
- If the record is genuinely ambiguous or incomplete, identify the uncertainty rather than
  guessing.

Continuation Synchronization Report

After reviewing the continuation DOCX for the first time, provide a concise Continuation
Synchronization Report before beginning substantive new work.

Include:

- the name of the continuation DOCX reviewed;
- the general subject or purpose of the continued conversation;
- the current FULL CHECK state;
- the most important established context needed to continue safely;
- the latest completed work or conclusion, when applicable;
- the exact point where the previous conversation stopped;
- important deferred or unresolved items relevant to continuing;
- whether any ambiguity currently prevents safe continuation.

Do not automatically begin the next major task merely because synchronization is complete.

After reporting the recovered state, wait for the user's instruction unless the user's
continuation message already contains a clear task to perform.

Treat the Continuation as One Chronological Conversation

Treat the continuation DOCX and the current chat as one chronological conversation when
historical checking is requested.

The continuation DOCX contains the older history. The current chat contains the newer history.

When reviewing historical information, follow the conversation forward sufficiently to determine
the most recent applicable state rather than treating an earlier statement in isolation.

Current Information Takes Precedence

The continuation DOCX is historical.

Information supplied or corrected by the user in the current chat is newer and takes precedence
when the two conflict.

Do not overwrite newer current-chat information with an older continuation-document version.

Later Decisions Supersede Earlier Conflicting Decisions

A later explicit decision supersedes an earlier conflicting decision unless the later record
explicitly says otherwise.

Do not silently restore an older requirement, conclusion, preference, or decision merely because
it appears in the continuation DOCX.

When checking an important historical statement, review the later history sufficiently to
determine whether that statement was changed, rejected, refined, implemented differently,
superseded, or otherwise qualified.

Then reconcile the continuation-document history with anything established later in the current
conversation.

The most recent explicit decision governs unless the record says otherwise.

Handling Uncertainty and Ambiguous History

If the historical record is genuinely ambiguous or incomplete, identify the uncertainty rather
than guessing.

When an exact historical fact remains uncertain after checking the applicable sources, say that
it is uncertain rather than filling the gap from assumption.

What Is FULL CHECK?

This continuation uses a persistent switch named FULL CHECK.

After initial synchronization, display the current FULL CHECK state at the top of every
response, using exactly one of:

FULL CHECK: ON

FULL CHECK: OFF

The selected state remains in effect until the user changes it.

Start a newly synchronized continuation with FULL CHECK: ON.

The user may switch between FULL CHECK: ON and FULL CHECK: OFF as often as needed. The most
recently selected state remains active until changed.

FULL CHECK: OFF

When FULL CHECK is OFF, answer using only the current ChatGPT conversation.

Do not routinely consult, search, or reread the continuation DOCX. This keeps ordinary
continuation work fast and prevents unnecessary reprocessing of a large historical document.

FULL CHECK: ON

When FULL CHECK is ON, use both the current conversation and the named continuation DOCX
whenever prior history could materially affect the answer.

Before treating an important decision, rule, requirement, preference, unresolved question,
implementation issue, or historical fact as new, unanswered, or missing, check the continuation
DOCX as necessary.

Do not stop at the first matching historical statement. Check the later history sufficiently to
determine whether that statement was changed, rejected, refined, implemented differently,
superseded, or otherwise qualified.

Then reconcile the continuation-document history with anything established later in the current
conversation. The most recent explicit decision governs unless the record says otherwise.

When an exact historical fact remains uncertain after checking both sources, say that it is
uncertain rather than filling the gap from assumption.

Switching FULL CHECK Modes

If the user says full check on or otherwise clearly requests FULL CHECK: ON, switch to FULL
CHECK: ON and keep it on until changed.

If the user says full check off or otherwise clearly requests FULL CHECK: OFF, switch to FULL
CHECK: OFF and keep it off until changed.

Explicit Historical Checks While FULL CHECK Is OFF

If the user explicitly asks you to check, search, verify, or compare something against the
continuation DOCX, do so for that request even while FULL CHECK is OFF.

This explicit historical check does not by itself change the persistent FULL CHECK setting.

After completing the requested historical check, continue using FULL CHECK: OFF unless the user
explicitly changes the setting.

Do Not Automatically Begin the Next Major Task

Do not automatically begin the next major task merely because synchronization is complete.

After reporting the recovered state, wait for the user's instruction unless the user's
continuation message already contains a clear task to perform.

Final Continuation Principle

The purpose of these instructions is to make a new ChatGPT conversation behave as closely as
practical to a continuation of the previous one.

The continuation DOCX preserves the detailed earlier history. The current chat carries the work
forward.

FULL CHECK: ON provides a deliberate mechanism for consulting both sources when historical
certainty matters, while FULL CHECK: OFF keeps normal work focused on the current conversation.

When historical certainty matters, check rather than guess.

When FULL CHECK is OFF and no historical check was requested, stay within the current
conversation.
```

[↑ Back to Document Navigator](#continuing-a-chat)

[← Back to Main Information Navigator](../README.md#chatgpt-conversation-bridge)