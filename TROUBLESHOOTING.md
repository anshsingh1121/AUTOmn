MASTER PROMPT — CREATE ONE FINAL INTERNSHIP PRESENTATION SLIDE

You are editing my EXISTING PowerPoint presentation.

Create ONLY ONE slide for the following project:

IMT Reconciliation & Update Automation

This is a project I developed to automate the monthly Master YTD reconciliation and update process.

IMPORTANT:
Do NOT create a new presentation.
Do NOT redesign the presentation template.
Do NOT modify the slide master.
Do NOT modify the existing background.
Do NOT modify existing theme elements.
Do NOT move, resize, recolor, replace, or disturb the existing First Citizens logo at the bottom-left.
Do NOT add another First Citizens logo.
Do NOT add any company logo.
Do NOT add any additional branding.
Do NOT add page numbers.
Do NOT add watermarks.
Do NOT add decorative corporate symbols.
Do NOT add unnecessary icons.
Do NOT add stock images.
Do NOT add unrelated graphics.

The existing presentation already contains the correct First Citizens corporate template and logo. Treat all existing template elements as LOCKED and untouched.

Only use the available content area of the existing slide.

The final result must look like it naturally belongs to the existing internship presentation.

==================================================
CORE OBJECTIVE
==================================================

Create a SINGLE high-quality executive/technical overview slide that communicates the complete project in approximately 20–30 seconds.

The story must be immediately understandable:

MANUAL MONTHLY PROCESS
        ↓
PYTHON-BASED AUTOMATION
        ↓
VALIDATION + RECONCILIATION
        ↓
BUSINESS RULES
        ↓
AUTOMATED OUTPUT
        ↓
SINGLE USER-FRIENDLY APPLICATION

The slide should communicate that this was NOT merely a Python script.

It was an end-to-end automation solution that:

• removes repetitive monthly manual work
• automates ServiceNow-to-Master YTD reconciliation
• applies business rules consistently
• generates the required outputs
• packages the complete codebase and dependencies into a single application
• makes the solution portable, platform-independent and usable by non-technical users

==================================================
SLIDE TITLE
==================================================

Use exactly:

IMT Reconciliation & Update Automation

Do not add a subtitle directly under the title unless necessary for visual balance.

==================================================
SECTION 1 — PROBLEM STATEMENT
==================================================

Place a compact PROBLEM section near the top of the slide.

Use this exact meaning, but format it as polished presentation text:

"Every month, updating the Master YTD requires downloading the ServiceNow monthly report and manually comparing rows and columns — a time-consuming, complex and error-prone process."

Keep it concise and highly readable.

Do NOT display this as a large paragraph.

Use:

PROBLEM

followed by the statement in a compact text/card area.

Immediately communicate the solution underneath or beside it:

"SOLUTION — A Python-based automation that performs the complete reconciliation and update workflow, packaged as a single application for portable, user-friendly execution by non-technical users."

The Problem and Solution should visually form a clear transition:

Manual & Error-Prone
        →
Automated & Standardized

Do NOT use large decorative arrows.

==================================================
SECTION 2 — MAIN AUTOMATION PIPELINE
==================================================

The central and largest portion of the slide must contain the actual end-to-end pipeline.

IMPORTANT:
Use the SAME terminology used in the actual project architecture.

Do NOT replace the terminology with generic AI/software terminology.

The pipeline must be:

INPUTS
↓
DATA LOADER
↓
VALIDATOR
↓
RECONCILER
↓
BUSINESS RULES
↓
OUTPUT WRITER
↓
OUTPUTS

Prefer a clean horizontal pipeline if it fits naturally within the template.

If the available content area is too narrow, use a compact left-to-right stepped flow.

Do NOT use a tall architecture diagram that consumes the entire slide vertically.

The pipeline should look like a business automation workflow, not a software engineering class diagram.

==================================================
INPUTS
==================================================

Show three compact input cards feeding the DATA LOADER:

MASTER YTD
(.xlsm / .xlsx)

SERVICENOW MONTHLY REPORT
(.xlsx / .csv)

IC LOOKUP
(Excel sheet)

Use these exact concepts.

Do not introduce additional input files.

Visually distinguish these as INPUTS rather than processing stages.

==================================================
STAGE 1 — DATA LOADER
==================================================

Use the exact heading:

DATA LOADER

Under it, use a concise description:

"Loads and standardizes Master YTD, ServiceNow report and IC Lookup data."

Do NOT show Python function names such as:

load_master_tracker()
load_servicenow()
load_ic_lookup()

Those are implementation details and are unnecessary for the presentation.

The audience should understand WHAT this stage does, not the function names.

==================================================
STAGE 2 — VALIDATOR
==================================================

Use the exact heading:

VALIDATOR

Use concise supporting text:

"Validates required columns, Number fields and IC lookup data before processing."

The visual message should be:

INPUT QUALITY CHECK
→
VALID DATA ENTERS RECONCILIATION

Do not overpopulate this box.

==================================================
STAGE 3 — RECONCILER
==================================================

This is one of the CORE stages of the project.

Use the exact heading:

RECONCILER

Clearly communicate that records are reconciled using the incident Number.

Use this compact logic:

MATCH BY NUMBER

Then show three outcomes:

UPDATED
Existing record matched and updated

NEW
Not present in Master YTD → created as new

HISTORICAL
Not present in current ServiceNow report → retained

The terminology UPDATED / NEW / HISTORICAL must be clearly visible.

A concise secondary statement may be:

"Matched by Number → Update / Create / Retain"

This stage should be visually prominent because it represents the central reconciliation logic.

Do NOT show Python implementation details.

==================================================
STAGE 4 — BUSINESS RULES
==================================================

Use the exact heading:

BUSINESS RULES

Show the key business rules implemented by the automation.

Use these four items:

• IC Determination
• Region Lookup
• Bank / SVB Determination
• Duration Calculation

Where useful, communicate the actual rule logic compactly:

IC Determination
IC → Proposed By for Priority 3 → blank

Region Lookup
IC → Region

Bank / SVB
Assignment Group → Bank / SVB

Duration
Calculate duration from applicable dates/timestamps

Do not turn these into lengthy explanations.

The audience should understand that the reconciled records are then enriched/transformed according to established business rules.

==================================================
STAGE 5 — OUTPUT WRITER
==================================================

Use the exact heading:

OUTPUT WRITER

Use this concise description:

"Updates the existing Excel template while preserving workbook structure, formulas and reporting format."

Show the key responsibilities:

• Write updated data
• Preserve workbook structure
• Preserve formulas / tables
• Update required sheets
• Save final outputs
• Generate audit information

IMPORTANT:

The actual implementation uses Excel/COM automation.

Do NOT make "COM" the visual focus.

If technically necessary, retain the project terminology:

OUTPUT WRITER (COM)

But make the business function the primary message, not the implementation technology.

Do NOT show library names such as openpyxl, pandas, pywin32, etc.

==================================================
SECTION 3 — OUTPUTS
==================================================

At the end of the pipeline, show three clean output cards:

UPDATED MASTER YTD
(.xlsm)

RECONCILIATION REPORT
(.xlsx)

AUDIT LOG
(.json)

Make it visually obvious that these are generated automatically by the pipeline.

Do not introduce other output types.

==================================================
SECTION 4 — APPLICATION PACKAGING
==================================================

This is an important part of the project and must be visible, but it should NOT compete with the main pipeline.

Create a compact callout near the bottom/right of the pipeline:

SINGLE APPLICATION

"Complete codebase + dependencies packaged into one portable application"

Then show four short benefits:

Portable
Platform-independent
User-friendly
For non-technical users

The message should communicate:

Python scripts
+
Dependencies
+
Configuration
+
Automation workflow
        ↓
SINGLE APPLICATION

The purpose is to show that the project was taken beyond development scripts and converted into a practical application that can be used without requiring users to manually manage the Python environment.

Do NOT show packaging implementation details such as PyInstaller unless absolutely necessary.

The audience only needs to understand the outcome:
a single, portable and user-friendly application.

==================================================
SECTION 5 — BUSINESS VALUE
==================================================

At the bottom of the content area, create a very compact VALUE strip.

Use four short outcomes:

AUTOMATED MONTHLY RECONCILIATION
REDUCED MANUAL EFFORT
CONSISTENT & ACCURATE PROCESS
AUDITABLE & REPEATABLE OUTPUT

Do not invent numerical improvement percentages.

Do not claim specific time savings unless those numbers already exist in the presentation.

Do not use words such as "100% accurate" or "zero errors."

==================================================
RECOMMENDED VISUAL HIERARCHY
==================================================

The slide should visually follow this structure:

------------------------------------------------------------

IMT Reconciliation & Update Automation

[ PROBLEM ]
Every month, updating the Master YTD requires downloading the
ServiceNow monthly report and manually comparing rows and columns —
a time-consuming, complex and error-prone process.

[ SOLUTION ]
Python-based end-to-end automation → packaged as a single application
for portable, user-friendly execution by non-technical users.

                 AUTOMATED WORKFLOW

[ INPUTS ]
Master YTD     ServiceNow Report     IC Lookup
       \              |                 /
        \             |                /
              [ DATA LOADER ]
                     ↓
              [ VALIDATOR ]
                     ↓
              [ RECONCILER ]
          Match by Number
       ┌────────┬────────┬──────────┐
     UPDATED     NEW    HISTORICAL
       └────────┴────────┴──────────┘
                     ↓
            [ BUSINESS RULES ]
       IC | Region | Bank/SVB | Duration
                     ↓
             [ OUTPUT WRITER ]
                     ↓
       ┌────────────┼──────────────┐
 Updated Master   Reconciliation   Audit
     YTD             Report         Log

          [ SINGLE APPLICATION ]
     Portable | Platform-independent
       User-friendly | Non-technical

[ AUTOMATED ] [ LESS MANUAL EFFORT ] [ CONSISTENT ] [ AUDITABLE ]

------------------------------------------------------------

Do NOT reproduce this literal ASCII layout on the slide.
Use it only as the conceptual design.

Optimize the exact placement based on the existing PowerPoint template.

==================================================
VISUAL DESIGN REQUIREMENTS
==================================================

The slide must feel like a premium corporate internship presentation.

Use the existing First Citizens presentation theme as the source of truth for:

• typography
• background
• spacing
• color palette
• visual hierarchy
• section treatment
• overall style

Do NOT introduce a completely new visual style.

Use subtle corporate blue accents only if they already fit the existing template.

Use clean rectangular/rounded cards for the pipeline.

Use thin professional connectors.

Use consistent card dimensions.

Use consistent typography.

Use strong alignment.

Use generous whitespace.

Use visual hierarchy rather than excessive graphics.

The pipeline should be the HERO visual element.

The Problem/Solution should be the HERO narrative.

The packaging and business-value sections should be secondary.

==================================================
STRICT TEMPLATE PROTECTION
==================================================

THIS IS EXTREMELY IMPORTANT.

The existing PowerPoint template is already correct.

Treat every existing template object as LOCKED.

Specifically:

DO NOT MOVE the First Citizens logo.
DO NOT RESIZE the First Citizens logo.
DO NOT CROP the First Citizens logo.
DO NOT RECOLOR the First Citizens logo.
DO NOT REPLACE the First Citizens logo.
DO NOT COVER the First Citizens logo.
DO NOT PLACE ANY OBJECT ON TOP OF THE LOGO.
DO NOT ADD ANOTHER LOGO.
DO NOT ADD A LOGO INSIDE THE PROJECT DIAGRAM.
DO NOT ADD A LOGO IN THE HEADER.
DO NOT ADD A LOGO IN THE FOOTER.

Do not create additional footer elements.

Do not alter the existing presentation dimensions.

Do not change slide orientation.

Do not change the slide master.

Do not change the background.

Do not change the existing template layout.

Only populate/design the project content within the safe content region.

If there is insufficient space, REDUCE SECONDARY CONTENT rather than modifying the template.

==================================================
EDITABILITY REQUIREMENT
==================================================

Everything newly created must remain editable in PowerPoint.

Use:

• native PowerPoint text boxes
• native PowerPoint rectangles/rounded rectangles
• native PowerPoint connectors/arrows
• native PowerPoint lines

Do NOT create the complete slide as a single image.

Do NOT create the pipeline as one flattened graphic.

Do NOT use a screenshot.

Do NOT use Mermaid.

Do NOT use an exported diagram image.

Every major pipeline stage must be independently editable.

==================================================
CONTENT DENSITY
==================================================

This is a 10-minute internship presentation.

The slide is ONE PROJECT among multiple projects.

Therefore:

DO NOT explain the complete codebase.

DO NOT show classes.

DO NOT show Python functions.

DO NOT show file paths.

DO NOT show package names.

DO NOT show technical implementation code.

DO NOT show model/library names.

DO NOT show excessive technical details.

The slide should communicate:

WHAT WAS THE PROBLEM?
WHAT DID I BUILD?
HOW DOES IT WORK?
WHAT DOES IT PRODUCE?
WHY IS IT USEFUL?

That is enough.

==================================================
TERMINOLOGY — DO NOT CHANGE
==================================================

Use these project terms exactly:

IMT Reconciliation & Update Automation

Master YTD

ServiceNow Monthly Report

IC Lookup

DATA LOADER

VALIDATOR

RECONCILER

UPDATED

NEW

HISTORICAL

BUSINESS RULES

IC Determination

Region Lookup

Bank / SVB Determination

Duration Calculation

OUTPUT WRITER

Updated Master YTD

Reconciliation Report

Audit Log

Single Application

Do not replace these with generic alternatives such as:

"Data Processing Engine"
"AI Engine"
"Intelligence Layer"
"Decision Engine"
"Data Transformation Layer"

This project should be represented accurately, not artificially made to sound like an AI system.

==================================================
IMPORTANT NARRATIVE
==================================================

The slide must make this progression immediately obvious:

BEFORE

Monthly manual process:
Download ServiceNow report
→ Compare rows/columns
→ Reconcile with Master YTD
→ Apply business rules
→ Update workbook
→ Risk of manual errors

AFTER

Automated Python solution:
Load
→ Validate
→ Reconcile
→ Apply business rules
→ Write outputs
→ Package as a single application

This "BEFORE → AFTER" transformation is the main story of the project.

==================================================
DESIGN PRIORITY
==================================================

Prioritize the following in order:

1. Project title
2. Problem statement
3. Python automation solution
4. Complete pipeline
5. Reconciliation outcomes
6. Business rules
7. Outputs
8. Single Application packaging
9. Business value

If space becomes limited:

FIRST remove explanatory sentences.

THEN shorten secondary descriptions.

THEN simplify business-value text.

DO NOT remove the core pipeline.

DO NOT remove UPDATED / NEW / HISTORICAL.

DO NOT remove the Single Application message.

DO NOT reduce the slide to tiny unreadable text.

==================================================
READABILITY REQUIREMENT
==================================================

The slide must be readable when projected in a meeting room.

Avoid:

• tiny fonts
• dense paragraphs
• excessive text
• overlapping boxes
• excessive arrows
• complicated branching
• unnecessary icons
• visual noise

A person seeing the slide for the first time should understand the project at a high level within approximately 10 seconds.

==================================================
FINAL QUALITY CHECK
==================================================

Before finalizing, perform a complete visual and content check.

Verify:

✓ Exactly ONE slide has been created/modified.

✓ Existing presentation template remains unchanged.

✓ Existing First Citizens logo at bottom-left remains exactly where it was.

✓ No additional logo has been introduced.

✓ No page number has been introduced.

✓ No watermark has been introduced.

✓ No unrelated symbols have been introduced.

✓ Project title is exactly:
"IMT Reconciliation & Update Automation"

✓ Problem clearly communicates the monthly manual Master YTD + ServiceNow reconciliation issue.

✓ Solution clearly communicates Python-based automation.

✓ Complete pipeline is visible:

Inputs
→ Data Loader
→ Validator
→ Reconciler
→ Business Rules
→ Output Writer
→ Outputs

✓ Reconciler clearly shows:

UPDATED
NEW
HISTORICAL

✓ Matching by Number is clearly communicated.

✓ Business Rules clearly include:

IC Determination
Region Lookup
Bank / SVB Determination
Duration Calculation

✓ Outputs clearly include:

Updated Master YTD
Reconciliation Report
Audit Log

✓ Single Application packaging is clearly communicated.

✓ Portable, platform-independent and user-friendly nature is communicated.

✓ Non-technical-user usability is communicated.

✓ No unsupported numerical business claims have been added.

✓ No unnecessary implementation details have been added.

✓ All newly created objects are editable.

✓ All connectors are properly aligned.

✓ No objects overlap.

✓ No content enters the existing footer/logo area.

✓ Slide has sufficient whitespace.

✓ Typography is presentation-readable.

✓ Visual hierarchy is clear.

✓ The slide looks like part of a professional First Citizens internship presentation.

✓ The slide communicates the COMPLETE PROJECT without becoming technically cluttered.

MOST IMPORTANT:

Do not optimize this slide for showing how much technical information can fit onto it.

Optimize it for showing the complete transformation:

MANUAL MONTHLY RECONCILIATION
        ↓
PYTHON AUTOMATION
        ↓
VALIDATE + RECONCILE
        ↓
APPLY BUSINESS RULES
        ↓
UPDATE MASTER + REPORT + AUDIT
        ↓
SINGLE PORTABLE APPLICATION
        ↓
LESS MANUAL EFFORT + CONSISTENT + AUDITABLE PROCESS

Create ONLY this final single slide.