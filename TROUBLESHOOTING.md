MASTER PROMPT — CREATE ONE NEW PROJECT SLIDE

Create ONE COMPLETELY NEW POWERPOINT SLIDE for my internship presentation.

PROJECT:
IMT Reconciliation & Update Automation

IMPORTANT CONTEXT:
I already have an existing PowerPoint presentation containing my First Citizens internship presentation, template, theme, and branding.

I DO NOT want you to edit, redesign, overwrite, or modify any existing slide.

Instead:

→ ADD ONE BRAND-NEW SLIDE to the existing presentation.
→ Keep ALL existing slides exactly as they are.
→ Do not modify any existing slide.
→ Do not modify the existing slide master.
→ Do not modify existing presentation-wide branding.
→ Do not disturb the existing First Citizens logo.
→ The new slide should visually belong to the same presentation and naturally match its existing First Citizens corporate theme.

The existing First Citizens logo/branding should NOT be recreated manually.
If the presentation's existing layout/theme automatically provides the logo on the new slide, preserve it exactly.
If it does not automatically appear, DO NOT manually add a new logo unless it is already part of the presentation's existing template/master.

DO NOT add any additional company logo, symbol, watermark, page number, decorative branding or unrelated graphics.

==================================================
PRIMARY OBJECTIVE
==================================================

Design ONE polished, premium corporate slide that explains the COMPLETE project at a high level.

The slide must communicate:

PROBLEM
↓
SOLUTION
↓
AUTOMATED PIPELINE
↓
OUTPUTS
↓
SINGLE APPLICATION / BUSINESS VALUE

The audience should understand the project within approximately 20–30 seconds.

This is an internship project presentation, NOT a technical architecture document.

Keep the slide visually strong, concise and presentation-friendly.

==================================================
SLIDE TITLE
==================================================

Use exactly:

IMT Reconciliation & Update Automation

Place the title prominently at the top, using the same typography and visual treatment as the existing presentation.

Do not add a long subtitle.

==================================================
1. PROBLEM
==================================================

Create a compact PROBLEM section near the top.

Use this wording:

"Every month, updating the Master YTD requires downloading the ServiceNow monthly report and manually comparing rows and columns — a time-consuming, complex and error-prone process."

Do not display this as a large paragraph.

Use a compact card/banner structure:

PROBLEM

Every month, updating the Master YTD requires downloading the ServiceNow monthly report and manually comparing rows and columns — a time-consuming, complex and error-prone process.

Immediately connect this to the solution.

==================================================
2. SOLUTION
==================================================

Create a visually clear SOLUTION statement immediately following the problem.

Use:

"SOLUTION — A Python-based automation that performs the complete reconciliation and update workflow and packages the solution into a single application for portable, user-friendly execution by non-technical users."

Keep this concise.

The visual narrative should clearly communicate:

MANUAL & ERROR-PRONE
        →
AUTOMATED & STANDARDIZED

Do not use oversized decorative arrows.

==================================================
3. MAIN VISUAL — AUTOMATED PIPELINE
==================================================

The main visual element of the slide must be the end-to-end automation pipeline.

Use native editable PowerPoint shapes and connectors.

Do NOT create the pipeline as a single image.

Do NOT use a screenshot.

Do NOT use Mermaid.

Do NOT flatten the diagram.

Every major stage must remain independently editable.

Use this exact high-level flow:

INPUTS
→
DATA LOADER
→
VALIDATOR
→
RECONCILER
→
BUSINESS RULES
→
OUTPUT WRITER
→
OUTPUTS

Prefer a clean horizontal flow across the central portion of the slide.

If required for readability, use a slightly stepped layout, but maintain a clear left-to-right progression.

The pipeline should be visually dominant.

==================================================
4. INPUTS
==================================================

Show three compact input cards:

MASTER YTD
(.xlsm / .xlsx)

SERVICENOW MONTHLY REPORT
(.xlsx / .csv)

IC LOOKUP
(Excel sheet)

These feed into:

DATA LOADER

Do not add any additional input source.

==================================================
5. DATA LOADER
==================================================

Heading:

DATA LOADER

Supporting text:

"Loads and standardizes Master YTD, ServiceNow report and IC Lookup data."

Do NOT show Python function names.

Do not show implementation-level code.

==================================================
6. VALIDATOR
==================================================

Heading:

VALIDATOR

Supporting text:

"Validates required columns, Number fields and IC lookup data before processing."

Keep it compact.

The purpose is simply to communicate that input data is checked before reconciliation.

==================================================
7. RECONCILER
==================================================

Heading:

RECONCILER

This should be one of the visually important stages.

Show:

MATCH BY NUMBER

Then clearly display the three reconciliation outcomes:

UPDATED
Existing record matched and updated

NEW
Not present in Master YTD → created as new

HISTORICAL
Not present in current ServiceNow report → retained

Use the terminology exactly:

UPDATED
NEW
HISTORICAL

A compact representation such as:

MATCH BY NUMBER
      ↓
UPDATED | NEW | HISTORICAL

is preferred.

Do not over-explain the reconciliation algorithm.

==================================================
8. BUSINESS RULES
==================================================

Heading:

BUSINESS RULES

Show the four important rules:

• IC Determination
• Region Lookup
• Bank / SVB Determination
• Duration Calculation

Where useful, show the logic very briefly:

IC:
IC → Proposed By for Priority 3 → blank

Region:
IC → Region Lookup

Bank / SVB:
Assignment Group → Bank / SVB

Duration:
Applicable dates/timestamps → Duration

Do not make this section text-heavy.

The audience only needs to understand that standardized business logic is applied after reconciliation.

==================================================
9. OUTPUT WRITER
==================================================

Heading:

OUTPUT WRITER

Supporting statement:

"Updates the existing Excel template while preserving workbook structure, formulas and reporting format."

Show a compact list:

• Write updated data
• Preserve workbook structure
• Preserve formulas / tables
• Save final output
• Generate audit information

Do not focus on implementation libraries.

Do not show code.

If the actual technical terminology needs to be retained, the heading may be:

OUTPUT WRITER (COM)

But prioritize what it accomplishes rather than the technology used.

==================================================
10. OUTPUTS
==================================================

Show three final output cards:

UPDATED MASTER YTD
(.xlsm)

RECONCILIATION REPORT
(.xlsx)

AUDIT LOG
(.json)

Make the flow visually obvious:

OUTPUT WRITER
      ↓
UPDATED MASTER YTD
RECONCILIATION REPORT
AUDIT LOG

==================================================
11. SINGLE APPLICATION / PACKAGING
==================================================

This is an important part of the project.

The Python solution was not left as a collection of scripts.

The complete codebase and dependencies were packaged into a single application to make it easier for non-technical users to run.

Create a compact callout near the lower portion of the slide:

SINGLE APPLICATION

"Complete codebase + dependencies packaged into one portable application"

Then show four concise benefits:

PORTABLE
PLATFORM-INDEPENDENT
USER-FRIENDLY
NON-TECHNICAL USER READY

Do not make this larger than the main pipeline.

The visual message should be:

Python Codebase
+
Dependencies
+
Automation Workflow
        ↓
SINGLE APPLICATION

Do not show PyInstaller or other packaging implementation details unless absolutely necessary.

==================================================
12. BUSINESS VALUE
==================================================

At the bottom of the content area, include a compact value strip:

AUTOMATED MONTHLY RECONCILIATION
REDUCED MANUAL EFFORT
CONSISTENT & ACCURATE PROCESS
AUDITABLE & REPEATABLE OUTPUT

Do not add numerical claims.

Do not invent percentage improvements.

==================================================
RECOMMENDED STORY
==================================================

The slide should visually communicate this:

        MANUAL MONTHLY PROCESS

Download ServiceNow Report
        ↓
Compare Rows / Columns
        ↓
Reconcile Master YTD
        ↓
Apply Business Rules
        ↓
Update Workbook
        ↓
Risk of Manual Errors

                 ↓

        PYTHON AUTOMATION

Master YTD + ServiceNow + IC Lookup
        ↓
Data Loader
        ↓
Validator
        ↓
Reconciler
        ↓
Business Rules
        ↓
Output Writer
        ↓
Updated Master + Report + Audit

                 ↓

        SINGLE APPLICATION

Portable | Platform-independent | User-friendly
                 ↓
        Non-technical Users

Do NOT literally reproduce this ASCII diagram.

Use it only as the conceptual structure.

==================================================
VISUAL DESIGN
==================================================

The new slide must match the visual language of the existing First Citizens internship presentation.

Use the existing presentation as the design reference for:

• typography
• colors
• spacing
• background
• section headings
• card treatment
• line weight
• visual hierarchy

The slide should feel as if it was designed as part of the same presentation from the beginning.

Use:

• clean corporate styling
• minimal design
• professional blue/white visual language consistent with the template
• subtle accents
• clean editable shapes
• precise alignment
• generous whitespace
• strong hierarchy

Avoid:

• flashy colors
• gradients
• excessive shadows
• cartoon graphics
• stock images
• decorative icons
• AI-generated illustrations
• unnecessary symbols
• excessive arrows
• visual clutter

The PIPELINE should be the primary visual.

The PROBLEM → SOLUTION transition should be the primary narrative.

The SINGLE APPLICATION callout should be secondary.

==================================================
IMPORTANT — DO NOT OVER-TECHNICALIZE
==================================================

This is ONE project slide in an internship presentation.

Do NOT include:

• Python code
• function names
• class names
• file paths
• package names
• library names
• source-code snippets
• detailed implementation architecture
• excessive technical terminology

Do not turn the slide into a developer documentation page.

The audience needs to understand:

WHAT WAS THE PROBLEM?
WHAT DID I BUILD?
HOW DOES IT WORK?
WHAT DOES IT PRODUCE?
WHY IS IT BETTER?

==================================================
EXACT PROJECT TERMINOLOGY
==================================================

Preserve these terms:

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

Do NOT rename these into artificial AI/ML terminology.

This is an automation/reconciliation project and should be represented accurately.

==================================================
LAYOUT
==================================================

Use approximately this hierarchy:

--------------------------------------------------

IMT Reconciliation & Update Automation

[ PROBLEM ]
Every month, updating the Master YTD requires downloading the
ServiceNow monthly report and manually comparing rows and columns —
a time-consuming, complex and error-prone process.

[ SOLUTION ]
Python-based end-to-end automation → Single Application

              AUTOMATED WORKFLOW

[INPUTS]
Master YTD | ServiceNow Monthly Report | IC Lookup

                     ↓

[DATA LOADER]
                     ↓
[VALIDATOR]
                     ↓
[RECONCILER]
MATCH BY NUMBER
UPDATED | NEW | HISTORICAL
                     ↓
[BUSINESS RULES]
IC | Region | Bank/SVB | Duration
                     ↓
[OUTPUT WRITER]
                     ↓
[OUTPUTS]
Updated Master YTD | Reconciliation Report | Audit Log

       [ SINGLE APPLICATION ]
Portable | Platform-independent | User-friendly
                     ↓
             Non-technical Users

[Automated] [Reduced Manual Effort] [Consistent] [Auditable]

--------------------------------------------------

Optimize the exact positioning based on the dimensions and visual structure of the existing presentation.

Do not blindly follow the ASCII layout if it produces an overcrowded slide.

==================================================
NEW SLIDE REQUIREMENT
==================================================

Again:

THIS MUST BE A NEW SLIDE.

Do NOT edit the content of any existing slide.

Do NOT replace an existing slide.

Do NOT redesign an existing slide.

Do NOT modify existing slide objects.

Do NOT modify the existing presentation's master/theme.

Add this project as a completely new slide while maintaining visual consistency with the presentation.

If the presentation already has a standard project-slide layout, use that visual language as the reference but create the new project content independently.

==================================================
LOGO / BRANDING REQUIREMENT
==================================================

Do NOT manually add any First Citizens logo.

Do NOT add any other company logo.

Do NOT add symbols representing First Citizens.

Do NOT add decorative branding.

If the existing PowerPoint theme/master automatically places the First Citizens logo on the newly created slide, leave it untouched.

The existing logo position and appearance must remain exactly consistent with the rest of the presentation.

==================================================
EDITABILITY
==================================================

All newly created content must be editable.

Use native PowerPoint:

• Text boxes
• Rectangles
• Rounded rectangles
• Lines
• Connectors
• Arrows

Do not create one large image containing the entire slide.

==================================================
FINAL QUALITY CHECK
==================================================

Before finalizing the new slide, verify:

✓ Exactly ONE new slide has been added.

✓ No existing slide has been modified.

✓ Existing presentation theme remains intact.

✓ Existing First Citizens branding remains intact.

✓ No additional logo has been introduced.

✓ No page number has been introduced.

✓ No watermark has been introduced.

✓ Project title is exactly:
"IMT Reconciliation & Update Automation"

✓ Problem statement clearly explains the monthly manual Master YTD + ServiceNow comparison process.

✓ Python-based automation is clearly presented as the solution.

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

✓ "Match by Number" is clearly communicated.

✓ Business Rules include:

IC Determination
Region Lookup
Bank / SVB Determination
Duration Calculation

✓ Outputs include:

Updated Master YTD
Reconciliation Report
Audit Log

✓ Single Application packaging is clearly communicated.

✓ Portable, platform-independent and user-friendly benefits are visible.

✓ Non-technical-user usability is clearly communicated.

✓ No unsupported numerical claims are added.

✓ No unnecessary implementation details are included.

✓ All newly created shapes/text are editable.

✓ No overlapping objects.

✓ No tiny unreadable text.

✓ Adequate whitespace.

✓ Strong visual hierarchy.

✓ The slide looks like a natural part of the existing First Citizens internship presentation.

MOST IMPORTANT:

The final slide should tell ONE clear story:

MANUAL MONTHLY RECONCILIATION
        ↓
PYTHON-BASED AUTOMATION
        ↓
LOAD + VALIDATE
        ↓
RECONCILE
        ↓
APPLY BUSINESS RULES
        ↓
UPDATE + REPORT + AUDIT
        ↓
SINGLE PORTABLE APPLICATION
        ↓
LESS MANUAL EFFORT + CONSISTENT + AUDITABLE

Create ONLY this ONE NEW PROJECT SLIDE.