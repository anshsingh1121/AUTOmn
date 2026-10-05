You are an expert enterprise presentation designer, solution architect, and technical storytelling specialist.

I need you to create ONE SINGLE, HIGH-IMPACT POWERPOINT SLIDE for my internship presentation at FIRST CITIZENS BANK.

PROJECT:
ServiceNow Incident Intelligence Platform

OBJECTIVE:
Combine the complete PROBLEM STATEMENT, SOLUTION STATEMENT, and END-TO-END AI PIPELINE into ONE visually powerful slide.

IMPORTANT:
This is an internship presentation, NOT a technical documentation slide.

The audience should understand within 10–15 seconds:
1. What business problem existed
2. What I built
3. How the AI platform works at a high level
4. What value it provides

Do NOT overload the slide with implementation-level details.

==================================================
1. CORE BUSINESS PROBLEM
==================================================

ServiceNow receives a large volume of IT incidents.

Traditionally, an L1/helpdesk agent manually reads each incident, understands the issue, determines the appropriate assignment group, and routes it to the correct team.

This creates four major problems:

• Incorrect Assignment
  Tickets can bounce between multiple teams, creating avoidable delay.

• Delayed Resolution
  Misrouted incidents remain in incorrect queues for hours or days.

• SLA Risk
  Priority incidents can lose valuable resolution time because of manual routing.

• Knowledge Loss
  Similar historical incidents are not consistently reused when new incidents occur.

DO NOT use exaggerated statistics unless they are explicitly available in the source material.

Instead of presenting a large paragraph, represent the problem visually as:

SERVICE NOW INCIDENT
        ↓
Manual Reading
        ↓
Manual Triage
        ↓
Assignment Group Selection
        ↓
Delay / Misrouting / SLA Risk / Knowledge Loss

Use 3–4 concise problem cards rather than a paragraph.

==================================================
2. SOLUTION STATEMENT
==================================================

The solution is:

"An AI-powered Incident Intelligence Platform that automatically understands a ServiceNow incident, predicts the most appropriate assignment team, estimates resolution characteristics, and retrieves similar historical incidents to support faster and more informed triage."

Core capabilities:

• Automated incident understanding
• Assignment-group prediction
• Resolution-time prediction
• Historical incident retrieval
• Explainable AI / feature importance
• End-to-end local/enterprise-ready ML pipeline

The central message should be:

"From Manual Incident Triage → AI-Assisted Incident Intelligence"

Make this transformation visually prominent.

==================================================
3. ACTUAL TECHNICAL PIPELINE
==================================================

The underlying implementation contains a detailed multi-stage pipeline.

Do NOT display all 12 stages individually on the main slide.

Instead, compress them into FIVE intelligent macro stages while preserving the actual architecture.

MACRO STAGE 1 — DATA INTELLIGENCE

Underlying stages:
• Data Validation
• ML Readiness Assessment
• Data Cleaning
• External Enrichment

Purpose:
Validate data quality, detect leakage/readiness issues, clean the dataset, and enrich incident context using enterprise information such as CMDB/shift data.

Represent as:

ServiceNow Data
→ Validate
→ Clean
→ Enrich

--------------------------------------------------

MACRO STAGE 2 — FEATURE & TEXT INTELLIGENCE

Underlying stages:
• Feature Engineering
• NLP Text Preprocessing
• Exploratory Data Analysis
• Train/Validation/Test Splitting

Purpose:
Convert raw incident information into ML-ready structured and textual features while maintaining zero-leakage preprocessing.

Represent as:

Structured Features + Incident Text
→ Feature Engineering
→ NLP Processing
→ EDA
→ Zero-Leakage Dataset

--------------------------------------------------

MACRO STAGE 3 — PREDICTIVE AI

Underlying stages:
• CatBoost Classification
• CatBoost Regression
• Hyperparameter Optimization

Models:
• Assignment Group Classification
• Resolution Time Regression

Represent visually as two parallel AI branches:

                    ┌→ Assignment Group
ML Features → AI    │   Classification
                    │
                    └→ Resolution Time
                        Regression

Use "CatBoost + HPO" as the technical label.

--------------------------------------------------

MACRO STAGE 4 — SEMANTIC MEMORY

Underlying stages:
• Embedding Generation
• FAISS Vector Index

Purpose:
Create a searchable semantic representation of historical incidents and retrieve relevant precedents.

Represent as:

Historical Incidents
→ Embeddings
→ FAISS Index
→ Similar Incident Retrieval

--------------------------------------------------

MACRO STAGE 5 — HYBRID INCIDENT INTELLIGENCE

Underlying stage:
• Hybrid Recommendation Engine

Purpose:
Combine predictive model outputs with historical precedents.

Final output:

NEW INCIDENT
        ↓
Predicted Assignment Group
+
Predicted Resolution Time
+
Similar Historical Incidents
+
Confidence / Supporting Intelligence

This is the final AI-assisted triage output.

==================================================
4. MOST IMPORTANT ARCHITECTURE VISUAL
==================================================

The center of the slide should contain ONE clean horizontal pipeline:

SERVICE NOW INCIDENT
        ↓
[ DATA INTELLIGENCE ]
        ↓
[ FEATURE + NLP INTELLIGENCE ]
        ↓
[ PREDICTIVE AI ]
        ↓
[ SEMANTIC MEMORY ]
        ↓
[ HYBRID INCIDENT INTELLIGENCE ]
        ↓
ASSIGNMENT GROUP + RESOLUTION ESTIMATE + HISTORICAL PRECEDENTS

Use arrows to make the flow immediately understandable.

Do NOT create a giant complicated architecture diagram.

Each macro stage should have:
• Short stage title
• One-line purpose
• 1–2 technology labels maximum
• Simple enterprise-style icon

==================================================
5. SLIDE STRUCTURE
==================================================

Use a 16:9 widescreen layout.

Recommended composition:

--------------------------------------------------
TOP
--------------------------------------------------

Title:

SERVICE NOW INCIDENT INTELLIGENCE PLATFORM

Subtitle:

AI-powered incident triage, prediction and historical intelligence

--------------------------------------------------
LEFT ~25%
--------------------------------------------------

Section:

THE CHALLENGE

Show:

Manual Incident
      ↓
Manual Triage
      ↓
Routing Decision

Then four small impact cards:

Incorrect Assignment
Delayed Resolution
SLA Risk
Knowledge Loss

Keep this section visually compact.

--------------------------------------------------
CENTER ~55%
--------------------------------------------------

Section:

AI-POWERED SOLUTION

Display the five-stage pipeline horizontally or as a clean left-to-right flow:

1. Data Intelligence
2. Feature + NLP Intelligence
3. Predictive AI
4. Semantic Memory
5. Hybrid Intelligence

This should be the visual focal point of the entire slide.

--------------------------------------------------
RIGHT ~20%
--------------------------------------------------

Section:

INTELLIGENCE OUTPUT

Use three large output cards:

🎯 Assignment Group
⏱ Resolution Time
🔎 Historical Precedents

Optionally include:

Confidence / Explainability

Do not add unsupported numerical claims.

--------------------------------------------------
BOTTOM
--------------------------------------------------

A single transformation statement:

MANUAL TRIAGE
        →
AI-ASSISTED INCIDENT INTELLIGENCE

Small footer:

First Citizens Bank | Internship Project

==================================================
6. VISUAL DESIGN — FIRST CITIZENS BANK
==================================================

The slide must feel like an enterprise banking presentation, not a generic AI startup presentation.

Use a visual language inspired by First Citizens Bank corporate materials:

• Deep navy / dark blue as primary brand color
• Clean white background or very light neutral background
• Blue hierarchy for sections and pipeline
• Restrained yellow/gold accent for important highlights
• Very limited use of red
• Strong whitespace
• Professional financial-services aesthetic
• Flat/vector enterprise icons
• Subtle rounded cards
• Thin separators
• No excessive gradients
• No neon colors
• No cyberpunk styling
• No futuristic robot imagery
• No generic AI stock images

If an approved First Citizens Bank logo asset is available in the PowerPoint environment, use it subtly in the footer/header.

If no approved logo asset is available:
DO NOT fabricate a logo.
Use clean text:
"First Citizens Bank"

Do not make the slide look like an advertisement.

It should look like an internal enterprise technology presentation.

==================================================
7. TYPOGRAPHY
==================================================

Use a modern corporate sans-serif font such as:

Aptos
Arial
Segoe UI

Title:
Bold, approximately 28–34 pt

Section headings:
18–22 pt

Pipeline stage names:
14–17 pt

Supporting text:
10–13 pt

Avoid tiny text.

Every important statement must be readable from presentation distance.

==================================================
8. ICONOGRAPHY
==================================================

Use simple consistent line/vector icons:

ServiceNow / ticket → incident icon
Data validation → database/check icon
Feature engineering → gear/data icon
AI prediction → brain/model icon
Semantic memory → search/database icon
Hybrid intelligence → connected nodes icon
Assignment group → target/team icon
Resolution time → clock icon
Historical precedents → document/search icon

Do NOT use random emoji.

Use one consistent icon style.

==================================================
9. CONTENT OPTIMIZATION
==================================================

CRITICAL:

Do not copy the existing markdown/documentation literally.

Rewrite and compress the content for executive presentation.

Avoid:

• Long paragraphs
• 12 individual pipeline boxes
• Large tables
• Excessive technical terminology
• Repeated descriptions
• Unsupported metrics
• Fake performance numbers
• Claims such as "100% accurate"
• Claims that the system is production deployed if it is not
• Claims of real-time production automation unless explicitly supported

Use technical keywords only where they add credibility:

ServiceNow
CatBoost
HPO
NLP
FAISS
Hybrid Recommendation
Zero-Leakage
Explainability

==================================================
10. IMPORTANT TECHNICAL ACCURACY
==================================================

The detailed implementation contains:

Data validation
ML readiness assessment
Data cleaning
External enrichment
Feature engineering
NLP preprocessing
EDA
Train/validation/test split
CatBoost classification
CatBoost regression
Hyperparameter optimization
Embedding generation
FAISS indexing
Hybrid recommendation

The final slide should represent ALL of these capabilities conceptually, but through five macro stages.

Do not invent additional ML models.

Do not introduce BERT, LLMs, transformers, RAG, XGBoost, LightGBM, Random Forest, or other technologies unless they are actually part of the final implementation being presented.

Use the CURRENT implementation as the source of truth.

==================================================
11. STORYTELLING
==================================================

The slide should visually tell this story:

BEFORE:

High-volume ServiceNow incidents
        ↓
Manual understanding
        ↓
Manual routing
        ↓
Delay + SLA risk + knowledge loss

AFTER:

ServiceNow Incident
        ↓
AI understands incident
        ↓
Predictive models
        ↓
Historical semantic search
        ↓
Hybrid intelligence
        ↓
Actionable triage recommendation

The viewer should immediately understand:

"Instead of manually figuring out where an incident belongs, the platform uses AI + historical knowledge to assist the routing decision."

==================================================
12. DESIGN PRIORITY
==================================================

Prioritize in this exact order:

1. STORY
2. READABILITY
3. VISUAL HIERARCHY
4. PIPELINE CLARITY
5. TECHNICAL CREDIBILITY
6. BRAND CONSISTENCY

Do not sacrifice readability merely to fit more information.

If content does not fit, REMOVE WORDS — do not shrink the font.

==================================================
13. FINAL SLIDE QUALITY CHECK
==================================================

Before finalizing, verify:

✓ One slide only
✓ 16:9
✓ Problem is immediately understandable
✓ Solution is immediately understandable
✓ Pipeline is visible end-to-end
✓ All major technical capabilities are represented
✓ No unnecessary 12-stage detail
✓ No unsupported metrics
✓ No fake claims
✓ First Citizens enterprise visual style
✓ Strong visual hierarchy
✓ No paragraph-heavy areas
✓ No tiny text
✓ No clutter
✓ Consistent icons
✓ Professional banking aesthetic
✓ Suitable for a 10-minute internship presentation

The final slide should look like a polished consulting/enterprise AI architecture slide, not a software documentation page.

If necessary, redesign the layout completely rather than mechanically placing the provided text into boxes.

FINAL OUTPUT:
Create the actual PowerPoint slide with the above content and visual design.

The slide must be presentation-ready.