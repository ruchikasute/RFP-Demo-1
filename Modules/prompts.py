

def get_executive_summary_and_objective_prompt(reference_text, condensed_rfp, num_interfaces=None):
    """
    Generates Crave-style Executive Summary and Objective based on the RFP.
    """
    interface_info = (
        f"The ~{num_interfaces} interfaces represent a scope of {num_interfaces} ICOs to migrate."
        if num_interfaces
        else "The project involves migration from the current SAP PI/PO integration platform to SAP Integration Suite."
    )

    return f"""
You are an expert SAP RFP proposal writer for **Crave InfoTech**.

Your writing must strictly follow **the tone, structure, and style of the reference text below**.
Do NOT use generic openings like "honored" or "delighted".

### 🔹 TASK:
1️Executive Summary – 300–350 words.  
   Follow Crave’s tone exactly:
   - Start with: **"Crave InfoTech is pleased to submit proposal for the PI/PO Integration Migration..."**
   - Acknowledge the client’s initiative and business goal.
   - Mention Crave’s SAP partnership, global presence, and expertise.
   - Keep paragraphs structured and professional (avoid sales tone).
   - Add we provide section and explain in detail

2️ Objective – a short 100-word paragraph + the below table:

| No. | Migration of ICOs from SAP PI/PO to SAP Integration Suite as per details below |
|------|--------------------------------------------------------------------------------|
| 1 | No of Interfaces to be migrated from SAP PI/PO to SAP Integration Suite: {num_interfaces} |

Then add:
**Interfaces Configuration Objects (ICOs) are listed in the Appendix.**

---

### 🔹 PROJECT CONTEXT:
{interface_info}

### 🔹 USE THIS STYLE AND TONE AS STRICT REFERENCE:
{reference_text}

### 🔹 CONDENSED RFP CONTENT:
{condensed_rfp}
"""

def get_scope_prereq_assumptions_prompt(reference_text, condensed_rfp, num_interfaces=None):
    """Compact, focused prompt to generate a concise 'Scope and Out of Scope' section."""
    interface_info = (
        f"Migration of approximately {num_interfaces} interfaces from SAP PI/PO to SAP Integration Suite."
        if num_interfaces
        else "Migration of interfaces from SAP PI/PO to SAP Integration Suite."
    )

    return f"""
You are an SAP proposal writer at Crave InfoTech.

Generate a concise, professional section covering:
- In Scope
- Migration Project Prerequisites
- Assumptions
- Out of Scope

Tone and structure should match a real client proposal — crisp, business-like, and easy to read.  
Each section should have 4–8 bullet points, each one line or two at most.  
Mention {interface_info} in the scope.  
Use 'the client' instead of any past customer name.  
Avoid unnecessary descriptions or closing summaries.

Reference Text:
{reference_text}

Condensed RFP:
{condensed_rfp}
"""

def get_resource_schedule_and_commercial_prompt(reference_text, condensed_rfp):
    """
    Concise prompt for generating the highly structured Resource Schedule and Commercials section.
    """
    return f"""
You are a senior SAP proposal writer at Crave InfoTech.

Generate the section titled **Resource Schedule and Commercials**. The output MUST strictly follow this exact structure using Markdown.

**STRUCTURE REQUIRED (Strict Replication):**

### Resource Schedule

    * Start with the intro sentence: "Crave Infotech proposes to deploy the following team with their indicative loading based on current understanding –"
    * Follow with a 4-column Markdown table (`Crave Infotech Resources`, `Location`, `Allocation`, `Resource Count`) containing the exact roles: Project Manager (Onshore, Fulltime), Integration Developer (Onshore, Fulltime), Integration Developer (Offshore, Fulltime), Business Analyst (Offshore, Fulltime).

    * Start with the intro sentence: "Recommended team from [Client Name, use 'the client'] who need to be available during the project execution –"
    * Follow with a 3-column Markdown table (`Client Resources`, `Allocation`, `Resource Count`) containing the exact roles: Project Manager, SAP IT/Basis, Solution Architect, Integration Specialist, Business Analysts/Functional SME, ABAP (If required).

### Commercials

    * Start with the paragraph: "We propose to execute this project on T&M basis. Following is the resource estimation and indicative of total cost:"
    * Include the fixed line: "Cost: $ (17 Weeks)"
4.  **Change Request:** Include the bolded paragraph: "**Any new enhancements or changes identified during the project phase will be considered a change request and will be estimated separately**"
5.  **Notes:** Use header `Note:` followed by a bulleted list of the two specified points (resource/fee estimates and onsite billing details).
6.  **Payment Terms:** Use header `Timesheet, Invoices and Payment Terms` followed by a bulleted list of the four specified payment/invoicing terms.

Reference Material:
{reference_text}

Condensed RFP:
{condensed_rfp}
"""


def get_communication_plan_prompt(reference_text, condensed_rfp):
    """
    Generates a concise, structured Communication Plan section prompt with Crave/client role clarity.
    """
    return f"""
You are an expert SAP proposal writer at Crave InfoTech.

Write a formal, client-ready **Communication Plan** section for an SAP migration or implementation RFP.  
Describe how Crave InfoTech and the client (use actual client name from context if available, e.g., Haceb) will manage communication, meetings, reporting, and escalation during the project.

The section must include:

1. 2–3 lines on the importance of clear and consistent communication for project success, stakeholder alignment, and timely decision-making.

2. **Exhibit: Daily Interaction** – Table with columns:  
   Activity | Communication Mode | Report Recipient/s | Frequency | Comments  
   Include a row for Weekly Status Report and others like Kick-off Meeting, Daily Stand-up, Steering Committee, etc.  
   Clearly mention which roles belong to **Crave InfoTech** and which belong to **the client** (e.g., “Crave InfoTech Project Manager”, “Haceb Project Manager”).

3. **Issue Resolution and Escalation Procedure** – Short paragraph describing the structured approach for logging, managing, and escalating issues.  
   Then include:
   - **Table: Issue Management** (Task | Timescale | Responsibility)  
   - **Bulleted list:** Issue reporting guidelines that mention role responsibilities on both Crave and client sides.

4. **Table: Issue Classification** – Columns: Problem Type | Definition | Reporting Process | Solution Responsible  
   Include rows for Low, Serious, and Critical.  
   Indicate clearly who is responsible on Crave side and on the client side for each case.

5. **Table: Escalation Process** – Columns: Issue Type | Escalation Point | Escalation Criteria | Governance Role (Project Core Group)  
   Include rows for Project Delivery, Contract, Unresolved Delivery Issue, and Program Management Issue.  
   Specify which escalation points or governance roles belong to Crave vs. the client.

6. Close with 1–2 lines summarizing how this structured plan ensures transparency, timely updates, and strong collaboration between Crave InfoTech and the client.

**Formatting rules:**
- Use markdown headers and tables exactly.
- Keep tone formal, enterprise-level, and realistic.
- Replace “the client” with actual client name (from context or reference) wherever possible.
- Keep total length around 700–900 words.
- Ensure all roles are clearly marked as Crave-side or Client-side.

### 🔹 USE THIS STYLE AND TONE AS STRICT REFERENCE:
{reference_text}

Condensed RFP:
{condensed_rfp}
"""

