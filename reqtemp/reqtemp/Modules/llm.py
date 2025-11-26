from openai import AzureOpenAI
import os

def get_llm_client():
    return AzureOpenAI(
        azure_endpoint=os.getenv("AZURE_OPENAI_FRFP_ENDPOINT"),
        api_key=os.getenv("AZURE_OPENAI_FRFP_KEY"),
        api_version=os.getenv("AZURE_OPENAI_FRFP_VERSION")
    )



def regenerate_section_llm(client, model_name, section_title, original_text, user_prompt):

    prompt = f"""
You are an expert SAP proposal writer.

PRIORITY RULE:
- USER INSTRUCTION ALWAYS COMES FIRST.

FORMATTING RULES:
- NEVER start the output with a heading.
- Bullets must use "- " at the start.
- Sub-bullets use two spaces "  - ".
- Keep formatting clean and consistent.

CONTENT RULES:
- If the user says "add", "expand", "more content", then EXPAND the section with more bullets, more depth, and SAP-specific detail.
- If the user says "shorten", "condense", "make concise", then SUMMARIZE while keeping all key meaning.
- If the user instruction doesn't specify size, rewrite cleanly with improved structure.
- You are allowed to produce long content when expanding.

Rewrite the section below according to the USER INSTRUCTION.

SECTION TITLE: {section_title}

ORIGINAL SECTION:
{original_text}

USER INSTRUCTION:
{user_prompt}

WRITE THE IMPROVED VERSION BELOW:
"""

    response = client.chat.completions.create(
        model=model_name,
        messages=[{"role":"user", "content":prompt}],
        temperature=0.5
    )

    return response.choices[0].message.content.strip()

# def regenerate_section_llm(client, model_name, section_title, original_text, user_prompt):

#     prompt = f"""
# You are an expert SAP proposal writer.

# PRIORITY RULE:
# - USER INSTRUCTION ALWAYS COMES FIRST.

# STRICT FORMATTING RULES:
# - NEVER start the output with a heading (###, ##, or numbered headings). Always begin with a paragraph or bullets.
# - Use ONLY markdown headings in this exact form:
#   - ### Title
#   - #### Sub-title
# - Bullets MUST use "- " at the start of the line.
# - Sub-bullets MUST be indented using exactly two spaces: "  - point".
# - NEVER generate any automatic numbering inside headings.
# - No paragraph should be long; break text into bullets or compact sections.
# - NEVER output plain unformatted text.

# CONDENSING RULES:
# - If the user says "add", "expand", or "more content", EXPAND the content with richer SAP-specific details.
# - If the user says "shorten", "condense", "make concise", then SUMMARIZE while keeping key meaning.
# - If the user gives no specific direction, improve clarity and structure only.


# Rewrite the section below according to the USER INSTRUCTION.

# SECTION TITLE: {section_title}

# ORIGINAL SECTION:
# {original_text}

# USER INSTRUCTION:
# {user_prompt}

# """



#     response = client.chat.completions.create(
#         model=model_name,
#         messages=[{"role":"user", "content":prompt}],
#         temperature=0.3
#     )

#     return response.choices[0].message.content.strip()
