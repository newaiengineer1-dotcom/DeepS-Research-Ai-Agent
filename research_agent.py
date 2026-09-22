import os

from crewai import Agent, Task, Crew, LLM


# ============================================================
# DEEPSEEK API
# ============================================================

DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY")

if not DEEPSEEK_API_KEY:
    raise RuntimeError(
        "DEEPSEEK_API_KEY is missing from Streamlit Secrets."
    )


llm = LLM(
    model="deepseek/deepseek-flash",
    api_key=DEEPSEEK_API_KEY,
    base_url="https://api.deepseek.com",
    temperature=0.2,
)


# ============================================================
# TEST FUNCTION
# ============================================================

def generate_long_research_report(
    topic,
    evidence="",
    max_pages=10,
):

    agent = Agent(
        role="Senior Research Analyst",
        goal="Produce a professional research report on the requested topic.",
        backstory=(
            "You are an experienced research analyst who produces "
            "structured, factual and professional research reports."
        ),
        llm=llm,
        verbose=False,
    )

    task = Task(
        description=f"""
Research and prepare a professional report about:

{topic}

Additional evidence:
{evidence}

Target length:
Approximately {max_pages} pages.

Create a structured report with:

1. Executive Summary
2. Introduction
3. Background
4. Market / Industry Analysis
5. Technology Analysis
6. Key Findings
7. Opportunities
8. Risks and Challenges
9. Future Outlook
10. Conclusions
11. References

Do not invent sources or facts.
Clearly identify assumptions and estimates.
""",
        expected_output="A professional Markdown research report.",
        agent=agent,
    )

    crew = Crew(
        agents=[agent],
        tasks=[task],
        verbose=False,
    )

    result = crew.kickoff()

    return str(result)
