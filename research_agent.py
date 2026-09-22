from crewai import Agent, Task, Crew, LLM
import os
import re


# ============================================================
# DEEPSEEK V4.1 FLASH
# ============================================================

DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY")

if not DEEPSEEK_API_KEY:
    raise RuntimeError(
        "DEEPSEEK_API_KEY is missing. "
        "Add DEEPSEEK_API_KEY to Streamlit Cloud Secrets."
    )

MODEL_NAME = "deepseek/deepseek-flash"

llm = LLM(
    model=MODEL_NAME,
    api_key=DEEPSEEK_API_KEY,
    base_url="https://api.deepseek.com",
    temperature=0.2,
)
