
import os
from pathlib import Path

from dotenv import load_dotenv
from langchain_openai import ChatOpenAI

print("Starting breakfast GenAI application...", flush=True)

# Load .env from the same directory as this script
env_path = Path(__file__).resolve().parent / ".env"
load_dotenv(dotenv_path=env_path)

# Check API key
if not os.getenv("OPENAI_API_KEY"):
    raise ValueError(
        "OPENAI_API_KEY is missing. Please check your .env file."
    )

print("API key loaded successfully.", flush=True)

# Initialize the LangChain model
llm = ChatOpenAI(
    model="gpt-4o-mini",
    temperature=0
)

# Prompt
prompt = """
Suggest exactly 5 healthy breakfast ideas.

Requirements:
- Number the ideas from 1 to 5.
- Write one idea per line.
- Include nutritious options containing protein or fiber.
- Do not include an introduction or conclusion.
"""

print("Requesting breakfast ideas from OpenAI...", flush=True)

# Call the model
reply = llm.invoke(prompt)

# Print generated text
print("\nYour 5 healthy breakfast ideas:\n", flush=True)
print(reply.content, flush=True)
