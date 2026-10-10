from pathlib import Path
from dotenv import load_dotenv
from langchain_openai import ChatOpenAI
import os

# Load .env from the folder containing this Python script
env_path = Path(__file__).resolve().parent / ".env"
load_dotenv(dotenv_path=env_path, override=True)

# Check configuration without printing secret values
print("Environment file exists:", env_path.exists())
print("OpenAI key loaded:", bool(os.getenv("OPENAI_API_KEY")))
print("LangSmith key loaded:", bool(os.getenv("LANGSMITH_API_KEY")))
print("LangSmith tracing:", os.getenv("LANGSMITH_TRACING"))
print("LangSmith project:", os.getenv("LANGSMITH_PROJECT"))

llm = ChatOpenAI(model="gpt-4o-mini")

prompts = [
    "Explain what LangChain is in simple terms.",
    "What is the difference between monitoring and observability?",
    "Explain how LangSmith helps debug an LLM application."
]

for prompt in prompts:
    try:
        response = llm.invoke(prompt)
        print(f"\nPrompt: {prompt}")
        print(f"Answer: {response.content}")
    except Exception as e:
        print(f"\nFailed for prompt: {prompt}")
        print(f"Error: {type(e).__name__}: {e}")
        break