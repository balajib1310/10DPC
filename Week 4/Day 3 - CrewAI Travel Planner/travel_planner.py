import os
from dotenv import load_dotenv
from crewai import Agent, Task, Crew, LLM

# Load environment variables from .env
load_dotenv()

# Check the OpenAI API key
if not os.getenv("OPENAI_API_KEY"):
    raise ValueError("OPENAI_API_KEY is missing from the .env file")

# Configure the AI model
llm = LLM(
    model=os.getenv("MODEL", "openai/gpt-4o-mini")
)

# Ask the user to enter a city
city = input("Which city would you like to visit? ").strip()

if not city:
    raise ValueError("Please enter a valid city name.")

# Create the Travel Planner Agent
planner = Agent(
    role="Expert Travel Planner",
    goal=(
        "Recommend exactly three interesting places to visit "
        "in the city selected by the user."
    ),
    backstory=(
        "You are an experienced travel planner who helps first-time "
        "visitors discover famous attractions and hidden gems. "
        "You provide practical, concise recommendations and explain "
        "why each place is worth visiting."
    ),
    llm=llm,
    verbose=True
)

# Create the Travel Planning Task
travel_task = Task(
    description=(
        "Create a travel guide for {city}. "
        "Recommend exactly 3 places to visit in this city. "
        "For each place, provide its name, a short description, "
        "and one practical travel tip. "
        "Keep the recommendations concise and useful for first-time visitors."
    ),
    expected_output=(
        "A numbered list of exactly 3 places to visit in {city}. "
        "Each recommendation must include the place name, "
        "a brief description, and one travel tip. "
        "Do not include any additional places."
    ),
    agent=planner
)

# Create the Crew
travel_crew = Crew(
    agents=[planner],
    tasks=[travel_task],
    verbose=True
)

# Execute the Crew with the user-selected city
result = travel_crew.kickoff(
    inputs={"city": city}
)

# Print the final travel recommendations
print("\n===== TRAVEL PLANNER RESULTS =====\n")
print(result)