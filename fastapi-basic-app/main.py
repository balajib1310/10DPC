
from fastapi import FastAPI

# Create the FastAPI application
app = FastAPI(title="My First FastAPI App")


# Root endpoint
@app.get("/")
def home():
    return {"message": "hi, FastAPI"}


# Greeting endpoint with a path parameter
@app.get("/greet/{name}")
def greet(name: str):
    return {"message": f"Hi, {name}! Welcome to FastAPI."}