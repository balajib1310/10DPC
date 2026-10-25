import os
import json

local_state = os.path.expandvars(
    r"%LOCALAPPDATA%\Google\Chrome\User Data\Local State"
)

with open(local_state, "r", encoding="utf-8") as f:
    data = json.load(f)

profiles = data["profile"]["info_cache"]

for directory, info in profiles.items():
    print(
        f"Directory: {directory} | "
        f"Name: {info.get('name')} | "
        f"Email: {info.get('user_name')}"
    )