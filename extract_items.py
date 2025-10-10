import os
import json
import openai

# Set API key directly
openai.api_key = os.environ.get("OPENAI_API_KEY")

# Read the parsed text
with open("outputs/parsed_offer1.txt", "r", encoding="utf-8") as file:
    raw_text = file.read()

# Set the prompts
system_prompt = "You are an expert data extractor. From unstructured supplier offer text, extract structured data as JSON."

user_prompt = f"""Extract items from this offer:

{raw_text}

Return JSON with this structure:
{{
  "items": [
    {{
      "name": "product name",
      "description": "description", 
      "quantity": "quantity with unit",
      "unit": "unit",
      "price": "price with currency"
    }}
  ]
}}

Return ONLY valid JSON, no other text."""

# Send request to GPT using the legacy method (more compatible)
response = openai.ChatCompletion.create(
    model="gpt-3.5-turbo",
    messages=[
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_prompt},
    ],
    temperature=0.0,
)

# Get GPT output
structured_data = response['choices'][0]['message']['content']

# Save result to file
os.makedirs("outputs", exist_ok=True)
with open("outputs/items_offer1.json", "w", encoding="utf-8") as f:
    f.write(structured_data)

print("✅ Structured data saved to: outputs/items_offer1.json")