import os
import json
import openai

# Set API key
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

# Send request to GPT (0.28 syntax)
try:
    print("🤖 Calling OpenAI to extract items...")
    
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
    
    print(f"📥 Received response from OpenAI ({len(structured_data)} chars)")
    
    # Clean the response - remove markdown code blocks if present
    structured_data = structured_data.strip()
    
    # Remove markdown code blocks
    if structured_data.startswith('```json'):
        structured_data = structured_data[7:]
        print("🧹 Removed ```json wrapper")
    elif structured_data.startswith('```'):
        structured_data = structured_data[3:]
        print("🧹 Removed ``` wrapper")
    
    if structured_data.endswith('```'):
        structured_data = structured_data[:-3]
        print("🧹 Removed trailing ```")
    
    structured_data = structured_data.strip()
    
    # Validate it's actual JSON
    try:
        json_test = json.loads(structured_data)
        items_count = len(json_test.get('items', []))
        print(f"✅ Validated JSON with {items_count} items")
    except json.JSONDecodeError as e:
        print(f"⚠️  JSON validation failed: {e}")