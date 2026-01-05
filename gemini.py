"""
Install the Google AI Python SDK

$ pip install google-generativeai

See the getting started guide for more information:
https://ai.google.dev/gemini-api/docs/get-started/python
"""

import os


import google.generativeai as genai
from apikey import GEMINI_API_KEY

genai.configure(api_key= GEMINI_API_KEY)

# Create the model
# See https://ai.google.dev/api/python/google/generativeai/GenerativeModel
generation_config = {
  "temperature": 1,
  "top_p": 0.95,
  "top_k": 64,
  "max_output_tokens": 8192,
  "response_mime_type": "text/plain",
}

model = genai.GenerativeModel(
  model_name="gemini-1.5-flash",
  generation_config=generation_config,
  # safety_settings = Adjust safety settings
  # See https://ai.google.dev/gemini-api/docs/safety-settings
)

chat_session = model.start_chat(
  history=[
    {
      "role": "user",
      "parts": [
        "AI\n",
      ],
    },
    {
      "role": "model",
      "parts": [
        "\"AI\" stands for **Artificial Intelligence**. It's a broad field of computer science that focuses on creating intelligent machines capable of performing tasks that typically require human intelligence",
      ],
    },
  ]
)

response = chat_session.send_message("essay on ai")

# print(response.text)

feedback = response.prompt_feedback
print(feedback)