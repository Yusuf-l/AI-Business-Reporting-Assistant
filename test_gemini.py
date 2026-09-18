from dotenv import load_dotenv
from google import genai
import os

load_dotenv()

api_key = os.getenv("GEMINI_API_KEY")

if not api_key:
    print("API key bulunamadı!")
    exit()

client = genai.Client(
    api_key=api_key
)

interaction = client.interactions.create(
    model="gemini-3.6-flash",
    input="Say hello in one sentence."
)

print("Gemini cevabı:")
print(interaction.output_text)