from openai import OpenAI
from dotenv import load_dotenv
import os

load_dotenv()
OPENAI_API_KEY=os.getenv('OPENAI_API_KEY')
client = OpenAI(api_key=OPENAI_API_KEY)

resp = client.responses.create(
    model="gpt-5-mini",
    input="Say hello"
)

print(resp.text)