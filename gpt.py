
import openai

def ChatGPTrequest(text):
    openai.api_key = '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'  # Replace 'YOUR_API_KEY' with your actual OpenAI API key

    response = openai.Completion.create(
        engine='text-davinci-003',  # Use the GPT-3.5 model
        prompt="text",  # Use the provided text as the prompt
        max_tokens=100
    )
    generated_text = response.choices[0].text.strip()
    print(generated_text)
    
ChatGPTrequest("сколько время")
