import openai

def get_gpt_response(prompt):
    response = openai.ChatCompletion.create(
        model="gpt-4-1106-preview",
        messages=[{"role": "system", "content": prompt}]
    )
    return response.choices[0].message['content']