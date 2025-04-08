from ollama import Client

# Connect to the Ollama server
client = Client(host="http://localhost:11434")

# Send a prompt to the mistral model
response = client.generate(model="mistral", prompt="What is the capital of France?")

# Print the response
print(response["response"])