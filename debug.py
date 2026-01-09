import google.generativeai as genai

# --- PASTE YOUR KEY INSIDE THE QUOTES BELOW ---
api_key = "AIzaSyCRfJyzrbCAm5Foc_or9R9k68nox2z0ScQ" 

genai.configure(api_key=api_key)

print("---------------")
print("CONNECTING TO GOOGLE...")
print("---------------")

try:
    # Ask Google for the list of available models
    for m in genai.list_models():
        if 'generateContent' in m.supported_generation_methods:
            print(f"✅ FOUND: {m.name}")
            
except Exception as e:
    print(f"❌ ERROR: {e}")

print("---------------")